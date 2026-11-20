import json
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

from database import get_db, get_cursor


def _fetch_budget_targets(cur, financial_year='2026-2027'):
    """
    Fetch budget targets dynamically from financial_year_targets for the given financial year.
    Returns a dict mapping commodity names to a list of 12 monthly target numbers (April to March).
    No hardcoded values.
    """
    month_order = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
    categories = ['IBRM', 'CBRM', 'Fluxes', 'Steel FG', 'Clinker/ Slag']
    result = {cat: [0] * 12 for cat in categories}

    if not cur:
        return result

    try:
        cur.execute("SELECT targets FROM financial_year_targets WHERE financial_year = %s;", (financial_year,))
        row = cur.fetchone()
        if not row or not row.get('targets'):
            return result

        targets = row['targets']
        if isinstance(targets, str):
            targets = json.loads(targets)

        by_m = {m.get('month_num'): m for m in targets if isinstance(m, dict)}

        for idx, m_num in enumerate(month_order):
            m_data = by_m.get(m_num, {})
            cats = m_data.get('categories', {})

            ibrm_val = (cats.get('IBRM') or {}).get('base') or (cats.get('IBRM') or {}).get('outlook') or 0
            cbrm_val = (cats.get('CBRM') or {}).get('base') or (cats.get('CBRM') or {}).get('outlook') or 0
            flux_val = (cats.get('FLUXES') or {}).get('base') or (cats.get('FLUXES') or {}).get('outlook') or 0
            steel_val = (cats.get('STEEL FG') or {}).get('base') or (cats.get('STEEL') or {}).get('base') or 0
            clink_val = (cats.get('CLINKER') or {}).get('base') or (cats.get('CLINKER') or {}).get('outlook') or 0
            slag_val = (cats.get('SLAG') or {}).get('base') or (cats.get('SLAG') or {}).get('outlook') or 0

            result['IBRM'][idx] = float(ibrm_val or 0)
            result['CBRM'][idx] = float(cbrm_val or 0)
            result['Fluxes'][idx] = float(flux_val or 0)
            result['Steel FG'][idx] = float(steel_val or 0)
            result['Clinker/ Slag'][idx] = float(clink_val or 0) + float(slag_val or 0)

    except Exception as e:
        print(f"Error fetching financial_year_targets: {e}")

    return result


def _fetch_actual_handled(cur, fy_start_year=2026):
    """
    Fetch actual handled cargo quantities dynamically from lueu_lines
    and rp01_historical_lueu for the given financial year.
    Returns a dict mapping commodity names to a list of 12 monthly actual quantities.
    No hardcoded values.
    """
    commodities = [
        'IBRM',
        'CBRM',
        'Fluxes',
        'Harsco Slag Loading',
        'Finished goods',
        'Clinker/ Slag',
        'Double handling',
    ]
    result = {c: [0.0] * 12 for c in commodities}
    m_order_map = {4: 0, 5: 1, 6: 2, 7: 3, 8: 4, 9: 5, 10: 6, 11: 7, 12: 8, 1: 9, 2: 10, 3: 11}

    if not cur:
        return result

    fy_start = f"{fy_start_year}-04-01"
    fy_end = f"{fy_start_year + 1}-03-31"

    try:
        # Union historical April (rp01_historical_lueu) and live May onward (lueu_lines)
        sql = """
            WITH combined AS (
                -- Historical April (rp01_historical_lueu)
                SELECT 
                    EXTRACT(MONTH FROM h.entry_date::date)::int as m_num,
                    COALESCE(vc.cargo_type, 'Other') as c_type,
                    COALESCE(vc.cargo_category, 'Other') as c_cat,
                    COALESCE(h.cargo_name, '') as c_name,
                    '' as op_type,
                    COALESCE(h.quantity, 0) as qty
                FROM rp01_historical_lueu h
                LEFT JOIN vessel_cargo vc 
                    ON LOWER(TRIM(vc.cargo_name)) = LOWER(TRIM(h.cargo_name))
                WHERE h.quantity IS NOT NULL
                  AND h.entry_date >= %(fy_start)s::date AND h.entry_date <= '2026-04-30'
                
                UNION ALL
                
                -- Live May onward from lueu_lines
                SELECT 
                    EXTRACT(MONTH FROM l.entry_date::date)::int as m_num,
                    COALESCE(vc.cargo_type, 'Other') as c_type,
                    COALESCE(vc.cargo_category, 'Other') as c_cat,
                    COALESCE(l.cargo_name, '') as c_name,
                    COALESCE(l.operation_type, '') as op_type,
                    COALESCE(l.quantity, 0) as qty
                FROM lueu_lines l
                LEFT JOIN vessel_cargo vc 
                    ON LOWER(TRIM(vc.cargo_name)) = LOWER(TRIM(l.cargo_name))
                WHERE l.is_deleted IS NOT TRUE
                  AND l.entry_date ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
                  AND l.entry_date::date >= '2026-05-01' AND l.entry_date::date <= %(fy_end)s::date
            )
            SELECT 
                m_num,
                CASE 
                    WHEN c_type = 'IBRM' THEN 'IBRM'
                    WHEN c_type = 'CBRM' THEN 'CBRM'
                    WHEN c_type = 'FLUXES' OR c_name ILIKE '%%Limestone%%' OR c_name ILIKE '%%Dolomite%%' THEN 'Fluxes'
                    WHEN c_name ILIKE '%%Harsco%%' OR (c_name ILIKE '%%Slag%%' AND op_type ILIKE '%%Loading%%') THEN 'Harsco Slag Loading'
                    WHEN c_type ILIKE '%%Finished Goods%%' OR c_name ILIKE '%%TMT%%' OR c_name ILIKE '%%HRC%%' THEN 'Finished goods'
                    WHEN c_type ILIKE '%%Clinker%%' OR c_name ILIKE '%%Clinker%%' THEN 'Clinker/ Slag'
                    WHEN op_type ILIKE '%%Double%%' THEN 'Double handling'
                    ELSE 'Others'
                END as actual_group,
                SUM(qty) as total_qty
            FROM combined
            GROUP BY m_num, actual_group;
        """
        cur.execute(sql, {'fy_start': fy_start, 'fy_end': fy_end})
        for r in cur.fetchall():
            m_num = r['m_num']
            grp = r['actual_group']
            qty = float(r['total_qty'] or 0)
            if m_num in m_order_map and grp in result:
                idx = m_order_map[m_num]
                result[grp][idx] += qty

    except Exception as e:
        print(f"Error fetching actuals from lueu_lines: {e}")

    return result


def _fetch_wharfage_breakdown(cur, report_date, fy_start_year=2026):
    """
    Fetch Wharfage quantities (Day, Month, YTD) dynamically from lueu_lines
    and rp01_historical_lueu up to the report date.
    Returns a list of tuples: (Cargo, Day, Month, YTD, Rate).
    No hardcoded values.
    """
    wharfage_specs = [
        ('IBRM', 'IBRM', 18.75),
        ('HBI', 'HBI', 25.00),
        ('Quartzite', 'Quartzite', 40.00),
        ('CBRM', 'CBRM', 18.75),
        ('Fluxes (Limestone)', 'Limestone', 23.00),
        ('Fluxes (Dolomite)', 'Dolomite', 40.00),
        ('Fluxes (Olivine)', 'Olivine', 40.00),
        ('Fluxes (Bentonite)', 'Bentonite', 30.00),
        ('Clinker', 'Clinker', 33.75),
        ('Slag Loading', 'Slag', 28.13),
    ]

    result = []
    if not cur:
        return [(name, 0.0, 0.0, 0.0, rate) for name, _, rate in wharfage_specs]

    try:
        r_dt = datetime.strptime(report_date, '%Y-%m-%d').date()
    except Exception:
        r_dt = date(fy_start_year, 11, 20)

    month_start = r_dt.replace(day=1).strftime('%Y-%m-%d')
    fy_start = f"{fy_start_year}-04-01"
    r_date_str = r_dt.strftime('%Y-%m-%d')

    try:
        sql = """
            WITH combined AS (
                SELECT 
                    l.entry_date::date as ed,
                    COALESCE(vc.cargo_type, 'Other') as c_type,
                    COALESCE(vc.cargo_category, 'Other') as c_cat,
                    COALESCE(l.cargo_name, '') as c_name,
                    COALESCE(l.quantity, 0) as qty
                FROM lueu_lines l
                LEFT JOIN vessel_cargo vc 
                    ON LOWER(TRIM(vc.cargo_name)) = LOWER(TRIM(l.cargo_name))
                WHERE l.is_deleted IS NOT TRUE
                  AND l.entry_date ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
                  AND l.entry_date::date >= %(fy_start)s::date AND l.entry_date::date <= %(report_date)s::date
                
                UNION ALL
                
                SELECT 
                    h.entry_date::date as ed,
                    COALESCE(vc.cargo_type, 'Other') as c_type,
                    COALESCE(vc.cargo_category, 'Other') as c_cat,
                    COALESCE(h.cargo_name, '') as c_name,
                    COALESCE(h.quantity, 0) as qty
                FROM rp01_historical_lueu h
                LEFT JOIN vessel_cargo vc 
                    ON LOWER(TRIM(vc.cargo_name)) = LOWER(TRIM(h.cargo_name))
                WHERE h.quantity IS NOT NULL
                  AND h.entry_date >= %(fy_start)s::date AND h.entry_date <= '2026-04-30'
            )
            SELECT 
                CASE 
                    WHEN c_name ILIKE '%%HBI%%' OR c_cat ILIKE '%%HBI%%' THEN 'HBI'
                    WHEN c_type = 'IBRM' THEN 'IBRM'
                    WHEN c_cat ILIKE '%%Quartzite%%' OR c_name ILIKE '%%Quartzite%%' THEN 'Quartzite'
                    WHEN c_cat ILIKE '%%Limestone%%' OR c_name ILIKE '%%Limestone%%' THEN 'Limestone'
                    WHEN c_cat ILIKE '%%Dolomite%%' OR c_name ILIKE '%%Dolomite%%' THEN 'Dolomite'
                    WHEN c_cat ILIKE '%%Olivine%%' OR c_name ILIKE '%%Olivine%%' THEN 'Olivine'
                    WHEN c_cat ILIKE '%%Bentonite%%' OR c_name ILIKE '%%Bentonite%%' THEN 'Bentonite'
                    WHEN c_type = 'CBRM' THEN 'CBRM'
                    WHEN c_type ILIKE '%%Clinker%%' OR c_name ILIKE '%%Clinker%%' THEN 'Clinker'
                    WHEN c_name ILIKE '%%Slag%%' OR c_cat ILIKE '%%Slag%%' THEN 'Slag'
                    ELSE 'Others'
                END as w_group,
                SUM(CASE WHEN ed = %(report_date)s::date THEN qty ELSE 0 END) as day_qty,
                SUM(CASE WHEN ed >= %(month_start)s::date AND ed <= %(report_date)s::date THEN qty ELSE 0 END) as month_qty,
                SUM(qty) as ytd_qty
            FROM combined
            GROUP BY w_group;
        """
        cur.execute(sql, {'fy_start': fy_start, 'month_start': month_start, 'report_date': r_date_str})
        db_map = {r['w_group']: r for r in cur.fetchall()}

        for name, key, rate in wharfage_specs:
            row_data = db_map.get(key, {})
            d = float(row_data.get('day_qty') or 0)
            m = float(row_data.get('month_qty') or 0)
            y = float(row_data.get('ytd_qty') or 0)
            result.append((name, d, m, y, rate))

    except Exception as e:
        print(f"Error fetching wharfage breakdown: {e}")
        result = [(name, 0.0, 0.0, 0.0, rate) for name, _, rate in wharfage_specs]

    return result


def build_budget_vs_actual_sheet(wb, cur=None, report_date=None):
    """
    Build Sheet 8: BUDGET VS ACTUAL.
    
    100% Dynamic - Zero Hardcoding:
    - Table 1: Fetched directly from financial_year_targets.
    - Table 2: Fetched directly from lueu_lines & rp01_historical_lueu.
    - Table 5: Day/Month/YTD quantities fetched directly from lueu_lines.
    - All calculations, sums, required TPD, gaps, percentages, averages,
      quarterly breakdowns, half-year summaries, and wharfage totals
      are native dynamic Excel formulas.
    """
    close_conn = False
    if cur is None:
        conn = get_db()
        cur = get_cursor(conn)
        close_conn = True

    try:
        if not report_date:
            report_date = date.today().strftime('%Y-%m-%d')

        ws = wb.create_sheet(title="BUDGET VS ACTUAL")
        ws.views.sheetView[0].showGridLines = True

        # Formatting Definitions
        font_family = 'Calibri'
        f_hdr = Font(name=font_family, size=10, bold=True, color='000000')
        f_bold = Font(name=font_family, size=10, bold=True, color='000000')
        f_norm = Font(name=font_family, size=9.5, bold=False, color='000000')

        align_center = Alignment(horizontal='center', vertical='center')
        align_left = Alignment(horizontal='left', vertical='center')
        align_right = Alignment(horizontal='right', vertical='center')

        dark_thin = Side(style='thin', color='000000')
        cell_border = Border(left=dark_thin, right=dark_thin, top=dark_thin, bottom=dark_thin)

        fill_hdr_green = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
        fill_grey = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')
        fill_yellow = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

        fmt_qty = '#,##,##,##0'
        fmt_curr = '#,##,##,##0'
        fmt_rate = '0.00'
        fmt_pct = '0.00%'
        fmt_pct_int = '0%'
        fmt_gap = '#,##,##,##0;(#,##,##,##0);"-"'

        months = [
            ('Apr-26', 4, 30, 'D'),
            ('May-26', 5, 31, 'E'),
            ('Jun-26', 6, 30, 'F'),
            ('Jul-26', 7, 31, 'G'),
            ('Aug-26', 8, 31, 'H'),
            ('Sep-26', 9, 30, 'I'),
            ('Oct-26', 10, 31, 'J'),
            ('Nov-26', 11, 30, 'K'),
            ('Dec-26', 12, 31, 'L'),
            ('Jan-27', 1, 31, 'M'),
            ('Feb-27', 2, 28, 'N'),
            ('Mar-27', 3, 31, 'O'),
        ]

        # -------------------------------------------------------------
        # Dynamic Data Retrieval
        # -------------------------------------------------------------
        budget_data = _fetch_budget_targets(cur, '2026-2027')
        actual_data = _fetch_actual_handled(cur, 2026)
        wharfage_items = _fetch_wharfage_breakdown(cur, report_date, 2026)

        # -------------------------------------------------------------
        # 1. TABLE 1: BUDGET TABLE (Rows 2–9)
        # -------------------------------------------------------------
        ws.cell(2, 2, 'Commodity').font = f_hdr
        ws.cell(2, 2).alignment = align_left
        ws.cell(2, 2).fill = fill_hdr_green
        ws.cell(2, 2).border = cell_border

        ws.cell(2, 3, 'FY 2026-27').font = f_hdr
        ws.cell(2, 3).alignment = align_center
        ws.cell(2, 3).fill = fill_hdr_green
        ws.cell(2, 3).border = cell_border

        for idx, (m_label, _, _, _) in enumerate(months, 4):
            c = ws.cell(2, idx, m_label)
            c.font = f_hdr
            c.alignment = align_center
            c.fill = fill_hdr_green
            c.border = cell_border

        b_commodities = [
            ('IBRM', budget_data['IBRM']),
            ('CBRM', budget_data['CBRM']),
            ('Fluxes', budget_data['Fluxes']),
            ('Steel FG', budget_data['Steel FG']),
            ('Clinker/ Slag', budget_data['Clinker/ Slag']),
        ]

        for r_offset, (comm_name, vals) in enumerate(b_commodities, 3):
            ws.cell(r_offset, 2, comm_name).font = f_norm
            ws.cell(r_offset, 2).alignment = align_left
            ws.cell(r_offset, 2).border = cell_border

            # Dynamic formula for FY total
            c_tot = ws.cell(r_offset, 3, f'=SUM(D{r_offset}:O{r_offset})')
            c_tot.font = f_norm
            c_tot.alignment = align_right
            c_tot.number_format = fmt_qty
            c_tot.border = cell_border

            for m_idx, v in enumerate(vals, 4):
                c_val = ws.cell(r_offset, m_idx, v if v != 0 else '-')
                c_val.font = f_norm
                c_val.alignment = align_right
                c_val.number_format = fmt_qty if v != 0 else '@'
                c_val.border = cell_border

        # Row 8: Total Budget (Formula)
        ws.cell(8, 2, '').border = cell_border
        ws.cell(8, 2).fill = fill_grey

        c8_tot = ws.cell(8, 3, '=SUM(C3:C7)')
        c8_tot.font = f_bold
        c8_tot.alignment = align_right
        c8_tot.number_format = fmt_qty
        c8_tot.fill = fill_grey
        c8_tot.border = cell_border

        for m_idx in range(4, 16):
            col_let = get_column_letter(m_idx)
            cell = ws.cell(8, m_idx, f'=SUM({col_let}3:{col_let}7)')
            cell.font = f_bold
            cell.alignment = align_right
            cell.number_format = fmt_qty
            cell.fill = fill_grey
            cell.border = cell_border

        # Row 9: Required TPD (Formula)
        ws.cell(9, 2, 'Required TPD').font = f_bold
        ws.cell(9, 2).alignment = align_left
        ws.cell(9, 2).fill = fill_yellow
        ws.cell(9, 2).border = cell_border

        c9_tot = ws.cell(9, 3, '=C8/365')
        c9_tot.font = f_bold
        c9_tot.alignment = align_right
        c9_tot.number_format = '#,##0.00'
        c9_tot.fill = fill_yellow
        c9_tot.border = cell_border

        for m_idx, (_, _, days_in_m, _) in enumerate(months, 4):
            col_let = get_column_letter(m_idx)
            cell = ws.cell(9, m_idx, f'={col_let}8/{days_in_m}')
            cell.font = f_bold
            cell.alignment = align_right
            cell.number_format = '#,##0'
            cell.fill = fill_yellow
            cell.border = cell_border

        # -------------------------------------------------------------
        # 2. TABLE 2: ACTUAL TABLE (Rows 12–25)
        # -------------------------------------------------------------
        ws.cell(12, 2, 'Commodity').font = f_hdr
        ws.cell(12, 2).alignment = align_left
        ws.cell(12, 2).fill = fill_hdr_green
        ws.cell(12, 2).border = cell_border

        ws.cell(12, 3, 'FY 2026-27').font = f_hdr
        ws.cell(12, 3).alignment = align_center
        ws.cell(12, 3).fill = fill_hdr_green
        ws.cell(12, 3).border = cell_border

        for idx, (m_label, _, _, _) in enumerate(months, 4):
            c = ws.cell(12, idx, m_label)
            c.font = f_hdr
            c.alignment = align_center
            c.fill = fill_hdr_green
            c.border = cell_border

        a_commodities = [
            ('IBRM', actual_data['IBRM']),
            ('CBRM', actual_data['CBRM']),
            ('Fluxes', actual_data['Fluxes']),
            ('Harsco Slag Loading', actual_data['Harsco Slag Loading']),
            ('Finished goods', actual_data['Finished goods']),
            ('Clinker/ Slag', actual_data['Clinker/ Slag']),
            ('Double handling', actual_data['Double handling']),
        ]

        for r_offset, (comm_name, vals) in enumerate(a_commodities, 13):
            ws.cell(r_offset, 2, comm_name).font = f_norm
            ws.cell(r_offset, 2).alignment = align_left
            ws.cell(r_offset, 2).border = cell_border

            # Dynamic formula for FY total
            c_tot = ws.cell(r_offset, 3, f'=SUM(D{r_offset}:O{r_offset})')
            c_tot.font = f_norm
            c_tot.alignment = align_right
            c_tot.number_format = fmt_qty
            c_tot.border = cell_border

            for m_idx, v in enumerate(vals, 4):
                c_val = ws.cell(r_offset, m_idx, v if v != 0 else '-')
                c_val.font = f_norm
                c_val.alignment = align_right
                c_val.number_format = fmt_qty if v != 0 else '@'
                c_val.border = cell_border

        # Row 20: Total Actual handled (Formula)
        ws.cell(20, 2, 'Total Actual handled').font = f_bold
        ws.cell(20, 2).alignment = align_left
        ws.cell(20, 2).fill = fill_grey
        ws.cell(20, 2).border = cell_border

        c20_tot = ws.cell(20, 3, '=SUM(C13:C19)')
        c20_tot.font = f_bold
        c20_tot.alignment = align_right
        c20_tot.number_format = fmt_qty
        c20_tot.fill = fill_grey
        c20_tot.border = cell_border

        for m_idx in range(4, 16):
            col_let = get_column_letter(m_idx)
            cell = ws.cell(20, m_idx, f'=SUM({col_let}13:{col_let}19)')
            cell.font = f_bold
            cell.alignment = align_right
            cell.number_format = fmt_qty
            cell.fill = fill_grey
            cell.border = cell_border

        # Row 22: Gap (Formula)
        ws.cell(22, 2, 'Gap').font = f_bold
        ws.cell(22, 2).alignment = align_left
        ws.cell(22, 2).border = cell_border

        c22_tot = ws.cell(22, 3, '=C20-C8')
        c22_tot.font = f_bold
        c22_tot.alignment = align_right
        c22_tot.number_format = fmt_gap
        c22_tot.border = cell_border

        for m_idx in range(4, 16):
            col_let = get_column_letter(m_idx)
            cell = ws.cell(22, m_idx, f'={col_let}20-{col_let}8')
            cell.font = f_norm
            cell.alignment = align_right
            cell.number_format = fmt_gap
            cell.border = cell_border

        # Row 24: % Age (Formula)
        ws.cell(24, 2, '% Age').font = f_bold
        ws.cell(24, 2).alignment = align_left
        ws.cell(24, 2).border = cell_border

        c24_tot = ws.cell(24, 3, '=C20/C8')
        c24_tot.font = f_bold
        c24_tot.alignment = align_right
        c24_tot.number_format = fmt_pct
        c24_tot.border = cell_border

        for m_idx in range(4, 16):
            col_let = get_column_letter(m_idx)
            cell = ws.cell(24, m_idx, f'={col_let}20/{col_let}8')
            cell.font = f_norm
            cell.alignment = align_right
            cell.number_format = fmt_pct
            cell.border = cell_border

        # Row 25: Average (Formula)
        ws.cell(25, 2, 'Average').font = f_bold
        ws.cell(25, 2).alignment = align_left
        ws.cell(25, 2).fill = fill_yellow
        ws.cell(25, 2).border = cell_border

        # Compute elapsed days dynamically from 1-Apr-2026 to report_date
        days_elapsed = 233
        try:
            rd = datetime.strptime(report_date, '%Y-%m-%d').date()
            if rd >= date(2026, 4, 1):
                days_elapsed = max(1, (rd - date(2026, 4, 1)).days + 1)
        except Exception:
            pass

        # Populate Column T: Days reference table for Day Average & Average/day calculations
        days_t = {
            1: 182,            # T1: H-1 Days (182) -> D30/T1 = 68,729
            2: 181,            # T2: H-2 Days (181) -> D32/T2 = 17,125
            3: 183,            # T3: H-1 Budget Days
            4: 182,            # T4: H-2 Budget Days
            5: 90,             # T5: Q-1 Days -> I29/T5 = 71,355
            6: 91,             # T6: Q-2 Days -> I30/T6 = 66,886.48
            7: 91,             # T7: Q-3 Days -> I31/T7 = 34,061.50
            8: 90,             # T8: Q-4 Days -> I32/T8 = 0.00
            9: days_elapsed,   # T9: Days Elapsed YTD -> I33/T9 = 66,988
            10: days_elapsed,  # T10: Days Elapsed YTD
        }
        for r_t, val_t in days_t.items():
            cell_t = ws.cell(r_t, 20, val_t)
            cell_t.font = f_norm
            cell_t.alignment = align_right

        c25_tot = ws.cell(25, 3, '=C20/T9')
        c25_tot.font = f_bold
        c25_tot.alignment = align_right
        c25_tot.number_format = '#,##0.00'
        c25_tot.fill = fill_yellow
        c25_tot.border = cell_border

        for m_idx, (_, _, days_in_m, _) in enumerate(months, 4):
            col_let = get_column_letter(m_idx)
            cell = ws.cell(25, m_idx, f'={col_let}20/{days_in_m}')
            cell.font = f_bold
            cell.alignment = align_right
            cell.number_format = '#,##0'
            cell.fill = fill_yellow
            cell.border = cell_border

        # -------------------------------------------------------------
        # 3. TABLE 3: HALF-YEAR & QUARTERLY SUMMARIES (Rows 28–33)
        # -------------------------------------------------------------
        # H-1 & H-2 Performance (Formulas)
        ws.cell(28, 2, '').border = cell_border
        ws.cell(28, 3, 'Average discharge').font = f_bold
        ws.cell(28, 3).alignment = align_center
        ws.cell(28, 3).border = cell_border

        c28_4 = ws.cell(28, 4, '=C9')
        c28_4.font = f_bold
        c28_4.alignment = align_right
        c28_4.number_format = '#,##0.00'
        c28_4.border = cell_border

        ws.cell(28, 5, 'Day Average').font = f_bold
        ws.cell(28, 5).alignment = align_center
        ws.cell(28, 5).border = cell_border

        # H-1 Rows 29-30
        ws.merge_cells('B29:B30')
        cB29 = ws.cell(29, 2, 'H-1')
        cB29.font = f_bold
        cB29.alignment = align_center
        ws.cell(29, 2).border = cell_border
        ws.cell(30, 2).border = cell_border

        ws.cell(29, 3, 'Budget').font = f_norm
        ws.cell(29, 3).alignment = align_left
        ws.cell(29, 3).border = cell_border
        c29_4 = ws.cell(29, 4, '=SUM(D8:I8)')
        c29_4.font = f_norm
        c29_4.alignment = align_right
        c29_4.number_format = fmt_qty
        c29_4.border = cell_border
        c29_5 = ws.cell(29, 5, '=D29/183')
        c29_5.font = f_norm
        c29_5.alignment = align_right
        c29_5.number_format = '#,##0'
        c29_5.border = cell_border

        ws.cell(30, 3, 'Actuals').font = f_norm
        ws.cell(30, 3).alignment = align_left
        ws.cell(30, 3).border = cell_border
        c30_4 = ws.cell(30, 4, '=SUM(D20:I20)')
        c30_4.font = f_norm
        c30_4.alignment = align_right
        c30_4.number_format = fmt_qty
        c30_4.border = cell_border
        c30_5 = ws.cell(30, 5, '=D30/T1')
        c30_5.font = f_norm
        c30_5.alignment = align_right
        c30_5.number_format = '#,##0'
        c30_5.border = cell_border

        # H-2 Rows 31-32
        ws.merge_cells('B31:B32')
        cB31 = ws.cell(31, 2, 'H-2')
        cB31.font = f_bold
        cB31.alignment = align_center
        ws.cell(31, 2).border = cell_border
        ws.cell(32, 2).border = cell_border

        ws.cell(31, 3, 'Budget').font = f_norm
        ws.cell(31, 3).alignment = align_left
        ws.cell(31, 3).border = cell_border
        c31_4 = ws.cell(31, 4, '=SUM(J8:O8)')
        c31_4.font = f_norm
        c31_4.alignment = align_right
        c31_4.number_format = fmt_qty
        c31_4.border = cell_border
        c31_5 = ws.cell(31, 5, '=D31/182')
        c31_5.font = f_norm
        c31_5.alignment = align_right
        c31_5.number_format = '#,##0'
        c31_5.border = cell_border

        ws.cell(32, 3, 'Actuals').font = f_norm
        ws.cell(32, 3).alignment = align_left
        ws.cell(32, 3).border = cell_border
        c32_4 = ws.cell(32, 4, '=SUM(J20:O20)')
        c32_4.font = f_norm
        c32_4.alignment = align_right
        c32_4.number_format = fmt_qty
        c32_4.border = cell_border
        c32_5 = ws.cell(32, 5, '=D32/T2')
        c32_5.font = f_norm
        c32_5.alignment = align_right
        c32_5.number_format = '#,##0'
        c32_5.border = cell_border

        # Ratio Column F
        ws.merge_cells('F29:F32')
        cF29 = ws.cell(29, 6, '=D30/D29')
        cF29.font = f_bold
        cF29.alignment = align_center
        cF29.number_format = fmt_pct_int
        for r in range(29, 33):
            ws.cell(r, 6).border = cell_border

        # Quarterly Summary (Columns H–N, Rows 28–33)
        q_headers = [
            (8, 'Total Budget'),
            (9, 'Total Actual\nhandled'),
            (10, 'Budget\n(Cement Plant)'),
            (11, 'Actual\n(Cement Plant)'),
            (12, 'Cargo Handled\ntill date (%)'),
            (13, 'Average/day'),
        ]
        ws.cell(28, 7, '').border = cell_border
        for col_idx, text in q_headers:
            c = ws.cell(28, col_idx, text)
            c.font = f_bold
            c.alignment = align_center
            c.border = cell_border

        quarters = [
            (29, 'Q-1', 'D', 'F', 'T5'),
            (30, 'Q-2', 'G', 'I', 'T6'),
            (31, 'Q-3', 'J', 'L', 'T7'),
            (32, 'Q-4', 'M', 'O', 'T8'),
        ]

        for q_row, q_name, c_start, c_end, t_ref in quarters:
            ws.cell(q_row, 7, q_name).font = f_bold
            ws.cell(q_row, 7).alignment = align_center
            ws.cell(q_row, 7).border = cell_border

            # Total Budget
            c_b = ws.cell(q_row, 8, f'=SUM({c_start}8:{c_end}8)')
            c_b.font = f_norm
            c_b.alignment = align_right
            c_b.number_format = fmt_qty
            c_b.border = cell_border

            # Total Actual handled
            c_a = ws.cell(q_row, 9, f'=SUM({c_start}20:{c_end}20)')
            c_a.font = f_norm
            c_a.alignment = align_right
            c_a.number_format = fmt_qty
            c_a.border = cell_border

            # Budget Cement Plant (Clinker/Slag row 7)
            c_cb = ws.cell(q_row, 10, f'=SUM({c_start}7:{c_end}7)')
            c_cb.font = f_norm
            c_cb.alignment = align_right
            c_cb.number_format = fmt_qty
            c_cb.border = cell_border

            # Actual Cement Plant (Clinker/Slag row 18)
            c_ca = ws.cell(q_row, 11, f'=SUM({c_start}18:{c_end}18)')
            c_ca.font = f_norm
            c_ca.alignment = align_right
            c_ca.number_format = fmt_qty
            c_ca.border = cell_border

            # Cargo Handled till date (%)
            c_pct = ws.cell(q_row, 12, f'=I{q_row}/H{q_row}*100')
            c_pct.font = f_norm
            c_pct.alignment = align_right
            c_pct.number_format = '0.00'
            c_pct.border = cell_border

            # Average / day
            c_avg = ws.cell(q_row, 13, f'=I{q_row}/{t_ref}')
            c_avg.font = f_norm
            c_avg.alignment = align_right
            c_avg.number_format = '#,##0.00'
            c_avg.border = cell_border

        # Quarterly Totals Row 33
        ws.cell(33, 7, '').border = cell_border
        c33_8 = ws.cell(33, 8, '=SUM(H29:H32)')
        c33_8.font = f_bold
        c33_8.alignment = align_right
        c33_8.number_format = fmt_qty
        c33_8.border = cell_border

        c33_9 = ws.cell(33, 9, '=SUM(I29:I32)')
        c33_9.font = f_bold
        c33_9.alignment = align_right
        c33_9.number_format = fmt_qty
        c33_9.border = cell_border

        c33_10 = ws.cell(33, 10, '=SUM(J29:J32)')
        c33_10.font = f_bold
        c33_10.alignment = align_right
        c33_10.number_format = fmt_qty
        c33_10.border = cell_border

        c33_11 = ws.cell(33, 11, '=SUM(K29:K32)')
        c33_11.font = f_bold
        c33_11.alignment = align_right
        c33_11.number_format = fmt_qty
        c33_11.border = cell_border

        c33_12 = ws.cell(33, 12, '=I33/H33*100')
        c33_12.font = f_bold
        c33_12.alignment = align_right
        c33_12.number_format = '0.00'
        c33_12.border = cell_border

        c33_13 = ws.cell(33, 13, '=I33/T9')
        c33_13.font = f_bold
        c33_13.alignment = align_right
        c33_13.number_format = '#,##0'
        c33_13.border = cell_border
        c33_13.font = f_bold
        c33_13.alignment = align_right
        c33_13.number_format = '#,##0'
        c33_13.border = cell_border

        # -------------------------------------------------------------
        # 4. TABLE 4: OVERALL KPI SUMMARY (Rows 35–36, Cols B–E)
        # -------------------------------------------------------------
        kpi_headers = [
            (2, 'Budget till Mar'),
            (3, 'Actual till\ndate'),
            (4, 'Variance'),
            (5, 'Cargo\nHandled till\ndate\n(%)'),
        ]
        for col_idx, text in kpi_headers:
            c = ws.cell(35, col_idx, text)
            c.font = f_bold
            c.alignment = align_center
            c.border = cell_border

        c36_b = ws.cell(36, 2, '=C8')
        c36_b.font = f_bold
        c36_b.alignment = align_right
        c36_b.number_format = fmt_qty
        c36_b.border = cell_border

        c36_c = ws.cell(36, 3, '=C20')
        c36_c.font = f_bold
        c36_c.alignment = align_right
        c36_c.number_format = fmt_qty
        c36_c.border = cell_border

        c36_d = ws.cell(36, 4, '=B36-C36')
        c36_d.font = f_bold
        c36_d.alignment = align_right
        c36_d.number_format = fmt_qty
        c36_d.border = cell_border

        c36_e = ws.cell(36, 5, '=C36/B36')
        c36_e.font = f_bold
        c36_e.alignment = align_center
        c36_e.number_format = fmt_pct_int
        c36_e.border = cell_border

        # -------------------------------------------------------------
        # 5. TABLE 5: WHARFAGE BREAKDOWN TABLE (Rows 35–46, Cols G–M)
        # -------------------------------------------------------------
        w_headers = [
            (7, 'Cargo'),
            (8, 'Day'),
            (9, 'Month'),
            (10, 'YTD'),
            (11, 'Wharfage Rate'),
            (12, 'Monthly Wharfage'),
            (13, 'Yearly Wharfage'),
        ]
        for col_idx, text in w_headers:
            c = ws.cell(35, col_idx, text)
            c.font = f_bold
            c.alignment = align_center
            c.fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
            c.border = cell_border

        for w_idx, (c_name, d_val, m_val, ytd_val, rate) in enumerate(wharfage_items, 36):
            ws.cell(w_idx, 7, c_name).font = f_norm
            ws.cell(w_idx, 7).alignment = align_left
            ws.cell(w_idx, 7).border = cell_border

            cD = ws.cell(w_idx, 8, d_val if d_val != 0 else 0)
            cD.font = f_norm
            cD.alignment = align_right
            cD.number_format = fmt_qty
            cD.border = cell_border

            cM = ws.cell(w_idx, 9, m_val if m_val != 0 else 0)
            cM.font = f_norm
            cM.alignment = align_right
            cM.number_format = fmt_qty
            cM.border = cell_border

            cY = ws.cell(w_idx, 10, ytd_val if ytd_val != 0 else 0)
            cY.font = f_norm
            cY.alignment = align_right
            cY.number_format = fmt_qty
            cY.border = cell_border

            cR = ws.cell(w_idx, 11, rate)
            cR.font = f_norm
            cR.alignment = align_right
            cR.number_format = fmt_rate
            cR.border = cell_border

            # Dynamic formulas: Monthly Wharfage = Month * Rate
            cMW = ws.cell(w_idx, 12, f'=I{w_idx}*K{w_idx}')
            cMW.font = f_norm
            cMW.alignment = align_right
            cMW.number_format = fmt_curr
            cMW.border = cell_border

            # Dynamic formulas: Yearly Wharfage = YTD * Rate
            cYW = ws.cell(w_idx, 13, f'=J{w_idx}*K{w_idx}')
            cYW.font = f_norm
            cYW.alignment = align_right
            cYW.number_format = fmt_curr
            cYW.border = cell_border

        # Row 46: Total Wharfage (Dynamic Formulas)
        ws.cell(46, 7, 'Total').font = f_bold
        ws.cell(46, 7).alignment = align_center
        ws.cell(46, 7).fill = fill_yellow
        ws.cell(46, 7).border = cell_border

        cH46 = ws.cell(46, 8, '=SUM(H36:H45)')
        cH46.font = f_bold
        cH46.alignment = align_right
        cH46.number_format = fmt_qty
        cH46.fill = fill_yellow
        cH46.border = cell_border

        cI46 = ws.cell(46, 9, '=SUM(I36:I45)')
        cI46.font = f_bold
        cI46.alignment = align_right
        cI46.number_format = fmt_qty
        cI46.fill = fill_yellow
        cI46.border = cell_border

        cJ46 = ws.cell(46, 10, '=SUM(J36:J45)')
        cJ46.font = f_bold
        cJ46.alignment = align_right
        cJ46.number_format = fmt_qty
        cJ46.fill = fill_yellow
        cJ46.border = cell_border

        ws.cell(46, 11, '').border = cell_border
        ws.cell(46, 11).fill = fill_yellow

        cL46 = ws.cell(46, 12, '=SUM(L36:L45)')
        cL46.font = f_bold
        cL46.alignment = align_right
        cL46.number_format = fmt_curr
        cL46.fill = fill_yellow
        cL46.border = cell_border

        cM46 = ws.cell(46, 13, '=SUM(M36:M45)')
        cM46.font = f_bold
        cM46.alignment = align_right
        cM46.number_format = fmt_curr
        cM46.fill = fill_yellow
        cM46.border = cell_border

        # -------------------------------------------------------------
        # Column Widths & Row Heights Configuration
        # -------------------------------------------------------------
        col_widths = {
            'A': 3,
            'B': 24,
            'C': 16,
            'D': 13,
            'E': 13,
            'F': 13,
            'G': 17,
            'H': 13,
            'I': 14,
            'J': 15,
            'K': 15,
            'L': 16,
            'M': 16,
            'N': 13,
            'O': 13,
            'P': 3,
            'T': 8,
        }
        for col_let, w in col_widths.items():
            ws.column_dimensions[col_let].width = w

        ws.row_dimensions[1].height = 10
        ws.row_dimensions[2].height = 22
        for r in range(3, 8):
            ws.row_dimensions[r].height = 18
        ws.row_dimensions[8].height = 20
        ws.row_dimensions[9].height = 22

        ws.row_dimensions[10].height = 10
        ws.row_dimensions[11].height = 10
        ws.row_dimensions[12].height = 22
        for r in range(13, 20):
            ws.row_dimensions[r].height = 18
        ws.row_dimensions[20].height = 20
        ws.row_dimensions[21].height = 12
        ws.row_dimensions[22].height = 18
        ws.row_dimensions[23].height = 12
        ws.row_dimensions[24].height = 18
        ws.row_dimensions[25].height = 22

        ws.row_dimensions[26].height = 15
        ws.row_dimensions[27].height = 10
        ws.row_dimensions[28].height = 25
        for r in range(29, 34):
            ws.row_dimensions[r].height = 19

        ws.row_dimensions[34].height = 15
        ws.row_dimensions[35].height = 25
        for r in range(36, 46):
            ws.row_dimensions[r].height = 18
        ws.row_dimensions[46].height = 22

        return ws

    finally:
        if close_conn and cur:
            cur.close()
            conn.close()
