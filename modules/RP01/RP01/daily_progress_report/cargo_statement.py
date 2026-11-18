import json
import os
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def map_cargo_to_column_key(cargo_type, cargo_category, cargo_name):
    """
    Map vessel_cargo cargo_type and cargo_category (along with cargo_name)
    to the corresponding Cargo Quantity Statement column key.
    """
    ct = (cargo_type or '').strip().upper()
    cc = (cargo_category or '').strip().upper()
    cn = (cargo_name or '').strip().upper()

    # IBRM
    if ct == 'IBRM':
        if 'PELLET FEED' in cn or 'FINE' in cc or 'FINE' in cn:
            return 'ore_fines'
        elif 'CLO' in cc or 'LUMP' in cc or 'LUMP' in cn:
            return 'lumps'
        elif 'PELLET' in cc or 'PELLET' in cn:
            return 'pellets'
        elif 'HBI' in cc or 'HBI' in cn:
            return 'hbi'
        elif 'SCRAP' in cc or 'SCRAP' in cn:
            return 'scrap'
        return 'ore_fines'

    # FLUXES
    if 'FLUX' in ct:
        if 'LIMESTONE' in cn or 'LIMESTONE' in cc:
            return 'limestone'
        elif 'DOLOMITE' in cn or 'DOLOMITE' in cc:
            return 'dolomite'
        elif 'PYROXINITE' in cn or 'PYROXINITE' in cc:
            return 'pyroxinite'
        elif 'OLIVINE' in cn or 'OLI' in cn or 'OLIVINE' in cc:
            return 'olivine'
        elif 'BENTONITE' in cn:
            return 'bentonite'
        elif 'QUARTZITE' in cn:
            return 'quartzite'
        elif 'BENTONITE' in cc:
            return 'bentonite'
        return 'limestone'

    # CBRM
    if ct == 'CBRM':
        if 'BREEZE' in cn or 'BREEZE' in cc:
            return 'coke_coke_breeze'
        elif cc == 'COKE' or 'COKE' in cn:
            return 'coke_coke_breeze'
        else:
            return 'coking_coal_anthracite_pci'

    # Finish Goods
    if 'FINISH' in ct or ct == 'FG':
        if 'PELLET' in cc or 'PELLET' in cn:
            return 'pellet_loading'
        elif 'TMT' in cc or 'TMT' in cn or 'BILLET' in cn:
            return 'tmt_bar_loading'
        elif 'HRC' in cc or 'HRC' in cn:
            return 'hrc_loading'
        elif 'LUMP' in cc or 'LUMP' in cn:
            return 'lumps_loading'
        elif 'HARSCO' in cc or 'SLAG' in cc or 'HARSCO' in cn:
            return 'harsco_slag_loading'
        elif 'COKE' in cc or 'COKE' in cn:
            return 'coke_coke_breeze'

    # SLAG
    if 'SLAG' in ct:
        if 'GBFS' in cn or 'GBFS' in cc:
            return 'gbfs_loading'
        elif 'GGBS' in cn or 'GGBS' in cc:
            return 'ggbs_loading'
        return 'harsco_slag_loading'

    # Other commodities & Finished goods
    if 'CLINKER' in ct or 'CLINKER' in cc or 'CLINKER' in cn:
        return 'clinker'
    if 'AMMONI' in cn or 'SULPHATE' in cn or 'AMMONI' in cc or 'SULPHATE' in cc:
        return 'ammonium_sulphate'
    if 'MOP' in cn or 'MOP' in cc:
        return 'mop'
    if 'SALAV' in cn and 'LUMP' in cn:
        return 'salav_lumps'

    return None


def fetch_live_cargo_data_from_july(cur):
    """
    Fetch cargo quantities for July 2026 onwards by joining vessel_cargo
    to categorize cargo_type and cargo_category dynamically.
    Returns: { month_label: { column_key: quantity } }
    """
    if not cur:
        return {}

    live_data = {}
    try:
        # 1. Query lueu_lines (jetty cargo discharge) joined with vessel_cargo
        cur.execute("""
            SELECT
                TO_CHAR(TO_DATE(BTRIM(l.entry_date), 'YYYY-MM-DD'), 'Mon-YY') AS month_label,
                LOWER(TRIM(l.cargo_name)) AS cargo_name,
                COALESCE(vc.cargo_type, 'Other') AS cargo_type,
                COALESCE(vc.cargo_category, 'Other') AS cargo_category,
                SUM(COALESCE(l.quantity, 0)) AS qty
            FROM lueu_lines l
            LEFT JOIN vessel_cargo vc
                ON LOWER(TRIM(vc.cargo_name)) = LOWER(TRIM(l.cargo_name))
            WHERE l.is_deleted IS NOT TRUE
              AND l.quantity IS NOT NULL
              AND NULLIF(BTRIM(l.entry_date), '') IS NOT NULL
              AND TO_DATE(BTRIM(l.entry_date), 'YYYY-MM-DD') >= '2026-07-01'
            GROUP BY 1, 2, 3, 4
        """)
        for r in cur.fetchall():
            m = r['month_label']
            k = map_cargo_to_column_key(r['cargo_type'], r['cargo_category'], r['cargo_name'])
            if k and float(r['qty'] or 0) > 0:
                if m not in live_data:
                    live_data[m] = {}
                live_data[m][k] = live_data[m].get(k, 0) + float(r['qty'])
    except Exception:
        pass

    try:
        # 2. Query mbc_header joined with vessel_cargo
        cur.execute("""
            SELECT
                TO_CHAR(TO_DATE(SUBSTRING(mh.doc_date FROM 1 FOR 10), 'YYYY-MM-DD'), 'Mon-YY') AS month_label,
                LOWER(TRIM(mh.cargo_name)) AS cargo_name,
                COALESCE(vc.cargo_type, mh.cargo_type) AS cargo_type,
                COALESCE(vc.cargo_category, 'Other') AS cargo_category,
                SUM(COALESCE(mh.bl_quantity, 0)) AS qty
            FROM mbc_header mh
            LEFT JOIN vessel_cargo vc
                ON LOWER(TRIM(vc.cargo_name)) = LOWER(TRIM(mh.cargo_name))
            WHERE mh.doc_date IS NOT NULL
              AND mh.doc_date ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}'
              AND TO_DATE(SUBSTRING(mh.doc_date FROM 1 FOR 10), 'YYYY-MM-DD') >= '2026-07-01'
            GROUP BY 1, 2, 3, 4
        """)
        for r in cur.fetchall():
            m = r['month_label']
            k = map_cargo_to_column_key(r['cargo_type'], r['cargo_category'], r['cargo_name'])
            if k and float(r['qty'] or 0) > 0:
                if m not in live_data:
                    live_data[m] = {}
                live_data[m][k] = live_data[m].get(k, 0) + float(r['qty'])
    except Exception:
        pass

    return live_data


def build_cargo_quantity_statement_sheet(wb, cur=None):
    """
    Build Sheet 6: Cargo Quantity Statement.
    Loads historical throughput statements from cargo_quantity_statement.json
    (FY 12-13 through FY 25-26).
    For the current year (FY 26-27):
      - April, May, June are hardcoded.
      - July onwards is dynamically fetched from the database using vessel_cargo.
    """
    curr_dir = os.path.dirname(os.path.abspath(__file__))
    json_path = os.path.join(curr_dir, 'cargo_quantity_statement.json')
    if not os.path.exists(json_path):
        json_path = os.path.join('modules', 'RP01', 'RP01', 'daily_progress_report', 'cargo_quantity_statement.json')

    with open(json_path, 'r', encoding='utf-8') as f:
        all_years = json.load(f)

    # Fetch live cargo handling data for July 2026 onwards if DB cursor is provided
    live_fy26_data = fetch_live_cargo_data_from_july(cur) if cur else {}

    ws = wb.create_sheet(title="Cargo Quantity Statement")
    ws.views.sheetView[0].showGridLines = True

    red_title_font = Font(name='Calibri', size=11, bold=True, color='FF0000')
    sub_title_font = Font(name='Calibri', size=9, bold=False, color='002060')
    header_font = Font(name='Calibri', size=9, bold=True, color='002060')
    data_font = Font(name='Calibri', size=9, bold=False, color='002060')
    subtotal_font = Font(name='Calibri', size=9, bold=True, color='002060')
    total_font = Font(name='Calibri', size=9, bold=True, color='FF0000')

    thin_side = Side(style='thin', color='000000')
    double_side = Side(style='double', color='000000')
    dashed_side = Side(style='dashed', color='7F7F7F')

    box_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    data_border = Border(left=thin_side, right=thin_side, top=dashed_side, bottom=dashed_side)
    total_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_side)

    cur_row = 1

    for year_data in all_years:
        title = year_data['title']
        subtitle = year_data.get('subtitle', 'Cargo Quantity (MT)')
        total_label = year_data['total_label']
        max_col = year_data['max_col']
        has_dh = year_data.get('has_double_handling', False)
        columns = year_data['columns']
        rows = [dict(r) for r in year_data['rows']]
        is_current_year = (year_data.get('financial_year') == 'FY 26-27')

        # If current year, dynamically update July onwards using live data from vessel_cargo (no hardcoded values)
        if is_current_year:
            for r_data in rows:
                m = r_data.get('month')
                if m not in ('Apr-26', 'May-26', 'Jun-26'):
                    # July onwards is strictly dynamic from DB
                    if live_fy26_data and m in live_fy26_data:
                        for k, live_val in live_fy26_data[m].items():
                            r_data[k] = live_val
                    else:
                        for k in list(r_data.keys()):
                            if k != 'month':
                                r_data[k] = 0

        title_r = cur_row
        sub_r = cur_row + 1
        h1_r = cur_row + 2
        h2_r = cur_row + 3

        # Title
        ws.merge_cells(start_row=title_r, start_column=1, end_row=title_r, end_column=max_col)
        tc = ws.cell(title_r, 1, title)
        tc.font = red_title_font
        tc.alignment = Alignment(horizontal='center', vertical='center')

        # Subtitle
        ws.merge_cells(start_row=sub_r, start_column=1, end_row=sub_r, end_column=max_col)
        sc = ws.cell(sub_r, 1, subtitle)
        sc.font = sub_title_font
        sc.alignment = Alignment(horizontal='center', vertical='center')

        # Groups
        ibrm_cols = [c for c in columns if c.get('group') == 'IBRM']
        flux_cols = [c for c in columns if c.get('group') == 'Fluxes']
        cbrm_cols = [c for c in columns if c.get('group') == 'CBRM']
        single_cols = [c for c in columns if not c.get('group') and c.get('key') != 'month']

        # Header Row 1: Month
        ws.merge_cells(start_row=h1_r, start_column=1, end_row=h2_r, end_column=1)
        ws.cell(h1_r, 1, 'Month')

        # Group headers
        if ibrm_cols:
            start_c = ibrm_cols[0]['index']
            end_c = ibrm_cols[-1]['index']
            ws.merge_cells(start_row=h1_r, start_column=start_c, end_row=h1_r, end_column=end_c)
            ws.cell(h1_r, start_c, 'IBRM')

        if flux_cols:
            start_c = flux_cols[0]['index']
            end_c = flux_cols[-1]['index']
            ws.merge_cells(start_row=h1_r, start_column=start_c, end_row=h1_r, end_column=end_c)
            ws.cell(h1_r, start_c, 'Fluxes')

        if cbrm_cols:
            start_c = cbrm_cols[0]['index']
            end_c = cbrm_cols[-1]['index']
            ws.merge_cells(start_row=h1_r, start_column=start_c, end_row=h1_r, end_column=end_c)
            ws.cell(h1_r, start_c, 'CBRM')

        # Single commodity headers
        for sc in single_cols:
            c_idx = sc['index']
            if c_idx <= max_col:
                ws.merge_cells(start_row=h1_r, start_column=c_idx, end_row=h2_r, end_column=c_idx)
                ws.cell(h1_r, c_idx, sc['label'])

        # Header Row 2: Subheaders for IBRM, Fluxes, CBRM
        for c in ibrm_cols + flux_cols + cbrm_cols:
            c_idx = c['index']
            ws.cell(h2_r, c_idx, c['label'])

        # Borders & Fonts for headers
        for r in range(h1_r, h2_r + 1):
            for c in range(1, max_col + 1):
                cell = ws.cell(r, c)
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = box_border

        ws.row_dimensions[title_r].height = 20
        ws.row_dimensions[sub_r].height = 15
        ws.row_dimensions[h1_r].height = 18
        ws.row_dimensions[h2_r].height = 26

        first_data_r = h2_r + 1
        for i, r_data in enumerate(rows):
            r = first_data_r + i
            ws.row_dimensions[r].height = 18

            # Month cell
            mc = ws.cell(r, 1, r_data['month'])
            mc.font = data_font
            mc.alignment = Alignment(horizontal='center', vertical='center')
            mc.border = data_border

            has_float = False
            for col_info in columns[1:]:
                c_idx = col_info['index']
                col_let = col_info['col_letter']
                k = col_info['key']
                val = r_data.get(k)
                cell = ws.cell(r, c_idx)
                cell.font = data_font
                cell.alignment = Alignment(horizontal='right', vertical='center')
                cell.border = data_border

                # Subtotals
                if c_idx == 8: # IBRM Total
                    cell.value = f'=SUM(B{r}:G{r})'
                    cell.font = subtotal_font
                    cell.number_format = '#,##0.00' if any(isinstance(r_data.get(c['key']), float) for c in ibrm_cols[:-1]) else '#,##0'
                elif c_idx == 15: # Fluxes Total
                    cell.value = f'=SUM(I{r}:N{r})'
                    cell.font = subtotal_font
                    cell.number_format = '#,##0.00' if any(isinstance(r_data.get(c['key']), float) for c in flux_cols[:-1]) else '#,##0'
                elif c_idx == 18: # CBRM Total
                    cell.value = f'=SUM(P{r}:Q{r})'
                    cell.font = subtotal_font
                    cell.number_format = '#,##0.00' if any(isinstance(r_data.get(c['key']), float) for c in cbrm_cols[:-1]) else '#,##0'
                elif has_dh and c_idx == 28: # Total Col AB in Double Handling layout
                    cell.value = f'=H{r}+O{r}+R{r}+SUM(S{r}:AA{r})'
                    cell.font = subtotal_font
                    cell.number_format = '#,##0.00' if has_float else '#,##0'
                elif has_dh and c_idx == 30: # Grand Total Col AD in Double Handling layout
                    cell.value = f'=AB{r}+AC{r}'
                    cell.font = subtotal_font
                    cell.number_format = '#,##0.00' if has_float else '#,##0'
                elif not has_dh and c_idx == max_col: # Grand Total in standard layout
                    prev_col = get_column_letter(max_col - 1)
                    cell.value = f'=H{r}+O{r}+R{r}+SUM(S{r}:{prev_col}{r})'
                    cell.font = subtotal_font
                    cell.number_format = '#,##0.00' if has_float else '#,##0'
                else:
                    if val is not None and val != 0:
                        cell.value = val
                        if isinstance(val, float) and val % 1 != 0:
                            cell.number_format = '#,##0.00'
                            has_float = True
                        else:
                            cell.number_format = '#,##0'

        last_data_r = first_data_r + len(rows) - 1

        # Total row
        tot_r = last_data_r + 1
        ws.row_dimensions[tot_r].height = 20
        tc = ws.cell(tot_r, 1, total_label)
        tc.font = total_font
        tc.alignment = Alignment(horizontal='center', vertical='center')
        tc.border = total_border

        gt_col_let = get_column_letter(max_col)
        for col_info in columns[1:]:
            c_idx = col_info['index']
            col_let = col_info['col_letter']
            k = col_info['key']
            cell = ws.cell(tot_r, c_idx, f'=SUM({col_let}{first_data_r}:{col_let}{last_data_r})')
            cell.font = total_font
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = total_border
            has_float = any(isinstance(r.get(k), float) for r in rows)
            cell.number_format = '#,##0.00' if has_float else '#,##0'

        # Percentage row
        pct_r = tot_r + 1
        ws.row_dimensions[pct_r].height = 18
        ws.cell(pct_r, 1).border = box_border

        for col_info in columns[1:]:
            c_idx = col_info['index']
            col_let = col_info['col_letter']
            k = col_info['key']
            cell = ws.cell(pct_r, c_idx)
            pct_val = year_data['percentages'].get(k)
            if pct_val:
                cell.value = f'={col_let}{tot_r}/{gt_col_let}{tot_r}'
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = '0.00%'
                cell.border = box_border
            else:
                cell.border = Border(top=thin_side, bottom=thin_side)

        cur_row = pct_r + 3

    col_widths = {
        'A': 12, 'B': 13, 'C': 12, 'D': 12, 'E': 12, 'F': 12, 'G': 12, 'H': 14,
        'I': 13, 'J': 10, 'K': 12, 'L': 12, 'M': 10, 'N': 12, 'O': 14,
        'P': 14, 'Q': 15, 'R': 14, 'S': 12, 'T': 13, 'U': 13, 'V': 12,
        'W': 13, 'X': 12, 'Y': 12, 'Z': 14, 'AA': 14, 'AB': 14, 'AC': 14, 'AD': 15
    }
    for col_let, width in col_widths.items():
        ws.column_dimensions[col_let].width = width

    return ws
