import json
import math
import os
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.trendline import Trendline
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import RichTextProperties


def match_category_to_column(cargo_category, cargo_type, columns):
    """
    Dynamically find the column key for a given cargo_category and cargo_type
    by matching directly against the table's column labels and group headers.
    No hardcoded mapping dictionary is used.
    """
    if not cargo_category:
        return None

    cat = cargo_category.strip().upper()
    grp = (cargo_type or '').strip().upper()
    data_cols = [c for c in columns if not c.get('is_subtotal') and c.get('key') != 'month']

    # 1. Exact match with column label (e.g. 'Ore Fines', 'HBI', 'Limestone', 'Dolomite')
    for col in data_cols:
        col_lbl = col['label'].strip().upper()
        col_grp = (col.get('group') or '').strip().upper()
        if cat == col_lbl:
            if not col_grp or not grp or col_grp == grp:
                return col['key']

    # 2. Match within group (e.g. 'Coke' within CBRM -> 'Coke / Coke Breeze')
    for col in data_cols:
        col_lbl = col['label'].strip().upper()
        col_grp = (col.get('group') or '').strip().upper()
        if col_grp and grp and col_grp == grp:
            if cat in col_lbl or col_lbl in cat:
                return col['key']

    # 3. Known industry synonyms & abbreviations
    for col in data_cols:
        col_lbl = col['label'].strip().upper()
        col_grp = (col.get('group') or '').strip().upper()
        # CLO (Calibrated Lump Ore) -> Lumps
        if cat == 'CLO' and 'LUMP' in col_lbl and (not col_grp or col_grp == 'IBRM'):
            return col['key']
        # HMGP / MGP -> MGP (Pig Iron) Loading
        if ('HMGP' in cat or 'MGP' in cat) and 'MGP' in col_lbl:
            return col['key']
        # Pyrite -> Ore Fines
        if 'PYRITE' in cat and 'ORE FINES' in col_lbl:
            return col['key']
        # Slag / GBFS
        if 'SLAG' in cat and ('SLAG' in col_lbl or 'GBFS' in col_lbl):
            return col['key']
        # Sulphate
        if 'SULPHATE' in cat and 'SULPHATE' in col_lbl:
            return col['key']

    # 4. Global label contains search
    for col in data_cols:
        col_lbl = col['label'].strip().upper()
        if cat in col_lbl:
            return col['key']

    return None


def fetch_live_cargo_data_from_july(cur, columns):
    """
    Fetch cargo quantities for July 2026 onwards by querying operational tables:
    1. lueu_lines (Jetty operations)
    2. mbc_header (MBC discharge operations not already logged in lueu_lines)
    Dynamically matches each record's cargo_category from vessel_cargo directly
    to the corresponding column in the statement.
    Returns: { month_label: { column_key: quantity } }
    """
    if not cur or not columns:
        return {}

    live_data = {}

    # 1. Query lueu_lines joined with vessel_cargo
    try:
        cur.execute("""
            SELECT
                TO_CHAR(TO_DATE(BTRIM(l.entry_date), 'YYYY-MM-DD'), 'Mon-YY') AS month_label,
                COALESCE(vc.cargo_type, 'Other') AS cargo_type,
                COALESCE(vc.cargo_category, 'Other') AS cargo_category,
                SUM(COALESCE(l.quantity, 0)) AS qty
            FROM lueu_lines l
            LEFT JOIN vessel_cargo vc
                ON LOWER(TRIM(vc.cargo_name)) = LOWER(TRIM(l.cargo_name))
            WHERE (l.is_deleted IS FALSE OR l.is_deleted IS NULL)
              AND l.quantity IS NOT NULL
              AND NULLIF(BTRIM(l.entry_date), '') IS NOT NULL
              AND TO_DATE(BTRIM(l.entry_date), 'YYYY-MM-DD') >= '2026-07-01'
            GROUP BY 1, 2, 3
        """)
        for r in cur.fetchall():
            m = r['month_label']
            k = match_category_to_column(r['cargo_category'], r['cargo_type'], columns)
            if k and float(r['qty'] or 0) > 0:
                if m not in live_data:
                    live_data[m] = {}
                live_data[m][k] = live_data[m].get(k, 0) + float(r['qty'])
    except Exception as e:
        print("Error fetching live cargo data from lueu_lines:", e)

    # 2. Query mbc_header (for any MBC shipment not logged in lueu_lines)
    try:
        cur.execute("""
            SELECT
                TO_CHAR(TO_DATE(SUBSTRING(mh.doc_date FROM 1 FOR 10), 'YYYY-MM-DD'), 'Mon-YY') AS month_label,
                COALESCE(vc.cargo_type, mh.cargo_type, 'Other') AS cargo_type,
                COALESCE(vc.cargo_category, 'Other') AS cargo_category,
                SUM(COALESCE(mh.bl_quantity, 0)) AS qty
            FROM mbc_header mh
            LEFT JOIN vessel_cargo vc
                ON LOWER(TRIM(vc.cargo_name)) = LOWER(TRIM(mh.cargo_name))
            WHERE mh.doc_date IS NOT NULL
              AND mh.doc_date ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}'
              AND TO_DATE(SUBSTRING(mh.doc_date FROM 1 FOR 10), 'YYYY-MM-DD') >= '2026-07-01'
              AND mh.id NOT IN (
                  SELECT DISTINCT source_id 
                  FROM lueu_lines 
                  WHERE source_type = 'MBC' AND source_id IS NOT NULL
              )
            GROUP BY 1, 2, 3
        """)
        for r in cur.fetchall():
            m = r['month_label']
            k = match_category_to_column(r['cargo_category'], r['cargo_type'], columns)
            if k and float(r['qty'] or 0) > 0:
                if m not in live_data:
                    live_data[m] = {}
                live_data[m][k] = live_data[m].get(k, 0) + float(r['qty'])
    except Exception as e:
        print("Error fetching live cargo data from mbc_header:", e)

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
    # Dynamically matched to FY 26-27 columns by cargo_category
    fy26_columns = all_years[-1]['columns'] if all_years else []
    live_fy26_data = fetch_live_cargo_data_from_july(cur, fy26_columns) if cur else {}

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


def build_cargo_quantity_graph_sheet(wb):
    """
    Build Sheet 7: Cargo Quantity Graph.
    Displays annual cargo throughput comparison table and column bar chart
    with trendline matching the executive dashboard summary.
    """
    ws = wb.create_sheet(title="Cargo Quantity Graph")
    ws.views.sheetView[0].showGridLines = True

    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    thin_side = Side(style='thin', color='000000')
    medium_side = Side(style='medium', color='000000')

    table_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    font_bold = Font(name='Calibri', size=10, bold=True, color='000000')
    font_normal = Font(name='Calibri', size=10, bold=False, color='000000')
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')

    # Table data
    data = [
        ('FY-12-13 (half year)', 3221320),
        ('FY-13-14', 7406280),
        ('FY-14-15', 7381918),
        ('FY-15-16 (B/F Shutdown)', 5955334),
        ('FY-16-17', 10045790),
        ('FY-17-18', 11624347),
        ('FY-18-19', 12668782),
        ('FY-19-20', 12839750),
        ('FY- 20-21 (Pandemic COVID19)', 12903667),
        ('FY- 21-22', 17206902),
        ('FY- 22-23', 24041938),
        ('FY- 23-24', 25144780),
        ('FY- 24-25', 23138410),
        ('FY- 25-26', 24532005),
        ('FY- 26-27', "='Cargo Quantity Statement'!AD291"),
    ]

    # Chart sizing & dynamic container boundary calculation
    max_chart_row = 18
    num_bars = max_chart_row - 4
    chart_width_cm = max(25.0, num_bars * 1.7 + 3.0)
    chart_height_cm = 12.0

    chart = BarChart()
    chart.type = 'col'
    chart.legend = None
    chart.width = chart_width_cm
    chart.height = chart_height_cm

    chart_start_col = 8  # Column H
    chart_start_row = 4  # Row 4

    # Dynamically compute container end column so yellow background fully
    # spans behind the graph and scales as graph width increases
    col_width_chars = 8
    col_width_cm = (col_width_chars * 7 + 5) * (2.54 / 96.0)
    cols_spanned = int(math.ceil(chart.width / col_width_cm))
    container_end_col = chart_start_col + cols_spanned + 2

    # Dynamically compute container end row so yellow background frames the bottom
    row_height_cm = 18.0 * (2.54 / 72.0)
    rows_spanned = int(math.ceil(chart.height / row_height_cm))
    chart_end_row = chart_start_row + rows_spanned
    table_end_row = 4 + len(data)
    container_end_row = max(table_end_row, chart_end_row) + 2

    # Draw dynamic yellow container E3 : (container_end_col, container_end_row)
    for r in range(3, container_end_row + 1):
        for c in range(5, container_end_col + 1):
            cell = ws.cell(r, c)
            cell.fill = yellow_fill
            top_b = medium_side if r == 3 else None
            bot_b = medium_side if r == container_end_row else None
            left_b = medium_side if c == 5 else None
            right_b = medium_side if c == container_end_col else None
            cell.border = Border(left=left_b, right=right_b, top=top_b, bottom=bot_b)

    # Table headers
    cE4 = ws.cell(4, 5, 'Year')
    cE4.font = font_bold
    cE4.alignment = align_left
    cE4.border = Border(left=medium_side, right=thin_side, top=thin_side, bottom=thin_side)

    cF4 = ws.cell(4, 6, 'MT')
    cF4.font = font_bold
    cF4.alignment = align_left
    cF4.border = table_border

    for idx, (yr, val) in enumerate(data, 5):
        is_covid = 'COVID19' in yr
        f_font = font_bold if is_covid else font_normal

        cE = ws.cell(idx, 5, yr)
        cE.font = f_font
        cE.alignment = align_left
        cE.border = Border(left=medium_side, right=thin_side, top=thin_side, bottom=thin_side)

        cF = ws.cell(idx, 6, val)
        cF.font = f_font
        cF.alignment = align_right
        cF.border = table_border
        if isinstance(val, int):
            if idx >= 17:
                cF.number_format = '#,##,##,##0'
            else:
                cF.number_format = '0'
        else:
            cF.number_format = '#,##,##,##0'

    # Set column widths & row heights dynamically
    ws.column_dimensions['E'].width = 28
    ws.column_dimensions['F'].width = 16
    ws.column_dimensions['G'].width = 4
    for c in range(8, container_end_col + 3):
        ws.column_dimensions[get_column_letter(c)].width = col_width_chars

    ws.row_dimensions[3].height = 15
    for r in range(4, container_end_row):
        ws.row_dimensions[r].height = 18
    ws.row_dimensions[container_end_row].height = 15

    # Configure Axes matching reference styling
    chart.y_axis.majorGridlines = ChartLines()
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = 30000000
    chart.y_axis.majorUnit = 5000000
    chart.y_axis.axPos = 'l'
    chart.y_axis.tickLblPos = 'nextTo'

    chart.x_axis.axPos = 'b'
    chart.x_axis.tickLblPos = 'low'
    chart.x_axis.txPr = RichText(bodyPr=RichTextProperties(rot=-2700000))

    data_ref = Reference(ws, min_col=6, min_row=4, max_row=max_chart_row)
    cats_ref = Reference(ws, min_col=5, min_row=5, max_row=max_chart_row)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)

    # Clean data labels: only show values at the top of the bars
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True
    chart.dataLabels.showCatName = False
    chart.dataLabels.showSerName = False
    chart.dataLabels.showLegendKey = False
    chart.dataLabels.showPercent = False
    chart.dataLabels.position = 'outEnd'

    # Dotted linear trendline (no equation or R-squared text)
    tl = Trendline()
    tl.trendlineType = 'linear'
    tl.dispEq = False
    tl.dispRSqr = False
    gp = GraphicalProperties()
    gp.line = LineProperties(prstDash='dot')
    tl.spPr = gp

    s = chart.series[0]
    s.graphicalProperties.solidFill = '5B9BD5'
    s.trendline = tl

    ws.add_chart(chart, 'H4')

    return ws
