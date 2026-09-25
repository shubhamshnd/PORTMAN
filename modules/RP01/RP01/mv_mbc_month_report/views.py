from flask import render_template, request, jsonify, session, redirect, url_for, Response
from functools import wraps
from datetime import date, datetime, timedelta
from collections import defaultdict
import io
import re

from .. import bp
from database import get_db, get_cursor

# ── Excel colour / style constants ─────────────────────────────────────────
XL_GREY     = 'C0C0C0'
XL_LAVEND   = 'CCCCFF'
XL_CYAN     = 'CCFFFF'
XL_WHITE    = 'FFFFFF'
XL_NAVY     = '1F4E78'
XL_YELLOW   = 'FFFF00'
XL_TITLE_SZ = 14
XL_NORM_SZ  = 10
XL_SMALL_SZ = 9

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

_thin   = Side(style='thin',   color='000000')
_med    = Side(style='medium', color='000000')
_bdr    = Border(left=_thin,  right=_thin,  top=_thin,  bottom=_thin)
_bdr_ml = Border(left=_med,   right=_thin,  top=_thin,  bottom=_thin)
_ctr    = Alignment(horizontal='center', vertical='center', wrap_text=True)
_left   = Alignment(horizontal='left',   vertical='center', wrap_text=True)
_right  = Alignment(horizontal='right',  vertical='center', wrap_text=True)


def _fill(hex_color):
    return PatternFill('solid', fgColor=hex_color)


def _font(bold=False, size=XL_NORM_SZ, color='000000'):
    return Font(name='Calibri', bold=bold, size=size, color=color)


def _safe_float(val):
    try:
        return float(val or 0)
    except (TypeError, ValueError):
        return 0.0


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


# ── Report cutoff ────────────────────────────────────────────────────────
# Matches the frontend's hardcoded validation (2026-05-01). Vessels/MBCs
# whose LAST discharge activity happened before this date are dropped
# from the report entirely, even if they still have outstanding BL qty.
REPORT_CUTOFF_DATE = date(2026, 5, 1)


def _fetch_mv_monthly_data(
    from_date=None,
    to_date=None,
    operation_type=None
):
    """
    Pivot daily quantity data by Mother Vessel OR MBC.

    CHANGES:
    1. Date axis is now the FULL calendar range from_date..to_date
       (previously only dates that had at least one row showed up).
       to_date is capped at today so the current month never shows
       future empty dates; a past month you pick still shows every
       day of the range you selected.
    2. Carry-forward fix: vessels/MBCs whose BL qty has not been
       fully discharged yet now show up in the report even if they
       have ZERO lueu_lines rows in the selected period (i.e.
       discharge hasn't started). They appear with dashes for every
       date and their outstanding qty sitting in Previous Month Qty.
    3. NEW: any vessel/MBC whose LAST activity is before
       REPORT_CUTOFF_DATE (2026-05-01) is dropped from the carry
       -forward list entirely — old, inactive vessels no longer
       linger in the report with dashes.
    """

    conn = get_db()
    cur = get_cursor(conn)

    # ---- resolve period, default to current month, cap at today ----
    today = date.today()

    if not to_date:
        to_date = today.strftime('%Y-%m-%d')
    if not from_date:
        from_date = today.replace(day=1).strftime('%Y-%m-%d')

    from_date_obj = datetime.strptime(from_date, '%Y-%m-%d').date()
    to_date_obj   = min(datetime.strptime(to_date, '%Y-%m-%d').date(), today)
    to_date = to_date_obj.strftime('%Y-%m-%d')

    sql = """
    SELECT

        l.entry_date,
        l.quantity,
        l.cargo_name AS lueu_cargo,
        l.operation_type,
        l.source_type,
        l.source_id,

        COALESCE(
            CASE
                WHEN l.source_type = 'VCN'
                    THEN CONCAT(v.vcn_doc_num, ' / ', v.vessel_name)

                WHEN l.source_type = 'MBC'
                    THEN CONCAT(m.doc_num, ' / ', m.mbc_name)

                ELSE COALESCE(l.barge_name, 'Unknown')
            END,
            'Unknown'
        ) AS header_name,

        CASE
            WHEN l.source_type = 'VCN'
                THEN l.cargo_name

            WHEN l.source_type = 'MBC'
                THEN COALESCE(m.cargo_name, l.cargo_name)

            ELSE l.cargo_name
        END AS cargo_name,

        CASE
            WHEN l.source_type = 'VCN'
                THEN COALESCE(
                    vcargo.cargo_type,
                    l.cargo_name
                )

            WHEN l.source_type = 'MBC'
                THEN COALESCE(
                    m.cargo_type,
                    m.cargo_name,
                    l.cargo_name
                )

            ELSE l.cargo_name
        END AS cargo_type,

        CASE
            WHEN l.source_type = 'VCN'
                THEN COALESCE(vc_total.bl_quantity, 0)

            WHEN l.source_type = 'MBC'
                THEN COALESCE(mc_total.quantity, 0)

            ELSE 0
        END AS bl_qty,

        CASE
            WHEN l.source_type = 'VCN'
                THEN 'Mother Vessel'

            WHEN l.source_type = 'MBC'
                 AND UPPER(COALESCE(mm.mbc_owner_name, '')) LIKE '%%SHIPPING%%'
                THEN 'JSW SHIPPING'

            WHEN l.source_type = 'MBC'
                 AND UPPER(COALESCE(mm.mbc_owner_name, '')) LIKE '%%INFRA%%'
                THEN 'JSW INFRA'

            WHEN l.source_type = 'MBC'
                THEN 'OTHERS'

            ELSE 'Unknown'
        END AS company

    FROM lueu_lines l

    LEFT JOIN vcn_header v
        ON l.source_type = 'VCN'
       AND l.source_id = v.id

    LEFT JOIN LATERAL (
        SELECT cargo_type
        FROM vessel_cargo
        WHERE cargo_name = l.cargo_name
        LIMIT 1
    ) vcargo ON l.source_type = 'VCN'

    LEFT JOIN (
        SELECT
            vcn_id,
            SUM(bl_quantity) AS bl_quantity
        FROM vcn_cargo_declaration
        GROUP BY vcn_id
    ) vc_total
        ON v.id = vc_total.vcn_id

    LEFT JOIN mbc_header m
        ON l.source_type = 'MBC'
       AND l.source_id = m.id

    LEFT JOIN mbc_master mm
        ON UPPER(TRIM(m.mbc_name)) = UPPER(TRIM(mm.mbc_name))

    LEFT JOIN (
        SELECT
            mbc_id,
            cargo_name,
            SUM(quantity) AS quantity
        FROM mbc_customer_details
        GROUP BY mbc_id, cargo_name
    ) mc
        ON m.id = mc.mbc_id
       AND mc.cargo_name = l.cargo_name

    LEFT JOIN (
        SELECT
            mbc_id,
            SUM(quantity) AS quantity
        FROM mbc_customer_details
        GROUP BY mbc_id
    ) mc_total
        ON m.id = mc_total.mbc_id

    WHERE l.is_deleted IS NOT TRUE
      AND l.quantity IS NOT NULL
      AND l.quantity > 0
      AND l.source_type IN ('VCN', 'MBC')
    """

    params = []

    if from_date:
        sql += " AND l.entry_date >= %s"
        params.append(from_date)

    if to_date:
        sql += " AND l.entry_date <= %s"
        params.append(to_date)

    if operation_type:
        sql += " AND l.operation_type = %s"
        params.append(operation_type)

    sql += """
        ORDER BY
            l.entry_date DESC,
            header_name
    """

    cur.execute(sql, params)
    rows = cur.fetchall()
    conn.close()

    vessel_meta = {}
    date_vessel_qty = defaultdict(lambda: defaultdict(float))

    for r in rows:

        v_name = re.sub(
            r'\s+', ' ',
            (r['header_name'] or '').strip()
        ).upper()

        dt = str(r['entry_date'])
        qty = _safe_float(r['quantity'])

        date_vessel_qty[dt][v_name] += qty

        if v_name not in vessel_meta:
            vessel_meta[v_name] = {
                'vessel_name':        v_name,
                'cargo_name':         r['cargo_name'] or r['lueu_cargo'] or '-',
                'cargo_type':         r['cargo_type'] or r['lueu_cargo'] or '-',

                'cargo_names': [],

                'bl_qty': _safe_float(r['bl_qty']),
                'company':            r['company'] or 'Mother Vessel',
                'source_type':        r['source_type'],
                'previous_month_qty': 0
            }

        cn = (r['cargo_name'] or '').strip()

        if cn and cn not in vessel_meta[v_name]['cargo_names']:
            vessel_meta[v_name]['cargo_names'].append(cn)

    # ─────────────────────────────────────────────
    # FULL CALENDAR DATE AXIS (from_date .. to_date)
    # ─────────────────────────────────────────────

    all_dates_list = []
    _d = from_date_obj
    while _d <= to_date_obj:
        all_dates_list.append(_d.strftime('%Y-%m-%d'))
        _d += timedelta(days=1)

    # ─────────────────────────────────────────────
    # Previous Month Continuing Qty
    # ─────────────────────────────────────────────

    prev_month_end   = from_date_obj - timedelta(days=1)
    prev_month_start = prev_month_end.replace(day=1)

    conn2 = get_db()
    cur2  = get_cursor(conn2)

    prev_sql = """
    SELECT
        COALESCE(
            CASE
                WHEN l.source_type = 'VCN'
                    THEN CONCAT(v.vcn_doc_num, ' / ', v.vessel_name)

                WHEN l.source_type = 'MBC'
                    THEN CONCAT(m.doc_num, ' / ', m.mbc_name)

                ELSE COALESCE(l.barge_name, 'Unknown')
            END,
            'Unknown'
        ) AS header_name,

        COALESCE(SUM(l.quantity), 0) AS prev_qty

    FROM lueu_lines l

    LEFT JOIN vcn_header v
        ON l.source_type = 'VCN'
       AND l.source_id = v.id

    LEFT JOIN mbc_header m
        ON l.source_type = 'MBC'
       AND l.source_id = m.id

    LEFT JOIN mbc_master mm
        ON UPPER(TRIM(m.mbc_name)) = UPPER(TRIM(mm.mbc_name))

    WHERE l.is_deleted IS NOT TRUE
      AND l.quantity IS NOT NULL
      AND l.quantity > 0
      AND l.source_type IN ('VCN', 'MBC')
      AND l.entry_date <= %s

    GROUP BY header_name
    """

    # NOTE: "<= prev_month_end" (all-time up to period start) instead of
    # a start/end window, so a vessel that has been sitting idle for
    # several months still carries its full outstanding qty forward.
    cur2.execute(prev_sql, [prev_month_end.strftime('%Y-%m-%d')])

    prev_rows = cur2.fetchall()
    conn2.close()

    prev_qty_map = {}

    for pr in prev_rows:
        vn = re.sub(
            r'\s+', ' ',
            (pr['header_name'] or '').strip()
        ).upper()
        prev_qty_map[vn] = _safe_float(pr['prev_qty'])

    for v_name in vessel_meta:
        vessel_meta[v_name]['previous_month_qty'] = prev_qty_map.get(v_name, 0)

    # ─────────────────────────────────────────────
    # CARRY-FORWARD: BL exists but discharge hasn't
    # started / hasn't finished, and it has ZERO rows
    # in this period so it never entered vessel_meta above.
    # ─────────────────────────────────────────────

    conn3 = get_db()
    cur3  = get_cursor(conn3)

    carry_sql = """
        SELECT
            CONCAT(v.vcn_doc_num, ' / ', v.vessel_name) AS header_name,
            'VCN' AS source_type,
            vcd.cargo_name AS cargo_name,
            COALESCE(vcargo.cargo_type, vcd.cargo_name) AS cargo_type,
            'Mother Vessel' AS company,
            COALESCE(vc_total.bl_quantity, 0) AS bl_qty,
            COALESCE(disc_total.disc_qty, 0)  AS discharged_qty,
            v.created_date                      AS created_date
        FROM vcn_header v
        LEFT JOIN vcn_cargo_declaration vcd
            ON vcd.vcn_id = v.id
        LEFT JOIN LATERAL (
            SELECT cargo_type FROM vessel_cargo
            WHERE cargo_name = vcd.cargo_name
            LIMIT 1
        ) vcargo ON true
        LEFT JOIN (
            SELECT vcn_id, SUM(bl_quantity) AS bl_quantity
            FROM vcn_cargo_declaration
            GROUP BY vcn_id
        ) vc_total ON vc_total.vcn_id = v.id
        LEFT JOIN (
            SELECT source_id,
                   SUM(quantity) AS disc_qty
            FROM lueu_lines
            WHERE is_deleted IS NOT TRUE AND source_type = 'VCN'
            GROUP BY source_id
        ) disc_total ON disc_total.source_id = v.id

        UNION ALL

        SELECT
            CONCAT(m.doc_num, ' / ', m.mbc_name) AS header_name,
            'MBC' AS source_type,
            COALESCE(m.cargo_name, mcd.cargo_name) AS cargo_name,
            COALESCE(m.cargo_type, m.cargo_name)   AS cargo_type,
            CASE
                WHEN UPPER(COALESCE(mm.mbc_owner_name,'')) LIKE '%%SHIPPING%%' THEN 'JSW SHIPPING'
                WHEN UPPER(COALESCE(mm.mbc_owner_name,'')) LIKE '%%INFRA%%'    THEN 'JSW INFRA'
                ELSE 'OTHERS'
            END AS company,
            COALESCE(mc_total.quantity, 0) AS bl_qty,
            COALESCE(disc_total.disc_qty, 0) AS discharged_qty,
            m.created_date                     AS created_date
        FROM mbc_header m
        LEFT JOIN mbc_master mm
            ON UPPER(TRIM(m.mbc_name)) = UPPER(TRIM(mm.mbc_name))
        LEFT JOIN mbc_customer_details mcd
            ON mcd.mbc_id = m.id
        LEFT JOIN (
            SELECT mbc_id, SUM(quantity) AS quantity
            FROM mbc_customer_details
            GROUP BY mbc_id
        ) mc_total ON mc_total.mbc_id = m.id
        LEFT JOIN (
            SELECT source_id,
                   SUM(quantity) AS disc_qty
            FROM lueu_lines
            WHERE is_deleted IS NOT TRUE AND source_type = 'MBC'
            GROUP BY source_id
        ) disc_total ON disc_total.source_id = m.id
    """

    cur3.execute(carry_sql)
    carry_rows = cur3.fetchall()
    conn3.close()

    for cr in carry_rows:
        v_name = re.sub(
            r'\s+', ' ',
            (cr['header_name'] or '').strip()
        ).upper()

        # already present via current-period lueu_lines rows -> skip
        if v_name in vessel_meta:
            continue

        bl_qty        = _safe_float(cr['bl_qty'])
        discharged    = _safe_float(cr['discharged_qty'])
        outstanding   = round(bl_qty - discharged, 2)

        # only carry it forward if there is still undischarged qty
        if bl_qty <= 0 or outstanding <= 0:
            continue

        # Drop vessels/MBCs that were CREATED before the report cutoff
        # (2026-05-01) — old vessels/MBCs entered into the system before
        # this date never carry forward, no matter how much BL is
        # outstanding. created_date is stored as TEXT (YYYY-MM-DD).
        created_date_raw = cr.get('created_date')

        if not created_date_raw:
            continue

        try:
            created_date_obj = datetime.strptime(
                str(created_date_raw).strip()[:10], '%Y-%m-%d'
            ).date()
        except Exception:
            continue   # unparseable -> skip rather than risk wrong inclusion

        if created_date_obj < REPORT_CUTOFF_DATE:
            continue

        cn = (cr['cargo_name'] or '').strip()

        vessel_meta[v_name] = {
            'vessel_name':        v_name,
            'cargo_name':         cn or '-',
            'cargo_type':         cr['cargo_type'] or cn or '-',
            'cargo_names':        [cn] if cn else [],
            'bl_qty':             bl_qty,
            'company':            cr['company'] or 'Mother Vessel',
            'source_type':        cr['source_type'],
            'previous_month_qty': discharged,   # everything discharged so far, carried in
        }

        # zero quantity for every date in the current axis -> shows as "-"

    # ─────────────────────────────────────────────
    # Sorting & Final Output
    # ─────────────────────────────────────────────

    vessel_totals = {}

    for v_name in vessel_meta:
        total = 0
        for dt in all_dates_list:
            total += date_vessel_qty[dt].get(v_name, 0)
        vessel_totals[v_name] = total

    # keep a vessel if it had discharge activity in period OR is
    # a valid carry-forward candidate (BL still outstanding)
    vessels_with_data = [
        v for v in vessel_meta.values()
        if vessel_totals.get(v['vessel_name'], 0) > 0
        or round(v['bl_qty'] - v['previous_month_qty'], 2) > 0
    ]

    def _vessel_sort_key(v):
        name = v['vessel_name'].upper()
        if v['source_type'] == 'VCN':
            return (0, name)
        else:
            m = re.search(r'MBC(\d+)', name)
            num = int(m.group(1)) if m else 999999999
            return (1, num)

    for v in vessel_meta.values():
        if v['cargo_names']:
            v['cargo_name'] = ' / '.join(v['cargo_names'])

    sorted_vessels = sorted(vessels_with_data, key=_vessel_sort_key)

    return {
        'vessels': sorted_vessels,
        'dates':   all_dates_list,
        'data': {
            dt: dict(date_vessel_qty[dt])
            for dt in all_dates_list
        }
    }

# ── Cargo Summary fetch ─────────────────────────────────────────────────────

def _fetch_cargo_summary(from_date=None, to_date=None):

    conn = get_db()
    cur  = get_cursor(conn)

    sql = """
        SELECT
            l.quantity,
            l.source_type,

            CASE
                WHEN l.source_type = 'VCN'
                    THEN COALESCE(
                        vcargo.cargo_type,
                        l.cargo_name
                    )

                WHEN l.source_type = 'MBC'
                    THEN COALESCE(
                        m.cargo_type,
                        m.cargo_name,
                        l.cargo_name
                    )

                ELSE COALESCE(l.cargo_name, 'Unknown')
            END AS cargo_type,

            CASE
                WHEN l.source_type = 'VCN'
                    THEN 'MV'

                WHEN l.source_type = 'MBC'
                     AND UPPER(COALESCE(mm.mbc_owner_name, '')) LIKE '%%SHIPPING%%'
                    THEN 'MBC-Shipping'

                WHEN l.source_type = 'MBC'
                     AND UPPER(COALESCE(mm.mbc_owner_name, '')) LIKE '%%INFRA%%'
                    THEN 'MBC-Infra'

                WHEN l.source_type = 'MBC'
                    THEN 'Other MBC'

                ELSE 'Double Handling'
            END AS op_category

        FROM lueu_lines l

        LEFT JOIN vcn_header v
            ON l.source_type = 'VCN'
           AND l.source_id = v.id

        LEFT JOIN LATERAL (
            SELECT cargo_type
            FROM vessel_cargo
            WHERE cargo_name = l.cargo_name
            LIMIT 1
        ) vcargo ON l.source_type = 'VCN'

        LEFT JOIN mbc_header m
            ON l.source_type = 'MBC'
           AND l.source_id = m.id

        LEFT JOIN mbc_master mm
            ON UPPER(TRIM(m.mbc_name)) = UPPER(TRIM(mm.mbc_name))

        WHERE l.is_deleted IS NOT TRUE
          AND l.quantity IS NOT NULL
          AND l.quantity > 0
          AND l.source_type IN ('VCN', 'MBC')
    """

    params = []

    if from_date:
        sql += " AND l.entry_date >= %s"
        params.append(from_date)

    if to_date:
        sql += " AND l.entry_date <= %s"
        params.append(to_date)

    sql += " ORDER BY cargo_type, op_category"

    cur.execute(sql, params)
    rows = cur.fetchall()
    conn.close()

    cargo_op_qty = defaultdict(lambda: defaultdict(float))
    all_cargos   = set()

    # Fixed ordered list — no 'Total' here
    all_ops = [
        'MV',
        'MBC-Shipping',
        'MBC-Infra',
        'Other MBC',
        'Double Handling'
    ]

    for r in rows:
        cargo = (r['cargo_type'] or 'Unknown').strip()
        op    = (r['op_category'] or 'Other MBC')
        qty   = _safe_float(r['quantity'])
        all_cargos.add(cargo)
        cargo_op_qty[cargo][op] += qty

    result_data = {}

    # ✅ Loop is complete BEFORE return
    for cargo in all_cargos:
        row = dict(cargo_op_qty[cargo])
        row['Total'] = (
            row.get('MV', 0)
            + row.get('MBC-Shipping', 0)
            + row.get('MBC-Infra', 0)
            + row.get('Other MBC', 0)
            + row.get('Double Handling', 0)
        )
        result_data[cargo] = row

    # ✅ 'Total' added ONCE, outside the loop
    all_ops_with_total = all_ops + ['Total']

    return {
        'cargos':     sorted(all_cargos),
        'operations': all_ops_with_total,
        'data':       result_data,
    }


# ── Financial-year month axis (Apr → Mar) ───────────────────────────────────

def _add_months(d, n):
    """Pure-stdlib month-add helper (no dateutil dependency needed)."""
    month = d.month - 1 + n
    year = d.year + month // 12
    month = month % 12 + 1
    return date(year, month, 1)


def _get_fy_months(reference_date=None):
    """
    Returns 12 (month_start, month_end, label) tuples for the financial
    year (Apr -> Mar) that `reference_date` falls in. Mirrors the fixed
    Apr-25 .. Mar-26 row axis seen in the MBC / YTD sheets.
    """
    if reference_date is None:
        reference_date = date.today()

    if reference_date.month >= 4:
        fy_start = date(reference_date.year, 4, 1)
    else:
        fy_start = date(reference_date.year - 1, 4, 1)

    months = []
    for i in range(12):
        m_start = _add_months(fy_start, i)
        m_end   = _add_months(fy_start, i + 1) - timedelta(days=1)
        label   = m_start.strftime('%b-%y')
        months.append((m_start, m_end, label))
    return months


# ── Dynamic Load Ports fetcher ────────────────────────────────────────────

def _normalize_port_name(port_name):
    """Normalize load port aliases (e.g. merge 'JSW JAIGAD' and 'JSW JAIGAD PORT')."""
    if not port_name:
        return ""
    p = str(port_name).strip()
    if 'JAIGAD' in p.upper():
        return 'JSW JAIGAD PORT'
    return p


def _fetch_load_ports(source_type=None):
    """
    Dynamically fetches distinct load ports from vcn_header and/or mbc_header.
    If source_type == 'MBC', fetches only from mbc_header.
    If source_type == 'VCN', fetches only from vcn_header.
    """
    conn = get_db()
    cur  = get_cursor(conn)

    if source_type == 'MBC':
        sql = """
            SELECT DISTINCT TRIM(load_port) AS port
            FROM mbc_header
            WHERE load_port IS NOT NULL AND TRIM(load_port) != ''
            ORDER BY port
        """
    elif source_type == 'VCN':
        sql = """
            SELECT DISTINCT TRIM(load_port) AS port
            FROM vcn_header
            WHERE load_port IS NOT NULL AND TRIM(load_port) != ''
            ORDER BY port
        """
    else:
        sql = """
            SELECT DISTINCT TRIM(load_port) AS port
            FROM (
                SELECT load_port FROM vcn_header WHERE load_port IS NOT NULL AND TRIM(load_port) != ''
                UNION
                SELECT load_port FROM mbc_header WHERE load_port IS NOT NULL AND TRIM(load_port) != ''
            ) combined_ports
            ORDER BY port
        """
    ports = []
    seen = set()
    try:
        cur.execute(sql)
        rows = cur.fetchall()
        for r in rows:
            p = r.get('port') if isinstance(r, dict) else r[0]
            if p and p.strip():
                norm = _normalize_port_name(p)
                if norm and norm.upper() not in seen:
                    seen.add(norm.upper())
                    ports.append(norm)
    except Exception:
        pass
    finally:
        conn.close()

    if ports:
        return ports

    return [
        'Jaigad', 'Salav', 'Haldia', 'Hazira', 'Kandla',
        'Goa', 'Mundra', 'Dharamtar', 'Karanja', 'New Manglore',
    ]


# ── Unified vessel/MBC "call" ledger (feeds Yearly, MBC status, YTD) ───────

def _fetch_vessel_call_ledger():
    """
    One row per Mother Vessel (VCN) call or MBC call, with load port,
    company/stevedore category, cargo, BL qty, discharged qty, and the
    activity date range pulled from lueu_lines. This is the single source
    of truth the other three sheets are built from.
    """
    conn = get_db()
    cur  = get_cursor(conn)

    sql = """
    SELECT
        'VCN' AS source_type,
        v.id AS source_id,
        CONCAT(v.vcn_doc_num, ' / ', v.vessel_name) AS header_name,
        v.load_port AS load_port,
        'Mother Vessel' AS company,
        COALESCE(vcargo.cargo_type, vcd.cargo_name) AS cargo_type,
        vcd.cargo_name AS cargo_name,
        COALESCE(vc_total.bl_quantity, 0) AS bl_qty,
        COALESCE(disc.disc_qty, 0) AS discharged_qty,
        disc.min_date AS first_activity_date,
        disc.max_date AS last_activity_date,
        v.doc_date AS doc_date,
        vn.vessel_run_type AS vessel_run_type,
        vn.eta AS eta,
        vn.etd AS etd,
        ld.discharge_commenced AS discharge_commenced,
        ld.discharge_completed AS discharge_completed

    FROM vcn_header v

    LEFT JOIN vcn_cargo_declaration vcd
        ON vcd.vcn_id = v.id

    LEFT JOIN LATERAL (
        SELECT cargo_type FROM vessel_cargo
        WHERE cargo_name = vcd.cargo_name
        LIMIT 1
    ) vcargo ON true

    LEFT JOIN (
        SELECT vcn_id, SUM(bl_quantity) AS bl_quantity
        FROM vcn_cargo_declaration
        GROUP BY vcn_id
    ) vc_total ON vc_total.vcn_id = v.id

    LEFT JOIN (
        SELECT source_id,
               SUM(quantity)   AS disc_qty,
               MIN(entry_date) AS min_date,
               MAX(entry_date) AS max_date
        FROM (
            SELECT source_id, quantity, entry_date::text AS entry_date
            FROM lueu_lines
            WHERE is_deleted IS NOT TRUE AND source_type = 'VCN'

            UNION ALL

            SELECT v.id AS source_id, h.quantity, h.entry_date::text AS entry_date
            FROM rp01_historical_lueu h
            JOIN vcn_header v ON (UPPER(TRIM(h.source_display)) = UPPER(TRIM(v.vessel_name)) OR UPPER(TRIM(h.source_display)) LIKE CONCAT('%%', UPPER(TRIM(v.vessel_name)), '%%'))
        ) vcn_disc_combined
        GROUP BY source_id
    ) disc ON disc.source_id = v.id

    LEFT JOIN vcn_nominations vn ON vn.vcn_id = v.id
    LEFT JOIN ldud_header ld ON ld.vcn_id = v.id

    UNION ALL

    SELECT
        'MBC' AS source_type,
        m.id AS source_id,
        CONCAT(m.doc_num, ' / ', m.mbc_name) AS header_name,
        m.load_port AS load_port,
        CASE
            WHEN UPPER(COALESCE(mm.mbc_owner_name, '')) LIKE '%%SHIPPING%%' THEN 'JSW SHIPPING'
            WHEN UPPER(COALESCE(mm.mbc_owner_name, '')) LIKE '%%INFRA%%'    THEN 'JSW INFRA'
            ELSE 'OTHERS'
        END AS company,
        COALESCE(m.cargo_type, m.cargo_name) AS cargo_type,
        m.cargo_name AS cargo_name,
        COALESCE(mc_total.quantity, 0) AS bl_qty,
        COALESCE(disc.disc_qty, 0) AS discharged_qty,
        disc.min_date AS first_activity_date,
        disc.max_date AS last_activity_date,
        m.doc_date AS doc_date,
        NULL AS vessel_run_type,
        mlp.arrived_load_port AS eta,
        mlp.cast_off_load_port AS etd,
        mdp.unloading_commenced AS discharge_commenced,
        mdp.unloading_completed AS discharge_completed

    FROM mbc_header m

    LEFT JOIN mbc_master mm
        ON UPPER(TRIM(m.mbc_name)) = UPPER(TRIM(mm.mbc_name))

    LEFT JOIN (
        SELECT mbc_id, SUM(quantity) AS quantity
        FROM mbc_customer_details
        GROUP BY mbc_id
    ) mc_total ON mc_total.mbc_id = m.id

    LEFT JOIN (
        SELECT source_id,
               SUM(quantity)   AS disc_qty,
               MIN(entry_date) AS min_date,
               MAX(entry_date) AS max_date
        FROM (
            SELECT source_id, quantity, entry_date::text AS entry_date
            FROM lueu_lines
            WHERE is_deleted IS NOT TRUE AND source_type = 'MBC'

            UNION ALL

            SELECT m.id AS source_id, h.quantity, h.entry_date::text AS entry_date
            FROM rp01_historical_lueu h
            JOIN mbc_header m ON UPPER(TRIM(h.source_display)) = UPPER(TRIM(m.mbc_name))
        ) mbc_disc_combined
        GROUP BY source_id
    ) disc ON disc.source_id = m.id

    LEFT JOIN mbc_load_port_lines mlp ON mlp.mbc_id = m.id
    LEFT JOIN mbc_discharge_port_lines mdp ON mdp.mbc_id = m.id
    """

    cur.execute(sql)
    rows = cur.fetchall()
    conn.close()

    calls = []
    for r in rows:
        d = dict(r)

        # resolve the month-bucketing date: prefer first discharge activity,
        # fall back to the document date if the call has no lueu_lines yet
        bucket_raw = d.get('first_activity_date') or d.get('doc_date')
        bucket_date = None
        if bucket_raw:
            try:
                bucket_date = datetime.strptime(
                    str(bucket_raw).strip()[:10], '%Y-%m-%d'
                ).date()
            except Exception:
                bucket_date = None
        d['bucket_date'] = bucket_date

        d['outstanding_qty'] = round(
            _safe_float(d.get('bl_qty')) - _safe_float(d.get('discharged_qty')), 2
        )
        calls.append(d)

    return calls


def _fetch_direct_barge_count_by_month():
    """
    lueu_lines and rp01_historical_lueu rows with no VCN/MBC parent
    are treated as direct barge activity.

    Existing direct-barge counts are preserved.

    If a month has no direct-barge count, live VCN barge activity
    from lueu_lines is used for that month.

    Returns {(year, month): count}.
    """
    conn = get_db()
    cur = get_cursor(conn)

    # ---------------------------------------------------------
    # Existing direct barge logic
    # DO NOT change this logic
    # ---------------------------------------------------------
    cur.execute("""
        SELECT entry_date::text AS entry_date
        FROM lueu_lines
        WHERE is_deleted IS NOT TRUE
          AND (
                source_type IS NULL
                OR TRIM(source_type) = ''
                OR source_type NOT IN ('VCN', 'MBC')
              )

        UNION ALL

        SELECT entry_date::text AS entry_date
        FROM rp01_historical_lueu
        WHERE (
                source_display IS NULL
                OR TRIM(source_display) = ''
                OR (
                    mv_mbc IS NOT NULL
                    AND UPPER(TRIM(mv_mbc))
                        NOT IN ('MV', 'VCN', 'MBC')
                )
              )
    """)

    rows = cur.fetchall()

    direct_counts = defaultdict(int)

    for r in rows:
        raw = r.get('entry_date') if isinstance(r, dict) else r[0]

        if not raw:
            continue

        try:
            d = datetime.strptime(
                str(raw).strip()[:10],
                '%Y-%m-%d'
            ).date()
        except Exception:
            continue

        direct_counts[(d.year, d.month)] += 1

    # ---------------------------------------------------------
    # LIVE VCN BARGE DATA
    # Only used when existing direct-barge count is zero.
    # ---------------------------------------------------------
    cur.execute("""
        SELECT
            entry_date::text AS entry_date,
            source_id,
            UPPER(TRIM(barge_name)) AS barge_name
        FROM lueu_lines
        WHERE is_deleted IS NOT TRUE
          AND UPPER(TRIM(source_type)) = 'VCN'
          AND COALESCE(TRIM(barge_name), '') <> ''
          AND entry_date::text ~
              '^[0-9]{4}-(0[1-9]|1[0-2])-[0-9]{2}'
    """)

    live_rows = cur.fetchall()
    conn.close()

    live_barges = defaultdict(set)

    for r in live_rows:
        if isinstance(r, dict):
            raw = r.get('entry_date')
            source_id = r.get('source_id')
            barge_name = r.get('barge_name')
        else:
            raw = r[0]
            source_id = r[1]
            barge_name = r[2]

        if not raw or not barge_name:
            continue

        try:
            d = datetime.strptime(
                str(raw).strip()[:10],
                '%Y-%m-%d'
            ).date()
        except Exception:
            continue

        # Unique VCN + barge combination
        live_barges[(d.year, d.month)].add(
            (
                source_id,
                str(barge_name).strip().upper()
            )
        )

    # ---------------------------------------------------------
    # FINAL COUNTS
    #
    # Existing Apr/May values remain unchanged.
    # Live VCN barges are used only for months where the
    # existing direct-barge count is zero.
    # ---------------------------------------------------------
    counts = defaultdict(int)

    for key, value in direct_counts.items():
        counts[key] = value

    for key, barges in live_barges.items():
        if counts.get(key, 0) == 0:
            counts[key] = len(barges)

    return counts


# ── MBC Handling Status fetch ───────────────────────────────────────────────

def _fetch_mbc_handling_status(reference_date=None, ports=None):
    """
    Recreates the 'MBC' sheet: monthly count of calls per load port
    (left block, MBC calls only) and monthly count of MBC calls per
    company category + direct barge activity (right block).
    """
    if not ports:
        ports = _fetch_load_ports(source_type='MBC')

    calls  = _fetch_vessel_call_ledger()
    months = _get_fy_months(reference_date)
    barge_counts = _fetch_direct_barge_count_by_month()

    port_rows = []
    company_rows = []

    for m_start, m_end, label in months:

        mbc_month_calls = [
            c for c in calls
            if c['source_type'] == 'MBC'
            and c['bucket_date'] and m_start <= c['bucket_date'] <= m_end
        ]

        port_counts = {port: 0 for port in ports}
        for c in mbc_month_calls:
            lp = _normalize_port_name(c.get('load_port'))
            if lp:
                for port in ports:
                    if port.upper() == lp.upper():
                        port_counts[port] += 1
                        break
        port_total = sum(port_counts.values())

        port_rows.append({
            'label': label,
            'month_start': m_start,
            'ports': port_counts,
            'total': port_total,
        })

        shipping_ct = sum(1 for c in mbc_month_calls if c['company'] == 'JSW SHIPPING')
        infra_ct    = sum(1 for c in mbc_month_calls if c['company'] == 'JSW INFRA')
        other_ct    = sum(1 for c in mbc_month_calls if c['company'] == 'OTHERS')
        barge_ct    = barge_counts.get((m_start.year, m_start.month), 0)
        comp_total  = shipping_ct + infra_ct + other_ct
        grand_total = comp_total + barge_ct

        company_rows.append({
            'label': label,
            'month_start': m_start,
            'jsw_shipping': shipping_ct,
            'jsw_infra': infra_ct,
            'other': other_ct,
            'barges': barge_ct,
            'total': comp_total,
            'grand_total': grand_total,
        })

    return {
        'ports': ports,
        'port_rows': port_rows,
        'company_rows': company_rows,
    }


# ── YTD Cargo-type x Month pivot fetch ──────────────────────────────────────

def _fetch_ytd_cargo_pivot(reference_date=None):
    """
    Recreates the 'YTD' sheet as specified in the template:
    - Rows: Operation categories / MBC load port categories:
        1. 'Mother Vessels Cargo' (VCN calls)
        2. 'Jaigad-Other MBC'
        3. 'Jsw Infrastructre'
        4. 'Shipping'
        5. 'Salav -Other'
        6. 'Goa-Other'
        7. any other dynamic category / Double Handling
    - Columns: 12 Financial Year months (Apr -> Mar).
    - Sub-columns under each month: All distinct Cargo Types (e.g. IBRM, CBRM, FLUXES, Clinker...).
    - Total (YTD) block on far right summarizing cargo totals across FY.
    - Two summary footer rows per month:
        - Per-cargo totals row (Yellow)
        - Monthly grand total row (Yellow, merged across month's cargo sub-columns).
    """
    months = _get_fy_months(reference_date)
    today = date.today()
    month_bounds = [(m_start, m_end, label) for m_start, m_end, label in months if m_start <= today]
    if not month_bounds:
        month_bounds = [(m_start, m_end, label) for m_start, m_end, label in months[:1]]

    from_date = month_bounds[0][0].strftime('%Y-%m-%d')
    to_date   = min(month_bounds[-1][1], today).strftime('%Y-%m-%d')

    conn = get_db()
    cur  = get_cursor(conn)

    sql = """
        SELECT
            l.entry_date::text AS entry_date,
            l.quantity,
            l.source_type,

            COALESCE(
                CASE WHEN l.source_type = 'VCN' THEN vcargo.cargo_type ELSE NULL END,
                CASE WHEN l.source_type = 'MBC' THEN COALESCE(m.cargo_type, m.cargo_name) ELSE NULL END,
                l.cargo_name,
                'Unknown'
            ) AS cargo_type,

            CASE
                WHEN l.source_type = 'VCN' THEN 'Mother Vessels Cargo'
                WHEN l.source_type = 'MBC' AND UPPER(COALESCE(mm.mbc_owner_name, '')) LIKE '%%SHIPPING%%'
                    THEN 'Shipping'
                WHEN l.source_type = 'MBC' AND UPPER(COALESCE(mm.mbc_owner_name, '')) LIKE '%%INFRA%%'
                    THEN 'Jsw Infrastructre'
                WHEN l.source_type = 'MBC' THEN
                    CASE
                        WHEN UPPER(TRIM(COALESCE(m.load_port, ''))) LIKE '%%JAIGAD%%' THEN 'Jaigad-Other MBC'
                        WHEN UPPER(TRIM(COALESCE(m.load_port, ''))) LIKE '%%SALAV%%'  THEN 'Salav -Other'
                        WHEN UPPER(TRIM(COALESCE(m.load_port, ''))) LIKE '%%GOA%%'    THEN 'Goa-Other'
                        WHEN NULLIF(TRIM(COALESCE(m.load_port, '')), '') IS NOT NULL
                            THEN CONCAT(TRIM(m.load_port), '-Other')
                        ELSE 'Other MBC'
                    END
                ELSE 'Double Handling'
            END AS row_category

        FROM lueu_lines l

        LEFT JOIN vcn_header v
            ON l.source_type = 'VCN' AND l.source_id = v.id
        LEFT JOIN LATERAL (
            SELECT cargo_type FROM vessel_cargo
            WHERE cargo_name = l.cargo_name
            LIMIT 1
        ) vcargo ON l.source_type = 'VCN'
        LEFT JOIN mbc_header m
            ON l.source_type = 'MBC' AND l.source_id = m.id
        LEFT JOIN mbc_master mm
            ON UPPER(TRIM(m.mbc_name)) = UPPER(TRIM(mm.mbc_name))

        WHERE l.is_deleted IS NOT TRUE
          AND l.quantity IS NOT NULL
          AND l.quantity > 0
          AND l.entry_date >= %s
          AND l.entry_date <= %s

        UNION ALL

        SELECT
            h.entry_date::text AS entry_date,
            h.quantity,
            'HISTORICAL' AS source_type,
            COALESCE(vcargo.cargo_type, h.cargo_name, 'Unknown') AS cargo_type,
            CASE
                WHEN UPPER(TRIM(COALESCE(h.mv_mbc, ''))) IN ('MV', 'VCN') THEN 'Mother Vessels Cargo'
                WHEN UPPER(TRIM(COALESCE(h.mv_mbc, ''))) = 'MBC' AND UPPER(COALESCE(mm.mbc_owner_name, '')) LIKE '%%SHIPPING%%' THEN 'Shipping'
                WHEN UPPER(TRIM(COALESCE(h.mv_mbc, ''))) = 'MBC' AND UPPER(COALESCE(mm.mbc_owner_name, '')) LIKE '%%INFRA%%' THEN 'Jsw Infrastructre'
                WHEN UPPER(TRIM(COALESCE(h.mv_mbc, ''))) = 'MBC' THEN 'Jaigad-Other MBC'
                ELSE 'Double Handling'
            END AS row_category
        FROM rp01_historical_lueu h
        LEFT JOIN LATERAL (
            SELECT cargo_type FROM vessel_cargo
            WHERE UPPER(TRIM(cargo_name)) = UPPER(TRIM(h.cargo_name))
            LIMIT 1
        ) vcargo ON true
        LEFT JOIN mbc_master mm
            ON UPPER(TRIM(h.source_display)) = UPPER(TRIM(mm.mbc_name))
        WHERE h.quantity IS NOT NULL AND h.quantity > 0
          AND h.entry_date::text >= %s AND h.entry_date::text <= %s
    """

    cur.execute(sql, [from_date, to_date, from_date, to_date])
    rows = cur.fetchall()
    conn.close()

    # Preferred row order matching template
    preferred_rows = [
        'Mother Vessels Cargo',
        'Jaigad-Other MBC',
        'Jsw Infrastructre',
        'Shipping',
        'Salav -Other',
        'Goa-Other',
    ]

    # data[row_category][ (year, month) ][cargo_type] = qty
    data = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    all_cargos = set()
    found_rows = set()

    for r in rows:
        raw_dt = r['entry_date']
        try:
            dt = datetime.strptime(str(raw_dt).strip()[:10], '%Y-%m-%d').date()
        except Exception:
            continue

        cargo = (r['cargo_type'] or 'Unknown').strip()
        cat   = r['row_category']
        qty   = _safe_float(r['quantity'])

        all_cargos.add(cargo)
        found_rows.add(cat)
        data[cat][(dt.year, dt.month)][cargo] += qty

    categories = [c for c in preferred_rows]
    for c in sorted(found_rows):
        if c not in categories:
            categories.append(c)

    return {
        'cargos':     sorted(all_cargos),
        'months':     month_bounds,
        'categories': categories,
        'data':       data,
    }


# ── Excel helpers ───────────────────────────────────────────────────────────

def _mbdr(ws, row, c1, c2, fill=XL_WHITE):
    """Apply perimeter thin borders + fill to every cell in a merged range."""
    for ci in range(c1, c2 + 1):
        b = Border(
            left   = _thin if ci == c1 else None,
            right  = _thin if ci == c2 else None,
            top    = _thin,
            bottom = _thin,
        )
        try:
            ws.cell(row, ci).border = b
            ws.cell(row, ci).fill   = _fill(fill)
        except AttributeError:
            pass


# ── Excel sheet writer ──────────────────────────────────────────────────────

def _write_mv_monthly_sheet(ws, report_data):

    vessels = report_data.get('vessels', [])
    dates   = report_data.get('dates', [])
    data    = report_data.get('data', {})

    if not vessels:
        ws['A1'] = "No data available"
        return

    # ----------------------------------------------------
    # COLUMN WIDTHS
    # ----------------------------------------------------

    ws.column_dimensions['A'].width = 18

    total_cols = len(vessels) + 4

    for c in range(2, total_cols + 1):
        ws.column_dimensions[get_column_letter(c)].width = 16

    # ----------------------------------------------------
    # HEADER ROWS
    # ----------------------------------------------------

    header_rows = [
        ("Cargo Name", "cargo_name"),
        ("Cargo Type", "cargo_type"),
        ("BL Qty",     "bl_qty"),
        ("Company",    "company"),
    ]

    current_row = 1

    for title, key in header_rows:

        cell = ws.cell(current_row, 1, title)
        cell.font      = _font(bold=True)
        cell.fill      = _fill(XL_GREY)
        cell.alignment = _ctr
        cell.border    = _bdr

        for idx, v in enumerate(vessels, start=2):

            value = v.get(key, '-')

            if key == 'cargo_name':
                value = str(value).replace('<br>', ' / ')

            c = ws.cell(current_row, idx, value)
            c.font      = _font(bold=False)
            c.alignment = _ctr
            c.border    = _bdr

        current_row += 1

    # ----------------------------------------------------
    # VESSEL HEADER ROW
    # ----------------------------------------------------

    c = ws.cell(current_row, 1, "Date")
    c.font      = _font(bold=True)
    c.fill      = _fill(XL_GREY)
    c.alignment = _ctr
    c.border    = _bdr

    for idx, v in enumerate(vessels, start=2):
        cell = ws.cell(current_row, idx, v['vessel_name'])
        cell.fill      = _fill(XL_YELLOW)
        cell.font      = _font(bold=True, color='C00000')
        cell.alignment = _ctr
        cell.border    = _bdr

    mv_col    = len(vessels) + 2
    mbc_col   = len(vessels) + 3
    grand_col = len(vessels) + 4

    for col, text, color in [
        (mv_col,    "MV Total",    "1E40AF"),
        (mbc_col,   "MBC Total",   "166534"),
        (grand_col, "Grand Total", "B45309"),
    ]:
        cell = ws.cell(current_row, col, text)
        cell.fill      = _fill(color)
        cell.font      = _font(bold=True, color='FFFFFF')
        cell.alignment = _ctr
        cell.border    = _bdr

    current_row += 1

    # ----------------------------------------------------
    # DATE ROWS
    # ----------------------------------------------------

    for dt in dates:

        row_total_mv  = 0
        row_total_mbc = 0
        row_total_all = 0

        # ✅ Convert string to date object so Excel formats it correctly
        try:
            dt_obj = datetime.strptime(dt, '%Y-%m-%d').date()
        except Exception:
            dt_obj = dt   # fallback to string if parse fails

        dcell = ws.cell(current_row, 1, dt_obj)   # ✅ pass date object, not string
        dcell.number_format = 'DD-MM-YYYY'         # ✅ formats as 01-05-2026
        dcell.font          = _font(bold=True)
        dcell.alignment     = _ctr
        dcell.border        = _bdr

        for idx, v in enumerate(vessels, start=2):
            qty  = data.get(dt, {}).get(v['vessel_name'], 0)
            cell = ws.cell(current_row, idx, qty if qty > 0 else "")
            cell.number_format = '#,##0.00'
            cell.alignment     = _ctr
            cell.border        = _bdr

            if v['source_type'] == 'VCN':
                row_total_mv += qty
            else:
                row_total_mbc += qty
            row_total_all += qty

        for col, val, clr in [
            (mv_col,    row_total_mv,  'DBEAFE'),
            (mbc_col,   row_total_mbc, 'DCFCE7'),
            (grand_col, row_total_all, 'FEF9C3'),
        ]:
            c = ws.cell(current_row, col, val)
            c.fill         = _fill(clr)
            c.font         = _font(bold=True)
            c.number_format = '#,##0.00'
            c.alignment    = _ctr
            c.border       = _bdr

        current_row += 1

    # ----------------------------------------------------
    # PREVIOUS MONTH QTY ROW
    # ----------------------------------------------------

    prev_row = current_row

    c = ws.cell(prev_row, 1, "Previous Month Qty")
    c.fill      = _fill('D9EAD3')
    c.font      = _font(bold=True)
    c.alignment = _ctr
    c.border    = _bdr

    mv_prev = mbc_prev = grand_prev = 0

    for idx, v in enumerate(vessels, start=2):
        qty  = v.get('previous_month_qty', 0)
        cell = ws.cell(prev_row, idx, qty)
        cell.fill          = _fill('D9EAD3')
        cell.font          = _font(bold=True)
        cell.number_format = '#,##0.00'
        cell.alignment     = _ctr
        cell.border        = _bdr

        if v['source_type'] == 'VCN':
            mv_prev += qty
        else:
            mbc_prev += qty
        grand_prev += qty

    for col, val in [
        (mv_col,    mv_prev),
        (mbc_col,   mbc_prev),
        (grand_col, grand_prev),
    ]:
        cc = ws.cell(prev_row, col, val)
        cc.fill          = _fill('B6D7A8')
        cc.font          = _font(bold=True)
        cc.number_format = '#,##0.00'
        cc.alignment     = _ctr
        cc.border        = _bdr

    current_row += 1

    # ----------------------------------------------------
    # TOTAL QTY ROW
    # ----------------------------------------------------

    total_row = current_row

    c = ws.cell(total_row, 1, "Total Qty")
    c.fill      = _fill(XL_YELLOW)
    c.font      = _font(bold=True)
    c.alignment = _ctr
    c.border    = _bdr

    mv_total = mbc_total = all_total = 0
    vessel_totals = {}

    for idx, v in enumerate(vessels, start=2):
        total = sum(data.get(dt, {}).get(v['vessel_name'], 0) for dt in dates)
        vessel_totals[v['vessel_name']] = total

        cell = ws.cell(total_row, idx, total)
        cell.fill          = _fill(XL_YELLOW)
        cell.font          = _font(bold=True)
        cell.number_format = '#,##0.00'
        cell.alignment     = _ctr
        cell.border        = _bdr

        if v['source_type'] == 'VCN':
            mv_total += total
        else:
            mbc_total += total
        all_total += total

    for col, val, clr in [
        (mv_col,    mv_total,  '1E40AF'),
        (mbc_col,   mbc_total, '166534'),
        (grand_col, all_total, 'B45309'),
    ]:
        cc = ws.cell(total_row, col, val)
        cc.fill          = _fill(clr)
        cc.font          = _font(bold=True, color='FFFFFF')
        cc.number_format = '#,##0.00'
        cc.alignment     = _ctr
        cc.border        = _bdr

    current_row += 1

    # ----------------------------------------------------
    # BL QTY ROW
    # ----------------------------------------------------

    bl_row = current_row

    c = ws.cell(bl_row, 1, "BL Qty")
    c.font      = _font(bold=True)
    c.alignment = _ctr
    c.border    = _bdr

    mv_bl = mbc_bl = all_bl = 0

    for idx, v in enumerate(vessels, start=2):
        qty  = v.get('bl_qty', 0)
        cell = ws.cell(bl_row, idx, qty)
        cell.number_format = '#,##0.00'
        cell.alignment     = _ctr
        cell.border        = _bdr

        if v['source_type'] == 'VCN':
            mv_bl += qty
        else:
            mbc_bl += qty
        all_bl += qty

    for col, val, clr in [
        (mv_col,    mv_bl,  'DBEAFE'),
        (mbc_col,   mbc_bl, 'DCFCE7'),
        (grand_col, all_bl, 'FEF9C3'),
    ]:
        cc = ws.cell(bl_row, col, val)
        cc.fill          = _fill(clr)
        cc.font          = _font(bold=True)
        cc.number_format = '#,##0.00'
        cc.alignment     = _ctr
        cc.border        = _bdr

    current_row += 1

    # ----------------------------------------------------
    # DIFFERENCE ROW
    # ----------------------------------------------------

    diff_row = current_row

    c = ws.cell(diff_row, 1, "Difference")
    c.font      = _font(bold=True)
    c.alignment = _ctr
    c.border    = _bdr

    for idx, v in enumerate(vessels, start=2):
        diff = (
            v.get('bl_qty', 0)
            - v.get('previous_month_qty', 0)
            - vessel_totals[v['vessel_name']]
        )
        cell = ws.cell(diff_row, idx, diff)
        cell.number_format = '#,##0.00'
        cell.alignment     = _ctr
        cell.border        = _bdr
        if diff < 0:
            cell.font = _font(bold=True, color='FF0000')   # Red
        elif diff > 0:
            cell.font = _font(bold=True, color='008000')   # Green
        else:
            cell.font = _font(bold=True)

    # ----------------------------------------------------
    # FREEZE PANES
    # ----------------------------------------------------

    ws.freeze_panes = 'B6'


# ── Cargo Summary Excel sheet ───────────────────────────────────────────────

def _write_cargo_summary_sheet(ws, report_data):
    """Write Cargo-wise summary sheet — ONE Total column only."""

    cargos     = report_data.get('cargos', [])
    operations = report_data.get('operations', [])
    data       = report_data.get('data', {})

    if not cargos:
        ws['A1'] = "No data available"
        return

    # ✅ Strip 'Total' from operations — we write it manually as the last column
    ops = [o for o in operations if o != 'Total']

    # Columns: Cargo | op1 | op2 | ... | Total
    NC = len(ops) + 2   # +1 for Cargo label col, +1 for Total col

    # Column widths
    ws.column_dimensions['A'].width = 22
    for i in range(2, NC + 1):
        ws.column_dimensions[get_column_letter(i)].width = 16

    # ----------------------------------------------------
    # Title row
    # ----------------------------------------------------
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NC)
    c = ws['A1']
    c.value     = 'Cargo Summary'
    c.font      = _font(bold=True, size=14)
    c.alignment = _ctr
    c.fill      = _fill(XL_WHITE)
    c.border    = _bdr

    # ----------------------------------------------------
    # Header row  (yellow)
    # ----------------------------------------------------
    ws.row_dimensions[2].height = 24
    headers = ['Cargo'] + ops + ['Total']   # ✅ exactly ONE Total

    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=c_idx, value=h)
        cell.fill      = _fill(XL_YELLOW)
        cell.font      = _font(bold=True, size=11)
        cell.alignment = _ctr
        cell.border    = _bdr

    # ----------------------------------------------------
    # Data rows
    # ----------------------------------------------------
    row_idx = 3

    for cargo in cargos:
        ws.row_dimensions[row_idx].height = 20

        # Cargo label
        cell = ws.cell(row=row_idx, column=1, value=cargo)
        cell.font      = _font(bold=True, size=10)
        cell.alignment = _left
        cell.border    = _bdr

        # Operation columns
        row_total = 0
        for c_idx, op in enumerate(ops, start=2):
            qty = data.get(cargo, {}).get(op, 0)
            row_total += qty
            cell = ws.cell(row=row_idx, column=c_idx, value=qty if qty > 0 else 0)
            cell.alignment     = _ctr
            cell.border        = _bdr
            cell.number_format = '#,##0.00'
            cell.font          = _font(size=10)

        # ✅ Single Total column at NC
        cell = ws.cell(row=row_idx, column=NC, value=row_total)
        cell.alignment     = _ctr
        cell.border        = _bdr
        cell.number_format = '#,##0.00'
        cell.font          = _font(bold=True, size=10)

        row_idx += 1

    # ----------------------------------------------------
    # Grand Total row  (yellow)
    # ----------------------------------------------------
    ws.row_dimensions[row_idx].height = 24

    cell = ws.cell(row=row_idx, column=1, value='Grand Total')
    cell.fill      = _fill(XL_YELLOW)
    cell.font      = _font(bold=True, size=10)
    cell.alignment = _left
    cell.border    = _bdr

    grand_total = 0
    for c_idx, op in enumerate(ops, start=2):
        col_total = sum(data.get(cargo, {}).get(op, 0) for cargo in cargos)
        grand_total += col_total
        cell = ws.cell(row=row_idx, column=c_idx, value=col_total)
        cell.fill          = _fill(XL_YELLOW)
        cell.font          = _font(bold=True, size=10)
        cell.alignment     = _ctr
        cell.border        = _bdr
        cell.number_format = '#,##0.00'

    # ✅ Grand total at NC
    cell = ws.cell(row=row_idx, column=NC, value=grand_total)
    cell.fill          = _fill(XL_YELLOW)
    cell.font          = _font(bold=True, size=10)
    cell.alignment     = _ctr
    cell.border        = _bdr
    cell.number_format = '#,##0.00'

    # Freeze: keep Cargo col + title/header rows fixed
    ws.freeze_panes = 'B3'


# ── Excel writer: MBC Handling Status ───────────────────────────────────────

def _write_mbc_status_sheet(ws, status_data):

    port_rows    = status_data.get('port_rows', [])
    company_rows = status_data.get('company_rows', [])
    ports        = status_data.get('ports', [])

    if not port_rows:
        ws['A1'] = "No data available"
        return

    headers_left  = ['Sr. No', 'Month'] + ports + ['Total']
    headers_right = ['Month', 'JSW Shipping', 'JSW Infra', 'Other', 'Barges', 'Total', 'Grand Total']

    right_start_col = len(headers_left) + 2  # one blank spacer column
    total_cols = len(headers_left) + 1 + len(headers_right)

    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 12
    for i in range(3, total_cols + 1):
        ws.column_dimensions[get_column_letter(i)].width = 14

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    c = ws['A1']
    c.value = 'MBC Handling Status'
    c.font = _font(bold=True, size=XL_TITLE_SZ)
    c.alignment = _ctr
    c.fill = _fill(XL_WHITE)
    c.border = _bdr

    headers_left  = ['Sr. No', 'Month'] + ports + ['Total']
    headers_right = ['Month', 'JSW Shipping', 'JSW Infra', 'Other', 'Barges', 'Total', 'Grand Total']

    right_start_col = len(headers_left) + 2  # one blank spacer column

    for idx, h in enumerate(headers_left, start=1):
        cell = ws.cell(2, idx, h)
        cell.font = _font(bold=True)
        cell.fill = _fill(XL_GREY)
        cell.alignment = _ctr
        cell.border = _bdr

    for idx, h in enumerate(headers_right):
        cell = ws.cell(2, right_start_col + idx, h)
        cell.font = _font(bold=True)
        cell.fill = _fill(XL_GREY)
        cell.alignment = _ctr
        cell.border = _bdr

    row = 3
    port_col_totals = defaultdict(int)
    grand_port_total = 0

    for i, pr in enumerate(port_rows):
        ws.cell(row, 1, i + 1).alignment = _ctr
        ws.cell(row, 1).border = _bdr

        dcell = ws.cell(row, 2, pr['month_start'])
        dcell.number_format = 'MMM-YY'
        dcell.alignment = _ctr
        dcell.border = _bdr

        for j, port in enumerate(ports):
            val = pr['ports'].get(port, 0)
            port_col_totals[port] += val
            cell = ws.cell(row, 3 + j, val if val else "")
            cell.alignment = _ctr
            cell.border = _bdr

        tcell = ws.cell(row, 3 + len(ports), pr['total'])
        tcell.font = _font(bold=True)
        tcell.fill = _fill(XL_YELLOW)
        tcell.alignment = _ctr
        tcell.border = _bdr
        grand_port_total += pr['total']

        row += 1

    # column totals footer for the port block
    ws.cell(row, 2, 'Total').font = _font(bold=True)
    ws.cell(row, 2).border = _bdr
    for j, port in enumerate(ports):
        cell = ws.cell(row, 3 + j, port_col_totals[port])
        cell.font = _font(bold=True)
        cell.fill = _fill(XL_GREY)
        cell.alignment = _ctr
        cell.border = _bdr
    gcell = ws.cell(row, 3 + len(ports), grand_port_total)
    gcell.font = _font(bold=True)
    gcell.fill = _fill(XL_YELLOW)
    gcell.alignment = _ctr
    gcell.border = _bdr

    # company block (independent row loop, same row numbers as port block)
    row = 3
    company_totals = defaultdict(int)

    for cr in company_rows:
        dcell = ws.cell(row, right_start_col, cr['month_start'])
        dcell.number_format = 'MMM-YY'
        dcell.alignment = _ctr
        dcell.border = _bdr

        vals = [
            ('jsw_shipping', 'DBEAFE'),
            ('jsw_infra',    'DCFCE7'),
            ('other',        'FEF9C3'),
            ('barges',       'F3E8FF'),
        ]
        for k, (key, clr) in enumerate(vals):
            v = cr[key]
            company_totals[key] += v
            cell = ws.cell(row, right_start_col + 1 + k, v if v else "")
            cell.alignment = _ctr
            cell.border = _bdr

        tcell = ws.cell(row, right_start_col + 5, cr['total'])
        tcell.font = _font(bold=True)
        tcell.alignment = _ctr
        tcell.border = _bdr

        gcell = ws.cell(row, right_start_col + 6, cr['grand_total'])
        gcell.font = _font(bold=True)
        gcell.fill = _fill(XL_YELLOW)
        gcell.alignment = _ctr
        gcell.border = _bdr

        company_totals['total'] = company_totals.get('total', 0) + cr['total']
        company_totals['grand_total'] = company_totals.get('grand_total', 0) + cr['grand_total']

        row += 1

    ws.cell(row, right_start_col, 'Total').font = _font(bold=True)
    ws.cell(row, right_start_col).border = _bdr
    for k, key in enumerate(['jsw_shipping', 'jsw_infra', 'other', 'barges']):
        cell = ws.cell(row, right_start_col + 1 + k, company_totals[key])
        cell.font = _font(bold=True)
        cell.fill = _fill(XL_GREY)
        cell.alignment = _ctr
        cell.border = _bdr
    cell = ws.cell(row, right_start_col + 5, company_totals.get('total', 0))
    cell.font = _font(bold=True)
    cell.alignment = _ctr
    cell.border = _bdr
    cell = ws.cell(row, right_start_col + 6, company_totals.get('grand_total', 0))
    cell.font = _font(bold=True)
    cell.fill = _fill(XL_YELLOW)
    cell.alignment = _ctr
    cell.border = _bdr

    ws.freeze_panes = 'C3'


# ── Excel writer: YTD Cargo Pivot ───────────────────────────────────────────

def _write_ytd_pivot_sheet(ws, ytd_data):
    cargos     = ytd_data.get('cargos', [])
    months     = ytd_data.get('months', [])
    categories = ytd_data.get('categories', [])
    data       = ytd_data.get('data', {})

    if not cargos:
        ws['A1'] = "No data available"
        return

    n_cargos = len(cargos)

    # Column dimensions
    ws.column_dimensions['A'].width = 24
    total_cols = 1 + len(months) * n_cargos + n_cargos + 1

    for col_idx in range(2, total_cols + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 14

    # ── Row 1: Month Headers (Bright Yellow) ──────────────────────────────────
    ws.cell(1, 1, "").border = _bdr

    col = 2
    for m_start, m_end, label in months:
        end_col = col + n_cargos - 1
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=end_col)
        mcell = ws.cell(1, col, label)
        mcell.font = _font(bold=True)
        mcell.fill = _fill(XL_YELLOW)
        mcell.alignment = _ctr

        for c_i in range(col, end_col + 1):
            ws.cell(1, c_i).border = _bdr
            ws.cell(1, c_i).fill   = _fill(XL_YELLOW)

        col += n_cargos

    # Total (YTD) merged header on far right
    ytd_start_col = col
    ytd_end_col   = col + n_cargos  # includes n_cargos + 1 final total column
    ws.merge_cells(start_row=1, start_column=ytd_start_col, end_row=1, end_column=ytd_end_col)
    ycell = ws.cell(1, ytd_start_col, 'Total (YTD)')
    ycell.font = _font(bold=True)
    ycell.fill = _fill('FFF2CC')
    ycell.alignment = _ctr
    for c_i in range(ytd_start_col, ytd_end_col + 1):
        ws.cell(1, c_i).border = _bdr
        ws.cell(1, c_i).fill   = _fill('FFF2CC')

    # ── Row 2: Cargo Sub-headers (Bright Yellow) ─────────────────────────────
    ws.cell(2, 1, "").border = _bdr

    col = 2
    for m_start, m_end, label in months:
        for cargo in cargos:
            cell = ws.cell(2, col, cargo)
            cell.font = _font(bold=True, size=XL_SMALL_SZ)
            cell.fill = _fill(XL_YELLOW)
            cell.alignment = _ctr
            cell.border = _bdr
            col += 1

    # YTD block cargo sub-headers
    for cargo in cargos:
        cell = ws.cell(2, col, cargo)
        cell.font = _font(bold=True, size=XL_SMALL_SZ)
        cell.fill = _fill('FFF2CC')
        cell.alignment = _ctr
        cell.border = _bdr
        col += 1

    # Final Total sub-header column
    cell = ws.cell(2, col, "")
    cell.fill = _fill('FFF2CC')
    cell.border = _bdr

    # ── Rows 3+: Data Rows (Category x Cargo) ────────────────────────────────
    row = 3
    month_cargo_totals = defaultdict(float)
    ytd_cargo_totals   = defaultdict(float)
    ytd_grand_total    = 0.0

    for cat in categories:
        ws.cell(row, 1, cat).font = _font(bold=False)
        ws.cell(row, 1).alignment = _left
        ws.cell(row, 1).border = _bdr

        col = 2
        cat_ytd_row_total  = 0.0
        cat_ytd_cargo_sums = defaultdict(float)

        for m_idx, (m_start, m_end, label) in enumerate(months):
            m_key  = (m_start.year, m_start.month)
            m_data = data.get(cat, {}).get(m_key, {})

            for cargo in cargos:
                v = _safe_float(m_data.get(cargo, 0))
                cell = ws.cell(row, col, v if v > 0 else "")
                cell.number_format = '#,##0'
                cell.alignment = _right
                cell.border = _bdr

                month_cargo_totals[(m_idx, cargo)] += v
                cat_ytd_cargo_sums[cargo] += v
                cat_ytd_row_total += v

                col += 1

        # Write YTD block for this category
        for cargo in cargos:
            v = cat_ytd_cargo_sums[cargo]
            cell = ws.cell(row, col, v if v > 0 else "")
            cell.number_format = '#,##0'
            cell.alignment = _right
            cell.border = _bdr
            cell.fill = _fill('FFF2CC')
            ytd_cargo_totals[cargo] += v
            col += 1

        # Final category row total cell in YTD block
        tcell = ws.cell(row, col, cat_ytd_row_total if cat_ytd_row_total > 0 else "")
        tcell.font = _font(bold=False)
        tcell.number_format = '#,##0'
        tcell.alignment = _right
        tcell.border = _bdr
        tcell.fill = _fill('FFF2CC')

        ytd_grand_total += cat_ytd_row_total
        row += 1

    # ── Summary Footer Row 1: Per-cargo Totals Row ───────────────────────────
    ws.cell(row, 1, "").border = _bdr

    col = 2
    month_totals_by_m_idx = defaultdict(float)

    for m_idx, (m_start, m_end, label) in enumerate(months):
        m_tot = 0.0
        for cargo in cargos:
            v = month_cargo_totals[(m_idx, cargo)]
            cell = ws.cell(row, col, v if v > 0 else 0)
            cell.font = _font(bold=True)
            cell.fill = _fill(XL_YELLOW)
            cell.number_format = '#,##0'
            cell.alignment = _right
            cell.border = _bdr
            m_tot += v
            col += 1
        month_totals_by_m_idx[m_idx] = m_tot

    # YTD per-cargo summary cells
    for cargo in cargos:
        v = ytd_cargo_totals[cargo]
        cell = ws.cell(row, col, v if v > 0 else 0)
        cell.font = _font(bold=True)
        cell.fill = _fill('FFF2CC')
        cell.number_format = '#,##0'
        cell.alignment = _right
        cell.border = _bdr
        col += 1

    # Final YTD grand total cell in per-cargo totals row
    gcell = ws.cell(row, col, ytd_grand_total)
    gcell.font = _font(bold=True)
    gcell.fill = _fill('FFF2CC')
    gcell.number_format = '#,##0'
    gcell.alignment = _right
    gcell.border = _bdr

    row += 1

    # ── Summary Footer Row 2: Merged Monthly Grand Total Row ─────────────────
    ws.cell(row, 1, "").border = _bdr

    col = 2
    for m_idx, (m_start, m_end, label) in enumerate(months):
        end_col = col + n_cargos - 1
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=end_col)
        m_total_val = month_totals_by_m_idx[m_idx]

        mcell = ws.cell(row, col, m_total_val)
        mcell.font = _font(bold=True)
        mcell.fill = _fill(XL_YELLOW)
        mcell.number_format = '#,##0'
        mcell.alignment = _ctr

        for c_i in range(col, end_col + 1):
            ws.cell(row, c_i).border = _bdr
            ws.cell(row, c_i).fill   = _fill(XL_YELLOW)

        col += n_cargos

    # YTD block merged grand total row
    ws.merge_cells(start_row=row, start_column=ytd_start_col, end_row=row, end_column=ytd_end_col)
    ygtcell = ws.cell(row, ytd_start_col, ytd_grand_total)
    ygtcell.font = _font(bold=True)
    ygtcell.fill = _fill(XL_YELLOW)
    ygtcell.number_format = '#,##0'
    ygtcell.alignment = _ctr

    for c_i in range(ytd_start_col, ytd_end_col + 1):
        ws.cell(row, c_i).border = _bdr
        ws.cell(row, c_i).fill   = _fill(XL_YELLOW)

    ws.freeze_panes = 'B3'


# ── Excel writer: Yearly Ledger ─────────────────────────────────────────────

def _write_yearly_ledger_sheet(ws, calls):

    if not calls:
        ws['A1'] = "No data available"
        return

    mbc_calls = [c for c in calls if c.get('source_type') == 'MBC']
    vcn_calls = [c for c in calls if c.get('source_type') == 'VCN']

    headers = [
        'Sr No', 'Vessel / MBC', 'Type', 'Load Port', 'Company',
        'Cargo Type', 'Cargo Name', 'BL Qty', 'Discharged Qty',
        'Discharge Commenced', 'Discharge Completed',
    ]

    widths = [8, 32, 8, 16, 16, 16, 22, 14, 14, 18, 18]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    def _sort_key(c):
        d = c.get('bucket_date')
        return d or date.min

    sorted_mbc = sorted(mbc_calls, key=_sort_key, reverse=True)
    sorted_vcn = sorted(vcn_calls, key=_sort_key, reverse=True)

    row = 1

    # ── Table 1: MBC (Coastal) Calls Table First ────────────────────────────
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(headers))
    title_cell = ws.cell(row, 1, 'MBC (Coastal) Calls')
    title_cell.font = _font(bold=True, size=12)
    title_cell.fill = _fill(XL_GREY)
    title_cell.alignment = _ctr
    title_cell.border = _bdr
    row += 1

    for idx, h in enumerate(headers, start=1):
        cell = ws.cell(row, idx, h)
        cell.font = _font(bold=True)
        cell.fill = _fill(XL_GREY)
        cell.alignment = _ctr
        cell.border = _bdr
    row += 1

    mbc_bl_total = 0.0
    mbc_disc_total = 0.0

    for i, c in enumerate(sorted_mbc, start=1):
        bl = _safe_float(c.get('bl_qty'))
        disc = _safe_float(c.get('discharged_qty'))
        mbc_bl_total += bl
        mbc_disc_total += disc

        vals = [
            i,
            c.get('header_name') or '-',
            c.get('source_type'),
            c.get('load_port') or '-',
            c.get('company') or '-',
            c.get('cargo_type') or '-',
            c.get('cargo_name') or '-',
            bl,
            disc,
            c.get('discharge_commenced') or '-',
            c.get('discharge_completed') or '-',
        ]
        for idx, v in enumerate(vals, start=1):
            cell = ws.cell(row, idx, v)
            cell.alignment = _ctr
            cell.border = _bdr
            if idx in (8, 9):
                cell.number_format = '#,##0.00'
        row += 1

    # Total row for MBC table
    ws.cell(row, 2, 'Total MBC Calls').font = _font(bold=True)
    ws.cell(row, 2).border = _bdr
    for idx in range(1, len(headers) + 1):
        ws.cell(row, idx).border = _bdr
        ws.cell(row, idx).fill = _fill(XL_YELLOW)

    cell_bl = ws.cell(row, 8, mbc_bl_total)
    cell_bl.font = _font(bold=True)
    cell_bl.number_format = '#,##0.00'
    cell_bl.alignment = _ctr

    cell_disc = ws.cell(row, 9, mbc_disc_total)
    cell_disc.font = _font(bold=True)
    cell_disc.number_format = '#,##0.00'
    cell_disc.alignment = _ctr

    # ── Line space between tables ───────────────────────────────────────────
    row += 3

    # ── Table 2: Mother Vessel (VCN) Calls Table Second ──────────────────────
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(headers))
    title_cell = ws.cell(row, 1, 'Mother Vessel (VCN) Calls')
    title_cell.font = _font(bold=True, size=12)
    title_cell.fill = _fill(XL_GREY)
    title_cell.alignment = _ctr
    title_cell.border = _bdr
    row += 1

    for idx, h in enumerate(headers, start=1):
        cell = ws.cell(row, idx, h)
        cell.font = _font(bold=True)
        cell.fill = _fill(XL_GREY)
        cell.alignment = _ctr
        cell.border = _bdr
    row += 1

    vcn_bl_total = 0.0
    vcn_disc_total = 0.0

    for i, c in enumerate(sorted_vcn, start=1):
        bl = _safe_float(c.get('bl_qty'))
        disc = _safe_float(c.get('discharged_qty'))
        vcn_bl_total += bl
        vcn_disc_total += disc

        vals = [
            i,
            c.get('header_name') or '-',
            c.get('source_type'),
            c.get('load_port') or '-',
            c.get('company') or '-',
            c.get('cargo_type') or '-',
            c.get('cargo_name') or '-',
            bl,
            disc,
            c.get('discharge_commenced') or '-',
            c.get('discharge_completed') or '-',
        ]
        for idx, v in enumerate(vals, start=1):
            cell = ws.cell(row, idx, v)
            cell.alignment = _ctr
            cell.border = _bdr
            if idx in (8, 9):
                cell.number_format = '#,##0.00'
        row += 1

    # Total row for VCN table
    ws.cell(row, 2, 'Total Vessel Calls').font = _font(bold=True)
    ws.cell(row, 2).border = _bdr
    for idx in range(1, len(headers) + 1):
        ws.cell(row, idx).border = _bdr
        ws.cell(row, idx).fill = _fill(XL_YELLOW)

    cell_bl = ws.cell(row, 8, vcn_bl_total)
    cell_bl.font = _font(bold=True)
    cell_bl.number_format = '#,##0.00'
    cell_bl.alignment = _ctr

    cell_disc = ws.cell(row, 9, vcn_disc_total)
    cell_disc.font = _font(bold=True)
    cell_disc.number_format = '#,##0.00'
    cell_disc.alignment = _ctr

    ws.freeze_panes = 'A3'


# ── Excel builder ───────────────────────────────────────────────────────────

# REPORT_CUTOFF_DATE: used for historical reporting/filtering
def _build_mv_monthly_excel(report_data, cargo_data=None):
    from openpyxl import Workbook
    wb = Workbook()

    # Sheet 1: Vessel Matrix
    ws1 = wb.active
    ws1.title = 'Vessel Report'
    _write_mv_monthly_sheet(ws1, report_data)

    # Sheet 2: Cargo Summary
    if cargo_data:
        ws2 = wb.create_sheet('Cargo Summary')
        _write_cargo_summary_sheet(ws2, cargo_data)

    # Sheet 3: MBC Handling Status
    ws3 = wb.create_sheet('MBC')
    mbc_status = _fetch_mbc_handling_status()
    _write_mbc_status_sheet(ws3, mbc_status)

    # Sheet 4: YTD Cargo Summary
    ws4 = wb.create_sheet('YTD')
    ytd_data = _fetch_ytd_cargo_pivot()
    _write_ytd_pivot_sheet(ws4, ytd_data)

    # Sheet 5: Yearly Ledger
    ws5 = wb.create_sheet('Yearly')
    calls = _fetch_vessel_call_ledger()
    _write_yearly_ledger_sheet(ws5, calls)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── Routes ───────────────────────────────────────────────────────────────────

@bp.route('/module/RP01/mv-monthly-report/')
@login_required
def mv_monthly_report_index():
    return render_template(
        'mv_mbc_month_report/mv_mbc_report.html',
        username=session.get('username')
    )


@bp.route('/api/module/RP01/mv-monthly-report/data')
@login_required
def mv_monthly_report_data():
    from_date = request.args.get('from_date')
    to_date   = request.args.get('to_date')
    op_type   = request.args.get('operation_type')
    report    = _fetch_mv_monthly_data(from_date, to_date, op_type)
    return jsonify(report)


@bp.route('/api/module/RP01/mv-monthly-report/download')
@login_required
def mv_monthly_report_download():
    from_date = request.args.get('from_date')
    to_date   = request.args.get('to_date')
    op_type   = request.args.get('operation_type')

    vessel_report = _fetch_mv_monthly_data(from_date, to_date, op_type)
    cargo_report  = _fetch_cargo_summary(from_date, to_date)

    buf   = _build_mv_monthly_excel(vessel_report, cargo_report)
    fname = f'MV_Monthly_Loading_Report_{date.today().strftime("%Y%m%d")}.xlsx'
    return Response(
        buf.getvalue(),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': f'attachment; filename="{fname}"'},
    )


@bp.route('/api/module/RP01/mv-monthly-report/cargo-summary')
@login_required
def mv_monthly_cargo_summary():
    from_date = request.args.get('from_date')
    to_date   = request.args.get('to_date')
    report    = _fetch_cargo_summary(from_date, to_date)
    return jsonify(report)