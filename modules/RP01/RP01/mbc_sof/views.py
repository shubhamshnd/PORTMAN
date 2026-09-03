from flask import render_template, session, redirect, url_for
from functools import wraps
from datetime import datetime
from database import get_db, get_cursor
from .. import bp


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


# ---------------------------------------------------------------------------
# Timestamp helpers
# ---------------------------------------------------------------------------

def _parse(ts):
    if not ts:
        return None
    try:
        return datetime.fromisoformat(str(ts))
    except Exception:
        return None


def fmt_dt(ts):
    """'2026-02-08 14:40:00'  →  'ON 08.02.2026 AT 1440 HRS.'"""
    dt = _parse(ts)
    if not dt:
        return ''
    return f"ON {dt.strftime('%d.%m.%Y')} AT {dt.strftime('%H%M')} HRS."


def fmt_dt_display(ts):
    """Format timestamp for display in activity table: '20-04-2026 14:20'"""
    dt = _parse(ts)
    if not dt:
        return ''
    return dt.strftime('%d-%m-%Y %H:%M')


def fmt_date_display(ts):
    """Format date only: '20-04-2026'"""
    dt = _parse(ts)
    if not dt:
        return ''
    return dt.strftime('%d-%m-%Y')


def fmt_qty(value):
    if value is None or value == '':
        return ''
    try:
        text = f'{float(value):,.3f}'
    except (TypeError, ValueError):
        return str(value)
    return text[:-4] if text.endswith('.000') else text


# ---------------------------------------------------------------------------
# DB helpers
# ---------------------------------------------------------------------------

def _fetch_mbc_list():
    conn = get_db()
    cur  = get_cursor(conn)
    cur.execute("""
        SELECT id, doc_num, mbc_name, operation_type, cargo_name,
               bl_quantity, quantity_uom, doc_status, doc_date
        FROM mbc_header
        ORDER BY operation_type, doc_date DESC NULLS LAST
    """)
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    for r in rows:
        d = r.get('doc_date')
        if d:
            try:
                if hasattr(d, 'strftime'):
                    r['doc_date_display'] = d.strftime('%d.%m.%Y')
                else:
                    dt = datetime.fromisoformat(str(d)[:10])
                    r['doc_date_display'] = dt.strftime('%d.%m.%Y')
            except Exception:
                r['doc_date_display'] = str(d)
        else:
            r['doc_date_display'] = '—'
        r['bl_quantity_display'] = fmt_qty(r.get('bl_quantity')) if r.get('bl_quantity') else ''
    return rows


def _fetch_mbc_sof_data(mbc_id):
    """Fetch header + relevant lines for one MBC."""
    conn = get_db()
    cur  = get_cursor(conn)

    cur.execute("SELECT * FROM mbc_header WHERE id = %s", (mbc_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return None, {}, {}, {}
    header  = dict(row)
    op_type = (header.get('operation_type') or '').lower()

    load_port      = {}
    discharge_port = {}
    export_load    = {}

    if op_type == 'export':
        cur.execute(
            "SELECT * FROM mbc_export_load_port_lines WHERE mbc_id = %s",
            (mbc_id,)
        )
        r = cur.fetchone()
        if r:
            export_load = dict(r)
    else:
        cur.execute(
            "SELECT * FROM mbc_load_port_lines WHERE mbc_id = %s",
            (mbc_id,)
        )
        r = cur.fetchone()
        if r:
            load_port = dict(r)

        cur.execute(
            "SELECT * FROM mbc_discharge_port_lines WHERE mbc_id = %s",
            (mbc_id,)
        )
        r = cur.fetchone()
        if r:
            discharge_port = dict(r)

    conn.close()
    return header, load_port, discharge_port, export_load


# ---------------------------------------------------------------------------
# Activity row builders — matching Excel SOF format
# ---------------------------------------------------------------------------

# def _build_import_activities(load_port, discharge_port):
#     activities = []

#     lp_source = 'Captured from MBC Operation - Load Port Details'

#     load_port_activities = [
#         ('Arrived Load Port',  'arrived_load_port'),
#         ('Alongside Berth',    'alongside_berth'),
#         ('Loading Commenced',  'loading_commenced'),
#         ('Loading Completed',  'loading_completed'),
#         ('Cast Off Load Port', 'cast_off_load_port'),
#         ('ETA at Gull Island', 'eta'),
#     ]

#     for label, key in load_port_activities:
#         val = fmt_dt_display(load_port.get(key))

#         activities.append({
#             'label': label,
#             'datetime': val if val else '',
#             'remark': '',
#             'source': lp_source
#         })

#     dp_source = 'Captured from MBC Operation - Discharge Port Details'

#     discharge_activities = [
#         ('Arrival at Gull Island',                  'arrival_gull_island'),
#         ('Departure from Gull Island',              'departure_gull_island'),
#         ('Arrived Yellow Crane',                    'arrived_yellow_crane'),
#         ('Vessel Arrival at Port',                  'vessel_arrival_port'),
#         ('Vessel All Made Fast at Unloading Berth', 'vessel_all_made_fast'),
#         ('Unloading Commenced',                     'unloading_commenced'),
#         ('Cleaning Commenced',                      'cleaning_commenced'),
#         ('Cleaning Completed',                      'unloading_completed'),
#         ('Unloading Completed',                     'unloading_completed'),
#         ('Vessel Cast Off from Dharamtar Jetty',    'vessel_cast_off'),
#         ('Sailed Out From Discharge Port',               'sailed_out_load_port'),
#     ]

#     for label, key in discharge_activities:
#         val = fmt_dt_display(discharge_port.get(key))

#         activities.append({
#             'label': label,
#             'datetime': val if val else '',
#             'remark': '',
#             'source': dp_source
#         })

#     # ADD THESE ONLY ONCE AT THE END
#     activities.append({
#         'label': 'Vessel Unloaded By',
#         'datetime': discharge_port.get('vessel_unloaded_by', '') or '',
#         'remark': '',
#         'source': dp_source,
#     })

#     activities.append({
#         'label': 'Vessel Unloading Berth',
#         'datetime': discharge_port.get('vessel_unloading_berth', '') or '',
#         'remark': '',
#         'source': dp_source,
#     })

#     return activities
def _build_import_activities(load_port, discharge_port):
    activities = []

    lp_source = 'Captured from MBC Operation - Load Port Details'

    load_port_activities = [
        ('Arrived Load Port',  'arrived_load_port'),
        ('Alongside Berth',    'alongside_berth'),
        ('Loading Commenced',  'loading_commenced'),
        ('Loading Completed',  'loading_completed'),
        ('Cast Off Load Port', 'cast_off_load_port'),
        ('ETA at Gull Island', 'eta'),
    ]

    for label, key in load_port_activities:
        val = fmt_dt_display(load_port.get(key))

        activities.append({
            'label': label,
            'datetime': val if val else '-',
            'remark': '',
            'source': lp_source
        })

    dp_source = 'Captured from MBC Operation - Discharge Port Details'

    discharge_activities = [
        ('Arrival at Gull Island',                  'arrival_gull_island'),
        ('Departure from Gull Island',              'departure_gull_island'),
        ('Arrived Yellow Crane',                    'arrived_yellow_crane'),
        ('Vessel Arrival at Port',                  'vessel_arrival_port'),
        ('Vessel All Made Fast at Unloading Berth', 'vessel_all_made_fast'),
        ('Unloading Commenced',                     'unloading_commenced'),
        ('Cleaning Commenced',                      'cleaning_commenced'),
        ('Cleaning Completed',                      'unloading_completed'),
        ('Unloading Completed',                     'unloading_completed'),
        ('Vessel Cast Off from Dharamtar Jetty',    'vessel_cast_off'),
        ('Sailed Out From Discharge Port',          'sailed_out_load_port'),
    ]

    for label, key in discharge_activities:
        val = fmt_dt_display(discharge_port.get(key))

        activities.append({
            'label': label,
            'datetime': val if val else '-',
            'remark': '',
            'source': dp_source
        })

    # Always show these rows.
    # If value is blank, display "-"
    activities.append({
        'label': 'Vessel Unloaded By',
        'datetime': discharge_port.get('vessel_unloaded_by', '') or '-',
        'remark': '',
        'source': dp_source,
    })

    activities.append({
        'label': 'Vessel Unloading Berth',
        'datetime': discharge_port.get('vessel_unloading_berth', '') or '-',
        'remark': '',
        'source': dp_source,
    })

    return activities
# def _build_export_activities(export_load):
#     source = 'Captured from MBC Operation - Export Load Port Details'
#     activities = []
#     for label, key in [
#         ('Arrived at Port',      'arrived_at_port'),
#         ('Alongside at Berth',   'alongside_at_berth'),
#         ('Loading Commenced',    'loading_commenced'),
#         ('Loading Completed',    'loading_completed'),
#         ('Cast Off From Berth',  'cast_off_from_berth'),
#         ('Sailed Out From Port', 'sailed_out_from_port'),
#         ('ETA at Gull Island',   'eta_at_gull_island'),
#     ]:
#         val = fmt_dt_display(export_load.get(key))
#         if val:
#             activities.append({'label': label, 'datetime': val, 'remark': '', 'source': source})

#     for label, key in [
#         ('Vessel Unloaded By', 'unloaded_by'),
#         ('Berth Master',       'berth_master'),
#     ]:
#         v = export_load.get(key) or ''
#         if v:
#             activities.append({'label': label, 'datetime': v, 'remark': '', 'source': source})

#     return activities

def _build_export_activities(export_load):
    source = 'Captured from MBC Operation - Export Load Port Details'
    activities = []

    # Always show these activities.
    # If the timestamp is blank, display "-"
    for label, key in [
        ('Arrived at Port',      'arrived_at_port'),
        ('Alongside at Berth',   'alongside_at_berth'),
        ('Loading Commenced',    'loading_commenced'),
        ('Loading Completed',    'loading_completed'),
        ('Cast Off From Berth',  'cast_off_from_berth'),
        ('Sailed Out From Port', 'sailed_out_from_port'),
        ('ETA at Gull Island',   'eta_at_gull_island'),
    ]:
        val = fmt_dt_display(export_load.get(key))

        activities.append({
            'label': label,
            'datetime': val if val else '-',
            'remark': '',
            'source': source
        })

    # These fields are also shown only when available
    for label, key in [
        ('Vessel Unloaded By', 'unloaded_by'),
        ('Berth Master',       'berth_master'),
    ]:
        v = export_load.get(key) or ''

        if v:
            activities.append({
                'label': label,
                'datetime': v,
                'remark': '',
                'source': source
            })

    return activities


# ---------------------------------------------------------------------------
# Delay builder with full debug logging
# ---------------------------------------------------------------------------

def _parse_time_only(t):
    """Parse 'HH:MM' or 'HH:MM:SS' time-only strings into total seconds from midnight."""
    if not t:
        return None
    try:
        parts = str(t).strip().split(':')
        hours   = int(parts[0])
        minutes = int(parts[1]) if len(parts) > 1 else 0
        seconds = int(parts[2]) if len(parts) > 2 else 0
        return hours * 3600 + minutes * 60 + seconds
    except Exception:
        return None


def _build_delays(mbc_id):
    conn = get_db()
    cur  = get_cursor(conn)

    query = """
        SELECT
            delay_name,
            from_time,
            to_time
        FROM lueu_lines
        WHERE source_id = %s
          AND delay_name IS NOT NULL
    """

    cur.execute(query, (mbc_id,))
    rows = cur.fetchall()
    conn.close()

    print(f"[DEBUG] _build_delays called — mbc_id={mbc_id}")
    print(f"[DEBUG] Total rows fetched from lueu_lines: {len(rows)}")
    for row in rows:
        print(f"[DEBUG]   delay_name={row['delay_name']!r}  "
              f"from_time={row['from_time']!r}  "
              f"to_time={row['to_time']!r}")

    maintenance_hours = 0.0
    rhms_hours        = 0.0
    master_hours      = 0.0
    weather_hours     = 0.0

    for row in rows:
        delay_name = (row['delay_name'] or '').strip().lower()

        from_secs = _parse_time_only(row.get('from_time'))
        to_secs   = _parse_time_only(row.get('to_time'))

        if from_secs is None or to_secs is None:
            print(f"[DEBUG]   SKIPPED (could not parse times): "
                  f"from={row.get('from_time')!r}  to={row.get('to_time')!r}")
            continue

        # Handle overnight crossing (e.g. from=23:00, to=01:00)
        if to_secs < from_secs:
            to_secs += 24 * 3600

        duration = (to_secs - from_secs) / 3600  # convert seconds → hours
        print(f"[DEBUG]   delay={delay_name!r}  duration={duration:.4f} hrs")

        if 'breakdown' in delay_name or 'lt' == delay_name:
            maintenance_hours += duration
        elif 'stop' in delay_name:
            rhms_hours += duration
        elif 'discharge stopped by master' in delay_name:
            master_hours += duration
        elif 'bad weather' in delay_name:
            weather_hours += duration
        else:
            print(f"[DEBUG]   WARNING: no bucket match for delay={delay_name!r}")

    print(f"[DEBUG] Totals — maintenance={maintenance_hours:.2f}  "
          f"rhms={rhms_hours:.2f}  "
          f"master={master_hours:.2f}  "
          f"weather={weather_hours:.2f}")

    return [
        {
            'description':    'Unloading System Not Available',
            'duration':       f"{maintenance_hours:.2f}" if maintenance_hours else '',
            'responsibility': 'DPPL',
        },
        {
            'description':    'Receiving System Not Available',
            'duration':       f"{rhms_hours:.2f}" if rhms_hours else '',
            'responsibility': 'Steel Plant (RMHS)',
        },
        {
            'description':    'Discharge Stopped by Master',
            'duration':       f"{master_hours:.2f}" if master_hours else '',
            'responsibility': 'Vessel Account',
        },
        {
            'description':    'Bad Weather',
            'duration':       f"{weather_hours:.2f}" if weather_hours else '',
            'responsibility': 'Force Majeure',
        },
    ]


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@bp.route('/module/RP01/mbc-sof/')
@login_required
def mbc_sof_list():
    records        = _fetch_mbc_list()
    import_records = [r for r in records if (r.get('operation_type') or '').lower() == 'import']
    export_records = [r for r in records if (r.get('operation_type') or '').lower() != 'import']
    return render_template('mbc_sof/mbc_sof_list.html',
                           import_records=import_records,
                           export_records=export_records,
                           username=session.get('username'))


@bp.route('/module/RP01/mbc-sof/<int:mbc_id>')
@login_required
def mbc_sof_print(mbc_id):
    header, load_port, discharge_port, export_load = _fetch_mbc_sof_data(mbc_id)
    if not header:
        return "Record not found", 404

    op_type = (header.get('operation_type') or '').lower()

    if op_type == 'export':
        activities = _build_export_activities(export_load)
    else:
        activities = _build_import_activities(load_port, discharge_port)

    delays = _build_delays(mbc_id)

    mbc_name            = header.get('mbc_name', '')
    mbc_no              = header.get('doc_num', '')
    cargo_name          = header.get('cargo_name', '') or header.get('cargo_type', '')
    bl_qty              = fmt_qty(header.get('bl_quantity') or 0)
    uom                 = header.get('quantity_uom', 'MT')
    # Port display:
    # IMPORT  -> hard-coded Discharge Port
    # EXPORT  -> keep Discharge Port blank
    load_port_name = header.get('load_port', '') or ''

    if op_type == 'export':
        discharge_port_name = ''
    else:
        discharge_port_name = 'JSW DHARAMTAR PORT'

    doc_date_display = fmt_date_display(header.get('doc_date'))

    return render_template('mbc_sof/mbc_sof_print.html',
                           header=header,
                           mbc_name=mbc_name,
                           mbc_no=mbc_no,
                           op_type=header.get('operation_type', ''),
                           activities=activities,
                           delays=delays,
                           cargo_name=cargo_name,
                           bl_qty=bl_qty,
                           uom=uom,
                           load_port_name=load_port_name,
                           discharge_port_name=discharge_port_name,
                           doc_date_display=doc_date_display,
                           mbc_id=mbc_id)
