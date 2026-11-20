from flask import Blueprint, render_template, request, jsonify, session, redirect, url_for, Response
from functools import wraps
from database import get_module_config, get_user_permissions, get_db, get_cursor
from datetime import datetime, timedelta
import re
import os
import csv
import io

bp = Blueprint('AUD01', __name__, template_folder='.')
MODULE_CODE = 'AUD01'
MODULE_INFO = {'code': 'AUD01', 'name': 'Audit Logs'}

# Nginx combined log format:
# 127.0.0.1 - - [16/Mar/2026:15:13:08 +0530] "GET /module/LDUD01/ HTTP/1.1" 200 4321 "..." "..."
_NGINX_RE = re.compile(
    r'^(\S+)\s+\S+\s+\S+\s+\[([^\]]+)\]\s+"(\S+)\s+([^"]+)\s+HTTP[^"]*"\s+(\d+)\s+(\d+)'
)

_MONTH_MAP = {
    'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5,  'Jun': 6,
    'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12,
}
_NGINX_DT_RE = re.compile(r'(\d{2})/(\w{3})/(\d{4}):(\d{2}):(\d{2}):(\d{2})')

def _parse_nginx_ts(ts_str):
    m = _NGINX_DT_RE.search(ts_str)
    if not m:
        return None
    day, mon, year, h, mi, s = m.groups()
    try:
        return datetime(int(year), _MONTH_MAP.get(mon, 1), int(day), int(h), int(mi), int(s))
    except ValueError:
        return None

# Static-asset extensions and path prefixes to ignore
_SKIP_EXTS = {'.png', '.jpg', '.jpeg', '.gif', '.ico', '.css', '.js',
              '.woff', '.woff2', '.ttf', '.svg', '.map', '.eot'}
_SKIP_PREFIXES = ('/static/', '/favicon')

_MODULE_NAMES = {
    'VC01': 'Vessel Creation', 'VCN01': 'Vessel Call Number',
    'LDUD01': 'Loading Unloading', 'MBC01': 'MBC Operation',
    'LUEU01': 'Load/Unload Equip Utilization', 'SRV01': 'Service Recording',
    'FIN01': 'Billing', 'FINV01': 'Invoicing', 'FCAM01': 'Customer Agreements',
    'FSTM01': 'Service Type Master', 'FGRM01': 'GST Rate Master',
    'FCRM01': 'Currency Master', 'FDCN01': 'Debit/Credit Note',
    'FSAP01': 'SAP Integration', 'FLOG01': 'Integration Logs',
    'VTM01': 'Vessel Type Master', 'VCM01': 'Vessel Country Master',
    'VAM01': 'Vessel Agent Master', 'VCUM01': 'Vessel Currency',
    'VRT01': 'Run Types', 'VDM01': 'Vessel Delay Master',
    'PDM01': 'Port Delay Master', 'VCG01': 'Cargo Master',
    'VQM01': 'Quantity UOM', 'VHO01': 'Vessel Holds',
    'VEM01': 'Equipment Master', 'VBM01': 'Barge Master',
    'VSDM01': 'Vessel Stevedore Master', 'MBCM01': 'MBC Master',
    'PBM01': 'Port Berth Master', 'PPL01': 'Port Payloader Master',
    'CRM01': 'Conveyor Route Master', 'VANM01': 'Anchorage Master',
    'VPM01': 'Port Master', 'TM01': 'Tide Master',
    'VCDS01': 'VCN Doc Series', 'MBCDS01': 'MBC Doc Series',
    'INVDS01': 'Invoice Doc Series', 'GSTCFG': 'GST API Config',
    'PSM01': 'PSM', 'PSMM01': 'PSMM', 'PSOM01': 'PSOM',
    'RP01': 'Reports', 'AUD01': 'Audit Logs', 'ADMIN': 'Admin Panel',
}

_API_MODULE_RE  = re.compile(r'^/api/module/([A-Z0-9]+)(?:/|$)')
_PAGE_MODULE_RE = re.compile(r'^/module/([A-Z0-9]+)(?:/|$)')


def _detect_module(path):
    m = _API_MODULE_RE.match(path)
    if m:
        code = m.group(1)
        return code, _MODULE_NAMES.get(code, code)
    m = _PAGE_MODULE_RE.match(path)
    if m:
        code = m.group(1)
        return code, _MODULE_NAMES.get(code, code)
    if path.startswith('/admin'):
        return 'ADMIN', 'Admin Panel'
    if path.startswith('/login') or path.startswith('/auth/'):
        return 'AUTH', 'Authentication'
    if path.startswith('/logout'):
        return 'AUTH', 'Authentication'
    if path.startswith('/home'):
        return 'HOME', 'Home'
    if path.startswith('/api/'):
        return 'APP', 'App API'
    return 'OTHER', path


def _should_skip(path):
    if any(path.startswith(p) for p in _SKIP_PREFIXES):
        return True
    _, ext = os.path.splitext(path.split('?')[0])
    return ext.lower() in _SKIP_EXTS


def _get_log_path():
    cfg = get_module_config('AUD01') or {}
    return cfg.get('nginx_log_path', '').strip()


def _parse_nginx_log(log_path, from_dt=None, to_dt=None):
    """Parse nginx log and return entries within the given datetime range (inclusive).
    No date filter → returns all entries. Always returns newest-first."""
    if not log_path or not os.path.isfile(log_path):
        return [], f'Log file not found: {log_path}'

    entries = []
    try:
        with open(log_path, 'r', encoding='utf-8', errors='replace') as f:
            lines = f.readlines()
    except Exception as e:
        return [], str(e)

    for line in lines:
        line = line.strip()
        if not line:
            continue
        m = _NGINX_RE.match(line)
        if not m:
            continue
        ip, ts, method, path_qs, status, size = m.groups()
        path_only = path_qs.split('?')[0]
        if _should_skip(path_only):
            continue

        dt = _parse_nginx_ts(ts)
        if dt is None:
            continue

        if from_dt and dt < from_dt:
            continue
        if to_dt and dt > to_dt:
            continue

        mod_code, mod_name = _detect_module(path_only)
        entries.append({
            'ip': ip,
            'timestamp': ts,
            'date_iso': dt.strftime('%Y-%m-%d'),
            'method': method,
            'path': path_qs,
            'status': int(status),
            'size': int(size),
            'module_code': mod_code,
            'module_name': mod_name,
        })

    entries.reverse()
    return entries, None


def _calc_stats(entries):
    from collections import Counter, defaultdict
    status_counts  = Counter(str(e['status']) for e in entries)
    module_counts  = Counter(e['module_code'] for e in entries)
    ip_counts      = Counter(e['ip'] for e in entries)
    method_counts  = Counter(e['method'] for e in entries)

    date_groups = defaultdict(Counter)
    for e in entries:
        date_groups[e['date_iso']][str(e['status'])] += 1

    by_date = [{'date': d, **dict(counts)}
               for d, counts in sorted(date_groups.items())]

    return {
        'total': len(entries),
        'status_counts': dict(status_counts),
        'top_modules': module_counts.most_common(10),
        'top_ips': ip_counts.most_common(10),
        'method_counts': dict(method_counts),
        'by_date': by_date,
    }


def _parse_date_param(s, fallback):
    try:
        return datetime.strptime(s, '%Y-%m-%d') if s else fallback
    except ValueError:
        return fallback


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


def get_perms():
    if session.get('is_admin'):
        return {'can_read': 1, 'can_add': 1, 'can_edit': 1, 'can_delete': 1}
    return get_user_permissions(session.get('user_id'), MODULE_CODE)


@bp.route('/module/AUD01/')
@login_required
def view():
    perms = get_perms()
    if not perms.get('can_read'):
        return render_template('no_access.html'), 403
    return render_template('aud01.html', permissions=perms)


@bp.route('/api/module/AUD01/data')
@login_required
def get_data():
    if not get_perms().get('can_read'):
        return jsonify({'error': 'No permission'}), 403

    today   = datetime.now().replace(hour=23, minute=59, second=59, microsecond=0)
    default_from = (datetime.now() - timedelta(days=6)).replace(hour=0, minute=0, second=0, microsecond=0)

    to_dt   = _parse_date_param(request.args.get('to_date'),   today).replace(hour=23, minute=59, second=59)
    from_dt = _parse_date_param(request.args.get('from_date'), default_from).replace(hour=0, minute=0, second=0)

    # Hard cap: never show more than 31 days in the UI
    if (to_dt - from_dt).days > 31:
        from_dt = (to_dt - timedelta(days=30)).replace(hour=0, minute=0, second=0)

    log_path = _get_log_path()
    entries, err = _parse_nginx_log(log_path, from_dt=from_dt, to_dt=to_dt)
    if err and not entries:
        return jsonify({'error': err}), 500

    stats = _calc_stats(entries)
    return jsonify({
        'entries': entries,
        'stats': stats,
        'from_date': from_dt.strftime('%Y-%m-%d'),
        'to_date': to_dt.strftime('%Y-%m-%d'),
    })


# DB audit tables, merged into one timeline. record_id is resolved to the
# document number per module so the trail reads without a second lookup.
# FIN01 'Cargo released' rows carry a cargo line id, not a bill id.
# integration_logs has its own screen (FLOG01), so it isn't repeated here.
_TRAIL_SQL = '''
    SELECT * FROM (
        SELECT 'Approval Log' AS source, al.module_code, al.record_id, al.action,
               al.comment, al.actioned_by AS by_user, al.actioned_at AS ts,
               CASE al.module_code
                 WHEN 'MBC01'  THEN (SELECT doc_num || ' · ' || COALESCE(mbc_name, '')
                                       FROM mbc_header WHERE id = al.record_id)
                 WHEN 'VCN01'  THEN (SELECT vcn_doc_num || ' · ' || COALESCE(vessel_name, '')
                                       FROM vcn_header WHERE id = al.record_id)
                 WHEN 'LDUD01' THEN (SELECT doc_num FROM ldud_header WHERE id = al.record_id)
                 WHEN 'FINV01' THEN (SELECT invoice_number FROM invoice_header WHERE id = al.record_id)
                 WHEN 'FIN01'  THEN CASE WHEN al.action ILIKE 'Cargo released%%' THEN NULL
                                    ELSE (SELECT bill_number FROM bill_header WHERE id = al.record_id) END
               END AS ref
        FROM approval_log al
        UNION ALL
        SELECT 'Cutover Audit', 'CUTOVER', NULL, action, details::text,
               performed_by, performed_at, NULL
        FROM cutover_audit
    ) t
    WHERE ts BETWEEN %s AND %s
    ORDER BY ts DESC
    LIMIT %s
'''
_TRAIL_LIMIT = 5000  # ponytail: hard row cap; add paging if a month ever exceeds it


@bp.route('/api/module/AUD01/trail')
@login_required
def get_trail():
    if not get_perms().get('can_read'):
        return jsonify({'error': 'No permission'}), 403
    to_dt = _parse_date_param(request.args.get('to_date'), datetime.now()).replace(hour=23, minute=59, second=59)
    from_dt = _parse_date_param(request.args.get('from_date'), to_dt - timedelta(days=29)).replace(hour=0, minute=0, second=0)
    conn = get_db()
    try:
        cur = get_cursor(conn)
        cur.execute(_TRAIL_SQL, (from_dt, to_dt, _TRAIL_LIMIT + 1))
        rows = [dict(r) for r in cur.fetchall()]
    finally:
        conn.close()
    for r in rows:
        r['ts'] = r['ts'].strftime('%d-%m-%Y %H:%M:%S') if r['ts'] else ''
    return jsonify({
        'rows': rows[:_TRAIL_LIMIT],
        'truncated': len(rows) > _TRAIL_LIMIT,
        'from_date': from_dt.strftime('%Y-%m-%d'),
        'to_date': to_dt.strftime('%Y-%m-%d'),
    })


@bp.route('/api/module/AUD01/export')
@login_required
def export_csv():
    """Export ALL log entries (no date cap) as CSV download."""
    if not get_perms().get('can_read'):
        return jsonify({'error': 'No permission'}), 403

    log_path = _get_log_path()

    # Optional date filter from query params; default = all data
    from_dt = _parse_date_param(request.args.get('from_date'), None)
    to_dt   = _parse_date_param(request.args.get('to_date'),   None)
    if from_dt:
        from_dt = from_dt.replace(hour=0, minute=0, second=0)
    if to_dt:
        to_dt = to_dt.replace(hour=23, minute=59, second=59)

    entries, err = _parse_nginx_log(log_path, from_dt=from_dt, to_dt=to_dt)

    out = io.StringIO()
    writer = csv.writer(out)
    writer.writerow(['Timestamp', 'Date', 'IP', 'Method', 'Status', 'Module Code', 'Module Name', 'Path', 'Size (bytes)'])
    for e in entries:
        writer.writerow([
            e['timestamp'], e['date_iso'], e['ip'], e['method'],
            e['status'], e['module_code'], e['module_name'],
            e['path'], e['size'],
        ])

    fname = f"audit_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    return Response(
        out.getvalue(),
        mimetype='text/csv',
        headers={'Content-Disposition': f'attachment; filename={fname}'}
    )
