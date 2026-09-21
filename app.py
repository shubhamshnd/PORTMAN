from flask import Flask, render_template, request, redirect, url_for, session, jsonify, send_from_directory
from functools import wraps
from database import get_db, get_cursor
from config import SECRET_KEY, FLASK_ENV, SERVER_HOST, SERVER_PORT
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = SECRET_KEY


@app.route('/favicon.ico')
def favicon():
    return send_from_directory(
        app.static_folder + '/img/favicon_io',
        'favicon.ico',
        mimetype='image/vnd.microsoft.icon'
    )


@app.template_filter('indian_number')
def indian_number_filter(value):
    """Format a number in Indian comma style: 1,23,45,678.90"""
    try:
        f = float(value)
        negative = f < 0
        f = abs(f)
        int_part = int(f)
        dec_part = round((f - int_part) * 100)
        s = str(int_part)
        if len(s) <= 3:
            formatted = s
        else:
            last3 = s[-3:]
            rest = s[:-3]
            groups = []
            while rest:
                groups.append(rest[-2:])
                rest = rest[:-2]
            groups.reverse()
            formatted = ','.join(groups) + ',' + last3
        result = formatted + '.' + f'{dec_part:02d}'
        return ('-' if negative else '') + result
    except Exception:
        return str(value)

# ── App-level logging ─────────────────────────────────────────────────────────
import logging
import os
from logging.handlers import RotatingFileHandler

_log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
os.makedirs(_log_dir, exist_ok=True)
_log_file = os.path.join(_log_dir, 'app.log')

_file_handler = RotatingFileHandler(
    _log_file,
    maxBytes=10 * 1024 * 1024,  # 10 MB per file
    backupCount=10,
    encoding='utf-8',
)
_file_handler.setLevel(logging.WARNING)
_file_handler.setFormatter(logging.Formatter(
    '%(asctime)s [%(levelname)s] %(name)s %(module)s:%(lineno)d — %(message)s'
))

# Attach to Flask app logger and root logger (catches all libraries)
app.logger.setLevel(logging.WARNING)
app.logger.addHandler(_file_handler)
logging.getLogger().addHandler(_file_handler)
logging.getLogger().setLevel(logging.WARNING)

# Module registry
MODULES = {}

def register_module(code, name, blueprint):
    MODULES[code] = {'name': name}
    app.register_blueprint(blueprint)

# Import and register existing modules
from modules.VC01 import bp as vc01_bp, MODULE_INFO as vc01_info
from modules.VTM01 import bp as vtm01_bp, MODULE_INFO as vtm01_info
from modules.VCM01 import bp as vcm01_bp, MODULE_INFO as vcm01_info
from modules.VFM01 import bp as vfm01_bp, MODULE_INFO as vfm01_info
from modules.GM01 import bp as gm01_bp, MODULE_INFO as gm01_info
from modules.ADMIN import bp as admin_bp, MODULE_INFO as admin_info

# Import new master modules
from modules.VAM01 import bp as vam01_bp, MODULE_INFO as vam01_info
from modules.VCUM01 import bp as vcum01_bp, MODULE_INFO as vcum01_info
from modules.VCDS01 import bp as vcds01_bp, MODULE_INFO as vcds01_info
from modules.VTOD01 import bp as vtod01_bp, MODULE_INFO as vtod01_info
from modules.VRT01 import bp as vrt01_bp, MODULE_INFO as vrt01_info
from modules.VDM01 import bp as vdm01_bp, MODULE_INFO as vdm01_info
from modules.VCG01 import bp as vcg01_bp, MODULE_INFO as vcg01_info
from modules.VCN01 import bp as vcn01_bp, MODULE_INFO as vcn01_info
from modules.VQM01 import bp as vqm01_bp, MODULE_INFO as vqm01_info

from modules.VHO01 import bp as vho01_bp, MODULE_INFO as vho01_info
from modules.PDM01 import bp as pdm01_bp, MODULE_INFO as pdm01_info
from modules.VEM01 import bp as vem01_bp, MODULE_INFO as vem01_info
from modules.VBM01 import bp as vbm01_bp, MODULE_INFO as vbm01_info
from modules.VSDM01 import bp as vsdm01_bp, MODULE_INFO as vsdm01_info
from modules.LDUD01 import bp as ldud01_bp, MODULE_INFO as ldud01_info
from modules.MBCM01 import bp as mbcm01_bp, MODULE_INFO as mbcm01_info
from modules.PBM01 import bp as pbm01_bp, MODULE_INFO as pbm01_info
from modules.MBCDS01 import bp as mbcds01_bp, MODULE_INFO as mbcds01_info
from modules.INVDS01 import bp as invds01_bp, MODULE_INFO as invds01_info
from modules.CNDS01 import bp as cnds01_bp, MODULE_INFO as cnds01_info
from modules.MBC01 import bp as mbc01_bp, MODULE_INFO as mbc01_info
from modules.PPL01 import bp as ppl01_bp, MODULE_INFO as ppl01_info
from modules.LUEU01 import bp as eu01_bp, MODULE_INFO as eu01_info
from modules.CRM01 import bp as crm01_bp, MODULE_INFO as crm01_info
from modules.CARGO_STATS import bp as cargo_stats_bp, MODULE_INFO as cargo_stats_info

# Import finance modules
from modules.FCRM01 import bp as fcrm01_bp, MODULE_INFO as fcrm01_info
from modules.FGRM01 import bp as fgrm01_bp, MODULE_INFO as fgrm01_info
from modules.FSTM01 import bp as fstm01_bp, MODULE_INFO as fstm01_info
from modules.FCAM01 import bp as fcam01_bp, MODULE_INFO as fcam01_info
from modules.FIN01 import bp as fin01_bp, MODULE_INFO as fin01_info
from modules.SRV01 import bp as srv01_bp, MODULE_INFO as srv01_info
from modules.VANM01 import bp as vanm01_bp, MODULE_INFO as vanm01_info
from modules.VPM01 import bp as vpm01_bp, MODULE_INFO as vpm01_info
from modules.TM01 import bp as tm01_bp, MODULE_INFO as tm01_info
from modules.PSM01 import bp as psm01_bp, MODULE_INFO as psm01_info
from modules.PSMM01 import bp as psmm01_bp, MODULE_INFO as psmm01_info
from modules.PSOM01 import bp as psom01_bp, MODULE_INFO as psom01_info
from modules.BPOM01 import bp as bpom01_bp, MODULE_INFO as bpom01_info
from modules.TGM01 import bp as tgm01_bp, MODULE_INFO as tgm01_info
from modules.BOM01 import bp as bom01_bp, MODULE_INFO as bom01_info

# Import reports module
from modules.RP01 import bp as rp01_bp, MODULE_INFO as rp01_info
from modules.RP02 import bp as rp02_bp, MODULE_INFO as rp02_info
from modules.RP03 import bp as rp03_bp, MODULE_INFO as rp03_info

# Import audit logs module
from modules.AUD01 import bp as aud01_bp, MODULE_INFO as aud01_info

# Import accounts redesign modules
from modules.FINV01 import bp as finv01_bp, MODULE_INFO as finv01_info
from modules.GSTCFG import bp as gstcfg_bp, MODULE_INFO as gstcfg_info
from modules.FSAP01 import bp as fsap01_bp, MODULE_INFO as fsap01_info
from modules.FLOG01 import bp as flog01_bp, MODULE_INFO as flog01_info
from modules.FDCN01 import bp as fdcn01_bp, MODULE_INFO as fdcn01_info

# Register existing modules
register_module(vc01_info['code'], vc01_info['name'], vc01_bp)
register_module(vtm01_info['code'], vtm01_info['name'], vtm01_bp)
register_module(vcm01_info['code'], vcm01_info['name'], vcm01_bp)
register_module(vfm01_info['code'], vfm01_info['name'], vfm01_bp)
register_module(gm01_info['code'], gm01_info['name'], gm01_bp)
register_module(admin_info['code'], admin_info['name'], admin_bp)

# Register new master modules
register_module(vam01_info['code'], vam01_info['name'], vam01_bp)
register_module(vcum01_info['code'], vcum01_info['name'], vcum01_bp)
register_module(vcds01_info['code'], vcds01_info['name'], vcds01_bp)
register_module(vtod01_info['code'], vtod01_info['name'], vtod01_bp)
register_module(vrt01_info['code'], vrt01_info['name'], vrt01_bp)
register_module(vdm01_info['code'], vdm01_info['name'], vdm01_bp)
register_module(vcg01_info['code'], vcg01_info['name'], vcg01_bp)
register_module(vcn01_info['code'], vcn01_info['name'], vcn01_bp)
register_module(vqm01_info['code'], vqm01_info['name'], vqm01_bp)

register_module(vho01_info['code'], vho01_info['name'], vho01_bp)
register_module(pdm01_info['code'], pdm01_info['name'], pdm01_bp)
register_module(vem01_info['code'], vem01_info['name'], vem01_bp)
register_module(vbm01_info['code'], vbm01_info['name'], vbm01_bp)
register_module(vsdm01_info['code'], vsdm01_info['name'], vsdm01_bp)
register_module(ldud01_info['code'], ldud01_info['name'], ldud01_bp)
register_module(mbcm01_info['code'], mbcm01_info['name'], mbcm01_bp)
register_module(pbm01_info['code'], pbm01_info['name'], pbm01_bp)
register_module(mbcds01_info['code'], mbcds01_info['name'], mbcds01_bp)
register_module(invds01_info['code'], invds01_info['name'], invds01_bp)
register_module(cnds01_info['code'], cnds01_info['name'], cnds01_bp)
register_module(mbc01_info['code'], mbc01_info['name'], mbc01_bp)
register_module(ppl01_info['code'], ppl01_info['name'], ppl01_bp)
register_module(eu01_info['code'], eu01_info['name'], eu01_bp)
register_module(crm01_info['code'], crm01_info['name'], crm01_bp)
# Register finance modules
register_module(fcrm01_info['code'], fcrm01_info['name'], fcrm01_bp)
register_module(fgrm01_info['code'], fgrm01_info['name'], fgrm01_bp)
register_module(fstm01_info['code'], fstm01_info['name'], fstm01_bp)
register_module(fcam01_info['code'], fcam01_info['name'], fcam01_bp)
register_module(fin01_info['code'], fin01_info['name'], fin01_bp)
register_module(srv01_info['code'], srv01_info['name'], srv01_bp)
register_module(vanm01_info['code'], vanm01_info['name'], vanm01_bp)
register_module(vpm01_info['code'], vpm01_info['name'], vpm01_bp)
register_module(tm01_info['code'], tm01_info['name'], tm01_bp)
register_module(psm01_info['code'], psm01_info['name'], psm01_bp)
register_module(psmm01_info['code'], psmm01_info['name'], psmm01_bp)
register_module(psom01_info['code'], psom01_info['name'], psom01_bp)
register_module(bpom01_info['code'], bpom01_info['name'], bpom01_bp)
register_module(tgm01_info['code'], tgm01_info['name'], tgm01_bp)
register_module(bom01_info['code'], bom01_info['name'], bom01_bp)

# Register reports module
register_module(rp01_info['code'], rp01_info['name'], rp01_bp)
register_module(rp02_info['code'], rp02_info['name'], rp02_bp)
register_module(rp03_info['code'], rp03_info['name'], rp03_bp)

# Register audit logs module
register_module(aud01_info['code'], aud01_info['name'], aud01_bp)

# Register accounts redesign modules
register_module(finv01_info['code'], finv01_info['name'], finv01_bp)
register_module(gstcfg_info['code'], gstcfg_info['name'], gstcfg_bp)
register_module(fsap01_info['code'], fsap01_info['name'], fsap01_bp)
register_module(flog01_info['code'], flog01_info['name'], flog01_bp)
register_module(fdcn01_info['code'], fdcn01_info['name'], fdcn01_bp)
register_module(cargo_stats_info['code'], cargo_stats_info['name'], cargo_stats_bp)

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

# ── Error handlers ────────────────────────────────────────────────────────────
@app.errorhandler(Exception)
def handle_exception(e):
    """Catch all unhandled exceptions, log them, return a clean 500."""
    import traceback
    app.logger.error(
        'Unhandled exception on %s %s\n%s',
        request.method, request.path,
        traceback.format_exc()
    )
    # Don't swallow HTTP errors (404, 403 etc.) — let Flask handle them normally
    from werkzeug.exceptions import HTTPException
    if isinstance(e, HTTPException):
        return e
    return (
        '<h2>Internal Server Error</h2>'
        '<p>The error has been logged. Please contact the administrator.</p>',
        500,
    )

@app.errorhandler(500)
def internal_error(e):
    app.logger.error('500 error on %s %s: %s', request.method, request.path, str(e))
    return (
        '<h2>Internal Server Error</h2>'
        '<p>The error has been logged. Please contact the administrator.</p>',
        500,
    )

@app.after_request
def no_cache(response):
    """Prevent browser from caching authenticated pages."""
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

# ---- SAP inbound callback (SAP team posts results back) ----
import sap_inbound
sap_inbound.ensure_token_table()
app.add_url_rule('/api/sap/callback', view_func=sap_inbound.sap_callback_view, methods=['POST'])

# Self-heal: rebuild cargo/service billed tracking from the bills that exist,
# so cargo can never stay stuck as non-billable behind a rejected or deleted
# bill. Idempotent — safe every boot.
from modules.FIN01 import model as fin01_model
try:
    _repaired = fin01_model.reconcile_billed_tracking()
    if _repaired:
        app.logger.warning('FIN01 billed-tracking reconciliation repaired %s', _repaired)
    _stuck = fin01_model.stuck_billed_cargo()
    if _stuck:
        # Count and a sample only. The full list ran to 862 tuples on one line
        # every boot, which buried everything else in the log; anyone acting on
        # this needs the query, not a wall of ids.
        _sample = ', '.join('%s#%s' % (r['source_type'], r['id']) for r in _stuck[:5])
        app.logger.warning(
            'FIN01: %d cargo declaration(s) flagged billed with no live bill and no '
            'bill_id — cutover mark or stale tracking, needs a human. First few: %s%s',
            len(_stuck), _sample, ', ...' if len(_stuck) > 5 else '')
except Exception:
    app.logger.exception('FIN01 billed-tracking reconciliation failed at startup')


@app.route('/')
def index():
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        conn = get_db()
        cur = get_cursor(conn)
        cur.execute('SELECT * FROM users WHERE username = %s', (username,))
        user = cur.fetchone()
        conn.close()
        if user and check_password_hash(user['password'], password):
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['email'] = user['email'] or user['username']
            session['is_admin'] = bool(user['is_admin'])
            return redirect(url_for('home'))
        return render_template('login.html', error='Invalid credentials')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/auth/send-otp', methods=['POST'])
def send_otp():
    import random
    import string
    import secrets
    import datetime
    import threading
    from mail_service import queue_mail, process_mail_queue

    email = (request.json or {}).get('email', '').strip().lower()
    if not email:
        return jsonify({'error': 'Email required'}), 400

    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT id, username FROM users WHERE LOWER(email) = %s', [email])
    user = cur.fetchone()
    if not user:
        conn.close()
        return jsonify({'success': True})  # don't reveal if email exists

    cur.execute('UPDATE password_reset_tokens SET used=TRUE WHERE user_id=%s AND used=FALSE', [user['id']])

    otp_code = ''.join(random.choices(string.digits, k=6))
    reset_token = secrets.token_urlsafe(32)
    expires_at = datetime.datetime.now() + datetime.timedelta(minutes=15)

    cur.execute('''
        INSERT INTO password_reset_tokens (user_id, email, otp_code, reset_token, expires_at)
        VALUES (%s, %s, %s, %s, %s)
    ''', [user['id'], email, otp_code, reset_token, expires_at])
    conn.commit()
    conn.close()

    import urllib.parse
    login_url = request.host_url.rstrip('/') + '/login?reset=1&email=' + urllib.parse.quote(email)

    body_html = f"""
    <div style="font-family:'Segoe UI',Arial,sans-serif;max-width:480px;margin:0 auto;padding:32px 24px;background:#f7fafc;border-radius:10px;">
      <div style="text-align:center;margin-bottom:24px;">
        <h2 style="color:#2d3748;font-size:20px;margin:0;">Portbird - DPPL</h2>
        <p style="color:#718096;font-size:13px;margin:4px 0 0;">Password Reset Request</p>
      </div>
      <div style="background:#fff;border-radius:8px;padding:24px;border:1px solid #e2e8f0;">
        <p style="color:#2d3748;font-size:14px;margin:0 0 16px;">Hi <strong>{user['username']}</strong>,</p>
        <p style="color:#4a5568;font-size:13px;margin:0 0 20px;">We received a request to reset your password. Use the OTP below to set a new password.</p>
        <div style="text-align:center;background:#ebf8ff;border-radius:8px;padding:20px;margin:0 0 24px;">
          <p style="color:#2b6cb0;font-size:12px;font-weight:600;margin:0 0 8px;letter-spacing:1px;text-transform:uppercase;">Your One-Time Password</p>
          <div style="font-size:36px;font-weight:700;letter-spacing:10px;color:#1a365d;font-family:monospace;">{otp_code}</div>
          <p style="color:#718096;font-size:11px;margin:10px 0 0;">Valid for 15 minutes</p>
        </div>
        <div style="text-align:center;margin:0 0 20px;">
          <a href="{login_url}" style="display:inline-block;background:linear-gradient(135deg,#ec1c24,#2544a7);color:#fff;padding:12px 28px;border-radius:8px;text-decoration:none;font-size:14px;font-weight:600;letter-spacing:0.5px;">Enter OTP &amp; Reset Password</a>
        </div>
        <p style="color:#718096;font-size:11px;margin:0 0 6px;">Or copy this link into your browser:</p>
        <p style="color:#4a90d9;font-size:11px;word-break:break-all;margin:0 0 16px;">{login_url}</p>
        <p style="color:#a0aec0;font-size:11px;margin:0;">If you did not request this, you can safely ignore this email.</p>
      </div>
      <p style="text-align:center;color:#a0aec0;font-size:10px;margin:16px 0 0;">Portbird - DPPL &mdash; Port Management System</p>
    </div>
    """

    queue_mail(email, user['username'],
               'Password Reset OTP - Portbird DPPL', body_html, 'AUTH', user['id'])
    threading.Thread(target=process_mail_queue, daemon=True).start()

    return jsonify({'success': True})


@app.route('/auth/verify-otp', methods=['POST'])
def verify_otp():
    data = request.json or {}
    email = data.get('email', '').strip().lower()
    otp = data.get('otp', '').strip()
    if not email or not otp:
        return jsonify({'error': 'Email and OTP required'}), 400

    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT * FROM password_reset_tokens
        WHERE LOWER(email)=%s AND otp_code=%s AND used=FALSE AND expires_at > NOW()
        ORDER BY created_at DESC LIMIT 1
    ''', [email, otp])
    row = cur.fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'Invalid or expired OTP'}), 400

    cur.execute('UPDATE password_reset_tokens SET otp_verified=TRUE WHERE id=%s', [row['id']])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'reset_token': row['reset_token']})


@app.route('/auth/set-password', methods=['POST'])
def set_password_otp():
    data = request.json or {}
    reset_token = data.get('reset_token', '').strip()
    new_password = data.get('password', '').strip()
    if not reset_token or not new_password:
        return jsonify({'error': 'Token and password required'}), 400

    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT * FROM password_reset_tokens
        WHERE reset_token=%s AND otp_verified=TRUE AND used=FALSE AND expires_at > NOW()
        LIMIT 1
    ''', [reset_token])
    row = cur.fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'Invalid or expired reset session. Please start over.'}), 400

    cur.execute('UPDATE users SET password=%s WHERE id=%s', [generate_password_hash(new_password), row['user_id']])
    cur.execute('UPDATE password_reset_tokens SET used=TRUE WHERE id=%s', [row['id']])
    conn.commit()
    conn.close()
    return jsonify({'success': True})


@app.route('/home')
@login_required
def home():
    return render_template('home.html', modules=MODULES, username=session.get('username'), is_admin=session.get('is_admin'))

@app.route('/api/modules/search')
def search_modules():
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401
    query = request.args.get('q', '').lower()
    results = [{'code': k, 'name': v['name']} for k, v in MODULES.items()
               if query in k.lower() or query in v['name'].lower()]
    return jsonify(results)

# ── Mail queue scheduler ──────────────────────────────────────────────────────
from apscheduler.schedulers.background import BackgroundScheduler
from mail_service import process_mail_queue as _process_mail_queue
from mail_service import get_smtp_config as _get_smtp_cfg

def _mail_tick():
    """Runs every N minutes. No-op if mail is disabled."""
    try:
        _process_mail_queue()
    except Exception:
        pass

def _reschedule_mail_job():
    """Re-read schedule_minutes from DB and reschedule if needed."""
    try:
        cfg = _get_smtp_cfg()
        mins = max(1, int(cfg.get('schedule_minutes', 5))) if cfg else 5
        _mail_scheduler.reschedule_job('mail_queue', trigger='interval', minutes=mins)
    except Exception:
        pass

# ── Saved Filters API ────────────────────────────────────────────────────────
import json as _json

@app.route('/api/saved-filters/<module_code>', methods=['GET'])
@login_required
def get_saved_filters(module_code):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT id, name, filters_json FROM saved_filters WHERE module_code = %s ORDER BY name', [module_code])
    rows = cur.fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/saved-filters', methods=['POST'])
@login_required
def create_saved_filter():
    data = request.get_json() or {}
    name = (data.get('name') or '').strip()[:100]
    module_code = (data.get('module_code') or '').strip()[:20]
    filters = data.get('filters')
    if not name or not module_code or not filters:
        return jsonify({'error': 'Missing fields'}), 400
    filters_json = _json.dumps(filters)
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute(
        'INSERT INTO saved_filters (name, module_code, filters_json, created_by) VALUES (%s, %s, %s, %s) RETURNING id',
        [name, module_code, filters_json, session.get('username', '')]
    )
    row = cur.fetchone()
    conn.commit()
    conn.close()
    return jsonify({'id': row['id']}), 201

@app.route('/api/saved-filters/<int:filter_id>', methods=['DELETE'])
@login_required
def delete_saved_filter(filter_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('DELETE FROM saved_filters WHERE id = %s', [filter_id])
    conn.commit()
    conn.close()
    return jsonify({'ok': True})

def _sap_queue_tick():
    """Runs every 5 minutes. Retries due SAP postings; safe no-op if queue empty."""
    try:
        from sap_queue import process_sap_queue
        process_sap_queue()
    except Exception:
        pass

_mail_scheduler = BackgroundScheduler(daemon=True)
_mail_scheduler.add_job(
    _mail_tick,
    trigger='interval',
    minutes=5,
    id='mail_queue',
    replace_existing=True,
    max_instances=1,
)
_mail_scheduler.add_job(
    _sap_queue_tick,
    trigger='interval',
    minutes=5,
    id='sap_queue',
    replace_existing=True,
    max_instances=1,
)
_mail_scheduler.start()

if __name__ == '__main__':
    is_production = FLASK_ENV == 'production'

    if is_production:
        # Waitress is the production WSGI server for Windows.
        # It is multi-threaded, stable, and does not crash under load like Werkzeug.
        from waitress import serve
        import logging as _logging
        _logging.getLogger('waitress').addHandler(_file_handler)
        _logging.getLogger('waitress').setLevel(_logging.WARNING)
        print(f'Starting Waitress on {SERVER_HOST}:{SERVER_PORT} ...')
        serve(
            app,
            host=SERVER_HOST,
            port=SERVER_PORT,
            threads=8,               # concurrent request threads
            connection_limit=200,    # max open connections
            channel_timeout=60,      # drop hung connections after 60s
            log_socket_errors=True,
        )
    else:
        # Development only — Werkzeug with reloader
        app.run(
            host=SERVER_HOST,
            port=SERVER_PORT,
            debug=True,
            threaded=True,
        )
