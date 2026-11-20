import io
import json as _json
import mimetypes
import os
import psycopg2.errors
from flask import Blueprint, render_template, request, jsonify, session, redirect, url_for, send_file
from werkzeug.security import check_password_hash
from functools import wraps
from . import model
from database import get_user_permissions, get_module_config, get_db, get_cursor

bp = Blueprint('MBC01', __name__, template_folder='.')
MODULE_CODE = 'MBC01'
MODULE_INFO = {'code': 'MBC01', 'name': 'MBC Operation'}

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


def _is_approver():
    config = get_module_config(MODULE_CODE)
    return (str(config.get('approver_id', '')) == str(session.get('user_id'))
            or bool(session.get('is_admin')))


def _mbc_id_of(table, row_id):
    """Owning mbc_id for a sub-table row, so a line edit is gated by its parent."""
    if not row_id:
        return None
    conn = get_db()
    cur = get_cursor(conn)
    # table is one of the fixed names passed by the routes below, never input.
    cur.execute(f'SELECT mbc_id FROM {table} WHERE id=%s', [row_id])
    row = cur.fetchone()
    conn.close()
    return row['mbc_id'] if row else None


def _check_editable(mbc_id):
    """Gate every MBC write. Returns an error response, or None when allowed.

    Two locks, in order of severity:
      billed / invoiced — nobody may change it, not even admin. The cargo has
        left the operations desk; a retyped customer or quantity here silently
        rewrites what a customer was already charged.
      Approved — operations may not reopen it; an admin resets it from
        Admin → MBC Approval, which is itself refused once billed.
    """
    if not mbc_id:
        return None
    from modules.FIN01 import model as fin_model
    conn = get_db()
    cur = get_cursor(conn)
    try:
        state = fin_model.source_lock_state(cur, 'MBC', mbc_id)
    finally:
        conn.close()
    if state:
        return jsonify({'error': fin_model.lock_message(state, 'MBC')}), 403
    if model.get_doc_status(mbc_id) == 'Approved' and not _is_approver():
        return jsonify({'error': 'This MBC is Approved and locked for editing. '
                                 'Ask an admin to reset it from Admin → MBC Approval.'}), 403
    return None

@bp.route('/module/MBC01/')
@login_required
def view():
    perms = get_perms()
    if not perms.get('can_read'):
        return render_template('no_access.html'), 403
    return render_template('mbc01.html', permissions=perms)

@bp.route('/api/module/MBC01/data')
@login_required
def get_data():
    try:
        page = int(request.args.get('page', 1))
        size = int(request.args.get('size', 20))
    except (ValueError, TypeError):
        page, size = 1, 20
    try:
        filters = _json.loads(request.args.get('filters', '[]'))
    except _json.JSONDecodeError:
        filters = []
    rows, total = model.get_data(page, size, filters)
    return jsonify({'data': rows, 'last_page': (total + size - 1) // size, 'total': total})

@bp.route('/api/module/MBC01/save', methods=['POST'])
@login_required
def save():
    perms = get_perms()
    data = request.json
    is_new = not data.get('id')
    if is_new and not perms.get('can_add'):
        return jsonify({'error': 'No permission to add'}), 403
    if not is_new and not perms.get('can_edit'):
        return jsonify({'error': 'No permission to edit'}), 403

    if not is_new:
        blocked = _check_editable(data['id'])
        if blocked:
            return blocked
        data['doc_status'] = ('Approved' if model.get_doc_status(data['id']) == 'Approved'
                              else 'Draft')
    else:
        data['doc_status'] = 'Draft'

    row_id, doc_num = model.save_header(data)
    return jsonify({'id': row_id, 'doc_num': doc_num, 'doc_status': data.get('doc_status', 'Draft')})


@bp.route('/api/module/MBC01/approval_check/<int:record_id>')
@login_required
def approval_check(record_id):
    return jsonify(model.get_approval_eligibility(record_id))


@bp.route('/api/module/MBC01/approve', methods=['POST'])
@login_required
def approve():
    config = get_module_config('MBC01')
    is_approver = str(config.get('approver_id', '')) == str(session.get('user_id')) or session.get('is_admin')
    if not is_approver:
        return jsonify({'error': 'No permission to approve'}), 403
    data = request.json or {}
    record_id = data.get('id')
    password = (data.get('password') or '').strip()
    if not record_id:
        return jsonify({'error': 'Missing id'}), 400
    if not password:
        return jsonify({'error': 'Password is required'}), 400

    # Verify password server-side
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT id, password FROM users WHERE id=%s', [session.get('user_id')])
    user = cur.fetchone()
    conn.close()
    if not user or not check_password_hash(user['password'], password):
        return jsonify({'error': 'Incorrect password'}), 403

    eligibility = model.get_approval_eligibility(record_id)
    if not eligibility['eligible']:
        return jsonify({'error': 'Record not eligible for approval', 'missing': eligibility['missing']}), 400

    # Enforce: at least one Proof of Quantity document must be uploaded
    conn_doc = get_db()
    cur_doc = get_cursor(conn_doc)
    cur_doc.execute('SELECT COUNT(*) FROM mbc_proof_documents WHERE mbc_id=%s', [record_id])
    doc_count = cur_doc.fetchone()['count']
    conn_doc.close()
    if doc_count == 0:
        return jsonify({'error': 'At least one Proof of Quantity document must be uploaded before approving'}), 400

    model.approve_record(record_id, session.get('username'))
    return jsonify({'doc_status': 'Approved'})


@bp.route('/api/module/MBC01/send_back', methods=['POST'])
@login_required
def send_back():
    config = get_module_config('MBC01')
    is_approver = str(config.get('approver_id', '')) == str(session.get('user_id')) or session.get('is_admin')
    if not is_approver:
        return jsonify({'error': 'No permission'}), 403
    data = request.json
    record_id = data.get('id')
    comment = (data.get('comment') or '').strip()
    if not record_id:
        return jsonify({'error': 'Missing id'}), 400
    if not comment:
        return jsonify({'error': 'A reason is required when sending back to Draft'}), 400
    blocked = _check_editable(record_id)
    if blocked:
        return blocked
    model.send_back_to_draft(record_id, comment, session.get('username'))
    return jsonify({'doc_status': 'Draft'})


@bp.route('/api/module/MBC01/approval-log/<int:record_id>')
@login_required
def approval_log(record_id):
    return jsonify(model.get_approval_log(record_id))


@bp.route('/api/module/MBC01/delete', methods=['POST'])
@login_required
def delete():
    perms = get_perms()
    if not perms.get('can_delete'):
        return jsonify({'error': 'No permission to delete'}), 403
    blocked = _check_editable(request.json['id'])
    if blocked:
        return blocked
    model.delete_header(request.json['id'])
    return jsonify({'success': True})

# Load Port Lines sub-table endpoints
@bp.route('/api/module/MBC01/load_port/<int:mbc_id>')
@login_required
def get_load_port_lines(mbc_id):
    return jsonify(model.get_load_port_lines(mbc_id))

@bp.route('/api/module/MBC01/load_port/save', methods=['POST'])
@login_required
def save_load_port_line():
    perms = get_perms()
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'error': 'No permission'}), 403
    data = request.json
    blocked = _check_editable(data.get('mbc_id') or _mbc_id_of('mbc_load_port_lines', data.get('id')))
    if blocked:
        return blocked
    row_id = model.save_load_port_line(data)
    return jsonify({'id': row_id, 'success': True})

@bp.route('/api/module/MBC01/load_port/delete', methods=['POST'])
@login_required
def delete_load_port_line():
    perms = get_perms()
    if not perms.get('can_delete'):
        return jsonify({'error': 'No permission to delete'}), 403
    blocked = _check_editable(_mbc_id_of('mbc_load_port_lines', request.json['id']))
    if blocked:
        return blocked
    model.delete_load_port_line(request.json['id'])
    return jsonify({'success': True})

# Discharge Port Lines sub-table endpoints
@bp.route('/api/module/MBC01/discharge_port/<int:mbc_id>')
@login_required
def get_discharge_port_lines(mbc_id):
    return jsonify(model.get_discharge_port_lines(mbc_id))

@bp.route('/api/module/MBC01/discharge_port/save', methods=['POST'])
@login_required
def save_discharge_port_line():
    perms = get_perms()
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'error': 'No permission'}), 403
    data = request.json
    blocked = _check_editable(data.get('mbc_id') or _mbc_id_of('mbc_discharge_port_lines', data.get('id')))
    if blocked:
        return blocked
    row_id = model.save_discharge_port_line(data)
    return jsonify({'id': row_id, 'success': True})

@bp.route('/api/module/MBC01/discharge_port/delete', methods=['POST'])
@login_required
def delete_discharge_port_line():
    perms = get_perms()
    if not perms.get('can_delete'):
        return jsonify({'error': 'No permission to delete'}), 403
    blocked = _check_editable(_mbc_id_of('mbc_discharge_port_lines', request.json['id']))
    if blocked:
        return blocked
    model.delete_discharge_port_line(request.json['id'])
    return jsonify({'success': True})

# Cleaning Details sub-table endpoints
@bp.route('/api/module/MBC01/cleaning/<int:mbc_id>')
@login_required
def get_cleaning_details(mbc_id):
    return jsonify(model.get_cleaning_details(mbc_id))

@bp.route('/api/module/MBC01/cleaning/save', methods=['POST'])
@login_required
def save_cleaning_detail():
    perms = get_perms()
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'error': 'No permission'}), 403
    data = request.json
    blocked = _check_editable(data.get('mbc_id') or _mbc_id_of('mbc_cleaning_details', data.get('id')))
    if blocked:
        return blocked
    row_id = model.save_cleaning_detail(data)
    return jsonify({'id': row_id, 'success': True})

@bp.route('/api/module/MBC01/cleaning/delete', methods=['POST'])
@login_required
def delete_cleaning_detail():
    perms = get_perms()
    if not perms.get('can_delete'):
        return jsonify({'error': 'No permission to delete'}), 403
    blocked = _check_editable(_mbc_id_of('mbc_cleaning_details', request.json['id']))
    if blocked:
        return blocked
    model.delete_cleaning_detail(request.json['id'])
    return jsonify({'success': True})

# Export Load Port Lines sub-table endpoints
@bp.route('/api/module/MBC01/export_load_port/<int:mbc_id>')
@login_required
def get_export_load_port_lines(mbc_id):
    return jsonify(model.get_export_load_port_lines(mbc_id))

@bp.route('/api/module/MBC01/export_load_port/save', methods=['POST'])
@login_required
def save_export_load_port_line():
    perms = get_perms()
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'error': 'No permission'}), 403
    data = request.json
    blocked = _check_editable(data.get('mbc_id') or _mbc_id_of('mbc_export_load_port_lines', data.get('id')))
    if blocked:
        return blocked
    row_id = model.save_export_load_port_line(data)
    return jsonify({'id': row_id, 'success': True})

@bp.route('/api/module/MBC01/export_load_port/delete', methods=['POST'])
@login_required
def delete_export_load_port_line():
    perms = get_perms()
    if not perms.get('can_delete'):
        return jsonify({'error': 'No permission to delete'}), 403
    blocked = _check_editable(_mbc_id_of('mbc_export_load_port_lines', request.json['id']))
    if blocked:
        return blocked
    model.delete_export_load_port_line(request.json['id'])
    return jsonify({'success': True})

# Customer Details sub-table endpoints
@bp.route('/api/module/MBC01/customer_details/<int:mbc_id>')
@login_required
def get_customer_details(mbc_id):
    return jsonify(model.get_customer_details(mbc_id))

@bp.route('/api/module/MBC01/customer_details/save', methods=['POST'])
@login_required
def save_customer_detail():
    perms = get_perms()
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'error': 'No permission'}), 403
    data = request.json
    blocked = _check_editable(data.get('mbc_id') or _mbc_id_of('mbc_customer_details', data.get('id')))
    if blocked:
        return blocked
    row_id = model.save_customer_detail(data)
    return jsonify({'id': row_id, 'success': True})

@bp.route('/api/module/MBC01/customer_details/delete', methods=['POST'])
@login_required
def delete_customer_detail():
    perms = get_perms()
    if not perms.get('can_delete'):
        return jsonify({'error': 'No permission to delete'}), 403
    blocked = _check_editable(_mbc_id_of('mbc_customer_details', request.json['id']))
    if blocked:
        return blocked
    model.delete_customer_detail(request.json['id'])
    return jsonify({'success': True})


@bp.route('/api/module/MBC01/customer_details/change_line', methods=['POST'])
@login_required
def change_line():
    """Fix the customer and/or Material PO on an unbilled line after approval /
    partial billing. Bypasses the whole-MBC lock on purpose; the per-line
    billing lock and a mandatory, logged reason stand in for it."""
    if not get_perms().get('can_edit'):
        return jsonify({'error': 'No permission'}), 403
    data = request.json or {}
    row_id = data.get('id')
    new_customer = (data.get('customer_name') or '').strip()
    new_po = (data.get('material_po') or '').strip()
    reason = (data.get('reason') or '').strip()
    # PO stays mandatory: approval requires it on every line.
    if not row_id or not new_customer or not new_po or not reason:
        return jsonify({'error': 'Customer line, customer, Material PO and reason are required'}), 400
    try:
        ok, err = model.change_line(row_id, new_customer, new_po, reason, session.get('username'))
    except psycopg2.errors.ForeignKeyViolation:
        return jsonify({'error': f'Unknown customer: {new_customer}'}), 400
    if not ok:
        return jsonify({'error': err}), 403
    return jsonify({'success': True})


# ── Proof of Quantity Documents (stored in DB as BYTEA) ──────────────────────

ALLOWED_EXTENSIONS = {'.pdf', '.jpg', '.jpeg', '.png', '.xlsx', '.xls', '.csv', '.doc', '.docx'}


@bp.route('/api/module/MBC01/proof_docs/upload', methods=['POST'])
@login_required
def upload_proof_docs():
    mbc_id = request.form.get('mbc_id')
    if not mbc_id:
        return jsonify({'error': 'Missing mbc_id'}), 400
    # Upload replaces the existing proof — the evidence behind a billed
    # quantity must not be swappable after the fact.
    blocked = _check_editable(mbc_id)
    if blocked:
        return blocked

    # One proof document per MBC — take a single file and replace any existing one.
    files = request.files.getlist('files')
    f = files[0] if files else None
    if not f or not f.filename:
        return jsonify({'error': 'No file provided'}), 400

    original = f.filename
    ext = os.path.splitext(original)[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        return jsonify({'error': 'Invalid file type (allowed: pdf, jpg, png, xlsx, csv, doc)'}), 400
    file_bytes = f.read()
    if not file_bytes:
        return jsonify({'error': 'Empty file'}), 400
    mime_type = f.mimetype or mimetypes.guess_type(original)[0] or 'application/octet-stream'

    conn = get_db()
    cur = get_cursor(conn)
    # Replace: the old file is not retained.
    cur.execute('DELETE FROM mbc_proof_documents WHERE mbc_id=%s', [mbc_id])
    cur.execute('''
        INSERT INTO mbc_proof_documents (mbc_id, original_filename, file_bytes, mime_type, uploaded_by)
        VALUES (%s, %s, %s, %s, %s) RETURNING id, original_filename, uploaded_at
    ''', [mbc_id, original, file_bytes, mime_type, session.get('username')])
    row = cur.fetchone()
    conn.commit()
    conn.close()

    return jsonify({'success': True, 'docs': [{'id': row['id'],
                    'original_filename': row['original_filename'],
                    'uploaded_at': str(row['uploaded_at'])[:16]}]})


@bp.route('/api/module/MBC01/proof_docs/<int:mbc_id>')
@login_required
def list_proof_docs(mbc_id):
    # Metadata-only query — never select file_bytes here, keeps the list fast.
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT id, original_filename, uploaded_by, uploaded_at
        FROM mbc_proof_documents WHERE mbc_id=%s ORDER BY uploaded_at
    ''', [mbc_id])
    docs = [{'id': r['id'], 'original_filename': r['original_filename'],
              'uploaded_by': r['uploaded_by'], 'uploaded_at': str(r['uploaded_at'])[:16]}
            for r in cur.fetchall()]
    conn.close()
    return jsonify({'docs': docs})


@bp.route('/api/module/MBC01/proof_docs/file/<int:doc_id>')
@login_required
def serve_proof_doc(doc_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT original_filename, file_bytes, mime_type FROM mbc_proof_documents WHERE id=%s', [doc_id])
    row = cur.fetchone()
    conn.close()
    if not row or row['file_bytes'] is None:
        return 'Not found', 404
    ext = os.path.splitext(row['original_filename'])[1].lower()
    inline_types = {'.pdf', '.jpg', '.jpeg', '.png'}
    as_attachment = ext not in inline_types
    return send_file(
        io.BytesIO(bytes(row['file_bytes'])),
        download_name=row['original_filename'],
        mimetype=row['mime_type'] or 'application/octet-stream',
        as_attachment=as_attachment,
    )
