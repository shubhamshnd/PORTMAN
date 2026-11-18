from flask import Blueprint, render_template, request, jsonify, session, redirect, url_for
from functools import wraps
from . import model
from database import get_user_permissions

bp = Blueprint('FSAP01', __name__, template_folder='.')
MODULE_CODE = 'FSAP01'


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


@bp.route('/module/FSAP01/')
@login_required
def view():
    perms = get_perms()
    if not perms.get('can_read'):
        return render_template('no_access.html'), 403
    return render_template('fsap01.html', permissions=perms)


@bp.route('/api/module/FSAP01/sap-invoice-logs')
@login_required
def sap_invoice_logs():
    page = int(request.args.get('page', 1))
    size = int(request.args.get('size', 50))
    data, total = model.get_sap_invoice_logs(page, size)
    return jsonify({'data': data, 'last_page': (total + size - 1) // size, 'total': total})


@bp.route('/api/module/FSAP01/callback-logs')
@login_required
def callback_logs():
    page = int(request.args.get('page', 1))
    size = int(request.args.get('size', 50))
    data, total = model.get_callback_logs(page, size)
    return jsonify({'data': data, 'last_page': (total + size - 1) // size, 'total': total})


@bp.route('/api/module/FSAP01/callback-logs/<int:log_id>', methods=['DELETE'])
@login_required
def delete_callback_log(log_id):
    """Admin only: delete a callback log and restore the invoice/CN it changed."""
    if not session.get('is_admin'):
        return jsonify({'ok': False, 'error': 'Admin only'}), 403
    ok, msg = model.delete_callback_log(log_id)
    return jsonify({'ok': ok, 'message' if ok else 'error': msg}), (200 if ok else 404)


@bp.route('/api/module/FSAP01/outbound-logs')
@login_required
def outbound_logs():
    """All outbound SAP interactions — postings, reversals, CNs, playground, IRN."""
    page = int(request.args.get('page', 1))
    size = int(request.args.get('size', 50))
    type_filter = request.args.get('type') or None
    data, total = model.get_outbound_logs(page, size, type_filter)
    return jsonify({'data': data, 'last_page': (total + size - 1) // size, 'total': total})


@bp.route('/api/module/FSAP01/sap-queue')
@login_required
def sap_queue_list():
    """The async SAP outbound queue with retry/payload details."""
    page = int(request.args.get('page', 1))
    size = int(request.args.get('size', 50))
    status = request.args.get('status') or None
    data, total = model.get_sap_queue(page, size, status)
    return jsonify({'data': data, 'last_page': (total + size - 1) // size, 'total': total})


@bp.route('/api/module/FSAP01/sap-queue/manual-send', methods=['POST'])
@login_required
def sap_queue_manual_send():
    """Manually fire a queued/failed SAP job from the FSAP01 console."""
    perms = get_perms()
    if not perms.get('can_edit'):
        return jsonify({'ok': False, 'error': 'No permission'}), 403
    queue_id = (request.json or {}).get('queue_id')
    if not queue_id:
        return jsonify({'ok': False, 'error': 'queue_id required'}), 400
    import sap_queue
    return jsonify(sap_queue.manual_send(queue_id))


@bp.route('/api/module/FSAP01/sap-cn-logs')
@login_required
def sap_cn_logs():
    page = int(request.args.get('page', 1))
    size = int(request.args.get('size', 50))
    data, total = model.get_sap_cn_logs(page, size)
    return jsonify({'data': data, 'last_page': (total + size - 1) // size, 'total': total})


@bp.route('/api/module/FSAP01/gst-logs')
@login_required
def gst_logs():
    page = int(request.args.get('page', 1))
    size = int(request.args.get('size', 50))
    data, total = model.get_gst_logs(page, size)
    return jsonify({'data': data, 'last_page': (total + size - 1) // size, 'total': total})


# ── Staging table endpoints ───────────────────────────────────────────────────

@bp.route('/api/module/FSAP01/staging/push', methods=['POST'])
@login_required
def staging_push():
    """Write invoice data into invoice_sap_staging (processing_status = 'N')."""
    perms = get_perms()
    if not perms.get('can_add'):
        return jsonify({'ok': False, 'error': 'No permission'}), 403

    invoice_id = (request.json or {}).get('invoice_id')
    if not invoice_id:
        return jsonify({'ok': False, 'error': 'invoice_id required'}), 400

    from flask import session as _session
    result = model.push_invoice_to_staging(invoice_id, pushed_by=_session.get('username'))
    return jsonify(result)


@bp.route('/api/module/FSAP01/staging/sync', methods=['POST'])
@login_required
def staging_sync():
    """Read SAP response back from staging and update invoice_header."""
    invoice_id = (request.json or {}).get('invoice_id')
    if not invoice_id:
        return jsonify({'ok': False, 'error': 'invoice_id required'}), 400

    result = model.sync_staging_response(invoice_id)
    return jsonify(result)


@bp.route('/api/module/FSAP01/staging/rows/<int:invoice_id>')
@login_required
def staging_rows(invoice_id):
    """Return all staging rows for a specific invoice (for debug/audit)."""
    rows = model.get_staging_rows_for_invoice(invoice_id)
    return jsonify({'data': rows, 'total': len(rows)})


@bp.route('/api/module/FSAP01/staging/list')
@login_required
def staging_list():
    """Paginated list of all staging rows with optional status filter."""
    page   = int(request.args.get('page', 1))
    size   = int(request.args.get('size', 50))
    status = request.args.get('status') or None
    data, total = model.get_staging_rows(page, size, status)
    return jsonify({'data': data, 'last_page': (total + size - 1) // size, 'total': total})


# ── SAP adapter inbound callback ──────────────────────────────────────────────
# Note: the live /api/sap/callback handler now lives in sap_inbound.py and is
# registered from app.py. It uses Bearer token auth + the `Record[]` schema
# documented in docs/SAP_Callback_API.md. The legacy reference_text-based
# handler that used to live here was removed to stop it from shadowing the
# new endpoint at the same URL.
