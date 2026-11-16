from flask import Blueprint, render_template, request, jsonify, session
from functools import wraps
from database import get_user_permissions
from . import model

bp = Blueprint('BGT01', __name__, template_folder='.')
MODULE_CODE = 'BGT01'

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session and 'username' not in session:
            return jsonify({'error': 'Unauthorized'}), 401
        return f(*args, **kwargs)
    return decorated

def get_perms():
    if session.get('is_admin'):
        return {'can_read': 1, 'can_add': 1, 'can_edit': 1, 'can_delete': 1}
    perms = get_user_permissions(session.get('user_id'), MODULE_CODE)
    if not perms:
        return {'can_read': 1, 'can_add': 1, 'can_edit': 1, 'can_delete': 1}
    return perms

@bp.route('/module/BGT01/')
def view():
    if 'user_id' not in session and 'username' not in session:
        return 'Unauthorized', 401
    perms = get_perms()
    if not perms.get('can_read'):
        return render_template('no_access.html'), 403
    return render_template('bgt01.html', permissions=perms)

@bp.route('/api/module/BGT01/data', methods=['GET', 'POST'])
@login_required
def get_data():
    try:
        page = request.args.get('page', default=1, type=int)
        size = request.args.get('size', default=20, type=int)
        data, total = model.get_all_headers(page=page, size=size)
        last_page = max(1, (total + size - 1) // size) if size else 1
        return jsonify({
            'data': data,
            'total': total,
            'last_page': last_page
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@bp.route('/api/module/BGT01/available_vcns')
@login_required
def available_vcns():
    try:
        vcns = model.get_available_vcns()
        return jsonify(vcns)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@bp.route('/api/module/BGT01/barges_for_vcn/<int:vcn_id>')
@login_required
def barges_for_vcn(vcn_id):
    try:
        res = model.get_barges_for_vcn(vcn_id)
        return jsonify(res)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@bp.route('/api/module/BGT01/masters')
@login_required
def masters():
    try:
        data = model.get_masters()
        return jsonify(data)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@bp.route('/api/module/BGT01/save', methods=['POST'])
@login_required
def save_header_route():
    try:
        perms = get_perms()
        if not perms.get('can_add') and not perms.get('can_edit'):
            return jsonify({'error': 'Permission denied'}), 403
        data = request.json or {}
        if not data.get('created_by'):
            data['created_by'] = session.get('username') or 'admin'
        row_id = model.save_header(data)
        return jsonify({'success': True, 'id': row_id})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@bp.route('/api/module/BGT01/delete', methods=['POST'])
@login_required
def delete_header_route():
    try:
        perms = get_perms()
        if not perms.get('can_delete'):
            return jsonify({'error': 'Permission denied'}), 403
        header_id = (request.json or {}).get('id')
        if not header_id:
            return jsonify({'error': 'Missing ID'}), 400
        model.delete_header(header_id)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@bp.route('/api/module/BGT01/activities/<int:header_id>')
@login_required
def get_activities_route(header_id):
    try:
        activities = model.get_activities(header_id)
        return jsonify(activities)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@bp.route('/api/module/BGT01/activity/save', methods=['POST'])
@login_required
def save_activity_route():
    try:
        perms = get_perms()
        if not perms.get('can_add') and not perms.get('can_edit'):
            return jsonify({'error': 'Permission denied'}), 403
        data = request.json or {}
        row_id = model.save_activity(data)
        return jsonify({'success': True, 'id': row_id})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@bp.route('/api/module/BGT01/activity/delete', methods=['POST'])
@login_required
def delete_activity_route():
    try:
        perms = get_perms()
        if not perms.get('can_delete'):
            return jsonify({'error': 'Permission denied'}), 403
        act_id = (request.json or {}).get('id')
        if not act_id:
            return jsonify({'error': 'Missing ID'}), 400
        model.delete_activity(act_id)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

