from flask import Blueprint, render_template, request, jsonify, session, send_file
from functools import wraps
from database import get_user_permissions
from . import model
import io

bp = Blueprint('TGU01', __name__, template_folder='.')
MODULE_CODE = 'TGU01'

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

@bp.route('/module/TGU01/')
def view():
    if 'user_id' not in session and 'username' not in session:
        return 'Unauthorized', 401
    perms = get_perms()
    if not perms.get('can_read'):
        return render_template('no_access.html'), 403
    return render_template('tgu01.html', permissions=perms)

@bp.route('/api/module/TGU01/tugs', methods=['GET'])
@login_required
def get_tugs_route():
    try:
        tugs = model.get_all_tugs()
        return jsonify(tugs)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@bp.route('/api/module/TGU01/records', methods=['GET'])
@login_required
def get_records_route():
    try:
        tug_name = (request.args.get('tug') or request.args.get('tug_name') or '').strip()
        page = request.args.get('page', default=1, type=int)
        size = request.args.get('size', default=20, type=int)
        if not tug_name:
            return jsonify({'data': [], 'total': 0, 'last_page': 1})
            
        data, total = model.get_tug_records(tug_name, page=page, size=size)
        last_page = max(1, (total + size - 1) // size) if size else 1
        return jsonify({
            'data': data,
            'total': total,
            'last_page': last_page
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@bp.route('/api/module/TGU01/latest_rob', methods=['GET'])
@login_required
def get_latest_rob_route():
    try:
        tug_name = (request.args.get('tug') or request.args.get('tug_name') or '').strip()
        if not tug_name:
            return jsonify({})
        res = model.get_latest_rob(tug_name)
        return jsonify(res)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@bp.route('/api/module/TGU01/save', methods=['POST'])
@login_required
def save_record_route():
    try:
        perms = get_perms()
        if not perms.get('can_add') and not perms.get('can_edit'):
            return jsonify({'error': 'Permission denied'}), 403
            
        data = request.json or {}
        if not data.get('tug_name') or not data.get('entry_date'):
            return jsonify({'error': 'Tug Name and Date are required'}), 400
            
        if not data.get('created_by'):
            data['created_by'] = session.get('username') or 'admin'
            
        saved_id = model.save_record(data)
        return jsonify({'success': True, 'id': saved_id})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@bp.route('/api/module/TGU01/delete', methods=['POST'])
@login_required
def delete_record_route():
    try:
        perms = get_perms()
        if not perms.get('can_delete'):
            return jsonify({'error': 'Permission denied'}), 403
            
        payload = request.json or {}
        ids = payload.get('ids')
        if not ids:
            rec_id = payload.get('id')
            if rec_id:
                ids = [rec_id]
        if not ids:
            return jsonify({'error': 'Missing ID'}), 400
            
        for rid in ids:
            model.delete_record(rid)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def format_hh_mm(val):
    if val is None:
        return '00:00'
    s = str(val).strip()
    if not s or s in ('0', '0.0', '0.00', '0:00', '00:00'):
        return '00:00'
    if ':' in s:
        parts = s.split(':')
        h = int(parts[0]) if parts[0] else 0
        m = int(parts[1]) if len(parts) > 1 and parts[1] else 0
        return f"{h:02d}:{m:02d}"
    try:
        f = float(s)
        sign = '-' if f < 0 else ''
        abs_f = abs(f)
        h = int(abs_f)
        m = round((abs_f - h) * 60)
        if m == 60:
            h += 1
            m = 0
        return f"{sign}{h:02d}:{m:02d}"
    except Exception:
        return s

@bp.route('/api/module/TGU01/export', methods=['GET'])
@login_required
def export_excel():
    try:
        import openpyxl
        tug_name = (request.args.get('tug') or request.args.get('tug_name') or '').strip()
        data, _ = model.get_tug_records(tug_name, page=None, size=None) if tug_name else ([], 0)
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = tug_name[:31] if tug_name else "Tug Utilization"
        
        headers = [
            "Date", "HSD Pre ROB (LTR)", "Diesel Recd (LTR)", "M/E Hours (HH:MM)", "DG Hours (HH:MM)", 
            "M/E Cons (LTR)", "DG Cons (LTR)", "Total Hours (HH:MM)", "Total Cons (LTR)", "HSD Today ROB (LTR)",
            "FW Pre ROB (MT)", "FW Recd (MT)", "FW Cons (MT)", "Today ROB (MT)",
            "Port Initial", "Port Final", "Port Diff (HH:MM)",
            "Stbd Initial", "Stbd Final", "Stbd Diff (HH:MM)",
            "Shore Power (HH:MM)", "Remarks", "Created By"
        ]
        ws.append(headers)
        
        for r in data:
            ws.append([
                r.get('entry_date', ''),
                r.get('hsd_opening_rob', '0'),
                r.get('hsd_received', '0'),
                format_hh_mm(r.get('hsd_me_hours', '0')),
                format_hh_mm(r.get('hsd_dg_hours', '0')),
                r.get('hsd_me_consumption', '0'),
                r.get('hsd_dg_consumption', '0'),
                format_hh_mm(r.get('hsd_total_hours', '0')),
                r.get('hsd_total_consumption', '0'),
                r.get('hsd_closing_rob', '0'),
                r.get('fw_opening_rob', '0'),
                r.get('fw_received', '0'),
                r.get('fw_consumption', '0'),
                r.get('fw_closing_rob', '0'),
                r.get('port_initial', '0'),
                r.get('port_final', '0'),
                format_hh_mm(r.get('port_diff', '0')),
                r.get('stbd_initial', '0'),
                r.get('stbd_final', '0'),
                format_hh_mm(r.get('stbd_diff', '0')),
                format_hh_mm(r.get('shore_power_hours', '0')),
                r.get('remarks', ''),
                r.get('created_by', '')
            ])
            
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        filename = f"Tug_Utilization_{tug_name or 'All'}.xlsx".replace(" ", "_")
        return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        return jsonify({'error': str(e)}), 500

