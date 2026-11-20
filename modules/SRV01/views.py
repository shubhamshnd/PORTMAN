from flask import Blueprint, render_template, request, jsonify, session, redirect, url_for
from . import model
from database import get_user_permissions, get_db, get_cursor, get_module_config

bp = Blueprint('SRV01', __name__, template_folder='.')
MODULE_CODE = 'SRV01'
MODULE_INFO = {'code': 'SRV01', 'name': 'Service Recording'}


def _get_perms():
    if session.get('is_admin'):
        return {'can_read': 1, 'can_add': 1, 'can_edit': 1, 'can_delete': 1}
    return get_user_permissions(session.get('user_id'), 'SRV01')


@bp.route('/module/SRV01/')
def index():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    perms = _get_perms()
    if not perms or not perms.get('can_read'):
        return render_template('no_access.html'), 403

    page = int(request.args.get('page', 1))
    source_type = request.args.get('source_type')
    service_type_id = request.args.get('service_type_id', type=int)
    billed_status = request.args.get('billed_status')

    data, total = model.get_service_records(
        page, source_type=source_type,
        service_type_id=service_type_id,
        billed_status=billed_status
    )

    # Check if user is approver
    config = get_module_config('SRV01')
    user_id = session.get('user_id')
    is_approver = str(config.get('approver_id', '')) == str(user_id) or session.get('is_admin')

    return render_template('srv01.html',
                         data=data,
                         page=page,
                         last_page=(total + 19) // 20,
                         perms=perms,
                         is_approver=is_approver,
                         username=session.get('username'))


@bp.route('/api/module/SRV01/data')
def get_data():
    """Get paginated service records (AJAX)"""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401

    perms = _get_perms()
    if not perms or not perms.get('can_read'):
        return jsonify({'error': 'No read permission'}), 403

    page = request.args.get('page', 1, type=int)
    source_type = request.args.get('source_type')
    service_type_id = request.args.get('service_type_id', type=int)
    billed_status = request.args.get('billed_status')

    data, total = model.get_service_records(
        page, source_type=source_type,
        service_type_id=service_type_id,
        billed_status=billed_status
    )
    return jsonify({'data': data, 'total': total})


@bp.route('/api/module/SRV01/service-types')
def get_service_types():
    """Get service types that have custom fields configured"""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401

    data = model.get_service_types_with_fields()
    return jsonify({'data': data})


@bp.route('/api/module/SRV01/fields/<int:service_type_id>')
def get_fields(service_type_id):
    """Get field definitions for a service type"""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401

    fields = model.get_field_definitions(service_type_id)
    return jsonify({'data': fields})


@bp.route('/api/module/SRV01/record/<int:record_id>')
def get_record(record_id):
    """Get a service record with its field values"""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401

    perms = _get_perms()
    if not perms or not perms.get('can_read'):
        return jsonify({'error': 'No read permission'}), 403

    header, values = model.get_service_record_by_id(record_id)
    if not header:
        return jsonify({'error': 'Record not found'}), 404

    return jsonify({'header': header, 'values': values})


@bp.route('/api/module/SRV01/save', methods=['POST'])
def save():
    """Save service record + field values"""
    if 'user_id' not in session:
        return jsonify({'success': False, 'error': 'Not logged in'}), 401

    perms = _get_perms()
    data = request.json or {}

    if data.get('id') and not perms.get('can_edit'):
        return jsonify({'success': False, 'error': 'No edit permission'}), 403
    if data.get('id') and model.is_billed(data['id']):
        return jsonify({'success': False, 'error': 'Record is billed - cannot edit'}), 400
    if not data.get('id') and not perms.get('can_add'):
        return jsonify({'success': False, 'error': 'No add permission'}), 403

    # Build header data
    header_data = {
        'id': data.get('id'),
        'service_type_id': data.get('service_type_id'),
        'source_type': data.get('source_type'),
        'source_id': data.get('source_id'),
        'source_display': data.get('source_display'),
        'ref_source_type': data.get('ref_source_type'),
        'ref_source_id': data.get('ref_source_id'),
        'ref_source_display': data.get('ref_source_display'),
        'record_date': data.get('record_date'),
        'billable_quantity': data.get('billable_quantity'),
        'billable_uom': data.get('billable_uom'),
        'remarks': data.get('remarks'),
        'created_by': session.get('username')
    }

    # Set doc_status based on approval config
    config = get_module_config('SRV01')
    user_id = session.get('user_id')
    is_approver = str(config.get('approver_id', '')) == str(user_id)
    is_admin = session.get('is_admin')

    if not data.get('id'):  # New record
        if is_approver or is_admin:
            header_data['doc_status'] = 'Approved'
        elif config.get('approval_add'):
            header_data['doc_status'] = 'Pending'
        else:
            header_data['doc_status'] = 'Approved'
    else:
        # Keep existing status on edit (unless explicitly changed)
        header_data['doc_status'] = data.get('doc_status', 'Pending')

    field_values = data.get('field_values', [])

    record_id, record_number = model.save_service_record(header_data, field_values)
    return jsonify({'success': True, 'id': record_id, 'record_number': record_number})


@bp.route('/api/module/SRV01/delete', methods=['POST'])
def delete():
    """Delete a service record"""
    if 'user_id' not in session:
        return jsonify({'success': False, 'error': 'Not logged in'}), 401

    perms = _get_perms()
    if not perms.get('can_delete'):
        return jsonify({'success': False, 'error': 'No delete permission'}), 403

    record_id = (request.json or {}).get('id')
    if model.is_billed(record_id):
        return jsonify({'success': False, 'error': 'Record is billed - cannot delete'}), 400
    model.delete_service_record(record_id)
    return jsonify({'success': True})


@bp.route('/api/module/SRV01/approve', methods=['POST'])
def approve():
    """Approve a service record - only approver or admin"""
    if 'user_id' not in session:
        return jsonify({'success': False, 'error': 'Not logged in'})

    config = get_module_config('SRV01')
    user_id = session.get('user_id')
    is_approver = str(config.get('approver_id', '')) == str(user_id)
    is_admin = session.get('is_admin')

    if not is_approver and not is_admin:
        return jsonify({'success': False, 'error': 'Only approver or admin can approve'})

    record_id = request.json.get('id')
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''UPDATE service_records
        SET doc_status='Approved', approved_by=%s, approved_date=%s
        WHERE id=%s''',
        [session.get('username'), __import__('datetime').datetime.now().strftime('%Y-%m-%d'), record_id])
    conn.commit()
    conn.close()

    return jsonify({'success': True})


@bp.route('/api/module/SRV01/customer-options/<customer_type>')
def get_customer_options(customer_type):
    """Get customer or agent list for primary source selection"""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401

    conn = get_db()
    cur = get_cursor(conn)

    if customer_type == 'Customer':
        cur.execute("SELECT id, name FROM vessel_customers ORDER BY name")
    elif customer_type == 'Agent':
        cur.execute("SELECT id, name FROM vessel_agents WHERE is_active = 1 ORDER BY name")
    else:
        conn.close()
        return jsonify({'data': []})

    rows = cur.fetchall()
    conn.close()
    return jsonify({'data': [dict(r) for r in rows]})


@bp.route('/api/module/SRV01/source-options/<source_type>')
def get_source_options(source_type):
    """Get optional VCN/MBC reference options"""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401

    conn = get_db()
    cur = get_cursor(conn)

    if source_type == 'VCN':
        cur.execute('''
            SELECT h.id, h.vcn_doc_num, h.vessel_name, a.anchorage_arrival
            FROM vcn_header h
            LEFT JOIN vcn_anchorage a ON h.id = a.vcn_id
            ORDER BY h.id DESC
        ''')
        rows = cur.fetchall()
        conn.close()
        result = []
        for r in rows:
            r = dict(r)
            anchored = r.get('anchorage_arrival', '') or ''
            if anchored:
                anchored = str(anchored)[:16].replace('T', ' ')
            display = f"{r['vcn_doc_num']} / {r['vessel_name']} / {anchored}"
            result.append({'id': r['id'], 'display': display})
        return jsonify({'data': result})

    elif source_type == 'MBC':
        cur.execute('''
            SELECT id, doc_num, mbc_name, doc_date
            FROM mbc_header ORDER BY id DESC
        ''')
        rows = cur.fetchall()
        conn.close()
        result = []
        for r in rows:
            r = dict(r)
            display = f"{r['doc_num']} / {r['mbc_name']} / {r.get('doc_date', '')}"
            result.append({'id': r['id'], 'display': display})
        return jsonify({'data': result})

    conn.close()
    return jsonify({'data': []})
