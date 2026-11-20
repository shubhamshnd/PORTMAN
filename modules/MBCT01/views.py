# pyrefly: ignore [missing-import]
from flask import render_template, request, jsonify, session, Blueprint

bp = Blueprint('MBCT01', __name__, template_folder='.')
from database import get_db, get_cursor
from .model import (get_anchorage, save_anchorage, delete_anchorage, 
                    get_port_activity, save_port_activity, delete_port_activity,
                    get_tug_activity, save_tug_activity, save_tug_activity_batch, delete_tug_activity)
from database import get_user_permissions as get_permissions
from collections import defaultdict
import json

@bp.route('/module/MBCT01/')
def index():
    if 'username' not in session:
        return 'Unauthorized', 401
    
    if session.get('is_admin'):
        perms = {'can_read': 1, 'can_add': 1, 'can_edit': 1, 'can_delete': 1}
    else:
        perms = get_permissions(session.get('user_id'), 'MBCT01')
    if not perms or not perms.get('can_read'):
        return 'Forbidden', 403

    return render_template('mbct01.html', permissions=perms)

# For the main grid (fetching headers only, similar to MBC01)
@bp.route('/api/module/MBCT01/data', methods=['GET', 'POST'])
def get_data():
    try:
        conn = get_db()
        cur = get_cursor(conn)
        cur.execute('''SELECT h.id as mbct01_id, m.id, m.operation_type, m.doc_num, m.doc_series, 
                       m.doc_date, m.mbc_name, m.load_port, 
                       m.cargo_type, m.cargo_name, m.bl_quantity, m.quantity_uom, m.doc_status, 
                       h.created_by 
                       FROM mbct01_header h
                       JOIN mbc_header m ON h.mbc_id = m.id 
                       ORDER BY h.id DESC''')
        res = cur.fetchall()
        for r in res:
            cb = r.get('created_by')
            if isinstance(cb, dict):
                r['created_by'] = cb.get('username') or str(cb.get('user_id') or '')
            elif isinstance(cb, str) and cb.strip().startswith('{'):
                try:
                    parsed = json.loads(cb)
                    if isinstance(parsed, dict):
                        r['created_by'] = parsed.get('username') or str(parsed.get('user_id') or cb)
                except Exception:
                    pass
        return jsonify({"data": res, "total": len(res)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@bp.route('/api/module/MBCT01/available_docs', methods=['GET'])
def get_available_docs():
    try:
        conn = get_db()
        cur = get_cursor(conn)
        cur.execute('''SELECT m.id, m.doc_num, m.doc_date, m.mbc_name, m.operation_type, m.load_port, m.cargo_name, m.bl_quantity, m.doc_status
                       FROM mbc_header m
                       WHERE NOT EXISTS (SELECT 1 FROM mbct01_header h WHERE h.mbc_id = m.id)
                       ORDER BY m.doc_num DESC''')
        res = cur.fetchall()
        return jsonify(res)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@bp.route('/api/module/MBCT01/save', methods=['POST'])
def save_header():
    try:
        data = request.json
        mbc_id = data.get('mbc_id')
        if not mbc_id:
            return jsonify({"error": "No mbc_id provided"}), 400
            
        conn = get_db()
        cur = get_cursor(conn)

        user_id = session.get('user_id')
        username = session.get('username') or data.get('created_by') or 'admin'
        creator_info = {
            "user_id": user_id,
            "username": username
        }
        creator_json = json.dumps(creator_info)

        cur.execute("""
            INSERT INTO mbct01_header (mbc_id, created_by) 
            VALUES (%s, %s::jsonb) 
            ON CONFLICT (mbc_id) DO UPDATE SET 
                created_by = COALESCE(mbct01_header.created_by, EXCLUDED.created_by)
            RETURNING id, created_by
        """, (mbc_id, creator_json))
        res = cur.fetchone()
        new_id = res['id']
        conn.commit()
        
        return jsonify({"success": True, "id": new_id, "created_by": username})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@bp.route('/api/module/MBCT01/delete', methods=['POST'])
def delete_header():
    try:
        data = request.json
        mbc_id = data.get('id')
        if not mbc_id:
            return jsonify({"error": "No id provided"}), 400
            
        conn = get_db()
        cur = get_cursor(conn)
        cur.execute("DELETE FROM mbct01_header WHERE mbc_id=%s", (mbc_id,))
        conn.commit()
        
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# Anchorage
@bp.route('/api/module/MBCT01/anchorage/<int:mbc_id>', methods=['GET'])
def fetch_anchorage(mbc_id):
    rows = get_anchorage(mbc_id)
    for r in rows:
        for f in ['eta_gull', 'nor_tendered_green_buoy', 'anchor_dropped_gull', 'anchor_away_gull', 
                  'passing_gull_island', 'passing_high_tension', 'passing_yellow_crane']:
            if r.get(f): r[f] = r[f].strftime('%Y-%m-%dT%H:%M') if hasattr(r[f], 'strftime') else r[f]
    return jsonify(rows)

@bp.route('/api/module/MBCT01/anchorage/save', methods=['POST'])
def save_anchorage_route():
    try:
        data = request.json
        for f in ['eta_gull', 'nor_tendered_green_buoy', 'anchor_dropped_gull', 'anchor_away_gull', 
                  'passing_gull_island', 'passing_high_tension', 'passing_yellow_crane']:
            if data.get(f) == '': data[f] = None
        id = save_anchorage(data)
        return jsonify({"success": True, "id": id})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@bp.route('/api/module/MBCT01/anchorage/delete', methods=['POST'])
def delete_anchorage_route():
    try:
        delete_anchorage(request.json['id'])
        return jsonify({"success": True})
    except:
        return jsonify({"success": False})

# Port Activity
@bp.route('/api/module/MBCT01/port_activity/<int:mbc_id>', methods=['GET'])
def fetch_port_activity(mbc_id):
    rows = get_port_activity(mbc_id)
    for r in rows:
        for f in ['first_line_ashore_cast_off', 'all_made_fast_cast_off', 'shore_supply_starts', 'shore_supply_ends']:
            if r.get(f): r[f] = r[f].strftime('%Y-%m-%dT%H:%M') if hasattr(r[f], 'strftime') else r[f]
    return jsonify(rows)

@bp.route('/api/module/MBCT01/port_activity/save', methods=['POST'])
def save_port_activity_route():
    try:
        data = request.json
        for f in ['first_line_ashore_cast_off', 'all_made_fast_cast_off', 'shore_supply_starts', 'shore_supply_ends']:
            if data.get(f) == '': data[f] = None
        for f in ['activity_id', 'from_berth_anchorage_id', 'to_berth_anchorage_id']:
            if data.get(f) == '': data[f] = None
        id = save_port_activity(data)
        return jsonify({"success": True, "id": id})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@bp.route('/api/module/MBCT01/port_activity/delete', methods=['POST'])
def delete_port_activity_route():
    try:
        delete_port_activity(request.json['id'])
        return jsonify({"success": True})
    except:
        return jsonify({"success": False})

# Tug Activity
@bp.route('/api/module/MBCT01/tug_activity/<int:mbc_id>', methods=['GET'])
def fetch_tug_activity(mbc_id):
    # The UI passes mbcId, but the table uses port_activity_id. 
    # For now, we will query tug_activity joined with port_activity by mbc_id
    rows = get_tug_activity(mbc_id)
    for r in rows:
        for f in ['tug_made_fast', 'tug_cast_off', 'engine_start', 'engine_stop']:
            if r.get(f): r[f] = r[f].strftime('%Y-%m-%dT%H:%M') if hasattr(r[f], 'strftime') else r[f]
        if r.get('t_running_hours') is not None:
            r['t_running_hours'] = str(r['t_running_hours'])
        if r.get('fuel_used') is not None:
            r['fuel_used'] = str(r['fuel_used'])
    return jsonify(rows)

@bp.route('/api/module/MBCT01/tug_activity/save', methods=['POST'])
def save_tug_activity_route():
    try:
        data = request.json
        if isinstance(data, list):
            for d in data:
                for f in ['tug_made_fast', 'tug_cast_off', 'engine_start', 'engine_stop']:
                    if d.get(f) == '': d[f] = None
            ids = save_tug_activity_batch(data)
            return jsonify({"success": True, "ids": ids})
        elif isinstance(data, dict) and 'rows' in data and isinstance(data['rows'], list):
            for d in data['rows']:
                for f in ['tug_made_fast', 'tug_cast_off', 'engine_start', 'engine_stop']:
                    if d.get(f) == '': d[f] = None
                if not d.get('mbc_id'):
                    d['mbc_id'] = data.get('mbc_id')
            ids = save_tug_activity_batch(data['rows'])
            return jsonify({"success": True, "ids": ids})
        else:
            for f in ['tug_made_fast', 'tug_cast_off', 'engine_start', 'engine_stop']:
                if data.get(f) == '': data[f] = None
            id = save_tug_activity(data)
            return jsonify({"success": True, "id": id})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@bp.route('/api/module/MBCT01/tug_activity/delete', methods=['POST'])
def delete_tug_activity_route():
    try:
        rid = request.json.get('id')
        if not rid or rid == "null":
            return jsonify({"success": True})
        delete_tug_activity(rid)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500
