from database import get_db, get_cursor
import json

def get_available_vcns():
    """Fetch VCNs for header dropdown"""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT v.id, v.vcn_doc_num, v.vessel_name, v.doc_date, v.cargo_type, 
               COALESCE(string_agg(DISTINCT c.cargo_name, ', '), v.cargo_type, '') as cargo_name, 
               v.doc_status
        FROM vcn_header v
        LEFT JOIN vcn_cargo_declaration c ON c.vcn_id = v.id
        GROUP BY v.id
        ORDER BY v.id DESC
    ''')
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def get_barges_for_vcn(vcn_id):
    """Fetch barges nominated for this VCN (LDUD and VEX), with fallback to master barges"""
    conn = get_db()
    cur = get_cursor(conn)
    
    # 1. Check Import operations (LDUD)
    cur.execute('''
        SELECT DISTINCT b.barge_name, b.trip_number
        FROM ldud_barge_lines b
        JOIN ldud_header l ON b.ldud_id = l.id
        WHERE l.vcn_id = %s AND b.barge_name IS NOT NULL AND b.barge_name != ''
        ORDER BY b.barge_name, b.trip_number
    ''', (vcn_id,))
    rows = cur.fetchall()
    
    # 2. Check Export operations (VEX)
    if not rows:
        cur.execute('''
            SELECT DISTINCT b.barge_name, b.trip_number
            FROM vex_barge_lines b
            JOIN vex_header v ON b.vex_id = v.id
            WHERE v.vcn_id = %s AND b.barge_name IS NOT NULL AND b.barge_name != ''
            ORDER BY b.barge_name, b.trip_number
        ''', (vcn_id,))
        rows = cur.fetchall()

    result = []
    seen = set()
    for r in rows:
        trip = str(r['trip_number']).strip() if r['trip_number'] else ''
        label = f"{r['barge_name']} / {trip}" if trip else r['barge_name']
        if label not in seen:
            seen.add(label)
            result.append({
                'barge_name': r['barge_name'],
                'trip_number': trip,
                'label': label
            })
    
    # Also fetch all master barges so user has complete freedom if needed
    cur.execute('SELECT barge_name FROM barges ORDER BY barge_name')
    master_barges = [r['barge_name'] for r in cur.fetchall()]
    conn.close()
    
    return {
        'vcn_barges': result,
        'master_barges': master_barges
    }

def get_all_headers(page=None, size=None):
    conn = get_db()
    cur = get_cursor(conn)
    
    cur.execute('SELECT COUNT(*) as count FROM bgt01_header')
    total = cur.fetchone()['count']
    
    if page and size and size > 0:
        offset = (page - 1) * size
        cur.execute('''
            SELECT h.id, h.vcn_id, h.vcn_doc_num, h.vessel_name, h.barge_name, 
                   h.trip_number, h.doc_date, h.cargo_name, h.doc_status, 
                   h.created_by, h.created_at
            FROM bgt01_header h
            ORDER BY h.id DESC
            LIMIT %s OFFSET %s
        ''', (size, offset))
    else:
        cur.execute('''
            SELECT h.id, h.vcn_id, h.vcn_doc_num, h.vessel_name, h.barge_name, 
                   h.trip_number, h.doc_date, h.cargo_name, h.doc_status, 
                   h.created_by, h.created_at
            FROM bgt01_header h
            ORDER BY h.id DESC
        ''')
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows], total

def save_header(data):
    conn = get_db()
    cur = get_cursor(conn)
    row_id = data.get('id')
    
    cols = ['vcn_id', 'vcn_doc_num', 'vessel_name', 'barge_name', 'trip_number', 'doc_date', 'cargo_name', 'doc_status', 'created_by']
    vals = [data.get(c) for c in cols]
    
    if row_id:
        set_clause = ', '.join([f"{c}=%s" for c in cols])
        cur.execute(f"UPDATE bgt01_header SET {set_clause} WHERE id=%s RETURNING id", (*vals, row_id))
    else:
        cols_str = ', '.join(cols)
        placeholders = ', '.join(['%s'] * len(cols))
        cur.execute(f"INSERT INTO bgt01_header ({cols_str}) VALUES ({placeholders}) RETURNING id", vals)
    
    res = cur.fetchone()
    conn.commit()
    conn.close()
    return res['id'] if res else None

def delete_header(header_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("DELETE FROM bgt01_header WHERE id=%s", (header_id,))
    conn.commit()
    conn.close()

def get_activities(header_id):
    """Fetch movement activities and attached tug details for a given header"""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT id, header_id, movement_date, activity, 
               from_berth, to_berth, marine_officer, remark,
               COALESCE(tugs_data, '[]'::jsonb) as tugs
        FROM bgt01_movement_activity
        WHERE header_id = %s
        ORDER BY id ASC
    ''', (header_id,))
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def save_activity(data):
    """Save an activity and its nested tug timings directly in JSONB"""
    conn = get_db()
    cur = get_cursor(conn)
    activity_id = data.get('id')
    header_id = data.get('header_id')
    
    tugs = data.get('tugs', [])
    tugs_json = json.dumps(tugs) if isinstance(tugs, (list, dict)) else '[]'
    
    cols = ['header_id', 'movement_date', 'activity', 'from_berth', 'to_berth', 'marine_officer', 'remark', 'tugs_data']
    vals = [
        header_id,
        data.get('movement_date') or None,
        data.get('activity') or None,
        data.get('from_berth') or None,
        data.get('to_berth') or None,
        data.get('marine_officer') or None,
        data.get('remark') or None,
        tugs_json
    ]
    
    if activity_id and str(activity_id).isdigit():
        set_clause = ', '.join([f"{c}=%s" for c in cols[1:]])
        cur.execute(f"UPDATE bgt01_movement_activity SET {set_clause} WHERE id=%s RETURNING id", (*vals[1:], int(activity_id)))
        act_res = cur.fetchone()
        act_id = act_res['id']
    else:
        cols_str = ', '.join(cols)
        placeholders = ', '.join(['%s'] * len(cols))
        cur.execute(f"INSERT INTO bgt01_movement_activity ({cols_str}) VALUES ({placeholders}) RETURNING id", vals)
        act_res = cur.fetchone()
        act_id = act_res['id']
        
    conn.commit()
    conn.close()
    return act_id

def delete_activity(activity_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("DELETE FROM bgt01_movement_activity WHERE id=%s", (activity_id,))
    conn.commit()
    conn.close()

def get_masters():
    """Fetch all dropdown lists needed by BGT01"""
    conn = get_db()
    cur = get_cursor(conn)
    
    cur.execute("SELECT tug_name FROM tug_master ORDER BY tug_name")
    tugs = [r['tug_name'] for r in cur.fetchall()]
    
    cur.execute("SELECT name FROM marine_officer_master ORDER BY name")
    officers = [r['name'] for r in cur.fetchall()]
    
    cur.execute("SELECT berth_name FROM port_berth_master ORDER BY berth_name")
    berths = [r['berth_name'] for r in cur.fetchall()]
    
    cur.execute("SELECT name FROM activity_master ORDER BY name")
    activities = [r['name'] for r in cur.fetchall()]
    if not activities:
        activities = ['Berthing', 'Unberthing', 'Shifting', 'Other']
        
    conn.close()
    return {
        'tugs': tugs,
        'officers': officers,
        'berths': berths,
        'activities': activities
    }
