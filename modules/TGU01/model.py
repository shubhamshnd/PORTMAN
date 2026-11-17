from database import get_db, get_cursor

COLS = [
    'tug_name', 'entry_date',
    'hsd_opening_rob', 'hsd_received', 'hsd_me_hours', 'hsd_dg_hours',
    'hsd_me_consumption', 'hsd_dg_consumption', 'hsd_total_hours',
    'hsd_total_consumption', 'hsd_closing_rob',
    'fw_opening_rob', 'fw_received', 'fw_consumption', 'fw_closing_rob',
    'port_initial', 'port_final', 'port_diff',
    'stbd_initial', 'stbd_final', 'stbd_diff',
    'shore_power_hours', 'remarks', 'created_by'
]

def get_all_tugs():
    """Fetch all tugs from tug_master for the button view"""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT id, tug_name, company_name, tug_capacity, bhp 
        FROM tug_master 
        ORDER BY tug_name
    ''')
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def get_tug_records(tug_name, page=1, size=20):
    """Fetch daily logs for a specific tug with SQL pagination"""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT COUNT(*) as count FROM tgu01_daily_log WHERE tug_name = %s', (tug_name,))
    total = cur.fetchone()['count']
    
    if page and size and size > 0:
        offset = (page - 1) * size
        cur.execute('''
            SELECT * FROM tgu01_daily_log 
            WHERE tug_name = %s 
            ORDER BY entry_date DESC, id DESC 
            LIMIT %s OFFSET %s
        ''', (tug_name, size, offset))
    else:
        cur.execute('''
            SELECT * FROM tgu01_daily_log 
            WHERE tug_name = %s 
            ORDER BY entry_date DESC, id DESC
        ''', (tug_name,))
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows], total

def get_latest_rob(tug_name):
    """Fetch latest closing balances for auto-filling new day's opening balances"""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT entry_date, hsd_closing_rob, fw_closing_rob, port_final, stbd_final
        FROM tgu01_daily_log
        WHERE tug_name = %s
        ORDER BY entry_date DESC, id DESC
        LIMIT 1
    ''', (tug_name,))
    row = cur.fetchone()
    conn.close()
    if row:
        return dict(row)
    return {
        'hsd_closing_rob': '0',
        'fw_closing_rob': '0',
        'port_final': '0',
        'stbd_final': '0'
    }

def save_record(data):
    """Save or update a single daily log row"""
    conn = get_db()
    cur = get_cursor(conn)
    row_id = data.get('id')
    
    vals = [str(data.get(c) if data.get(c) is not None else '0').strip() for c in COLS]
    # Keep remarks and created_by as clean strings
    vals[COLS.index('remarks')] = data.get('remarks') or ''
    vals[COLS.index('created_by')] = data.get('created_by') or 'admin'
    vals[COLS.index('tug_name')] = data.get('tug_name') or ''
    vals[COLS.index('entry_date')] = data.get('entry_date') or ''
    
    if row_id and str(row_id).isdigit():
        set_clause = ', '.join([f"{c}=%s" for c in COLS])
        cur.execute(f"UPDATE tgu01_daily_log SET {set_clause} WHERE id=%s RETURNING id", (*vals, int(row_id)))
        res = cur.fetchone()
        saved_id = res['id'] if res else row_id
    else:
        cols_str = ', '.join(COLS)
        placeholders = ', '.join(['%s'] * len(COLS))
        cur.execute(f'''
            INSERT INTO tgu01_daily_log ({cols_str}) 
            VALUES ({placeholders}) 
            ON CONFLICT (tug_name, entry_date) DO UPDATE SET
            {', '.join([f"{c}=EXCLUDED.{c}" for c in COLS if c not in ('tug_name', 'entry_date')])}
            RETURNING id
        ''', vals)
        res = cur.fetchone()
        saved_id = res['id'] if res else None
        
    conn.commit()
    conn.close()
    return saved_id

def delete_record(record_id):
    """Delete a single daily log entry"""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("DELETE FROM tgu01_daily_log WHERE id = %s", (record_id,))
    conn.commit()
    conn.close()

