from database import get_db, get_cursor

TABLE = 'tug_master'

def get_all():
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute(f"SELECT tug_name FROM {TABLE} WHERE tug_name IS NOT NULL AND tug_name != '' ORDER BY tug_name ASC")
    rows = cur.fetchall()
    conn.close()
    return [r['tug_name'] for r in rows]

def get_data(page=1, size=20):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute(f'SELECT COUNT(*) FROM {TABLE}')
    total = cur.fetchone()['count']
    cur.execute(f'SELECT * FROM {TABLE} ORDER BY id DESC LIMIT %s OFFSET %s', (size, (page-1)*size))
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows], total

def save_data(data):
    conn = get_db()
    cur = get_cursor(conn)
    row_id = data.get('id')
    tug_name = data.get('tug_name', '')
    company_name = data.get('company_name', '')
    tug_capacity = data.get('tug_capacity', '')
    bhp = data.get('bhp', '')
    contact_no = data.get('contact_no', '')
    
    if row_id:
        cur.execute(f"UPDATE {TABLE} SET tug_name=%s, company_name=%s, tug_capacity=%s, bhp=%s, contact_no=%s WHERE id=%s", 
                    [tug_name, company_name, tug_capacity, bhp, contact_no, row_id])
    else:
        cur.execute(f"INSERT INTO {TABLE} (tug_name, company_name, tug_capacity, bhp, contact_no) VALUES (%s, %s, %s, %s, %s) RETURNING id", 
                    [tug_name, company_name, tug_capacity, bhp, contact_no])
        row_id = cur.fetchone()['id']
    conn.commit()
    conn.close()
    return row_id

def delete_data(row_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute(f'DELETE FROM {TABLE} WHERE id=%s', (row_id,))
    conn.commit()
    conn.close()
