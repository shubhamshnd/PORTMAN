from database import get_db, get_cursor

TABLE = 'activity_master'

def get_all():
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute(f"SELECT name FROM {TABLE} WHERE name IS NOT NULL AND name != '' ORDER BY name ASC")
    rows = cur.fetchall()
    conn.close()
    return [r['name'] for r in rows]

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
    name = data.get('name', '')
    
    if row_id:
        cur.execute(f"UPDATE {TABLE} SET name=%s WHERE id=%s", 
                    [name, row_id])
    else:
        cur.execute(f"INSERT INTO {TABLE} (name) VALUES (%s) RETURNING id", 
                    [name])
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
