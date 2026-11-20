from database import get_db, get_cursor
from datetime import datetime


def _nv(value):
    """Normalize empty strings to None so they bind as SQL NULL.

    Postgres rejects '' for timestamp/numeric columns, so blank form
    fields must be sent as NULL instead of an empty string.
    """
    if isinstance(value, str) and value.strip() == '':
        return None
    return value


def get_next_doc_num(doc_series):
    """Get next doc number for given series"""
    conn = get_db()
    cur = get_cursor(conn)
    prefix = doc_series.replace('-', '') if doc_series else 'MBC'
    cur.execute(
        "SELECT MAX(CAST(SUBSTR(doc_num, LENGTH(%s) + 1) AS INTEGER)) FROM mbc_header WHERE doc_num LIKE %s",
        [prefix, f"{prefix}%"]
    )
    result = cur.fetchone()['max']
    conn.close()
    next_num = (result or 0) + 1
    return f"{prefix}{next_num:04d}"


def get_data(page=1, size=20, filters=None):
    conn = get_db()
    cur = get_cursor(conn)

    allowed = {'operation_type','doc_num','mbc_name','cargo_type','doc_status',
               'doc_date','bl_quantity','doc_series'}
    where_clauses, params = [], []
    for f in (filters or []):
        field = f.get('field', '')
        if field not in allowed:
            continue
        ftype = f.get('type')
        if ftype == 'contains' and f.get('value'):
            where_clauses.append(f"{field} ILIKE %s")
            params.append(f"%{f['value']}%")
        elif ftype == 'multi' and f.get('values'):
            ph = ','.join(['%s'] * len(f['values']))
            where_clauses.append(f"{field} IN ({ph})")
            params.extend(f['values'])
        elif ftype == 'range':
            if f.get('from'):
                where_clauses.append(f"{field} >= %s")
                params.append(f['from'])
            if f.get('to'):
                where_clauses.append(f"{field} <= %s")
                params.append(f['to'])

    where_sql = f"WHERE {' AND '.join(where_clauses)}" if where_clauses else ""
    try:
        cur.execute(f'SELECT COUNT(*) FROM mbc_header {where_sql}', params)
        total = cur.fetchone()['count']
        cur.execute(f'''
            SELECT mh.*,
                   (SELECT COUNT(*) FROM mbc_customer_details cd WHERE cd.mbc_id = mh.id) AS _customer_count,
                   (SELECT pd.id FROM mbc_proof_documents pd WHERE pd.mbc_id = mh.id ORDER BY pd.id DESC LIMIT 1) AS _proof_doc_id,
                   (SELECT pd.original_filename FROM mbc_proof_documents pd WHERE pd.mbc_id = mh.id ORDER BY pd.id DESC LIMIT 1) AS _proof_filename
            FROM mbc_header mh
            {where_sql}
            ORDER BY mh.id DESC LIMIT %s OFFSET %s
        ''', params + [size, (page - 1) * size])
        rows = cur.fetchall()
        result = []
        for r in rows:
            d = dict(r)
            d['_eligible'] = bool(
                d.get('operation_type') and
                d.get('mbc_name') and
                d.get('cargo_name') and
                d.get('bl_quantity') and
                int(d.get('_customer_count') or 0) >= 1
            )
            result.append(d)
        return result, total
    finally:
        conn.close()


def save_header(data):
    conn = get_db()
    cur = get_cursor(conn)
    row_id = data.get('id')

    # Auto-set doc_date if not provided
    if not data.get('doc_date'):
        data['doc_date'] = datetime.now().strftime('%Y-%m-%d')

    # Strip client-side computed fields (prefixed with _) before writing to DB
    data = {k: v for k, v in data.items() if not k.startswith('_')}

    if row_id:
        cols = [k for k in data if k not in ['id', 'doc_num']]
        cur.execute(f"UPDATE mbc_header SET {', '.join([f'{c}=%s' for c in cols])} WHERE id=%s",
                   [data[c] for c in cols] + [row_id])
    else:
        data['doc_num'] = get_next_doc_num(data.get('doc_series', ''))
        cols = [k for k in data if k != 'id']
        cur.execute(f"INSERT INTO mbc_header ({', '.join(cols)}) VALUES ({', '.join(['%s']*len(cols))}) RETURNING id",
                   [data[c] for c in cols])
        row_id = cur.fetchone()['id']

    conn.commit()
    conn.close()
    return row_id, data.get('doc_num')


def delete_header(row_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('DELETE FROM mbc_header WHERE id=%s', (row_id,))
    conn.commit()
    conn.close()


# Load Port Lines sub-table operations
def get_load_port_lines(mbc_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT * FROM mbc_load_port_lines WHERE mbc_id=%s ORDER BY id DESC', (mbc_id,))
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def save_load_port_line(data):
    conn = get_db()
    cur = get_cursor(conn)

    if data.get('id'):
        cur.execute('''UPDATE mbc_load_port_lines SET
                      eta=%s, arrived_load_port=%s, alongside_berth=%s, loading_commenced=%s,
                      loading_completed=%s, cast_off_load_port=%s,
                      fwd_draft=%s, mid_draft=%s, aft_draft=%s
                      WHERE id=%s''',
                   [_nv(data.get('eta')), _nv(data.get('arrived_load_port')), _nv(data.get('alongside_berth')), _nv(data.get('loading_commenced')),
                    _nv(data.get('loading_completed')), _nv(data.get('cast_off_load_port')),
                    _nv(data.get('fwd_draft')), _nv(data.get('mid_draft')), _nv(data.get('aft_draft')), data['id']])
        row_id = data['id']
    else:
        cur.execute('''INSERT INTO mbc_load_port_lines
                      (mbc_id, eta, arrived_load_port, alongside_berth, loading_commenced, loading_completed, cast_off_load_port,
                       fwd_draft, mid_draft, aft_draft)
                      VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s) RETURNING id''',
                   [data['mbc_id'], _nv(data.get('eta')), _nv(data.get('arrived_load_port')), _nv(data.get('alongside_berth')),
                    _nv(data.get('loading_commenced')), _nv(data.get('loading_completed')), _nv(data.get('cast_off_load_port')),
                    _nv(data.get('fwd_draft')), _nv(data.get('mid_draft')), _nv(data.get('aft_draft'))])
        row_id = cur.fetchone()['id']
    conn.commit()
    conn.close()
    return row_id


def delete_load_port_line(row_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('DELETE FROM mbc_load_port_lines WHERE id=%s', (row_id,))
    conn.commit()
    conn.close()


# Discharge Port Lines sub-table operations
def get_discharge_port_lines(mbc_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT * FROM mbc_discharge_port_lines WHERE mbc_id=%s ORDER BY id DESC', (mbc_id,))
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def save_discharge_port_line(data):
    conn = get_db()
    cur = get_cursor(conn)

    if data.get('id'):
        cur.execute('''UPDATE mbc_discharge_port_lines SET
                      arrival_gull_island=%s, departure_gull_island=%s, arrived_yellow_crane=%s,
                      vessel_arrival_port=%s,
                      vessel_all_made_fast=%s, unloading_commenced=%s, cleaning_commenced=%s,
                      cleaning_completed=%s, unloading_completed=%s, vessel_cast_off=%s, sailed_out_load_port=%s,
                      vessel_unloaded_by=%s,
                      vessel_unloading_berth=%s, discharge_stop_shifting=%s, discharge_start_shifting=%s,
                      reached_load_port=%s
                      WHERE id=%s''',
                   [_nv(data.get('arrival_gull_island')), _nv(data.get('departure_gull_island')), _nv(data.get('arrived_yellow_crane')),
                    _nv(data.get('vessel_arrival_port')),
                    _nv(data.get('vessel_all_made_fast')), _nv(data.get('unloading_commenced')), _nv(data.get('cleaning_commenced')),
                    _nv(data.get('cleaning_completed')), _nv(data.get('unloading_completed')), _nv(data.get('vessel_cast_off')),
                    _nv(data.get('sailed_out_load_port')), _nv(data.get('vessel_unloaded_by')), _nv(data.get('vessel_unloading_berth')),
                    _nv(data.get('discharge_stop_shifting')), _nv(data.get('discharge_start_shifting')),
                    _nv(data.get('reached_load_port')), data['id']])
        row_id = data['id']
    else:
        cur.execute('''INSERT INTO mbc_discharge_port_lines
                      (mbc_id, arrival_gull_island, departure_gull_island, arrived_yellow_crane,
                       vessel_arrival_port,
                       vessel_all_made_fast, unloading_commenced, cleaning_commenced, cleaning_completed,
                       unloading_completed, vessel_cast_off, sailed_out_load_port, vessel_unloaded_by,
                       vessel_unloading_berth, discharge_stop_shifting, discharge_start_shifting,
                       reached_load_port)
                      VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s) RETURNING id''',
                   [data['mbc_id'], _nv(data.get('arrival_gull_island')), _nv(data.get('departure_gull_island')),
                    _nv(data.get('arrived_yellow_crane')), _nv(data.get('vessel_arrival_port')), _nv(data.get('vessel_all_made_fast')),
                    _nv(data.get('unloading_commenced')), _nv(data.get('cleaning_commenced')),
                    _nv(data.get('cleaning_completed')), _nv(data.get('unloading_completed')), _nv(data.get('vessel_cast_off')),
                    _nv(data.get('sailed_out_load_port')), _nv(data.get('vessel_unloaded_by')), _nv(data.get('vessel_unloading_berth')),
                    _nv(data.get('discharge_stop_shifting')), _nv(data.get('discharge_start_shifting')),
                    _nv(data.get('reached_load_port'))])
        row_id = cur.fetchone()['id']
    conn.commit()
    conn.close()
    return row_id


def delete_discharge_port_line(row_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('DELETE FROM mbc_discharge_port_lines WHERE id=%s', (row_id,))
    conn.commit()
    conn.close()


# Cleaning Details sub-table operations
def get_cleaning_details(mbc_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT * FROM mbc_cleaning_details WHERE mbc_id=%s ORDER BY id DESC', (mbc_id,))
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def save_cleaning_detail(data):
    conn = get_db()
    cur = get_cursor(conn)

    if data.get('id'):
        cur.execute('''UPDATE mbc_cleaning_details SET
                      payloader_name=%s, hmr_start=%s, hmr_end=%s,
                      diesel_start=%s, diesel_end=%s, start_time=%s, end_time=%s
                      WHERE id=%s''',
                   [data.get('payloader_name'), data.get('hmr_start'), data.get('hmr_end'),
                    data.get('diesel_start'), data.get('diesel_end'),
                    data.get('start_time'), data.get('end_time'), data['id']])
        row_id = data['id']
    else:
        cur.execute('''INSERT INTO mbc_cleaning_details
                      (mbc_id, payloader_name, hmr_start, hmr_end, diesel_start, diesel_end, start_time, end_time)
                      VALUES (%s, %s, %s, %s, %s, %s, %s, %s) RETURNING id''',
                   [data['mbc_id'], data.get('payloader_name'), data.get('hmr_start'), data.get('hmr_end'),
                    data.get('diesel_start'), data.get('diesel_end'),
                    data.get('start_time'), data.get('end_time')])
        row_id = cur.fetchone()['id']
    conn.commit()
    conn.close()
    return row_id


def delete_cleaning_detail(row_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('DELETE FROM mbc_cleaning_details WHERE id=%s', (row_id,))
    conn.commit()
    conn.close()


# Export Load Port Lines sub-table operations
def get_export_load_port_lines(mbc_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT * FROM mbc_export_load_port_lines WHERE mbc_id=%s ORDER BY id DESC', (mbc_id,))
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def save_export_load_port_line(data):
    conn = get_db()
    cur = get_cursor(conn)

    if data.get('id'):
        cur.execute('''UPDATE mbc_export_load_port_lines SET
                      arrived_at_port=%s, alongside_at_berth=%s, loading_commenced=%s,
                      loading_completed=%s, cast_off_from_berth=%s, sailed_out_from_port=%s,
                      eta_at_gull_island=%s, unloaded_by=%s, berth_master=%s,
                      fwd_draft=%s, mid_draft=%s, aft_draft=%s
                      WHERE id=%s''',
                   [data.get('arrived_at_port'), data.get('alongside_at_berth'), data.get('loading_commenced'),
                    data.get('loading_completed'), data.get('cast_off_from_berth'), data.get('sailed_out_from_port'),
                    data.get('eta_at_gull_island'), data.get('unloaded_by'), data.get('berth_master'),
                    data.get('fwd_draft'), data.get('mid_draft'), data.get('aft_draft'), data['id']])
        row_id = data['id']
    else:
        cur.execute('''INSERT INTO mbc_export_load_port_lines
                      (mbc_id, arrived_at_port, alongside_at_berth, loading_commenced, loading_completed,
                       cast_off_from_berth, sailed_out_from_port, eta_at_gull_island, unloaded_by, berth_master,
                       fwd_draft, mid_draft, aft_draft)
                      VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s) RETURNING id''',
                   [data['mbc_id'], data.get('arrived_at_port'), data.get('alongside_at_berth'),
                    data.get('loading_commenced'), data.get('loading_completed'),
                    data.get('cast_off_from_berth'), data.get('sailed_out_from_port'),
                    data.get('eta_at_gull_island'), data.get('unloaded_by'), data.get('berth_master'),
                    data.get('fwd_draft'), data.get('mid_draft'), data.get('aft_draft')])
        row_id = cur.fetchone()['id']
    conn.commit()
    conn.close()
    return row_id


def delete_export_load_port_line(row_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('DELETE FROM mbc_export_load_port_lines WHERE id=%s', (row_id,))
    conn.commit()
    conn.close()


# Customer Details sub-table operations
def _line_lock(cur, line):
    """Billing lock on one customer line: '', 'billed' or 'invoiced'.

    The line's own billed counters decide, because they are what FIN01 reads to
    offer cargo for billing. Cargo freed by Admin → Release Cargo keeps its old
    bill/invoice rows (the money is corrected by credit note) but its counters
    are zeroed — it is billable again, so its customer must be correctable too."""
    from modules.FIN01 import model as fin_model
    if not (line.get('is_billed') or float(line.get('billed_quantity') or 0) > 0
            or line.get('bill_id')):
        return ''
    return fin_model.source_lock_state(cur, 'MBC', line['mbc_id'], line_id=line['id']) or 'billed'


def get_customer_details(mbc_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT * FROM mbc_customer_details WHERE mbc_id=%s ORDER BY id DESC', (mbc_id,))
    rows = [dict(r) for r in cur.fetchall()]
    # ponytail: up to 2 queries per billed line; MBCs carry a handful of lines
    for r in rows:
        r['_lock'] = _line_lock(cur, r)
    conn.close()
    return rows


def save_customer_detail(data):
    conn = get_db()
    cur = get_cursor(conn)

    if data.get('id'):
        cur.execute('''UPDATE mbc_customer_details SET
                      customer_name=%s, cargo_name=%s, bill_of_coastal_goods_no=%s, quantity=%s, material_po=%s
                      WHERE id=%s''',
                   [data.get('customer_name'), data.get('cargo_name'),
                    data.get('bill_of_coastal_goods_no'),
                    data.get('quantity'), data.get('material_po'), data['id']])
        row_id = data['id']
    else:
        cur.execute('''INSERT INTO mbc_customer_details
                      (mbc_id, customer_name, cargo_name, bill_of_coastal_goods_no, quantity, material_po)
                      VALUES (%s, %s, %s, %s, %s, %s) RETURNING id''',
                   [data['mbc_id'], data.get('customer_name'), data.get('cargo_name'),
                    data.get('bill_of_coastal_goods_no'),
                    data.get('quantity'), data.get('material_po')])
        row_id = cur.fetchone()['id']
    conn.commit()
    conn.close()
    return row_id


def line_changes(line, new_customer, new_po):
    """What a Change Customer / PO request actually moves on one line.
    Returns (action, [descriptions]); action is '' when nothing differs."""
    changes, kinds = [], []
    if new_customer != (line['customer_name'] or ''):
        changes.append(f"Customer: {line['customer_name'] or '-'} → {new_customer}")
        kinds.append('Customer')
    if new_po != (line['material_po'] or ''):
        changes.append(f"PO: {line['material_po'] or '-'} → {new_po}")
        kinds.append('PO')
    return (' / '.join(kinds) + ' Changed') if kinds else '', changes


def change_line(row_id, new_customer, new_po, reason, username):
    """Correct the customer and/or Material PO on one unbilled customer line,
    even on an Approved MBC or one whose other lines are billed. The old/new
    values and the reason go to approval_log. Billing reads material_po at
    bill time, so the next bill and SAP posting pick up the new PO.

    Returns (ok, error_message)."""
    from modules.FIN01 import model as fin_model
    conn = get_db()
    cur = get_cursor(conn)
    try:
        cur.execute('''SELECT id, mbc_id, customer_name, material_po, is_billed, billed_quantity, bill_id
                       FROM mbc_customer_details WHERE id=%s FOR UPDATE''', (row_id,))
        line = cur.fetchone()
        if not line:
            return False, 'Customer line not found'
        state = _line_lock(cur, dict(line))
        if state:
            return False, fin_model.lock_message(state, 'customer line')
        action, changes = line_changes(line, new_customer, new_po)
        if not action:
            return False, 'Nothing changed: customer and PO are the same as the current ones'
        # FK to vessel_customers.name rejects unknown names.
        cur.execute('UPDATE mbc_customer_details SET customer_name=%s, material_po=%s WHERE id=%s',
                    (new_customer, new_po, row_id))
        cur.execute("""INSERT INTO approval_log (module_code, record_id, action, comment, actioned_by)
                       VALUES ('MBC01', %s, %s, %s, %s)""",
                    (line['mbc_id'], action,
                     f"Line #{row_id}: {'; '.join(changes)}. Reason: {reason}",
                     username))
        conn.commit()
        return True, None
    finally:
        conn.close()


def delete_customer_detail(row_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('DELETE FROM mbc_customer_details WHERE id=%s', (row_id,))
    conn.commit()
    conn.close()


# Approval functions
def get_doc_status(record_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT doc_status FROM mbc_header WHERE id=%s', (record_id,))
    row = cur.fetchone()
    conn.close()
    return row['doc_status'] if row else None


def get_approval_eligibility(mbc_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT operation_type, mbc_name, cargo_name, bl_quantity FROM mbc_header WHERE id=%s', (mbc_id,))
    header = cur.fetchone()
    if not header:
        conn.close()
        return {'eligible': False, 'missing': ['Record not found']}
    missing = []
    if not header['operation_type']:
        missing.append('Operation Type')
    if not header['mbc_name']:
        missing.append('MBC Name')
    if not header['cargo_name']:
        missing.append('Cargo Name')
    if not header['bl_quantity']:
        missing.append('BL Quantity')
    cur.execute('SELECT COUNT(*) as cnt FROM mbc_customer_details WHERE mbc_id=%s', (mbc_id,))
    if cur.fetchone()['cnt'] < 1:
        missing.append('Customer Details (minimum 1 entry required)')
    else:
        cur.execute(
            "SELECT COUNT(*) as cnt FROM mbc_customer_details WHERE mbc_id=%s AND (material_po IS NULL OR TRIM(material_po) = '')",
            (mbc_id,)
        )
        if cur.fetchone()['cnt'] > 0:
            missing.append('Material PO Number — required on every Customer Details row')
    cur.execute('SELECT COUNT(*) FROM mbc_proof_documents WHERE mbc_id=%s', (mbc_id,))
    if cur.fetchone()['count'] == 0:
        missing.append('Proof of Quantity — document must be uploaded (Proof of Qty column)')
    conn.close()
    return {'eligible': len(missing) == 0, 'missing': missing}


def approve_record(record_id, username):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("UPDATE mbc_header SET doc_status='Approved' WHERE id=%s", (record_id,))
    cur.execute("""INSERT INTO approval_log (module_code, record_id, action, comment, actioned_by)
                   VALUES ('MBC01', %s, 'Approved', NULL, %s)""", (record_id, username))
    conn.commit()
    conn.close()


def send_back_to_draft(record_id, comment, username):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("UPDATE mbc_header SET doc_status='Draft' WHERE id=%s", (record_id,))
    cur.execute("""INSERT INTO approval_log (module_code, record_id, action, comment, actioned_by)
                   VALUES ('MBC01', %s, 'Back to Draft', %s, %s)""", (record_id, comment, username))
    conn.commit()
    conn.close()


def get_approval_log(record_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("""SELECT action, comment, actioned_by,
                          to_char(actioned_at, 'DD-MM-YYYY HH24:MI') AS actioned_at
                   FROM approval_log WHERE module_code='MBC01' AND record_id=%s
                   ORDER BY actioned_at DESC""", (record_id,))
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]
