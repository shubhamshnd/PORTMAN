"""RP02 Revenue Register — shared constants + the backdated upload.

The register itself is built live from invoice_header / invoice_lines, so
anything invoiced before go-live has no rows there. Finance uploads those here
(CSV/Excel, same columns as the register grid) and the register reads both,
tagging the uploaded rows 'Backdated'. Mirrors the bill master upload pattern.
"""
from database import get_db, get_cursor
from .model import _blank, _norm_header, read_matrix

TABLE = 'rp02_revenue_backdated'

# (stored field, header) — one source of truth for the stored columns, the CSV
# template and the header map. Order matches the Revenue Register grid.
# Days/Bucket are NOT here on purpose: storing them would go stale overnight.
# See DERIVED_FIELDS below.
FIELDS = [
    ('invoice_no',     'Invoice No.'),
    ('group_type',     'Group/Non group'),
    ('revenue_type_1', 'Revenue Type 1'),
    ('revenue_type_2', 'Revenue Type 2'),
    ('cargo_volume',   'Cargo Volume'),
    ('invoice_date',   'Date'),
    ('cust_code',      'Cust Code'),
    ('customer_name',  'Customer Name'),
    ('gl_code',        'GL CODE'),
    ('grouping_label', 'Grouping'),
    ('qty',            'Qty - MT'),
    ('rate',           'Rate'),
    ('tax_category',   'Tax Category'),
    ('tax_rate',       'Tax Rate'),
    ('basic_value',    'Basic Value (Rs.)'),
    ('sgst',           'SGST (Rs.)'),
    ('cgst',           'CGST (Rs.)'),
    ('igst',           'IGST (Rs.)'),
    ('invoice_value',  'Invoice value (Rs.)'),
    ('gstin',          'GSTIN'),
    ('sap_doc_no',     'SAP Doc No'),
    ('sac_code',       'SAC CODE'),
    ('hsn_code',       'HSN CODE'),
    ('irn',            'IRN'),
    ('ack_date',       'Ack Date'),
    ('ack_no',         'Ack No'),
    ('barcode',        'BARCODE'),
    ('booking_status', 'BOOKING/PAYMENT STATUS'),
    ('tds_tcs',        'TDS/TCS'),
    ('net_receivable', 'NET RECEIVABLE'),
]

COLUMNS = [f for f, _ in FIELDS]
# Recomputed from the date on every read, never stored. The template still
# carries the two columns (the sheet finance keeps has them) and the parser
# ignores whatever is in them.
DERIVED_FIELDS = [('days', 'Days'), ('bucket', 'Bucket')]
DERIVED_HEADERS = [h for _, h in DERIVED_FIELDS]
TEMPLATE_HEADERS = ['Sr'] + [h for _, h in FIELDS] + DERIVED_HEADERS

# Every column is stored as text, verbatim. A backdated register is a
# spreadsheet: Cargo Volume holds YES/NO, amounts carry commas and stray notes,
# cells are blank. Nothing here is arithmetic, so nothing is coerced and no
# file is ever rejected over a cell's shape. The one exception is the date,
# normalised to YYYY-MM-DD where it parses, because the month key and the
# ageing columns read it.


def _norm(h):
    return _norm_header(h).rstrip('.')


HEADER_MAP = {_norm(h): f for f, h in FIELDS}
HEADER_MAP.update({
    'invoice number':    'invoice_no',
    'bill no':           'invoice_no',
    'revenue type':      'revenue_type_1',   # a repeated header fills 1 then 2
    'invoice date':      'invoice_date',
    'customer code':     'cust_code',
    'sap customer code': 'cust_code',
    'qty':               'qty',
    'quantity':          'qty',
    'basic value':       'basic_value',
    'invoice value':     'invoice_value',
    'sac code':          'sac_code',
    'hsn code':          'hsn_code',
    'net receivable':    'net_receivable',
})

# GL code → TDS/TCS %. Used by the live register and to fill uploaded rows
# that leave TDS blank.
TDS_PERCENT = {
    '4101076010': 2.0,
    '4101076030': 2.0,
    '4201090080': 2.0,
    '4101076100': 2.0,
    '4101071000': 0.1,
    '4201090210': 10.0,
    '4101076020': 2.0,
}

_AGE_BUCKETS = ((30, '0-30 days'), (60, '31-60 days'), (90, '61-90 days'),
                (180, '91-180 days'), (365, '181-365 days'),
                (730, '1-2 years'), (1095, '2-3 years'))


def age_bucket(days):
    """Ageing bucket for an invoice `days` old. Blank or future-dated → ''."""
    if days is None or days == '' or days < 0:
        return ''
    for limit, label in _AGE_BUCKETS:
        if days <= limit:
            return label
    return '> 3 years'


_DATE_FORMATS = ('%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%d.%m.%Y', '%Y/%m/%d',
                 '%d-%b-%Y', '%d-%b-%y', '%d %b %Y')


def parse_date(v):
    """Date cell → datetime.date, or None if blank/unparseable. Never raises —
    an odd date costs a row its ageing columns, not the whole upload."""
    from datetime import datetime as _dt
    if _blank(v):
        return None
    if hasattr(v, 'year'):                    # openpyxl hands back date/datetime
        return v.date() if hasattr(v, 'hour') else v
    s = str(v).strip()
    for cand in (s, s[:10]):
        for fmt in _DATE_FORMATS:
            try:
                return _dt.strptime(cand, fmt).date()
            except ValueError:
                pass
    return None


def _text(v):
    """Any cell → the text to store. Blank → None. Excel hands back floats and
    datetimes for cells that read as '1000' and '05-04-2025' in the sheet, so
    those two get spelled back the way the sheet shows them."""
    if _blank(v):
        return None
    if hasattr(v, 'year'):                    # date / datetime
        return v.strftime('%Y-%m-%d')
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _date_text(v):
    """Date cell → 'YYYY-MM-DD' where it parses, else the cell's own text."""
    d = parse_date(v)
    return d.strftime('%Y-%m-%d') if d else _text(v)


# ── Parsing ──────────────────────────────────────────────────────────────────
def _index(headers):
    """Field → column position. A repeated 'Revenue type' header fills 1 then 2."""
    idx = {}
    for i, h in enumerate(headers):
        f = HEADER_MAP.get(_norm(h))
        if f == 'revenue_type_1' and f in idx:
            f = 'revenue_type_2'
        if f and f not in idx:
            idx[f] = i
    return idx


def parse_rows(headers, raw_rows):
    """Map raw spreadsheet rows to field dicts → (rows, errors). Fully blank
    rows are skipped. Every cell is stored as text, so the only thing that can
    fail a row is a missing Invoice No. — it is the row's identity and what the
    register is read by. Row numbers are 1-based spreadsheet rows (header row
    included)."""
    idx = _index(headers)
    rows, errors = [], []

    def cell(r, key):
        i = idx.get(key)
        return r[i] if (i is not None and i < len(r)) else None

    for n, r in enumerate(raw_rows, start=2):   # +1 header, +1 to 1-base
        if all(_blank(c) for c in r):
            continue
        rec = {key: (_date_text(cell(r, key)) if key == 'invoice_date'
                     else _text(cell(r, key)))
               for key in COLUMNS}
        if rec.get('invoice_no'):
            rows.append(rec)
        else:
            errors.append({'row': n, 'message': 'Invoice No. is required'})
    return rows, errors


def parse_upload(file_storage):
    """Parse a backdated register upload (.csv/.xlsx) → (rows, errors). The
    header row is the first one carrying Invoice No. plus two more known
    columns, so title rows above it are tolerated."""
    matrix, errors = read_matrix(file_storage)
    if errors:
        return [], errors
    for i, row in enumerate(matrix):
        idx = _index([str(c if c is not None else '') for c in row])
        if 'invoice_no' in idx and len(idx) >= 3:
            headers = [str(c).strip() if c is not None else '' for c in row]
            return parse_rows(headers, matrix[i + 1:])
    return [], [{'row': 0, 'message': 'Could not find a header row — download the '
                                      'template and keep its column names'}]


TEMPLATE_EXAMPLE_ROWS = [
    ['1', 'DPPL/25-26/001', 'Non Group', 'Cargo Handling', 'Wharfage', 'YES',
     '05-04-2025', '10000123', 'JSW Steel Ltd', '4101076010', 'Cargo Handling Charges',
     '19500', '120.00', 'CGST+SGST', '18', '2340000.00', '210600.00', '210600.00',
     '0.00', '2761200.00', '27AAACJ4323N1ZS', '9100000123', '', '996751',
     '', '', '', '', 'Booked', '46800.00', '2714400.00', '', ''],
    ['2', 'DPPL/25-26/002', 'Group', 'Marine', 'Berth Hire', 'NO',
     '12-04-2025', '10000456', 'JSW Cement Ltd', '4101071000', 'Berth Hire Charges',
     '1', '85000.00', 'IGST', '18', '85000.00', '0.00', '0.00', '15300.00',
     '100300.00', '29AAACJ0616P1ZL', '9100000456', '', '996729',
     '', '', '', '', 'Pending', '85.00', '100215.00', '', ''],
]


def build_template_csv():
    """The backdated revenue register upload template as CSV text. Days and
    Bucket are there to match the register sheet; leave them blank."""
    import io, csv as _csv
    buf = io.StringIO()
    w = _csv.writer(buf)
    w.writerow(TEMPLATE_HEADERS)
    for r in TEMPLATE_EXAMPLE_ROWS:
        w.writerow(r)
    return buf.getvalue()


# ── Storage ──────────────────────────────────────────────────────────────────
def month_key(date_text):
    """The 'YYYY-MM' an upload replaces. Dates that did not parse keep their own
    first 7 characters, so re-uploading the same file still targets the same
    rows."""
    return (date_text or '')[:7]


def replace_all(rows, uploaded_by):
    """Clear the stored backdated rows and insert the file's, in one
    transaction. A backdated register is maintained as one whole sheet, so an
    upload is the new truth in full — not a merge. Returns (inserted, months)."""
    months = sorted({month_key(r['invoice_date']) for r in rows})
    conn = get_db()
    cur = get_cursor(conn)
    try:
        cur.execute('DELETE FROM %s' % TABLE)
        cols = COLUMNS + ['uploaded_by']
        sql = 'INSERT INTO %s (%s) VALUES (%s)' % (
            TABLE, ', '.join(cols), ', '.join(['%s'] * len(cols)))
        for r in rows:
            cur.execute(sql, [r.get(c) for c in COLUMNS] + [uploaded_by])
        conn.commit()
        return len(rows), months
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


replace_months = replace_all


def get_status():
    """{count, uploaded_at, months} for the stored backdated dataset."""
    conn = get_db()
    cur = get_cursor(conn)
    try:
        cur.execute('SELECT COUNT(*) AS c, MAX(uploaded_at) AS at FROM %s' % TABLE)
        r = cur.fetchone()
        cur.execute("SELECT DISTINCT left(COALESCE(invoice_date, ''), 7) AS m "
                    'FROM %s ORDER BY m' % TABLE)
        months = [x['m'] for x in cur.fetchall()]
        at = r['at']
        return {'count': r['c'], 'uploaded_at': at.isoformat() if at else None,
                'months': months}
    finally:
        conn.close()


def _row_out(d):
    """DB row → JSON-safe dict, with the two derived columns filled in.
    Everything stored is already text bar the timestamp."""
    out = dict(d)
    if out.get('uploaded_at') is not None:
        out['uploaded_at'] = str(out['uploaded_at'])
    out['days'], out['bucket'] = derive_ageing(out.get('invoice_date'))
    return out


def derive_ageing(date_text):
    """(days, bucket) for a stored date. Unparseable or blank → ('', '')."""
    from datetime import date as _date
    d = parse_date(date_text)
    if not d:
        return '', ''
    days = (_date.today() - d).days
    return days, age_bucket(days)


def get_rows(page=1, size=50, filters=None):
    """Paginated stored rows. `filters` is a list of {field, value}; each is
    ILIKE-matched against the FULL dataset and AND-combined."""
    conn = get_db()
    cur = get_cursor(conn)
    try:
        allowed = set(COLUMNS)
        where_parts, params = [], []
        for f in (filters or []):
            field = f.get('field')
            val = (f.get('value') or '').strip() if f.get('value') is not None else ''
            if field in allowed and val:
                where_parts.append('%s::TEXT ILIKE %%s' % field)
                params.append('%' + val + '%')
        where = ('WHERE ' + ' AND '.join(where_parts)) if where_parts else ''
        cur.execute('SELECT COUNT(*) AS c FROM %s %s' % (TABLE, where), params)
        total = cur.fetchone()['c']
        cur.execute('SELECT * FROM %s %s ORDER BY invoice_date DESC, id DESC '
                    'LIMIT %%s OFFSET %%s' % (TABLE, where),
                    params + [size, (page - 1) * size])
        return [_row_out(r) for r in cur.fetchall()], total
    finally:
        conn.close()


def update_row(row_id, data):
    """Update one stored row. {'success': True} or {'error': msg}. Same rule as
    the upload: text in, text out, Invoice No. is the one thing required."""
    clean = {k: (_date_text(data.get(k)) if k == 'invoice_date' else _text(data.get(k)))
             for k in COLUMNS}
    if not clean.get('invoice_no'):
        return {'error': 'Invoice No. is required'}

    conn = get_db()
    cur = get_cursor(conn)
    try:
        sets = ', '.join(['%s=%%s' % c for c in COLUMNS])
        cur.execute('UPDATE %s SET %s WHERE id=%%s' % (TABLE, sets),
                    [clean.get(c) for c in COLUMNS] + [row_id])
        conn.commit()
        return {'success': True}
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def delete_row(row_id):
    conn = get_db()
    cur = get_cursor(conn)
    try:
        cur.execute('DELETE FROM %s WHERE id=%%s' % TABLE, [row_id])
        conn.commit()
    finally:
        conn.close()


def get_register_rows(month=None, year=None):
    """Stored backdated rows shaped like the live Revenue Register rows."""
    where, params = [], []
    if month:
        where.append('substring(invoice_date, 6, 2) = %s')
        params.append('%02d' % int(month))
    if year:
        where.append('left(invoice_date, 4) = %s')
        params.append(str(int(year)))
    conn = get_db()
    cur = get_cursor(conn)
    try:
        cur.execute('SELECT * FROM %s %s ORDER BY invoice_date DESC, id DESC' % (
            TABLE, ('WHERE ' + ' AND '.join(where)) if where else ''), params)
        rows = cur.fetchall()
    finally:
        conn.close()

    out = []
    for r in rows:
        d = _row_out(r)
        rec = {k: (d[k] if d[k] is not None else '') for k in COLUMNS}
        rec['grouping'] = rec.pop('grouping_label')
        rec['date'] = rec.pop('invoice_date')
        rec['irn_date'] = rec['ack_date']
        rec['days'], rec['bucket'] = d['days'], d['bucket']
        rec['source'] = 'Backdated'
        out.append(rec)
    return out
