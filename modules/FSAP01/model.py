from database import get_db, get_cursor
from datetime import datetime
import sap_builder


def get_outbound_logs(page=1, size=50, type_filter=None):
    """Every OUTBOUND SAP interaction, regardless of doc type.

    Covers postings (source_type='Invoice'), FB08 reversals ('InvoiceReversal'),
    cancellation CNs ('InvoiceCreditNote'), the admin SAP playground
    ('Playground') and IRN fetches — i.e. everything written to integration_logs
    except the inbound callbacks. This is the single place that shows the raw
    request/response payload of a cancellation or a playground send, which the
    per-doc tabs filter out.
    """
    conn = get_db()
    cur = get_cursor(conn)
    where = ["integration_type <> 'SAP_INBOUND'"]
    params = []
    if type_filter:
        where.append('source_type = %s')
        params.append(type_filter)
    where_sql = 'WHERE ' + ' AND '.join(where)
    cur.execute(f'SELECT COUNT(*) AS cnt FROM integration_logs {where_sql}', params)
    total = cur.fetchone()['cnt']
    cur.execute(f'''
        SELECT id, integration_type, source_type, source_id, source_reference,
               request_url, request_body, response_status_code, response_body,
               status, error_message, duration_ms, created_by, created_date
        FROM integration_logs {where_sql}
        ORDER BY id DESC LIMIT %s OFFSET %s
    ''', params + [size, (page - 1) * size])
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows, total


def get_sap_queue(page=1, size=50, status_filter=None):
    """The async SAP outbound queue (pending / processing / sent / failed)."""
    conn = get_db()
    cur = get_cursor(conn)
    where = ''
    params = []
    if status_filter:
        where = 'WHERE status = %s'
        params.append(status_filter)
    cur.execute(f'SELECT COUNT(*) AS cnt FROM sap_outbound_queue {where}', params)
    total = cur.fetchone()['cnt']
    cur.execute(f'''
        SELECT id, job_type, invoice_id, reference_type, reference_number,
               status, retry_count, max_retries, next_attempt_at,
               sap_document_number, last_error, payload, created_by,
               created_date, updated_date
        FROM sap_outbound_queue {where}
        ORDER BY id DESC LIMIT %s OFFSET %s
    ''', params + [size, (page - 1) * size])
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows, total


def get_sap_invoice_logs(page=1, size=50):
    """Invoices with SAP posting data."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT COUNT(*) as cnt FROM invoice_header')
    total = cur.fetchone()['cnt']
    cur.execute('''
        SELECT ih.id, ih.invoice_number, ih.invoice_date, ih.financial_year,
               ih.customer_name, ih.customer_type,
               ih.total_amount, ih.invoice_status,
               ih.sap_document_number, ih.sap_posting_date, ih.sap_fiscal_year, ih.sap_company_code,
               ih.created_by, ih.created_date, ih.posted_by, ih.posted_date,
               log.id AS sap_log_id,
               log.status AS sap_log_status,
               log.error_message AS sap_error_message,
               log.request_url AS sap_request_url,
               log.request_body AS sap_request_body,
               log.response_status_code AS sap_response_status_code,
               log.response_body AS sap_response_body,
               log.duration_ms AS sap_duration_ms,
               log.created_date AS sap_log_date
        FROM invoice_header ih
        LEFT JOIN LATERAL (
            SELECT id, status, error_message, request_url, request_body,
                   response_status_code, response_body, duration_ms, created_date
            FROM integration_logs
            WHERE integration_type = 'SAP'
              AND source_type = 'Invoice'
              AND source_id = ih.id
            ORDER BY id DESC
            LIMIT 1
        ) log ON TRUE
        ORDER BY ih.id DESC LIMIT %s OFFSET %s
    ''', [size, (page - 1) * size])
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows], total


def get_callback_logs(page=1, size=50):
    """SAP inbound callbacks — integration_logs where integration_type = SAP_INBOUND."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("SELECT COUNT(*) as cnt FROM integration_logs WHERE integration_type = 'SAP_INBOUND'")
    total = cur.fetchone()['cnt']
    cur.execute('''
        SELECT il.id,
               il.created_date,
               il.status,
               il.source_reference  AS token_label,
               il.created_by,
               il.response_status_code,
               il.request_body,
               il.response_body
        FROM integration_logs il
        WHERE il.integration_type = 'SAP_INBOUND'
        ORDER BY il.id DESC
        LIMIT %s OFFSET %s
    ''', [size, (page - 1) * size])
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows, total


# Undo what sap_inbound._apply_record did for each record of one callback log.
# The log keeps no before-image, so each outcome is reversed from the payload,
# and only while the row still carries that callback's values.
_UNDO_CALLBACK_SQL = '''
CREATE TEMP TABLE cb ON COMMIT DROP AS
SELECT l.id = %(log_id)s                                            AS deleting,
       TRIM(q.rec->>'Reference')                                    AS ref,
       LEFT(NULLIF(TRIM(q.rec->>'Document_Number'), ''), 50)        AS sap_doc,
       LEFT(NULLIF(COALESCE(q.rec->>'IRN_No', q.rec->>'IRN'), ''), 64) AS irn,
       r.res->>'Message'                                            AS outcome
FROM integration_logs l
CROSS JOIN LATERAL jsonb_array_elements(CASE WHEN jsonb_typeof(l.request_body::jsonb->'Record') = 'array'
                                             THEN l.request_body::jsonb->'Record' ELSE '[]' END) WITH ORDINALITY q(rec, n)
JOIN LATERAL jsonb_array_elements(CASE WHEN jsonb_typeof(l.response_body::jsonb->'Record') = 'array'
                                       THEN l.response_body::jsonb->'Record' ELSE '[]' END) WITH ORDINALITY r(res, n) ON r.n = q.n
WHERE l.integration_type = 'SAP_INBOUND';

-- Reversal recorded: un-cancel and strip the reversal remark
UPDATE invoice_header i
SET invoice_status = CASE WHEN COALESCE(i.gst_irn, '') <> '' THEN 'Posted to GST' ELSE 'Posted to SAP' END,
    remarks = CASE WHEN i.remarks = c.note THEN ''
                   ELSE REPLACE(REPLACE(i.remarks, ' | ' || c.note, ''), c.note || ' | ', '') END
FROM (SELECT ref, 'Reversal SAP Doc: ' || sap_doc AS note
      FROM cb WHERE deleting AND outcome = 'Reversal recorded') c
WHERE i.invoice_number = c.ref
  AND i.invoice_status = 'Cancelled'
  AND STRPOS(i.remarks, c.note) > 0;

-- Stored SAP error: back to Posted to SAP
UPDATE invoice_header i
SET invoice_status = 'Posted to SAP', sap_error = NULL
FROM cb c
WHERE c.deleting AND c.outcome LIKE 'Stored SAP error: %%'
  AND i.invoice_number = c.ref
  AND i.invoice_status = 'SAP Failed'
  AND LEFT(i.sap_error, 80) = SUBSTRING(c.outcome FROM 19);

UPDATE fdcn_header f
SET doc_status = 'Posted to SAP', rejection_reason = NULL
FROM cb c
WHERE c.deleting AND c.outcome LIKE 'Stored SAP error: %%'
  AND f.doc_number = c.ref
  AND f.doc_status = 'SAP Failed'
  AND LEFT(f.rejection_reason, 80) = SUBSTRING(c.outcome FROM 19)
  AND NOT EXISTS (SELECT 1 FROM invoice_header i WHERE i.invoice_number = c.ref);

-- Updated: clear the SAP doc and/or IRN fields this callback filled,
-- unless another remaining callback also updated the same reference.
UPDATE invoice_header i SET
    sap_document_number = CASE WHEN c.sap_doc IS NOT NULL THEN NULL ELSE i.sap_document_number END,
    sap_posting_date    = CASE WHEN c.sap_doc IS NOT NULL THEN NULL ELSE i.sap_posting_date END,
    sap_company_code    = CASE WHEN c.sap_doc IS NOT NULL THEN NULL ELSE i.sap_company_code END,
    gst_irn             = CASE WHEN c.irn IS NOT NULL THEN NULL ELSE i.gst_irn END,
    gst_ack_number      = CASE WHEN c.irn IS NOT NULL THEN NULL ELSE i.gst_ack_number END,
    gst_ack_date        = CASE WHEN c.irn IS NOT NULL THEN NULL ELSE i.gst_ack_date END,
    gst_qr_code         = CASE WHEN c.irn IS NOT NULL THEN NULL ELSE i.gst_qr_code END,
    invoice_status      = CASE WHEN c.irn IS NOT NULL AND i.invoice_status = 'Posted to GST'
                               THEN 'Posted to SAP' ELSE i.invoice_status END
FROM cb c
WHERE c.deleting AND c.outcome LIKE 'Updated%%'
  AND i.invoice_number = c.ref
  AND (c.sap_doc IS NULL OR i.sap_document_number = c.sap_doc)
  AND (c.irn IS NULL OR i.gst_irn = c.irn)
  AND NOT EXISTS (SELECT 1 FROM cb k WHERE NOT k.deleting AND k.ref = c.ref AND k.outcome LIKE 'Updated%%');

UPDATE fdcn_header f SET
    sap_document_number = CASE WHEN c.sap_doc IS NOT NULL THEN NULL ELSE f.sap_document_number END,
    sap_posting_date    = CASE WHEN c.sap_doc IS NOT NULL THEN NULL ELSE f.sap_posting_date END,
    sap_company_code    = CASE WHEN c.sap_doc IS NOT NULL THEN NULL ELSE f.sap_company_code END,
    gst_irn             = CASE WHEN c.irn IS NOT NULL THEN NULL ELSE f.gst_irn END,
    gst_ack_number      = CASE WHEN c.irn IS NOT NULL THEN NULL ELSE f.gst_ack_number END,
    gst_ack_date        = CASE WHEN c.irn IS NOT NULL THEN NULL ELSE f.gst_ack_date END,
    gst_qr_code         = CASE WHEN c.irn IS NOT NULL THEN NULL ELSE f.gst_qr_code END,
    doc_status          = CASE WHEN c.irn IS NOT NULL AND f.doc_status = 'Posted to GST'
                               THEN 'Posted to SAP' ELSE f.doc_status END
FROM cb c
WHERE c.deleting AND c.outcome LIKE 'Updated%%'
  AND f.doc_number = c.ref
  AND (c.sap_doc IS NULL OR f.sap_document_number = c.sap_doc)
  AND (c.irn IS NULL OR f.gst_irn = c.irn)
  AND NOT EXISTS (SELECT 1 FROM cb k WHERE NOT k.deleting AND k.ref = c.ref AND k.outcome LIKE 'Updated%%')
  AND NOT EXISTS (SELECT 1 FROM invoice_header i WHERE i.invoice_number = c.ref);
'''


def delete_callback_log(log_id):
    """Delete one SAP_INBOUND log and restore the invoice/CN it changed. Return (ok, message)."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("SELECT id FROM integration_logs WHERE id=%s AND integration_type='SAP_INBOUND' FOR UPDATE",
                [log_id])
    if not cur.fetchone():
        conn.close()
        return False, 'Callback log not found'
    try:
        cur.execute(_UNDO_CALLBACK_SQL, {'log_id': log_id})
        cur.execute("SELECT ref, outcome FROM cb WHERE deleting")
        records = [f"{r['ref']} ({r['outcome']})" for r in cur.fetchall()]
        cur.execute("DELETE FROM integration_logs WHERE id=%s", [log_id])
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return True, 'Log deleted and its effect undone where the record was unchanged since: ' + ('; '.join(records) or 'no records')


def get_sap_cn_logs(page=1, size=50):
    """FDCN01 credit notes with SAP posting data."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("SELECT COUNT(*) as cnt FROM fdcn_header WHERE doc_type = 'CN'")
    total = cur.fetchone()['cnt']
    cur.execute('''
        SELECT h.id,
               h.doc_number AS credit_note_number,
               h.doc_date AS credit_note_date,
               h.financial_year,
               h.customer_name AS party_name,
               h.original_invoice_number,
               h.total_amount,
               h.doc_status AS credit_note_status,
               h.sap_document_number,
               h.sap_posting_date,
               h.sap_fiscal_year,
               h.sap_company_code,
               h.created_by,
               h.created_date,
               h.posted_by,
               h.posted_date,
               log.id AS sap_log_id,
               log.status AS sap_log_status,
               log.error_message AS sap_error_message,
               log.request_url AS sap_request_url,
               log.request_body AS sap_request_body,
               log.response_status_code AS sap_response_status_code,
               log.response_body AS sap_response_body,
               log.duration_ms AS sap_duration_ms,
               log.created_date AS sap_log_date
        FROM fdcn_header h
        LEFT JOIN LATERAL (
            SELECT id, status, error_message, request_url, request_body,
                   response_status_code, response_body, duration_ms, created_date
            FROM integration_logs
            WHERE integration_type = 'SAP'
              AND source_type = 'CreditNote'
              AND source_id = h.id
            ORDER BY id DESC
            LIMIT 1
        ) log ON TRUE
        WHERE h.doc_type = 'CN'
        ORDER BY h.id DESC LIMIT %s OFFSET %s
    ''', [size, (page - 1) * size])
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows], total


def get_gst_logs(page=1, size=50):
    """Invoices with GST IRN / e-invoice data."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT COUNT(*) as cnt FROM invoice_header')
    total = cur.fetchone()['cnt']
    cur.execute('''
        SELECT id, invoice_number, invoice_date, financial_year,
               customer_name, customer_type,
               total_amount, invoice_status,
               gst_irn, gst_ack_number, gst_ack_date,
               created_by, created_date
        FROM invoice_header
        ORDER BY id DESC LIMIT %s OFFSET %s
    ''', [size, (page - 1) * size])
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows], total


# ─────────────────────────────────────────────────────────────────────────────
# Staging table helpers
# ─────────────────────────────────────────────────────────────────────────────

def get_staging_rows(page=1, size=50, status_filter=None):
    """All staging rows with optional processing_status filter."""
    conn = get_db()
    cur = get_cursor(conn)
    where = 'WHERE processing_status = %s' if status_filter else ''
    params_count = [status_filter] if status_filter else []
    cur.execute(f'SELECT COUNT(*) as cnt FROM invoice_sap_staging {where}', params_count)
    total = cur.fetchone()['cnt']
    params = params_count + [size, (page - 1) * size]
    cur.execute(f'''
        SELECT id, invoice_id, invoice_line_id, line_number,
               reference_text, document_type, invoice_type, company_code,
               document_date, customer_code, invoice_amount, currency,
               processing_status, sap_document_number, sap_message,
               irn_number, ack_number, fiscal_year, fiscal_period,
               push_date, push_time, pushed_by, created_date, updated_date
        FROM invoice_sap_staging {where}
        ORDER BY id DESC LIMIT %s OFFSET %s
    ''', params)
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows], total


def push_invoice_to_staging(invoice_id, pushed_by=None):
    """
    Build staging rows for one invoice and insert into invoice_sap_staging.
    Clears any existing 'N' (New/un-pushed) rows for the same invoice first,
    so re-pushing refreshes the payload without creating duplicates.

    Returns  {'ok': bool, 'rows_inserted': int, 'error': str|None}
    """
    conn = get_db()
    cur = get_cursor(conn)

    # ── Fetch invoice header ──────────────────────────────────────────────────
    cur.execute('SELECT * FROM invoice_header WHERE id = %s', [invoice_id])
    invoice = cur.fetchone()
    if not invoice:
        conn.close()
        return {'ok': False, 'rows_inserted': 0, 'error': f'Invoice {invoice_id} not found'}
    invoice = dict(invoice)

    # ── Fetch invoice lines ───────────────────────────────────────────────────
    cur.execute('''
        SELECT il.*, fst.sap_gl_account, fst.sap_igst_gl, fst.sap_cgst_gl,
               fst.sap_sgst_gl, fst.sap_tds_gl, fst.sap_tcs_gl,
               fst.service_sale_flag, fst.sap_profit_center,
               fst.sap_tax_code as fst_tax_code
        FROM invoice_lines il
        LEFT JOIN finance_service_types fst ON fst.service_code = il.service_code
        WHERE il.invoice_id = %s
        ORDER BY il.line_number
    ''', [invoice_id])
    lines = [dict(r) for r in cur.fetchall()]

    if not lines:
        conn.close()
        return {'ok': False, 'rows_inserted': 0, 'error': 'No invoice lines found'}

    # ── Load SAP config defaults ──────────────────────────────────────────────
    cur.execute('SELECT * FROM sap_api_config WHERE is_active = 1 LIMIT 1')
    cfg_row = cur.fetchone()
    cfg = dict(cfg_row) if cfg_row else {}

    company_code   = cfg.get('company_code') or '5171'
    business_place = cfg.get('business_place') or company_code
    section_code   = cfg.get('section_code')   or company_code
    payment_term   = (cfg.get('default_payment_term') or cfg.get('payment_term') or '')[:4]
    plant_code     = cfg.get('plant_code') or ''
    round_off_gl   = cfg.get('round_off_gl') or ''
    tds_gl_default = cfg.get('tds_gl') or ''
    tcs_gl_default = cfg.get('tcs_gl') or ''

    # Customer company code override (inter-company)
    cust_company = sap_builder._get_customer_company_code(
        invoice.get('customer_type'), invoice.get('customer_id')
    )
    company = cust_company or company_code

    inv_date  = invoice.get('invoice_date')
    ref_text  = (invoice.get('invoice_number') or '')[:16]
    inv_type  = 'C' if invoice.get('is_cancelled') else 'I'
    cancel_f  = 'X' if invoice.get('is_cancelled') else ''
    nat_tx    = sap_builder._nature_of_transaction(invoice.get('customer_gstin'))
    total_amt = sap_builder._total_invoice_amount(invoice, lines)
    service_sale = sap_builder._service_sale_flag(
        lines,
        {l.get('service_code'): {'service_sale_flag': l.get('service_sale_flag')} for l in lines}
    )

    round_off = float(invoice.get('round_off') or 0)

    # ── Delete any un-processed rows for this invoice (safe refresh) ──────────
    cur.execute(
        "DELETE FROM invoice_sap_staging WHERE invoice_id = %s AND processing_status = 'N'",
        [invoice_id]
    )

    # ── Insert one row per line ───────────────────────────────────────────────
    rows_inserted = 0
    for idx, line in enumerate(lines):
        gl_account   = line.get('sap_gl_account') or line.get('gl_code') or ''
        profit_center = line.get('sap_profit_center') or line.get('profit_center') or cfg.get('profit_center') or ''
        tds_gl        = line.get('sap_tds_gl')  or tds_gl_default
        tcs_gl        = line.get('sap_tcs_gl')  or tcs_gl_default
        hsn_sac       = (line.get('sac_code') or line.get('hsn_sac') or '')[:16]

        igst_amt = float(line.get('igst_amount') or 0)
        cgst_amt = float(line.get('cgst_amount') or 0)
        sgst_amt = float(line.get('sgst_amount') or 0)
        tds_amt  = float(line.get('tds_amount')  or 0)
        tcs_amt  = float(line.get('tcs_amount')  or 0)

        # Tax_Code is empty when no GST applies on the line (TC-09 spec).
        # Otherwise pick by transaction type.
        if igst_amt > 0:
            tax_code = cfg.get('igst_tax_code') or cfg.get('tax_code') or ''
        elif (cgst_amt + sgst_amt) > 0:
            tax_code = cfg.get('cgst_tax_code') or cfg.get('tax_code') or ''
        else:
            tax_code = ''

        # GST GLs only when ANY GST applies on this line (all 3 sent together).
        any_gst = (igst_amt + cgst_amt + sgst_amt) > 0
        igst_gl = (line.get('sap_igst_gl') or '') if any_gst else ''
        cgst_gl = (line.get('sap_cgst_gl') or '') if any_gst else ''
        sgst_gl = (line.get('sap_sgst_gl') or '') if any_gst else ''

        # Round-off on the FIRST line, positive value — matches SAP-validated
        # payload (TC-04, DPPL/26-27/41 multi-line).
        rnd_val = round_off if idx == 0 else 0.0

        cur.execute('''
            INSERT INTO invoice_sap_staging (
                invoice_id, invoice_line_id, line_number,
                invoice_type, company_code, document_date, posting_date,
                reference_text, document_type, cancellation_flag,
                nature_of_transaction, service_sale,
                customer_code, invoice_amount, currency,
                business_place, section_code, payment_term, baseline_date, header_text,
                gl_account, gl_amount, plant, profit_center, text_description, tax_code,
                igst_gl, igst_amount, sgst_gl, sgst_amount, cgst_gl, cgst_amount,
                hsn_sac_code, uom, unit_price, quantity,
                tds_gl, tds_amount, tcs_gl, tcs_amount,
                round_off_gl, round_off_value,
                processing_status, pushed_by, created_date
            ) VALUES (
                %s, %s, %s,
                %s, %s, %s, %s,
                %s, %s, %s,
                %s, %s,
                %s, %s, %s,
                %s, %s, %s, %s, %s,
                %s, %s, %s, %s, %s, %s,
                %s, %s, %s, %s, %s, %s,
                %s, %s, %s, %s,
                %s, %s, %s, %s,
                %s, %s,
                'N', %s, NOW()
            )
        ''', [
            invoice_id, line.get('id'), line.get('line_number') or (idx + 1),
            inv_type, company, inv_date, inv_date,
            ref_text, 'DR', cancel_f,
            nat_tx, service_sale,
            (invoice.get('customer_gl_code') or '')[:10],
            total_amt, invoice.get('currency_code') or 'INR',
            business_place, section_code, payment_term, inv_date,
            ref_text[:25],
            gl_account[:10] if gl_account else '',
            float(line.get('line_amount') or 0),
            plant_code,
            profit_center[:10] if profit_center else '',
            (line.get('service_name') or '')[:25],
            tax_code[:2] if tax_code else '',
            igst_gl[:10] if igst_gl else '', igst_amt if igst_amt else None,
            sgst_gl[:10] if sgst_gl else '', sgst_amt if sgst_amt else None,
            cgst_gl[:10] if cgst_gl else '', cgst_amt if cgst_amt else None,
            hsn_sac,
            (line.get('uom') or '')[:3],
            float(line.get('rate') or 0),
            float(line.get('quantity') or 0),
            tds_gl or None, tds_amt if tds_amt else None,
            tcs_gl or None, tcs_amt if tcs_amt else None,
            round_off_gl or None, rnd_val if rnd_val else None,
            pushed_by,
        ])
        rows_inserted += 1

    conn.commit()
    conn.close()
    return {'ok': True, 'rows_inserted': rows_inserted, 'error': None}


def sync_staging_response(invoice_id):
    """
    Read back SAP response fields from the staging table and update
    invoice_header with document number, fiscal year, IRN, etc.

    Called after the adapter has processed the rows (status = 'Y' or 'E').
    Returns  {'ok': bool, 'status': str, 'sap_document_number': str, 'message': str}
    """
    conn = get_db()
    cur = get_cursor(conn)

    # Take the first posted row — all header lines share the same response
    cur.execute('''
        SELECT processing_status, sap_document_number, sap_message,
               fiscal_year, fiscal_period, irn_number, ack_number, irn_date, qr_code
        FROM invoice_sap_staging
        WHERE invoice_id = %s
        ORDER BY
            CASE processing_status WHEN 'Y' THEN 0 WHEN 'E' THEN 1 ELSE 2 END,
            id DESC
        LIMIT 1
    ''', [invoice_id])
    row = cur.fetchone()

    if not row:
        conn.close()
        return {'ok': False, 'status': 'N', 'sap_document_number': None,
                'message': 'No staging rows found for this invoice'}

    row = dict(row)
    status   = row.get('processing_status') or 'N'
    doc_no   = row.get('sap_document_number') or ''
    message  = row.get('sap_message') or ''
    fy       = row.get('fiscal_year') or ''
    irn      = row.get('irn_number') or ''
    ack_no   = row.get('ack_number') or ''
    irn_date = row.get('irn_date')
    qr_code  = row.get('qr_code') or ''

    if status == 'Y':
        inv_status = 'SAP Posted'
        cur.execute('''
            UPDATE invoice_header
            SET sap_document_number = %s,
                sap_posting_date    = NOW(),
                sap_fiscal_year     = %s,
                sap_error           = NULL,
                gst_irn             = CASE WHEN %s <> '' THEN %s ELSE gst_irn END,
                gst_ack_number      = CASE WHEN %s <> '' THEN %s ELSE gst_ack_number END,
                gst_ack_date        = CASE WHEN %s IS NOT NULL THEN %s ELSE gst_ack_date END,
                gst_qr_code         = CASE WHEN %s <> '' THEN %s ELSE gst_qr_code END,
                invoice_status      = %s
            WHERE id = %s
        ''', [
            doc_no, fy,
            irn, irn,
            ack_no, ack_no,
            irn_date, irn_date,
            qr_code, qr_code,
            inv_status, invoice_id,
        ])
        conn.commit()

    elif status == 'E':
        cur.execute(
            "UPDATE invoice_header SET invoice_status='SAP Failed', sap_error=%s WHERE id=%s",
            [message, invoice_id]
        )
        conn.commit()

    conn.close()
    return {
        'ok': status == 'Y',
        'status': status,
        'sap_document_number': doc_no,
        'message': message or ('Posted successfully' if status == 'Y' else 'Pending or error'),
    }


def get_staging_rows_for_invoice(invoice_id):
    """Return all staging rows for one invoice (for debug/view)."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute(
        'SELECT * FROM invoice_sap_staging WHERE invoice_id = %s ORDER BY line_number',
        [invoice_id]
    )
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]


# ─────────────────────────────────────────────────────────────────────────────
# SAP adapter callback — SAP posts response back matched by reference_text
# ─────────────────────────────────────────────────────────────────────────────

def sap_callback(payload):
    """
    Called by the SAP adapter after it processes staging rows.

    Expected payload keys:
        reference_text      – our invoice number (VARCHAR 16, primary match key)
        processing_status   – 'Y' (posted) | 'E' (error) | 'R' (reversed)
        sap_document_number – SAP FI document number
        fiscal_year         – e.g. '2025'
        fiscal_period       – e.g. '01'
        sap_message         – free-text message / error description
        irn_number          – GST IRN
        ack_number          – GST acknowledgement number
        irn_date            – DATE string (YYYY-MM-DD)
        qr_code             – GST QR code string

    Returns {'ok': bool, 'rows_updated': int, 'error': str|None}
    """
    ref = (payload.get('reference_text') or '').strip()
    if not ref:
        return {'ok': False, 'rows_updated': 0, 'error': 'reference_text is required'}

    status    = (payload.get('processing_status') or '').strip().upper()
    if status not in ('Y', 'E', 'R'):
        return {'ok': False, 'rows_updated': 0,
                'error': f"processing_status must be Y/E/R, got '{status}'"}

    doc_no    = (payload.get('sap_document_number') or '').strip()
    fy        = (payload.get('fiscal_year')          or '').strip()
    fp        = (payload.get('fiscal_period')        or '').strip()
    message   = (payload.get('sap_message')          or '').strip()
    irn       = (payload.get('irn_number')           or '').strip()
    ack_no    = (payload.get('ack_number')            or '').strip()
    irn_date  = payload.get('irn_date')   or None
    qr_code   = (payload.get('qr_code')              or '').strip()

    conn = get_db()
    cur  = get_cursor(conn)

    # ── Update staging rows ───────────────────────────────────────────────────
    cur.execute('''
        UPDATE invoice_sap_staging
        SET processing_status   = %s,
            sap_document_number = %s,
            fiscal_year         = %s,
            fiscal_period       = %s,
            sap_message         = %s,
            irn_number          = CASE WHEN %s <> '' THEN %s ELSE irn_number END,
            ack_number          = CASE WHEN %s <> '' THEN %s ELSE ack_number END,
            irn_date            = CASE WHEN %s IS NOT NULL THEN %s ELSE irn_date END,
            qr_code             = CASE WHEN %s <> '' THEN %s ELSE qr_code END,
            updated_date        = NOW()
        WHERE reference_text = %s
    ''', [
        status, doc_no or None, fy or None, fp or None, message or None,
        irn, irn,
        ack_no, ack_no,
        irn_date, irn_date,
        qr_code, qr_code,
        ref,
    ])
    rows_updated = cur.rowcount

    if rows_updated == 0:
        conn.close()
        return {'ok': False, 'rows_updated': 0,
                'error': f"No staging rows found for reference_text '{ref}'"}

    # ── Resolve invoice_id from staging ──────────────────────────────────────
    cur.execute(
        'SELECT DISTINCT invoice_id FROM invoice_sap_staging WHERE reference_text = %s',
        [ref]
    )
    inv_rows = [r['invoice_id'] for r in cur.fetchall()]

    # ── Update invoice_header for each matched invoice_id ────────────────────
    for invoice_id in inv_rows:
        if status == 'Y':
            cur.execute('''
                UPDATE invoice_header
                SET sap_document_number = %s,
                    sap_posting_date    = NOW(),
                    sap_fiscal_year     = %s,
                    sap_error           = NULL,
                    gst_irn             = CASE WHEN %s <> '' THEN %s ELSE gst_irn END,
                    gst_ack_number      = CASE WHEN %s <> '' THEN %s ELSE gst_ack_number END,
                    gst_ack_date        = CASE WHEN %s IS NOT NULL THEN %s ELSE gst_ack_date END,
                    gst_qr_code         = CASE WHEN %s <> '' THEN %s ELSE gst_qr_code END,
                    invoice_status      = 'SAP Posted'
                WHERE id = %s
            ''', [
                doc_no, fy,
                irn, irn,
                ack_no, ack_no,
                irn_date, irn_date,
                qr_code, qr_code,
                invoice_id,
            ])
        elif status == 'E':
            cur.execute(
                "UPDATE invoice_header SET invoice_status='SAP Failed', sap_error=%s WHERE id=%s",
                [message or 'SAP adapter error', invoice_id]
            )
        elif status == 'R':
            cur.execute(
                "UPDATE invoice_header SET invoice_status='SAP Reversed', sap_error=NULL WHERE id=%s",
                [invoice_id]
            )

    conn.commit()
    conn.close()
    return {'ok': True, 'rows_updated': rows_updated, 'error': None}
