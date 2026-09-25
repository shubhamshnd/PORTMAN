"""
SAP outbound queue — async posting with automatic retry.

Decouples invoice generation / cancellation from the synchronous SAP call.
Mirrors the mail_service queue: a posting is saved (payload + metadata) and a
background worker attempts it, retrying up to MAX_RETRIES times spaced
RETRY_INTERVAL_MIN apart. If every retry fails the row is left 'failed' and a
user can trigger a manual send. Every attempt is logged to integration_logs by
sap_client, so FSAP01 shows the full history.

Status lifecycle:  pending -> processing -> sent
                                          -> pending (retry, next_attempt_at += 5m)
                                          -> failed (retries exhausted; manual send)
"""
import json
import threading
from datetime import datetime, timedelta

from database import get_db, get_cursor
import sap_client

MAX_RETRIES = 10
RETRY_INTERVAL_MIN = 5
BATCH = 25


# ---------------------------------------------------------------------------
# Enqueue + trigger
# ---------------------------------------------------------------------------
def enqueue(job_type, reference_type, reference_id, reference_number, payload,
            invoice_id=None, created_by=None):
    """Queue a SAP posting. Returns the queue row id.

    Idempotent per (invoice_id, job_type): if an active (not-sent) job already
    exists it is reused instead of inserting a duplicate, so double-clicks
    cannot double-post. Fires an immediate attempt in the background.
    """
    now = datetime.now()
    conn = get_db()
    cur = get_cursor(conn)
    if invoice_id is not None:
        cur.execute("""SELECT id FROM sap_outbound_queue
                       WHERE invoice_id=%s AND job_type=%s AND status <> 'sent'
                       ORDER BY id DESC LIMIT 1""", [invoice_id, job_type])
        existing = cur.fetchone()
        if existing:
            conn.close()
            trigger()
            return existing['id']
    cur.execute("""INSERT INTO sap_outbound_queue
        (job_type, invoice_id, reference_type, reference_id, reference_number,
         payload, status, retry_count, max_retries, next_attempt_at,
         created_by, created_date, updated_date)
        VALUES (%s,%s,%s,%s,%s,%s,'pending',0,%s,%s,%s,%s,%s) RETURNING id""",
        [job_type, invoice_id, reference_type, reference_id, reference_number,
         json.dumps(payload), MAX_RETRIES, now, created_by, now, now])
    qid = cur.fetchone()['id']
    conn.commit()
    conn.close()
    trigger()
    return qid


def trigger():
    """Kick a worker pass in the background (non-blocking)."""
    threading.Thread(target=process_sap_queue, daemon=True).start()


# ---------------------------------------------------------------------------
# Worker
# ---------------------------------------------------------------------------
def process_sap_queue():
    """Scheduler entry point. Attempts every due pending item once."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("""SELECT id FROM sap_outbound_queue
        WHERE status='pending' AND retry_count < max_retries
          AND (next_attempt_at IS NULL OR next_attempt_at <= %s)
        ORDER BY id LIMIT %s""", [datetime.now(), BATCH])
    ids = [r['id'] for r in cur.fetchall()]
    conn.close()
    for qid in ids:
        row = _claim(qid)
        if row:
            _attempt(row)


def _claim(qid):
    """Atomically flip pending->processing so two threads can't both run it."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("""UPDATE sap_outbound_queue SET status='processing', updated_date=%s
                   WHERE id=%s AND status='pending' RETURNING *""",
                [datetime.now(), qid])
    row = cur.fetchone()
    conn.commit()
    conn.close()
    return dict(row) if row else None


def _already_in_sap(row):
    """A timed-out POST may still have landed in SAP. If the inbound callback has
    since stamped a SAP document number on the invoice, re-posting would duplicate it."""
    if row['job_type'] != 'post' or not row.get('invoice_id'):
        return None
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("SELECT sap_document_number FROM invoice_header WHERE id=%s", [row['invoice_id']])
    r = cur.fetchone()
    conn.close()
    return (r and (r['sap_document_number'] or '').strip()) or None


def _attempt(row):
    sap_doc = _already_in_sap(row)
    if sap_doc:
        _mark_sent(row, sap_doc)
        return
    payload = json.loads(row['payload'])
    result = sap_client.post_invoice_to_sap(
        payload, row['reference_type'], row['reference_id'] or 0,
        row['reference_number'], row['created_by'])

    if result.get('ok'):
        sap_doc = result.get('sap_document_number') or ''
        try:
            _apply_success(row, sap_doc)
        except Exception as e:
            # SAP accepted but our bookkeeping failed — don't silently retry a
            # second SAP post. Park it as failed with a clear note for manual fix.
            _mark_failed(row, f"SAP posted ({sap_doc}) but post-processing failed: {e}",
                         final=True)
            return
        _mark_sent(row, sap_doc)
    else:
        final = (row['retry_count'] + 1) >= row['max_retries']
        _mark_failed(row, result.get('message'), final=final)


def _mark_sent(row, sap_doc):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("""UPDATE sap_outbound_queue
        SET status='sent', sap_document_number=%s, last_error=NULL, updated_date=%s
        WHERE id=%s""", [sap_doc, datetime.now(), row['id']])
    conn.commit()
    conn.close()


def _mark_failed(row, msg, final):
    """Bump retry_count; back to pending (retry in 5m) unless retries exhausted."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("""UPDATE sap_outbound_queue
        SET retry_count = retry_count + 1,
            status = CASE WHEN %s THEN 'failed' ELSE 'pending' END,
            next_attempt_at = %s,
            last_error = %s,
            updated_date = %s
        WHERE id=%s""",
        [final, datetime.now() + timedelta(minutes=RETRY_INTERVAL_MIN),
         (msg or '')[:1000], datetime.now(), row['id']])
    conn.commit()
    conn.close()
    # Reflect a fully-exhausted POST on the invoice so the UI shows Manual Send.
    if final and row['job_type'] == 'post' and row.get('invoice_id'):
        _set_invoice_status(row['invoice_id'], 'SAP Failed', msg)


# ---------------------------------------------------------------------------
# Manual send (one-shot, bypasses the retry-count gate)
# ---------------------------------------------------------------------------
def manual_send(queue_id):
    row = _claim_for_manual(queue_id)
    if not row:
        return {'ok': False, 'error': 'Queue item not found or already sent'}
    sap_doc = _already_in_sap(row)
    if sap_doc:
        _mark_sent(row, sap_doc)
        return {'ok': True, 'sap_document_number': sap_doc}
    payload = json.loads(row['payload'])
    result = sap_client.post_invoice_to_sap(
        payload, row['reference_type'], row['reference_id'] or 0,
        row['reference_number'], row['created_by'])
    if result.get('ok'):
        sap_doc = result.get('sap_document_number') or ''
        try:
            _apply_success(row, sap_doc)
        except Exception as e:
            _mark_failed(row, f"SAP posted ({sap_doc}) but post-processing failed: {e}",
                         final=True)
            return {'ok': False, 'error': f'SAP posted but bookkeeping failed: {e}'}
        _mark_sent(row, sap_doc)
        return {'ok': True, 'sap_document_number': sap_doc}
    # Keep it failed (manual attempts don't re-arm the auto-retry loop).
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("""UPDATE sap_outbound_queue SET status='failed', last_error=%s, updated_date=%s
                   WHERE id=%s""", [(result.get('message') or '')[:1000], datetime.now(), queue_id])
    conn.commit()
    conn.close()
    return {'ok': False, 'error': result.get('message')}


def _claim_for_manual(queue_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("""UPDATE sap_outbound_queue SET status='processing', updated_date=%s
                   WHERE id=%s AND status IN ('failed','pending') RETURNING *""",
                [datetime.now(), queue_id])
    row = cur.fetchone()
    conn.commit()
    conn.close()
    return dict(row) if row else None


# ---------------------------------------------------------------------------
# Success side-effects (mirror the original synchronous view handlers)
# ---------------------------------------------------------------------------
def _apply_success(row, sap_doc):
    jt = row['job_type']
    if jt == 'post':
        _apply_post_success(row, sap_doc)
    elif jt == 'reversal':
        _apply_reversal_success(row, sap_doc)
    elif jt == 'credit_note':
        _apply_cn_success(row, sap_doc)


def _apply_post_success(row, sap_doc):
    # Staging push only — sap_document_number / posting_date arrive via callback.
    now_ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("""UPDATE invoice_header
        SET posted_by=%s, posted_date=%s, invoice_status='Posted to SAP', sap_error=NULL
        WHERE id=%s""", [row['created_by'], now_ts, row['invoice_id']])
    conn.commit()
    conn.close()


def _apply_reversal_success(row, sap_doc):
    from modules.FIN01 import model
    now_ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    invoice = model.get_invoice_by_id(row['invoice_id']) or {}
    original_doc = invoice.get('sap_document_number') or ''
    note = f"SAP FB08 reversal posted. Original: {original_doc}; Reversal: {sap_doc}"
    conn = get_db()
    cur = get_cursor(conn)
    unbilled = model.unbill_invoice_sources(cur, row['invoice_id'])
    if unbilled:
        note += f". Bills unbilled: {', '.join(unbilled)}"
    cur.execute("""UPDATE invoice_header
        SET invoice_status='Cancelled', posted_by=%s, posted_date=%s,
            remarks = CASE WHEN COALESCE(remarks,'')='' THEN %s ELSE remarks || ' | ' || %s END
        WHERE id=%s""", [row['created_by'], now_ts, note, note, row['invoice_id']])
    conn.commit()
    conn.close()


def _apply_cn_success(row, sap_doc):
    from modules.FIN01 import model
    from modules.FDCN01 import model as fdcn_model
    username = row['created_by']
    invoice_id = row['invoice_id']
    invoice = model.get_invoice_by_id(invoice_id) or {}

    fdcn_id, cn_doc_number = fdcn_model.create_cancellation_cn(invoice_id, username)
    if sap_doc:
        fdcn_model.update_sap_details(fdcn_id, sap_doc, username)

    now_ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    original_sap_doc = invoice.get('sap_document_number') or ''
    note = (f"Cancelled via CN {cn_doc_number}. "
            f"SAP original: {original_sap_doc}; SAP CN: {sap_doc}")
    conn = get_db()
    cur = get_cursor(conn)
    unbilled = model.unbill_invoice_sources(cur, invoice_id)
    if unbilled:
        note += f". Bills unbilled: {', '.join(unbilled)}"
    cur.execute("""UPDATE invoice_header
        SET invoice_status='Cancelled', posted_by=%s, posted_date=%s,
            remarks = CASE WHEN COALESCE(remarks,'')='' THEN %s ELSE remarks || ' | ' || %s END
        WHERE id=%s""", [username, now_ts, note, note, invoice_id])
    conn.commit()
    conn.close()


def _set_invoice_status(invoice_id, status, error=None):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("UPDATE invoice_header SET invoice_status=%s, sap_error=%s WHERE id=%s",
                [status, (error or '')[:1000], invoice_id])
    conn.commit()
    conn.close()


# ---------------------------------------------------------------------------
# Read helpers for FSAP01 / FINV01 UI
# ---------------------------------------------------------------------------
def get_active_jobs_map(invoice_ids):
    """{invoice_id: latest non-sent queue job} for a batch of invoices (UI badges)."""
    if not invoice_ids:
        return {}
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("""SELECT DISTINCT ON (invoice_id)
                          invoice_id, id, job_type, status, retry_count, max_retries,
                          next_attempt_at, last_error
                   FROM sap_outbound_queue
                   WHERE invoice_id = ANY(%s) AND status <> 'sent'
                   ORDER BY invoice_id, id DESC""", [list(invoice_ids)])
    out = {r['invoice_id']: dict(r) for r in cur.fetchall()}
    conn.close()
    return out


def get_queue_for_invoice(invoice_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("""SELECT id, job_type, status, retry_count, max_retries,
                          next_attempt_at, last_error, sap_document_number, updated_date
                   FROM sap_outbound_queue WHERE invoice_id=%s
                   ORDER BY id DESC""", [invoice_id])
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows
