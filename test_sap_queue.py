"""Self-check for the SAP outbound queue retry state machine.

Stubs the SAP HTTP call so no real SAP is touched. Verifies:
  - a failed attempt increments retry_count and stays 'pending' until exhausted
  - after max_retries the row flips to 'failed'
  - a later success flips it to 'sent'
Run: python test_sap_queue.py
"""
import sap_queue
import sap_client
from database import get_db, get_cursor


def _row(qid):
    conn = get_db(); cur = get_cursor(conn)
    cur.execute('SELECT * FROM sap_outbound_queue WHERE id=%s', [qid])
    r = cur.fetchone(); conn.close()
    return dict(r)


def _cleanup(qid):
    conn = get_db(); cur = get_cursor(conn)
    cur.execute('DELETE FROM sap_outbound_queue WHERE id=%s', [qid])
    conn.commit(); conn.close()


def run():
    # Don't let enqueue's background trigger race the synchronous worker calls.
    sap_queue.trigger = lambda: None

    outcome = {'ok': False}
    sap_client.post_invoice_to_sap = (
        lambda payload, rtype, rid, rnum, by=None:
        {'ok': outcome['ok'], 'sap_document_number': 'DOC1' if outcome['ok'] else None,
         'message': 'stub fail', 'log_id': 0})

    # job_type 'test' has no side-effects in _apply_success; invoice_id=None
    # sidesteps the dedupe index and the invoice-status reflection.
    qid = sap_queue.enqueue('test', 'Test', 0, 'SELFCHECK', {'x': 1}, invoice_id=None)
    try:
        # Force the per-item 5-min spacing out of the way each pass.
        for expected in range(1, sap_queue.MAX_RETRIES + 1):
            _due_now(qid)
            sap_queue.process_sap_queue()
            r = _row(qid)
            assert r['retry_count'] == expected, (expected, r['retry_count'])
            if expected < sap_queue.MAX_RETRIES:
                assert r['status'] == 'pending', r['status']
            else:
                assert r['status'] == 'failed', r['status']

        # Now SAP recovers; a manual send must post and mark it sent.
        outcome['ok'] = True
        res = sap_queue.manual_send(qid)
        assert res['ok'], res
        r = _row(qid)
        assert r['status'] == 'sent', r['status']
        assert r['sap_document_number'] == 'DOC1', r['sap_document_number']
        print('PASS: retry -> failed -> manual send -> sent')
    finally:
        _cleanup(qid)

    # A timed-out post that SAP actually booked (callback stamped a doc no.)
    # must be marked sent without a second SAP post.
    calls = []
    sap_client.post_invoice_to_sap = lambda *a, **k: calls.append(1) or {'ok': True}
    sap_queue._already_in_sap = lambda row: 'CALLBACKDOC'
    qid = sap_queue.enqueue('post', 'Invoice', 0, 'SELFCHECK2', {'x': 1}, invoice_id=None)
    try:
        sap_queue.process_sap_queue()
        r = _row(qid)
        assert not calls, 'reposted an invoice already in SAP'
        assert r['status'] == 'sent' and r['sap_document_number'] == 'CALLBACKDOC', r
        print('PASS: already-in-SAP retry skipped')
    finally:
        _cleanup(qid)


def _due_now(qid):
    conn = get_db(); cur = get_cursor(conn)
    cur.execute("UPDATE sap_outbound_queue SET next_attempt_at=NULL WHERE id=%s", [qid])
    conn.commit(); conn.close()


if __name__ == '__main__':
    run()
