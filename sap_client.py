"""
SAP REST API Client — OAuth2 token management + DynaportInvoice posting.

Uses the active sap_api_config row for credentials/endpoints.
Writes every request/response to integration_logs.
"""
import requests
import json
import time
from datetime import datetime
from database import get_db, get_cursor
from modules.SAPCFG.model import get_active_config


# ---------------------------------------------------------------------------
# Token cache (module-level; refreshed when expired)
# ---------------------------------------------------------------------------
_token_cache = {'access_token': None, 'expires_at': 0}


def _get_oauth_token(config, force_refresh=False):
    """
    Obtain or reuse a Bearer token via client_credentials grant.
    PORTBIRD spec: credentials are passed as URL query parameters, not form body.
    """
    global _token_cache
    now = time.time()

    if not force_refresh and _token_cache['access_token'] and now < _token_cache['expires_at']:
        return _token_cache['access_token']

    base = config['base_url'].rstrip('/')
    token_url = config.get('token_url') or f"{base}/RESTAdapter/OAuthServer"
    resp = requests.post(token_url, params={
        'client_id':     config['client_id'],
        'client_secret': config['client_secret'],
        'grant_type':    'client_credentials',
    }, timeout=120)
    resp.raise_for_status()
    body = resp.json()

    _token_cache['access_token'] = body['access_token']
    # SAP tokens are typically valid for 3600s; subtract 60s buffer
    _token_cache['expires_at'] = now + body.get('expires_in', 3600) - 60
    return _token_cache['access_token']


# ---------------------------------------------------------------------------
# Integration log helpers
# ---------------------------------------------------------------------------
def _write_log(integration_type, source_type, source_id, source_reference,
               request_body, response_body, status, error_message=None, created_by=None,
               request_url=None, response_status_code=None, duration_ms=None):
    """Insert a row into integration_logs (matches migration column names)."""
    conn = get_db()
    cur = get_cursor(conn)
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    cur.execute('''INSERT INTO integration_logs
        (integration_type, source_type, source_id, source_reference,
         request_url, request_body, response_status_code, response_body,
         status, error_message, duration_ms, created_by, created_date)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        RETURNING id''',
        [integration_type, source_type, source_id, source_reference,
         request_url,
         json.dumps(request_body) if request_body else None,
         response_status_code,
         json.dumps(response_body) if response_body else None,
         status, error_message, duration_ms, created_by, now])
    log_id = cur.fetchone()['id']
    conn.commit()
    conn.close()
    return log_id


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def post_invoice_to_sap(payload, reference_type, reference_id, reference_number, created_by=None):
    """
    POST the DynaportInvoice JSON to SAP PI and log the result.

    Parameters
    ----------
    payload : dict
        The complete {"Record_Header": [...]} dict built by sap_builder.
    reference_type : str
        'Invoice' or 'CreditNote'.
    reference_id : int
        The invoice_header.id or credit_note_header.id.
    reference_number : str
        The invoice_number or credit_note_number.
    created_by : str, optional
        Username for audit trail.

    Returns
    -------
    dict  {"ok": bool, "sap_document_number": str|None, "message": str, "log_id": int}
    """
    config = get_active_config()
    if not config:
        log_id = _write_log('SAP', reference_type, reference_id, reference_number,
                            payload, None, 'Error',
                            'No active SAP configuration found', created_by)
        return {'ok': False, 'sap_document_number': None,
                'message': 'No active SAP configuration found', 'log_id': log_id}

    try:
        token = _get_oauth_token(config)
    except Exception as e:
        log_id = _write_log('SAP', reference_type, reference_id, reference_number,
                            payload, None, 'Error',
                            f'Token error: {str(e)}', created_by)
        return {'ok': False, 'sap_document_number': None,
                'message': f'SAP token error: {str(e)}', 'log_id': log_id}

    url = f"{config['base_url'].rstrip('/')}/RESTAdapter/DynaportInvoice"
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json',
    }

    try:
        started_at = time.time()
        resp = requests.post(url, json=payload, headers=headers, timeout=120)
        duration_ms = int((time.time() - started_at) * 1000)
        resp_body = resp.json() if resp.headers.get('content-type', '').startswith('application/json') else {'raw': resp.text}

        if resp.ok:
            # SAP typically returns the document number in the response
            sap_doc_no = resp_body.get('Document_Number') or resp_body.get('document_number')
            log_id = _write_log('SAP', reference_type, reference_id, reference_number,
                                payload, resp_body, 'Success', None, created_by,
                                request_url=url,
                                response_status_code=resp.status_code,
                                duration_ms=duration_ms)
            return {'ok': True, 'sap_document_number': sap_doc_no,
                    'message': 'Posted to SAP successfully', 'log_id': log_id,
                    'request_url': url, 'http_status': resp.status_code,
                    'response_body': resp_body, 'duration_ms': duration_ms}
        else:
            error_msg = resp_body.get('message') or resp_body.get('error') or resp.text
            log_id = _write_log('SAP', reference_type, reference_id, reference_number,
                                payload, resp_body, 'Error',
                                f'HTTP {resp.status_code}: {error_msg}', created_by,
                                request_url=url,
                                response_status_code=resp.status_code,
                                duration_ms=duration_ms)
            return {'ok': False, 'sap_document_number': None,
                    'message': f'SAP returned {resp.status_code}: {error_msg}', 'log_id': log_id,
                    'request_url': url, 'http_status': resp.status_code,
                    'response_body': resp_body, 'duration_ms': duration_ms}

    except requests.RequestException as e:
        log_id = _write_log('SAP', reference_type, reference_id, reference_number,
                            payload, None, 'Error',
                            f'Request failed: {str(e)}', created_by,
                            request_url=url)
        return {'ok': False, 'sap_document_number': None,
                'message': f'SAP request failed: {str(e)}', 'log_id': log_id}


# fetch_irn_from_sap was removed: SAP PI has no IRN fetch endpoint — IRN/ack
# details arrive via the inbound callback (sap_inbound.py) instead.
