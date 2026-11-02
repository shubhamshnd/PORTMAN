from flask import render_template, request, redirect, url_for, session, jsonify, make_response as flask_make_response
from datetime import datetime, timedelta
from . import bp
from modules.FIN01 import model  # reuse FIN01 model for invoice functions
from database import get_user_permissions, get_db, get_cursor, get_module_config
from mail_service import notify_module_approver, get_module_approver_info, build_approval_mail_html
import sap_builder
import sap_client
import sap_queue
import logging

log = logging.getLogger(__name__)

MODULE_CODE = 'FINV01'


def _get_default_invds_series(cur):
    """Resolve the default invoice doc series from INVDS01 storage."""
    try:
        cur.execute('''
            SELECT id, name, prefix, is_default
            FROM invoice_doc_series
            WHERE is_default = TRUE
            ORDER BY id
            LIMIT 1
        ''')
        row = cur.fetchone()
        if row:
            return dict(row)
        cur.execute('''
            SELECT id, name, prefix, is_default
            FROM invoice_doc_series
            ORDER BY id
            LIMIT 1
        ''')
        row = cur.fetchone()
        return dict(row) if row else None
    except Exception:
        return None


def _get_customer_master_snapshot(cur, customer_type, customer_id):
    """Fetch billing master data used for invoice header/print fallbacks."""
    if not customer_id:
        return {}

    table = 'vessel_agents' if str(customer_type or '').lower() == 'agent' else 'vessel_customers'
    try:
        cur.execute(f'''
            SELECT name, gstin, gst_state_code, gst_state_name, pan, cin,
                   billing_address, city, pincode, contact_email, contact_phone,
                   virtual_account_number
            FROM {table}
            WHERE id = %s
        ''', [customer_id])
        row = cur.fetchone()
        return dict(row) if row else {}
    except Exception:
        cur.connection.rollback()
        cur.execute(f'''
            SELECT name, gstin, gst_state_code, gst_state_name, pan,
                   billing_address, city, pincode, contact_email, contact_phone,
                   virtual_account_number
            FROM {table}
            WHERE id = %s
        ''', [customer_id])
        row = cur.fetchone()
        result = dict(row) if row else {}
        result.setdefault('cin', '')
        return result


def _queue_invoice_review_request(invoice_id, invoice_number, customer_name, total_amount, invoice_status):
    info = get_module_approver_info(MODULE_CODE, fallback_module='FIN01')
    if not info.get('approval_add'):
        return
    invoice_url = request.host_url.rstrip('/') + url_for('FINV01.print_invoice', invoice_id=invoice_id)
    info = get_module_approver_info(MODULE_CODE, fallback_module='FIN01')
    notify_module_approver(
        module_code=MODULE_CODE,
        ref_id=invoice_id,
        subject=f"[Portbird DPPL] Invoice {invoice_number} — Review Required",
        fallback_module='FIN01',
        body_html=build_approval_mail_html(
            approver_name=info.get('username'),
            action_label=invoice_status or 'Generated',
            subtitle='Invoice — Review Required',
            details=[
                ('Invoice No',    invoice_number or '—'),
                ('Customer',      customer_name or '—'),
                ('Total Amount',  f'₹ {float(total_amount or 0):,.2f}'),
                ('Status',        invoice_status or 'Generated'),
            ],
            action_url=invoice_url,
            action_btn_label='View Invoice',
            submitted_by=session.get('username'),
            badge_color='#2544a7',
        ),
    )


def _parse_datetime(value):
    if not value:
        return None
    txt = str(value).strip()
    formats = (
        ('%Y-%m-%d %H:%M:%S', 19),
        ('%Y-%m-%dT%H:%M:%S', 19),
        ('%Y-%m-%d', 10),
    )
    for fmt, width in formats:
        try:
            return datetime.strptime(txt[:width], fmt)
        except Exception:
            continue
    return None


def get_perms():
    if session.get('is_admin'):
        return {'can_read': 1, 'can_add': 1, 'can_edit': 1, 'can_delete': 1}
    # Fall back to FIN01 permissions if FINV01 not set up yet
    perms = get_user_permissions(session.get('user_id'), MODULE_CODE)
    if not perms.get('can_read'):
        perms = get_user_permissions(session.get('user_id'), 'FIN01')
    return perms


# ===== Invoice List =====

@bp.route('/module/FINV01/')
def index():
    return redirect(url_for('FINV01.invoices'))


@bp.route('/module/FINV01/doc-series')
def doc_series_page():
    """Invoice doc series master management page"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    perms = get_perms()
    return render_template('finv01_doc_series.html',
                           perms=perms,
                           username=session.get('username'),
                           module_code='FINV01')


@bp.route('/module/FINV01/invoices')
def invoices():
    """List all invoices"""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    perms = get_perms()
    page = int(request.args.get('page', 1))
    status_filter = request.args.get('status')
    data, total = model.get_invoice_data(page, status_filter=status_filter)

    now = datetime.now()
    queue_map = sap_queue.get_active_jobs_map([r['id'] for r in data if r.get('id')])
    for row in data:
        # 24h FB08 window anchors on sap_posting_date (set by /api/sap/callback
        # when SAP confirms the posting), falling back to posted_date (our staging
        # push time) so Cancel is available before SAP has acknowledged.
        posted_dt = (_parse_datetime(row.get('sap_posting_date'))
                     or _parse_datetime(row.get('posted_date')))
        row['within_cancel_window'] = bool(
            posted_dt and (now - posted_dt) <= timedelta(hours=24)
        )
        row['sap_queue'] = queue_map.get(row.get('id'))

    return render_template('finv01_invoices.html',
                         data=data,
                         page=page,
                         last_page=(total + 19) // 20,
                         status_filter=status_filter,
                         perms=perms,
                         username=session.get('username'),
                         module_code='FINV01')


# ===== Generate Invoice from Bills =====

@bp.route('/module/FINV01/generate')
def generate_invoice():
    """Generate invoice from approved bills"""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    perms = get_perms()

    # Get all approved bills not yet invoiced
    approved_bills, _ = model.get_bill_data(page=1, size=1000, status_filter='Approved')

    from datetime import datetime
    current_date = datetime.now().strftime('%Y-%m-%d')

    return render_template('finv01_generate_invoice.html',
                         approved_bills=approved_bills,
                         current_date=current_date,
                         perms=perms,
                         username=session.get('username'),
                         module_code='FINV01')


def _enqueue_invoice_post(invoice_id, invoice_number):
    """Queue the invoice for async SAP posting (non-blocking).

    The payload is built and saved now so the posting survives SAP being down
    (e.g. Thursday migration windows). The background worker retries it every
    5 minutes until SAP acks "Invoice saved in staging table"; SAP's latest
    error shows in FINV01. Returns immediately — no synchronous SAP call.
    """
    try:
        invoice = model.get_invoice_by_id(invoice_id)
        invoice_lines = model.get_invoice_lines(invoice_id)
        payload = sap_builder.build_invoice_payload(invoice, invoice_lines)
        qid = sap_queue.enqueue(
            'post', 'Invoice', invoice_id, invoice_number, payload,
            invoice_id=invoice_id, created_by=session.get('username'))
        conn = get_db()
        cur = get_cursor(conn)
        cur.execute('''UPDATE invoice_header
            SET invoice_status='Queued for SAP', sap_error=NULL WHERE id=%s''',
            [invoice_id])
        conn.commit()
        conn.close()
        return {'ok': True, 'queued': True, 'queue_id': qid,
                'message': 'Queued for SAP posting'}
    except Exception as e:
        log.exception('Failed to queue invoice %s for SAP', invoice_number)
        try:
            conn = get_db()
            cur = get_cursor(conn)
            cur.execute('''UPDATE invoice_header
                SET invoice_status='SAP Failed', sap_error=%s WHERE id=%s''',
                [str(e), invoice_id])
            conn.commit()
            conn.close()
        except Exception:
            pass
        return {'ok': False, 'message': str(e)}


@bp.route('/api/module/FINV01/invoice/create', methods=['POST'])
def create_invoice():
    """Create invoice from selected bills"""
    if 'user_id' not in session:
        return jsonify({'success': False, 'error': 'Not logged in'})

    perms = get_perms()
    if not perms.get('can_add'):
        return jsonify({'success': False, 'error': 'No permission'})

    data = request.json or {}
    bill_ids = data.get('bill_ids', [])
    if not bill_ids:
        return jsonify({'success': False, 'error': 'No bills selected'})

    conn_seq = get_db()
    cur_seq = get_cursor(conn_seq)

    # Two independent party guards, both server-side because the page-level
    # checks are advisory only (a stale tab or a direct POST bypasses them).
    #   1. every selected bill must be the same customer — no merging parties
    #      into one invoice, whatever the page allowed;
    #   2. every cargo line must still belong to its bill's customer — catches
    #      a declaration whose customer_name was retro-edited after billing.
    parties = model.get_bill_parties(cur_seq, bill_ids)
    if not model.single_party(parties):
        conn_seq.close()
        names = ', '.join(sorted({(p['customer_name'] or '?') for p in parties}))
        return jsonify({'success': False,
                        'error': f'Selected bills belong to different parties ({names}). '
                                 f'An invoice can only cover one customer — '
                                 f'invoice each party separately.'})

    mismatches = model.find_cargo_party_mismatch(cur_seq, bill_ids)
    if mismatches:
        conn_seq.close()
        detail = '; '.join(
            f"{m['bill_number']}: \"{m['cargo_name'] or 'cargo'}\" belongs to "
            f"{m['cargo_customer']}, not {m['bill_customer']}"
            for m in mismatches[:5])
        return jsonify({'success': False,
                        'error': f'Cargo on these bills belongs to another party — {detail}. '
                                 f'The declaration was changed after the bill was made. '
                                 f'Fix the bill before invoicing.'})

    default_series = _get_default_invds_series(cur_seq) or {}

    cur_seq.execute('SELECT bill_date, customer_type, customer_id FROM bill_header WHERE id=%s', [bill_ids[0]])
    first_bill = cur_seq.fetchone()

    customer_type = data.get('customer_type') or (first_bill['customer_type'] if first_bill else '')
    customer_id = data.get('customer_id') or (first_bill['customer_id'] if first_bill else None)
    customer_master = _get_customer_master_snapshot(cur_seq, customer_type, customer_id)

    doc_series_prefix = (
        default_series.get('prefix') or
        (data.get('doc_series_prefix') or 'INV')
    ).strip().rstrip('/').upper()
    doc_series_name = default_series.get('name') or data.get('doc_series_name', '')

    invoice_date = data.get('invoice_date', '')
    if first_bill and first_bill.get('bill_date'):
        invoice_date = str(first_bill['bill_date'])[:10]

    fy_suffix = model.get_financial_year(invoice_date) if invoice_date else ''
    next_seq = model.next_invoice_seq(cur_seq, doc_series_prefix, fy_suffix)
    conn_seq.close()
    invoice_number_override = f'{doc_series_prefix}/{next_seq}'

    invoice_data = {
        'invoice_date': invoice_date,
        'invoice_series': doc_series_prefix,
        'doc_series': doc_series_prefix,
        'doc_series_seq': next_seq,
        'customer_type': data.get('customer_type'),
        'customer_id': data.get('customer_id'),
        'customer_name': data.get('customer_name'),
        'customer_gstin': data.get('customer_gstin'),
        'customer_gst_state_code': data.get('customer_gst_state_code'),
        'customer_gl_code': data.get('customer_gl_code'),
        'customer_pan': data.get('customer_pan') or customer_master.get('pan'),
        'customer_cin': data.get('customer_cin') or customer_master.get('cin'),
        'billing_address': data.get('billing_address') or customer_master.get('billing_address'),
        'customer_city': data.get('customer_city') or customer_master.get('city'),
        'customer_pincode': data.get('customer_pincode') or customer_master.get('pincode'),
        'customer_phone': data.get('customer_phone') or customer_master.get('contact_phone'),
        'customer_email': data.get('customer_email') or customer_master.get('contact_email'),
        'ship_to_name': data.get('ship_to_name'),
        'ship_to_address': data.get('ship_to_address'),
        'ship_to_gstin': data.get('ship_to_gstin'),
        'ship_to_state_code': data.get('ship_to_state_code'),
        'currency_code': data.get('currency_code', 'INR'),
        'exchange_rate': data.get('exchange_rate', 1.0),
        'subtotal': data.get('subtotal'),
        'cgst_amount': data.get('cgst_amount'),
        'sgst_amount': data.get('sgst_amount'),
        'igst_amount': data.get('igst_amount'),
        'tds_amount': data.get('tds_amount', 0),
        'tcs_amount': data.get('tcs_amount', 0),
        'round_off': data.get('round_off', 0),
        'total_amount': data.get('total_amount'),
        'amount_in_words': data.get('amount_in_words'),
        'payment_terms': data.get('payment_terms'),
        'due_date': data.get('due_date'),
        'vessel_name': data.get('vessel_name'),
        'vessel_call_no': data.get('vessel_call_no'),
        'commodity': data.get('commodity'),
        'date_of_berthing': data.get('date_of_berthing'),
        'date_of_sailing': data.get('date_of_sailing'),
        'grt_of_vessel': data.get('grt_of_vessel'),
        'no_of_days': data.get('no_of_days'),
        'cargo_quantity': data.get('cargo_quantity'),
        'no_of_hrs': data.get('no_of_hrs'),
        'created_by': session.get('username'),
        'created_date': __import__('datetime').datetime.now().strftime('%Y-%m-%d'),
        'remarks': data.get('remarks'),
        'virtual_account_id': data.get('virtual_account_id'),
        '_invoice_number_override': invoice_number_override,
    }

    invoice_id, invoice_number = model.create_invoice_from_bills(bill_ids, invoice_data)

    # Queue for SAP posting (async, retried by the background worker) — no loader.
    sap_result = _enqueue_invoice_post(invoice_id, invoice_number)
    invoice = model.get_invoice_by_id(invoice_id) or {}
    _queue_invoice_review_request(
        invoice_id,
        invoice_number,
        invoice.get('customer_name') or invoice_data.get('customer_name'),
        invoice.get('total_amount') or invoice_data.get('total_amount'),
        invoice.get('invoice_status'),
    )

    return jsonify({
        'success': True,
        'id': invoice_id,
        'invoice_number': invoice_number,
        'sap_status': invoice.get('invoice_status') or 'Queued for SAP',
        'sap_message': sap_result.get('message', ''),
    })


# ===== Bill lines (for generate invoice page) =====

@bp.route('/api/module/FINV01/bill-lines/<int:bill_id>')
def get_bill_lines_api(bill_id):
    """Get bill lines for expand in invoice generation page"""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401
    lines = model.get_bill_lines(bill_id)
    return jsonify({'lines': lines})


def _fmt_lueu_timestamp(date_val, time_val):
    """Format LUEU date + time fields as DD-MM-YYYY HH:MM."""
    date_txt = ''
    if date_val:
        s = str(date_val).replace('T', ' ').strip()
        parts = s.split(' ')
        d = parts[0].split('-')
        if len(d) == 3:
            date_txt = f'{d[2]}-{d[1]}-{d[0]}'
        else:
            date_txt = str(date_val)

    time_txt = ''
    if time_val:
        t = str(time_val).replace('T', ' ').strip().split(' ')[-1]
        chunks = t.split(':')
        if len(chunks) >= 2:
            time_txt = f'{chunks[0]}:{chunks[1]}'
        else:
            time_txt = t

    if date_txt and time_txt:
        return f'{date_txt} {time_txt}'
    return date_txt or time_txt


_CH_CODES = ('CHGL01', 'CHGU01')


def _build_display_lines(invoice_lines):
    """
    For the invoice print items table, merge cargo handling lines by rate:
      - All cargo lines at the same rate  → one merged row (summed qty + amount)
      - Cargo lines at different rates    → one row per distinct rate (summed within each)
      - Non-cargo lines                   → passed through unchanged in original order

    Non-cargo lines appear first (original order), then cargo row(s) sorted by rate.
    `invoice_lines` itself is never modified — the caller still uses it for the
    SAC summary and cargo appendix.
    """
    non_cargo = []
    cargo_by_rate = {}   # {rate_key: accumulator dict}

    for line in invoice_lines:
        if line.get('service_code') in _CH_CODES:
            rate_key = round(float(line.get('rate') or 0), 4)
            if rate_key not in cargo_by_rate:
                cargo_by_rate[rate_key] = {
                    'service_code': line['service_code'],
                    'service_name': 'Cargo Handling Services',
                    'sac_code':     line.get('sac_code') or '',
                    'rate':         rate_key,
                    'quantity':     0.0,
                    'line_amount':  0.0,
                    'uom':          line.get('uom') or '',
                }
            cargo_by_rate[rate_key]['quantity']   += float(line.get('quantity')   or 0)
            cargo_by_rate[rate_key]['line_amount'] += float(line.get('line_amount') or 0)
        else:
            non_cargo.append(line)

    cargo_rows = sorted(cargo_by_rate.values(), key=lambda r: r['rate'])
    return non_cargo + cargo_rows


def _get_cargo_handling_details(invoice_id):
    """
    Build cargo appendix rows for invoice print.
    Source: bill_lines.cargo_source_type / cargo_source_id
      VCN_IMPORT  -> vcn_cargo_declaration  -> ldud_anchorage for timing
      VCN_EXPORT  -> vcn_export_cargo_declaration -> ldud_anchorage for timing
      MBC         -> mbc_customer_details   -> mbc_header dates
    """
    conn = get_db()
    cur = get_cursor(conn)
    rows = []
    seen = set()
    try:
        # Get all cargo source references for this invoice
        cur.execute('''
            SELECT DISTINCT bl.cargo_source_type, bl.cargo_source_id,
                            SUM(bl.quantity) OVER (
                                PARTITION BY bl.cargo_source_type, bl.cargo_source_id
                            ) AS billed_qty
            FROM invoice_bill_mapping ibm
            JOIN bill_lines bl ON bl.bill_id = ibm.bill_id
            WHERE ibm.invoice_id = %s
              AND bl.cargo_source_type IS NOT NULL
              AND bl.cargo_source_id IS NOT NULL
        ''', [invoice_id])
        sources = [dict(r) for r in cur.fetchall()]

        for src in sources:
            cstype = src['cargo_source_type']
            csid   = src['cargo_source_id']
            key    = (cstype, csid)
            if key in seen:
                continue
            seen.add(key)

            billed_qty = float(src.get('billed_qty') or 0)

            if cstype in ('VCN_IMPORT', 'VCN_EXPORT'):
                table = 'vcn_cargo_declaration' if cstype == 'VCN_IMPORT' else 'vcn_export_cargo_declaration'
                cur.execute(f'''
                    SELECT cd.vcn_id, cd.cargo_name, cd.bl_no, cd.bl_date,
                           cd.bl_quantity, cd.quantity_uom, cd.customer_name,
                           vh.vcn_doc_num, vh.vessel_name,
                           lh.material_po_number
                    FROM {table} cd
                    JOIN vcn_header vh ON cd.vcn_id = vh.id
                    LEFT JOIN ldud_header lh ON lh.vcn_id = vh.id
                    WHERE cd.id = %s
                ''', [csid])
                decl = cur.fetchone()
                if not decl:
                    continue

                vcn_id = decl['vcn_id']

                # Disch/Load Started = MIN(discharge_started), Completed = MAX(discharge_commenced) from ldud_anchorage
                cur.execute('''
                    SELECT MIN(la.discharge_started)   AS start_dt,
                           MAX(la.discharge_commenced) AS end_dt
                    FROM ldud_anchorage la
                    JOIN ldud_header lh ON la.ldud_id = lh.id
                    WHERE lh.vcn_id = %s
                ''', [vcn_id])
                timing = cur.fetchone()

                def _ts(val):
                    if not val:
                        return ''
                    s = str(val).strip()
                    date_s = s[:10]
                    time_s = s[11:16] if len(s) >= 16 else ''
                    if len(date_s) == 10:
                        date_s = date_s[8:10] + '-' + date_s[5:7] + '-' + date_s[0:4]
                    return (date_s + ' ' + time_s).strip()

                rows.append({
                    'source_type':       'VCN',
                    'source_id':         vcn_id,
                    'vessel_name':       decl['vessel_name'] or '',
                    'vcn_doc_num':       decl['vcn_doc_num'] or '',
                    'consignee':         decl['customer_name'] or '',
                    'cargo':             decl['cargo_name'] or '',
                    'bl_no':             decl['bl_no'] or '',
                    'bl_date':           str(decl['bl_date'] or '')[:10],
                    'quantity':          billed_qty,
                    'uom':               decl['quantity_uom'] or 'MT',
                    'source_type_label': 'MV',
                    'start':             _ts(timing['start_dt'] if timing else None),
                    'end':               _ts(timing['end_dt']   if timing else None),
                    'material_po':       decl.get('material_po_number') or '',
                })

            elif cstype == 'MBC':
                cur.execute('''
                    SELECT cd.mbc_id, cd.cargo_name, cd.bill_of_coastal_goods_no,
                           cd.quantity, cd.customer_name, cd.material_po,
                           mh.doc_num, mh.mbc_name, mh.doc_date,
                           mh.operation_type
                    FROM mbc_customer_details cd
                    JOIN mbc_header mh ON cd.mbc_id = mh.id
                    WHERE cd.id = %s
                ''', [csid])
                decl = cur.fetchone()
                if not decl:
                    continue

                mbc_id = decl['mbc_id']
                op_type = (decl.get('operation_type') or '').lower()

                def _ts(val):
                    if not val:
                        return ''
                    s = str(val).strip()
                    date_s = s[:10]
                    time_s = s[11:16] if len(s) >= 16 else ''
                    if len(date_s) == 10:
                        date_s = date_s[8:10] + '-' + date_s[5:7] + '-' + date_s[0:4]
                    return (date_s + ' ' + time_s).strip()

                if 'export' in op_type:
                    # Load port: loading_commenced → start, loading_completed → end
                    cur.execute('''
                        SELECT MIN(loading_commenced) AS start_dt,
                               MAX(loading_completed)  AS end_dt
                        FROM mbc_load_port_lines
                        WHERE mbc_id = %s
                    ''', [mbc_id])
                else:
                    # Discharge port: unloading_commenced → start, unloading_completed → end
                    cur.execute('''
                        SELECT MIN(unloading_commenced) AS start_dt,
                               MAX(unloading_completed)  AS end_dt
                        FROM mbc_discharge_port_lines
                        WHERE mbc_id = %s
                    ''', [mbc_id])

                mbc_timing = cur.fetchone()

                rows.append({
                    'source_type':       'MBC',
                    'source_id':         mbc_id,
                    'vessel_name':       decl['mbc_name'] or '',
                    'vcn_doc_num':       decl['doc_num'] or '',
                    'consignee':         decl['customer_name'] or '',
                    'cargo':             decl['cargo_name'] or '',
                    'bl_no':             decl['bill_of_coastal_goods_no'] or '',
                    'bl_date':           str(decl['doc_date'] or '')[:10],
                    'quantity':          billed_qty,
                    'uom':               'MT',
                    'source_type_label': 'MBC',
                    'start':             _ts(mbc_timing['start_dt'] if mbc_timing else None),
                    'end':               _ts(mbc_timing['end_dt']   if mbc_timing else None),
                    'material_po':       decl.get('material_po') or '',
                })

        return rows
    except Exception as e:
        log.error(f'[CARGO] Error for invoice {invoice_id}: {e}', exc_info=True)
        return rows
    finally:
        conn.close()


# ===== Print Invoice =====

@bp.route('/module/FINV01/invoice/print/<int:invoice_id>')
def print_invoice(invoice_id):
    """Print invoice"""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    perms = get_perms()
    if not perms.get('can_read'):
        return render_template('no_access.html'), 403

    log.info(f'[PRINT] === Printing invoice {invoice_id} ===')
    invoice = model.get_invoice_by_id(invoice_id)
    if not invoice:
        log.warning(f'[PRINT] Invoice {invoice_id} not found')
        return "Invoice not found", 404

    invoice_lines = model.get_invoice_lines(invoice_id)
    sac_summary = model.get_invoice_sac_summary(invoice_id)
    log.info(f'[PRINT] Invoice {invoice_id}: {len(invoice_lines)} lines, vessel_name={invoice.get("vessel_name")}')

    # If vessel details are missing, re-fetch live from the bill chain
    if not (invoice.get('vessel_name') or '').strip():
        log.info(f'[PRINT] Invoice {invoice_id}: vessel_name missing, re-fetching from bill chain')
        try:
            conn_v = get_db()
            cur_v = get_cursor(conn_v)
            cur_v.execute(
                'SELECT bill_id FROM invoice_bill_mapping WHERE invoice_id=%s ORDER BY id LIMIT 1',
                [invoice_id]
            )
            bm = cur_v.fetchone()
            conn_v.close()
            if bm:
                log.info(f'[PRINT] Invoice {invoice_id}: fetching vessel details from bill {bm["bill_id"]}')
                vd_resp = get_bill_vessel_details(bm['bill_id'])
                vd = vd_resp.get_json() if hasattr(vd_resp, 'get_json') else {}
                log.info(f'[PRINT] Invoice {invoice_id}: vessel details response: {vd}')
                invoice = dict(invoice)
                for fld in ('vessel_name', 'vessel_call_no', 'commodity',
                            'date_of_berthing', 'date_of_sailing', 'grt_of_vessel', 'cargo_quantity'):
                    if vd.get(fld) and not (invoice.get(fld) or ''):
                        invoice[fld] = vd[fld]
            else:
                log.warning(f'[PRINT] Invoice {invoice_id}: no bill mapping found')
        except Exception as e:
            log.error(f'[PRINT] Invoice {invoice_id}: error re-fetching vessel details: {e}')

    # Port config for header GSTIN etc.
    config = get_module_config('FIN01')
    port_config = {
        'seller_gstin': config.get('seller_gstin', ''),
        'seller_legal_name': config.get('seller_legal_name', 'JSW Dharamtar Port Pvt. Ltd.'),
    }

    invoice = dict(invoice)

    # Payment bank: use selected virtual account if set, else fall back to legacy VA / admin bank.
    payment_bank = None
    invoice_display_date = str(invoice.get('invoice_date') or '')[:10]
    conn_b = get_db()
    cur_b = get_cursor(conn_b)
    try:
        customer_master = _get_customer_master_snapshot(
            cur_b,
            invoice.get('customer_type'),
            invoice.get('customer_id')
        )
        if customer_master:
            invoice['customer_pan'] = invoice.get('customer_pan') or customer_master.get('pan') or ''
            invoice['customer_cin'] = invoice.get('customer_cin') or customer_master.get('cin') or ''

        cur_b.execute('''
            SELECT b.bill_date
            FROM invoice_bill_mapping ibm
            JOIN bill_header b ON b.id = ibm.bill_id
            WHERE ibm.invoice_id = %s
            ORDER BY ibm.id
            LIMIT 1
        ''', [invoice_id])
        bill_row = cur_b.fetchone()
        if bill_row and bill_row.get('bill_date'):
            invoice_display_date = str(bill_row['bill_date'])[:10]

        cur_b.execute('SELECT * FROM port_bank_accounts ORDER BY id LIMIT 1')
        base_row = cur_b.fetchone()
        base_bank = dict(base_row) if base_row else None

        va_id = invoice.get('virtual_account_id')
        if va_id and va_id > 0:
            # Specific virtual account chosen at invoice time
            cur_b.execute('SELECT * FROM customer_virtual_accounts WHERE id=%s AND is_active=1', [va_id])
            va_row = cur_b.fetchone()
            if va_row:
                payment_bank = dict(base_bank) if base_bank else {}
                payment_bank.update({
                    'account_holder_name': va_row['account_holder_name'] or payment_bank.get('account_holder_name') or '',
                    'account_number': va_row['account_number'] or payment_bank.get('account_number') or '',
                    'ifsc_code': va_row['ifsc_code'] or payment_bank.get('ifsc_code') or '',
                    'bank_name': va_row['bank_name'] or payment_bank.get('bank_name') or '',
                    'branch_name': va_row['branch_name'] or payment_bank.get('branch_name') or '',
                })
        elif va_id == -1:
            # User explicitly chose admin port bank account
            payment_bank = base_bank
        else:
            # va_id is null: use virtual_account_number from customer master overlaid on admin bank
            legacy_va = (customer_master.get('virtual_account_number') or '').strip() if customer_master else ''
            if legacy_va:
                payment_bank = dict(base_bank) if base_bank else {}
                payment_bank['account_number'] = legacy_va
            else:
                payment_bank = base_bank
        if not payment_bank:
            payment_bank = base_bank
    except Exception:
        payment_bank = None
    conn_b.close()
    invoice['invoice_date'] = invoice_display_date

    current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # Build display lines: cargo handling grouped by rate for the print table
    display_lines = _build_display_lines(invoice_lines)

    # Fetch cargo handling details by tracing bill chain
    cargo_details = _get_cargo_handling_details(invoice_id)
    log.info(f'[PRINT] Invoice {invoice_id}: {len(cargo_details)} cargo detail rows')

    # Check if any service type used in this invoice has triplicate enabled
    show_triplicate = False
    try:
        conn_t = get_db()
        cur_t = get_cursor(conn_t)
        cur_t.execute('''
            SELECT COUNT(*) as cnt FROM invoice_lines il
            JOIN finance_service_types fst ON fst.service_code = il.service_code
            WHERE il.invoice_id = %s AND fst.is_triplicate = 1
        ''', [invoice_id])
        row_t = cur_t.fetchone()
        show_triplicate = (row_t['cnt'] > 0) if row_t else False
        conn_t.close()
    except Exception:
        pass

    return render_template('finv01_invoice_print.html',
                         invoice=invoice,
                         invoice_lines=invoice_lines,
                         display_lines=display_lines,
                         sac_summary=sac_summary,
                         port_config=port_config,
                         payment_bank=payment_bank,
                         current_datetime=current_datetime,
                         cargo_details=cargo_details,
                         show_triplicate=show_triplicate)


# ===== GSTR-1 B2B Export =====

@bp.route('/api/module/FINV01/export/gstr1-b2b', methods=['POST'])
def export_gstr1_b2b():
    """Export selected invoices as GSTR-1 B2B JSON"""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401

    from datetime import datetime
    data = request.json
    invoice_ids = data.get('invoice_ids', [])
    supplier_gstin = data.get('supplier_gstin', '')
    filing_period = data.get('filing_period', '')

    if not invoice_ids:
        return jsonify({'error': 'No invoices selected'}), 400

    b2b = {}
    for inv_id in invoice_ids:
        invoice = model.get_invoice_by_id(inv_id)
        if not invoice:
            continue
        lines = model.get_invoice_lines(inv_id)
        ctin = invoice.get('customer_gstin', '')
        if not ctin:
            continue

        if ctin not in b2b:
            b2b[ctin] = {'ctin': ctin, 'inv': []}

        rate_groups = {}
        for line in lines:
            cgst = float(line.get('cgst_rate') or 0)
            sgst = float(line.get('sgst_rate') or 0)
            igst = float(line.get('igst_rate') or 0)
            rt = igst if igst > 0 else (cgst + sgst)
            if rt not in rate_groups:
                rate_groups[rt] = {'txval': 0, 'camt': 0, 'samt': 0, 'iamt': 0, 'csamt': 0}
            rate_groups[rt]['txval'] += float(line.get('line_amount') or 0)
            rate_groups[rt]['camt'] += float(line.get('cgst_amount') or 0)
            rate_groups[rt]['samt'] += float(line.get('sgst_amount') or 0)
            rate_groups[rt]['iamt'] += float(line.get('igst_amount') or 0)

        itms = [{'num': i+1, 'itm_det': {
            'rt': round(rt, 2), 'txval': round(v['txval'], 2),
            'camt': round(v['camt'], 2), 'samt': round(v['samt'], 2),
            'iamt': round(v['iamt'], 2), 'csamt': round(v['csamt'], 2)
        }} for i, (rt, v) in enumerate(rate_groups.items())]

        inv_date = invoice.get('invoice_date')
        if hasattr(inv_date, 'strftime'):
            inv_date = inv_date.strftime('%d-%m-%Y')
        else:
            inv_date = str(inv_date or '')
            if '-' in inv_date:
                parts = inv_date.split('-')
                if len(parts) == 3 and len(parts[0]) == 4:
                    inv_date = f"{parts[2]}-{parts[1]}-{parts[0]}"

        b2b[ctin]['inv'].append({
            'inum': invoice.get('invoice_number', ''),
            'idt': inv_date,
            'val': round(float(invoice.get('total_amount') or 0), 2),
            'pos': invoice.get('customer_gst_state_code', ''),
            'rchrg': 'N', 'inv_typ': 'R', 'itms': itms
        })

    return jsonify({'gstin': supplier_gstin, 'fp': filing_period, 'b2b': list(b2b.values())})


# ===== SAP Integration =====

@bp.route('/api/module/FINV01/invoice/retry-sap', methods=['POST'])
def retry_sap():
    """Retry SAP posting for a failed invoice"""
    if 'user_id' not in session:
        return jsonify({'success': False, 'error': 'Not logged in'}), 401

    perms = get_perms()
    if not perms.get('can_edit'):
        return jsonify({'success': False, 'error': 'No permission'}), 403

    invoice_id = request.json.get('invoice_id')
    invoice = model.get_invoice_by_id(invoice_id)
    if not invoice:
        return jsonify({'success': False, 'error': 'Invoice not found'}), 404

    if invoice.get('sap_document_number'):
        return jsonify({'success': False, 'error': 'Invoice already posted to SAP'})

    if invoice.get('invoice_status') not in ('SAP Failed', 'Generated', 'Queued for SAP'):
        return jsonify({'success': False, 'error': f"Cannot retry — status is '{invoice.get('invoice_status')}'"})

    # If a post job is already queued/failed, do a one-shot manual send; otherwise
    # queue it fresh. Either way SAP is hit by the worker, not synchronously here.
    post_rows = [q for q in sap_queue.get_queue_for_invoice(invoice_id)
                 if q['job_type'] == 'post' and q['status'] != 'sent']
    if post_rows:
        result = sap_queue.manual_send(post_rows[0]['id'])
        if result.get('ok'):
            return jsonify({'success': True,
                            'sap_document_number': result.get('sap_document_number'),
                            'message': 'Posted to SAP.'})
        return jsonify({'success': False, 'error': result.get('error') or 'SAP send failed'})

    result = _enqueue_invoice_post(invoice_id, invoice['invoice_number'])
    return jsonify({
        'success': result.get('ok', False),
        'queued': result.get('queued', False),
        'message': result.get('message', ''),
    })


@bp.route('/api/module/FINV01/invoice/cancel-sap', methods=['POST'])
def cancel_invoice_sap():
    """Cancel/reverse an SAP-posted invoice (FB08 rule: within 24 hours only)"""
    if 'user_id' not in session:
        return jsonify({'success': False, 'error': 'Not logged in'}), 401

    perms = get_perms()
    if not perms.get('can_edit'):
        return jsonify({'success': False, 'error': 'No permission'}), 403

    invoice_id = request.json.get('invoice_id')
    invoice = model.get_invoice_by_id(invoice_id)
    if not invoice:
        return jsonify({'success': False, 'error': 'Invoice not found'}), 404

    if invoice.get('invoice_status') == 'Cancelled':
        return jsonify({'success': False, 'error': 'Invoice is already cancelled'})

    # The SAP push is a staging push — sap_document_number arrives later via
    # /api/sap/callback. Gate on invoice_status instead so cancellation works
    # even before SAP has acknowledged; SAP matches the reversal by Reference
    # (invoice number), not by document number.
    if invoice.get('invoice_status') not in ('Posted to SAP', 'Posted to GST'):
        return jsonify({'success': False, 'error': 'Invoice is not posted to SAP'})

    # No hard 24h block here: the FB08 reversal must stay available past the
    # window because the SAP team sometimes cancels documents manually on
    # their side, and PORTMAN still needs to record the cancellation and
    # unbill the cargo. The UI offers Create CN as the preferred path once
    # the window has expired.

    invoice_lines = model.get_invoice_lines(invoice_id)
    payload = sap_builder.build_invoice_reversal_payload(invoice, invoice_lines)
    # Queue the reversal — the worker posts to SAP (with retries) and, on
    # success, unbills the cargo and marks the invoice Cancelled.
    qid = sap_queue.enqueue(
        'reversal', 'InvoiceReversal', invoice_id, invoice['invoice_number'],
        payload, invoice_id=invoice_id, created_by=session.get('username'))

    return jsonify({
        'success': True,
        'queued': True,
        'queue_id': qid,
        'message': 'FB08 reversal queued for SAP. The invoice will be marked '
                   'Cancelled and its cargo unbilled once SAP confirms.'
    })




@bp.route('/api/module/FINV01/invoice/create-cancellation-cn', methods=['POST'])
def create_cancellation_cn():
    """Cancel an SAP-posted invoice via a Credit Note after the 24hr FB08
    window has expired.

    Two-layer design:
    1. The SAP payload sent is the *original invoice's* payload with
       Invoice_Credit='C' / Document_type='DG' — Reference stays the
       original invoice_number so SAP can match it against the original.
    2. On SAP success we also create an FDCN01 'CN' record with its own
       unique doc_number (DPPLCN/...) — this is purely a Portbird-side
       artifact for tracking + the FDCN invoice-print reference. It is NOT
       what's pushed to SAP.

    Posting-first / FDCN-second ordering keeps us idempotent: if SAP rejects,
    no orphan FDCN row is left behind and the user can retry."""
    if 'user_id' not in session:
        return jsonify({'success': False, 'error': 'Not logged in'}), 401

    perms = get_perms()
    if not perms.get('can_edit'):
        return jsonify({'success': False, 'error': 'No permission'}), 403

    invoice_id = request.json.get('invoice_id')
    invoice = model.get_invoice_by_id(invoice_id)
    if not invoice:
        return jsonify({'success': False, 'error': 'Invoice not found'}), 404

    if invoice.get('invoice_status') == 'Cancelled':
        return jsonify({'success': False, 'error': 'Invoice is already cancelled'})

    # Gate on invoice_status (not sap_document_number) — the staging push
    # leaves sap_document_number empty until the SAP callback arrives.
    if invoice.get('invoice_status') not in ('Posted to SAP', 'Posted to GST'):
        return jsonify({'success': False, 'error': 'Invoice not posted to SAP — use direct cancellation instead'})

    # Block duplicate cancellation CN against the same invoice.
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("""SELECT doc_number FROM fdcn_header
        WHERE original_invoice_id = %s AND doc_type = 'CN'
          AND creation_type = 'cancellation'
          AND doc_status NOT IN ('Rejected', 'Cancelled')
        LIMIT 1""", [invoice_id])
    existing = cur.fetchone()
    conn.close()
    if existing:
        return jsonify({
            'success': False,
            'error': f'Cancellation CN already exists: {existing["doc_number"]}'
        })

    # Build the SAP payload from the ORIGINAL invoice (sap_builder flips
    # Invoice_Credit→C and Document_type→DG only) and queue it. The worker
    # posts to SAP (with retries) and, on success, creates the Portbird-side
    # FDCN 'CN' record, marks the invoice Cancelled, and unbills the cargo.
    invoice_lines = model.get_invoice_lines(invoice_id)
    payload = sap_builder.build_invoice_credit_note_payload(invoice, invoice_lines)
    qid = sap_queue.enqueue(
        'credit_note', 'InvoiceCreditNote', invoice_id, invoice['invoice_number'],
        payload, invoice_id=invoice_id, created_by=session.get('username'))

    return jsonify({
        'success': True,
        'queued': True,
        'queue_id': qid,
        'message': 'Cancellation Credit Note queued for SAP. Once SAP confirms, '
                   'the CN is created and the invoice marked Cancelled.'
    })


@bp.route('/api/module/FINV01/sap-queue/manual-send', methods=['POST'])
def sap_queue_manual_send():
    """One-shot manual send of a queued/failed SAP job (post / reversal / CN)."""
    if 'user_id' not in session:
        return jsonify({'success': False, 'error': 'Not logged in'}), 401
    perms = get_perms()
    if not perms.get('can_edit'):
        return jsonify({'success': False, 'error': 'No permission'}), 403

    data = request.json or {}
    queue_id = data.get('queue_id')
    if not queue_id and data.get('invoice_id'):
        rows = [q for q in sap_queue.get_queue_for_invoice(data['invoice_id'])
                if q['status'] != 'sent']
        if rows:
            queue_id = rows[0]['id']
    if not queue_id:
        return jsonify({'success': False, 'error': 'No queued SAP job found'})

    result = sap_queue.manual_send(queue_id)
    return jsonify({
        'success': result.get('ok', False),
        'sap_document_number': result.get('sap_document_number'),
        'error': result.get('error'),
    })


def _build_sample_payload(invoice, lines, cancel=False):
    """Fallback sample payload when SAP config is not yet configured.
    Mirrors sap_builder field set exactly — pulls SAP customer code from the
    customer master and service GL fields from FSTM01 so the exported JSON
    reflects populated master data even without an active SAP API config."""
    inv_date = invoice.get('invoice_date') or ''
    if inv_date:
        try:
            inv_date = datetime.strptime(str(inv_date)[:10], '%Y-%m-%d').strftime('%d.%m.%Y')
        except ValueError:
            pass

    total = sap_builder._total_invoice_amount(invoice, lines)
    gstin = invoice.get('customer_gstin') or ''

    cust_info = sap_builder._get_customer_sap_info(
        invoice.get('customer_type'), invoice.get('customer_id')
    )
    company = cust_info.get('company_code') or '5130'
    customer_code = cust_info.get('sap_customer_code') or invoice.get('customer_gl_code') or ''

    svc_codes = {l.get('service_code') for l in lines if l.get('service_code')}
    svc_map   = sap_builder._get_service_gl_map(svc_codes)

    inv_num = invoice.get('invoice_number') or ''
    reference = inv_num

    items = []
    # Reversal payload uses the same items as the original invoice, only Cancellation_Flag changes.
    if True:
        for l in lines:
            svc = svc_map.get(l.get('service_code') or '', {})
            cgst = float(l.get('cgst_amount') or 0)
            sgst = float(l.get('sgst_amount') or 0)
            igst = float(l.get('igst_amount') or 0)
            unit_price = l.get('unit_price') if l.get('unit_price') is not None else l.get('rate')
            qty = l.get('quantity')
            items.append({
                'Reference':        reference[:16],
                'GL_account':       (svc.get('sap_gl_account') or l.get('gl_code') or '')[:10],
                'Amount':           f'{float(l.get("line_amount") or 0):.2f}',
                'Tax_Code':         '',
                'Cost_Center':      l.get('cost_center') or svc.get('sap_cost_center') or '',
                'Plant':            company,
                'Text':             (l.get('service_name') or '')[:25],
                'Profit_Center':    l.get('profit_center') or svc.get('sap_profit_center') or '',
                'HSN_SAC':          (l.get('sac_code') or l.get('hsn_sac') or '')[:16],
                'CGST_AMT':         f'{cgst:.2f}' if cgst else '',
                'SGST_AMT':         f'{sgst:.2f}' if sgst else '',
                'IGST_AMT':         f'{igst:.2f}' if igst else '',
                'IGST_GL':          (svc.get('sap_igst_gl') or '')[:10] if svc.get('sap_igst_gl') else '',
                'SGST_GL':          (svc.get('sap_sgst_gl') or '')[:10] if svc.get('sap_sgst_gl') else '',
                'CGST_GL':          (svc.get('sap_cgst_gl') or '')[:10] if svc.get('sap_cgst_gl') else '',
                'UOM':              l.get('uom') or svc.get('uom') or '',
                'Unit_Price':       f'{float(unit_price):.2f}' if unit_price is not None else '',
                'Quantity':         f'{float(qty):.3f}' if qty is not None else '',
                'TDS_GL':           (svc.get('sap_tds_gl') or '') if (int(l.get('tds_applicable') or 0) or float(l.get('tds_amount') or 0) > 0) else '',
                'TDS_amount':       f'{float(l.get("tds_amount") or 0):.2f}' if l.get('tds_amount') else '',
                'TCS_GL':           (svc.get('sap_tcs_gl') or '') if (int(l.get('tcs_applicable') or 0) or float(l.get('tcs_amount') or 0) > 0) else '',
                'TCS_amount':       f'{-float(l.get("tcs_amount") or 0):.2f}' if l.get('tcs_amount') else '',
                'Round_off_GL':     '',
                'Round_off_Value':  '',
            })

    record = {
        'Invoice_Credit':        'I',
        'Company_code':          company,
        'Invoice_date':          inv_date,
        'Posting_Date':          inv_date,
        'Reference':             (reference if not cancel else reference)[:16],
        'Document_type':         'DR',
        'Customer_Code':         customer_code[:10],
        'Invoice_Amount':        f'{total:.2f}',
        'Business_place':        company,
        'Section_code':          company,
        'Text':                  reference[:25],
        'Document_Header_Text':  (f"REV {inv_num}" if cancel else inv_num)[:25],
        'Payment_Term':          '',
        'Credit_Control_Area':   company,
        'Cancellation_Flag':     'X' if cancel else '',
        'Nature_of_transaction': 'B2B' if gstin else 'B2C',
        'Service_Sale':          sap_builder._service_sale_flag(lines, svc_map),
        'Currency':              invoice.get('currency_code') or 'INR',
        'Payment_term':          '',
        'Baseline_Date':         inv_date,
    }
    record['ITEM'] = items
    return {'Record_Header': [record]}


# ===== SAP JSON Export (temporary — for SAP team review) =====

@bp.route('/api/module/FINV01/invoice/export-sap-json/<int:invoice_id>')
def export_sap_json(invoice_id):
    """Export SAP payload JSON for an invoice (posting + cancellation samples)"""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401

    invoice = model.get_invoice_by_id(invoice_id)
    if not invoice:
        return jsonify({'error': 'Invoice not found'}), 404

    invoice_lines = model.get_invoice_lines(invoice_id)

    # Build posting payload (Y1) — use defaults if SAP config not yet set up
    try:
        posting_payload = sap_builder.build_invoice_payload(invoice, invoice_lines)
        cancellation_payload = sap_builder.build_invoice_reversal_payload(invoice, invoice_lines)
    except ValueError:
        # No SAP config — build with hardcoded defaults for sample review
        posting_payload = _build_sample_payload(invoice, invoice_lines, cancel=False)
        cancellation_payload = _build_sample_payload(invoice, invoice_lines, cancel=True)

    # Enrich with customer & service master details for SAP team reference
    enriched = {
        '_info': 'Sample SAP payloads generated from PORTMAN for SAP team review',
        '_invoice_number': invoice.get('invoice_number'),
        '_customer_name': invoice.get('customer_name'),
        '_customer_gstin': invoice.get('customer_gstin'),
        '_invoice_date': invoice.get('invoice_date'),
        '_total_amount': float(invoice.get('total_amount') or 0),
        'posting_payload': posting_payload,
        'cancellation_payload': cancellation_payload,
    }

    import json
    response = flask_make_response(json.dumps(enriched, indent=2, ensure_ascii=False))
    response.headers['Content-Type'] = 'application/json'
    response.headers['Content-Disposition'] = f'attachment; filename=SAP_Payload_{invoice.get("invoice_number", invoice_id)}.json'
    return response


# ===== Customer/Agent lookup for virtual account + full billing details =====

@bp.route('/api/module/FINV01/customer-bank/<customer_type>/<int:customer_id>')
def get_customer_bank(customer_type, customer_id):
    """Return full billing details + virtual bank accounts from customer_virtual_accounts"""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401
    conn = get_db()
    cur = get_cursor(conn)
    try:
        result = _get_customer_master_snapshot(cur, customer_type, customer_id)

        # Admin port bank account (always the fallback)
        cur.execute('SELECT * FROM port_bank_accounts ORDER BY id LIMIT 1')
        admin_row = cur.fetchone()
        result['admin_bank'] = dict(admin_row) if admin_row else None

        # Get virtual bank accounts from customer_virtual_accounts table
        party_type = 'Agent' if customer_type.lower() == 'agent' else 'Customer'
        cur.execute('''
            SELECT id, account_number, ifsc_code, bank_name, branch_name, account_holder_name
            FROM customer_virtual_accounts
            WHERE party_type=%s AND party_id=%s AND is_active=1
            ORDER BY id
        ''', [party_type, customer_id])
        va_rows = cur.fetchall()
        result['virtual_accounts'] = [dict(r) for r in va_rows]
    except Exception as e:
        log.error(f'[CUSTOMER-BANK] Error: {e}')
        result = {}
    conn.close()
    return jsonify(result)


# ===== Bill → Vessel details (LDUD / VCN / MBC lookup) =====

@bp.route('/api/module/FINV01/bill-vessel-details/<int:bill_id>')
def get_bill_vessel_details(bill_id):
    """Trace a bill back to vessel/cargo details via cargo declarations, then bill_header fallback.
    Collects ALL linked sources and concatenates unique values with commas."""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401
    conn = get_db()
    cur = get_cursor(conn)
    result = {}
    try:
        # Primary path: bill_lines.cargo_source_type / cargo_source_id
        cur.execute('''
            SELECT DISTINCT bl.cargo_source_type, bl.cargo_source_id,
                            bl.service_description AS eu_cargo, SUM(bl.quantity) AS eu_qty
            FROM bill_lines bl
            WHERE bl.bill_id = %s AND bl.cargo_source_type IS NOT NULL AND bl.cargo_source_id IS NOT NULL
            GROUP BY bl.cargo_source_type, bl.cargo_source_id, bl.service_description
        ''', [bill_id])
        cargo_rows = cur.fetchall()

        if cargo_rows:
            vessel_names, vcn_docs, commodities = [], [], []
            berthing_dates, sailing_dates, grts = [], [], []
            total_qty = 0

            for cr in cargo_rows:
                cstype = (cr['cargo_source_type'] or '').upper()
                csid   = cr['cargo_source_id']
                cargo_name = cr['eu_cargo'] or ''
                qty = float(cr['eu_qty'] or 0)
                total_qty += qty
                if cargo_name and cargo_name not in commodities:
                    commodities.append(cargo_name)

                if cstype in ('VCN_IMPORT', 'VCN_EXPORT'):
                    table = 'vcn_cargo_declaration' if cstype == 'VCN_IMPORT' else 'vcn_export_cargo_declaration'
                    cur.execute(f'''
                        SELECT cd.vcn_id, cd.cargo_name,
                               vh.vessel_name, vh.vcn_doc_num, vh.vessel_master_doc,
                               lh.nor_tendered, lh.discharge_completed
                        FROM {table} cd
                        JOIN vcn_header vh ON vh.id = cd.vcn_id
                        LEFT JOIN ldud_header lh ON lh.vcn_id = cd.vcn_id
                        WHERE cd.id = %s
                        ORDER BY lh.id DESC LIMIT 1
                    ''', [csid])
                    row = cur.fetchone()
                    if row:
                        if row['vessel_name'] and row['vessel_name'] not in vessel_names:
                            vessel_names.append(row['vessel_name'])
                        if row['vcn_doc_num'] and row['vcn_doc_num'] not in vcn_docs:
                            vcn_docs.append(row['vcn_doc_num'])
                        if row['cargo_name'] and row['cargo_name'] not in commodities:
                            commodities.append(row['cargo_name'])
                        if row['nor_tendered']:
                            d = _fmt_date(row['nor_tendered'])
                            if d and d not in berthing_dates:
                                berthing_dates.append(d)
                        if row['discharge_completed']:
                            d = _fmt_date(row['discharge_completed'])
                            if d and d not in sailing_dates:
                                sailing_dates.append(d)
                        if row.get('vessel_master_doc'):
                            cur.execute('SELECT gt FROM vessels WHERE doc_num=%s LIMIT 1',
                                        [row['vessel_master_doc']])
                            vrow = cur.fetchone()
                            if vrow and vrow['gt'] and str(vrow['gt']) not in grts:
                                grts.append(str(vrow['gt']))

                elif cstype == 'MBC':
                    cur.execute('''
                        SELECT cd.cargo_name, mh.mbc_name
                        FROM mbc_customer_details cd
                        JOIN mbc_header mh ON mh.id = cd.mbc_id
                        WHERE cd.id = %s
                    ''', [csid])
                    row = cur.fetchone()
                    if row:
                        if row['mbc_name'] and row['mbc_name'] not in vessel_names:
                            vessel_names.append(row['mbc_name'])
                        if row['cargo_name'] and row['cargo_name'] not in commodities:
                            commodities.append(row['cargo_name'])

            result['vessel_name']      = ', '.join(vessel_names)
            result['vessel_call_no']   = ', '.join(vcn_docs)
            result['date_of_berthing'] = ', '.join(berthing_dates)
            result['date_of_sailing']  = ', '.join(sailing_dates)
            result['grt_of_vessel']    = ', '.join(grts) if grts else ''
            result['commodity']        = ', '.join(commodities)
            result['cargo_quantity']   = total_qty if total_qty else None

        else:
            # Fallback: bill_header.source_type / source_id
            cur.execute('SELECT source_type, source_id FROM bill_header WHERE id=%s', [bill_id])
            bill = cur.fetchone()
            if not bill:
                conn.close()
                return jsonify(result)

            src_type = (bill['source_type'] or '').upper()
            src_id   = bill['source_id']

            if src_type == 'VCN' and src_id:
                cur.execute('''
                    SELECT v.vessel_name, v.vcn_doc_num, v.vessel_master_doc,
                           l.nor_tendered, l.discharge_completed
                    FROM vcn_header v
                    LEFT JOIN ldud_header l ON l.vcn_id = v.id
                    WHERE v.id = %s
                    ORDER BY l.id DESC LIMIT 1
                ''', [src_id])
                row = cur.fetchone()
                if row:
                    result['vessel_name']      = row['vessel_name'] or ''
                    result['vessel_call_no']   = row['vcn_doc_num'] or ''
                    result['date_of_berthing'] = _fmt_date(row['nor_tendered'])
                    result['date_of_sailing']  = _fmt_date(row['discharge_completed'])
                    if row.get('vessel_master_doc'):
                        cur.execute('SELECT gt FROM vessels WHERE doc_num=%s LIMIT 1',
                                    [row['vessel_master_doc']])
                        vrow = cur.fetchone()
                        result['grt_of_vessel'] = vrow['gt'] if vrow else ''
                cur.execute('''
                    SELECT cargo_name, bl_quantity FROM vcn_cargo_declaration
                    WHERE vcn_id = %s ORDER BY id LIMIT 1
                ''', [src_id])
                cargo = cur.fetchone()
                if cargo:
                    result['commodity']      = cargo['cargo_name'] or ''
                    result['cargo_quantity'] = cargo['bl_quantity']

            elif src_type == 'MBC' and src_id:
                cur.execute('''
                    SELECT mbc_name, cargo_name, bl_quantity
                    FROM mbc_header WHERE id = %s
                ''', [src_id])
                row = cur.fetchone()
                if row:
                    result['vessel_name']    = row['mbc_name'] or ''
                    result['commodity']      = row['cargo_name'] or ''
                    result['cargo_quantity'] = row['bl_quantity']

    except Exception:
        pass
    conn.close()
    return jsonify(result)


def _fmt_date(val):
    """Extract YYYY-MM-DD from a datetime/text value"""
    if not val:
        return ''
    s = str(val)
    if len(s) >= 10:
        return s[:10]
    return s


# ===== Invoice Doc Series =====

def _ensure_invoice_doc_series(cur):
    cur.execute('''
        CREATE TABLE IF NOT EXISTS invoice_doc_series (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL,
            prefix TEXT NOT NULL,
            is_default BOOLEAN DEFAULT FALSE
        )
    ''')


@bp.route('/api/module/FINV01/doc-series')
def get_doc_series():
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401
    conn = get_db()
    cur = get_cursor(conn)
    _ensure_invoice_doc_series(cur)
    conn.commit()
    cur.execute('SELECT * FROM invoice_doc_series ORDER BY name')
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return jsonify(rows)


@bp.route('/api/module/FINV01/doc-series/save', methods=['POST'])
def save_doc_series():
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401
    perms = get_perms()
    if not perms.get('can_edit') and not perms.get('can_add'):
        return jsonify({'error': 'No permission'}), 403
    data = request.json
    conn = get_db()
    cur = get_cursor(conn)
    _ensure_invoice_doc_series(cur)
    is_default = bool(data.get('is_default', False))
    if is_default:
        cur.execute('UPDATE invoice_doc_series SET is_default=FALSE WHERE is_default=TRUE')
    if data.get('id'):
        cur.execute('UPDATE invoice_doc_series SET name=%s, prefix=%s, is_default=%s WHERE id=%s',
                    [data['name'], data['prefix'], is_default, data['id']])
        row_id = data['id']
    else:
        cur.execute('INSERT INTO invoice_doc_series (name, prefix, is_default) VALUES (%s,%s,%s) RETURNING id',
                    [data['name'], data['prefix'], is_default])
        row_id = cur.fetchone()['id']
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'id': row_id})


@bp.route('/api/module/FINV01/doc-series/delete', methods=['POST'])
def delete_doc_series():
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401
    perms = get_perms()
    if not perms.get('can_delete'):
        return jsonify({'error': 'No permission'}), 403
    row_id = request.json.get('id')
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('DELETE FROM invoice_doc_series WHERE id=%s', [row_id])
    conn.commit()
    conn.close()
    return jsonify({'success': True})


# ===== Port Config (shared) =====

@bp.route('/api/module/FINV01/port-config')
def get_port_config():
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401
    config = get_module_config('FIN01')
    conn = get_db()
    cur = get_cursor(conn)
    # Fetch the primary port bank account for invoice footer
    bank = {}
    try:
        _ensure_invoice_doc_series(cur)  # ensure tables exist
        cur.execute('''
            SELECT * FROM port_bank_accounts ORDER BY id LIMIT 1
        ''')
        row = cur.fetchone()
        if row:
            bank = dict(row)
    except Exception:
        pass
    conn.commit()
    conn.close()
    return jsonify({
        'port_gst_state_code': config.get('port_gst_state_code', ''),
        'port_gstin': config.get('port_gstin', ''),
        'seller_gstin': config.get('seller_gstin', ''),
        'seller_legal_name': config.get('seller_legal_name', ''),
        'bank': bank,
    })
