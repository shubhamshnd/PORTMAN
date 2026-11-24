"""RP03 — Arrival Planner: when does each barge / MBC reach Dharamtar.

Nothing in the schema stores an ETA. mbc_load_port_lines.eta exists but the
MBC01 form never writes it (0 of 253 rows filled), so an arrival time has to be
derived: take the latest milestone a trip has stamped and add the median
observed time from that milestone to arrival, measured over this port's own
completed trips.

Two sources, in this order:
  1. The v6 model (eta_predict.py): one regressor per milestone fed the tide
     (fitted from TM01 tide_master) and the queue of other open trips. Replayed
     over 19 Sep - 22 Nov 2026 it beat the median on every barge milestone.
  2. The median table below, whenever the model has no answer or cannot load.
p10/p90 ride along so the board can show a band rather than a false point
estimate. eta_service.py logs every ETA and retrains weekly; see eta_learn.py.
"""
import logging
from collections import defaultdict
from datetime import datetime, timedelta
from statistics import median, quantiles

from database import get_db, get_cursor

log = logging.getLogger(__name__)

# Milestone chains, earliest first. A trip is placed at the LAST one it stamped.
BARGE_STAGES = [
    ('trip_start',           'Trip start'),
    ('along_side_vessel',    'Alongside mother vessel'),
    ('commenced_loading',    'Loading commenced'),
    ('completed_loading',    'Loading completed'),
    ('cast_off_mv',          'Cast off mother vessel'),
    ('anchored_gull_island', 'Anchored Gull Island'),
    ('aweigh_gull_island',   'Aweigh Gull Island'),
    ('amf_at_port',          'All made fast at port'),
]
# sailed_out_load_port is excluded on purpose: it is stamped AFTER arrival in
# 914 of 915 rows — the field is populated wrong upstream, so it would drag
# every MBC's ETA into the past.
MBC_STAGES = [
    ('arrived_load_port',      'Arrived load port'),
    ('lp_alongside_berth',     'Alongside load berth'),
    ('lp_loading_commenced',   'Loading commenced'),
    ('lp_loading_completed',   'Loading completed'),
    ('cast_off_load_port',     'Cast off load port'),
    ('arrival_gull_island',    'Arrived Gull Island'),
    ('departure_gull_island',  'Departed Gull Island'),
]

BARGE_ARRIVAL, BARGE_DONE = 'along_side_berth', 'completed_discharge_berth'
MBC_ARRIVAL,   MBC_DONE   = 'vessel_arrival_port', 'unloading_completed'

MIN_N = 30          # below this a stage has no trustworthy median
MAX_H = 240.0       # a leg longer than 10 days is a data error, not a voyage
STALE_DAYS = 7      # a trip untouched this long is an abandoned document
OPS_HOUR = 7        # the operational day runs 07:00 → 07:00
HORIZON_DAYS = 2    # today + tomorrow
SLOT_MIN = 30       # grid resolution
SLOTS_PER_DAY = 24 * 60 // SLOT_MIN
NO_PRIORITY = 9999  # unranked trips sort after ranked ones inside a slot
# ponytail: a berth double-book is called on ETAs closer together than this.
# There is no discharge-duration model — tune the constant, or measure per-berth
# occupancy from completed trips if the false alarms start costing attention.
BERTH_GAP_H = 6


def _dt(val):
    """Parse the mixed TEXT/TIMESTAMP datetimes these tables hold."""
    if not val:
        return None
    if isinstance(val, datetime):
        return val
    try:
        return datetime.fromisoformat(str(val).strip())
    except Exception:
        return None


def _int(val):
    """Form fields arrive as strings; blank means 'not set', not zero."""
    try:
        return int(str(val).strip())
    except (TypeError, ValueError):
        return None


def _plans():
    """{(kind, trip_id): plan row} — the only table this module owns."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT p.trip_kind, p.trip_id, p.berth_id, p.route_id, p.priority,
               p.planned_at, p.remarks, p.updated_by, p.updated_at,
               b.berth_name, r.route_name
        FROM rp03_arrival_plan p
        LEFT JOIN port_berth_master b ON b.id = p.berth_id
        LEFT JOIN conveyor_routes   r ON r.id = p.route_id
    ''')
    out = {(r['trip_kind'], r['trip_id']): dict(r) for r in cur.fetchall()}
    cur.close()
    conn.close()
    return out


def _options():
    """Berth and route choices for the assign form, straight from the masters."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT id, berth_name FROM port_berth_master '
                'ORDER BY berth_sequence NULLS LAST, berth_name')
    berths = [dict(r) for r in cur.fetchall()]
    cur.execute('SELECT id, route_name FROM conveyor_routes '
                'WHERE is_active = 1 ORDER BY route_name')
    routes = [dict(r) for r in cur.fetchall()]
    cur.close()
    conn.close()
    return berths, routes


def save_plan(form, user):
    """Upsert one trip's plan. Everything blank deletes the row — a plan with
    no berth, route, priority, pin or remark is the same as no plan."""
    kind = (form.get('kind') or '').strip().upper()
    if kind not in ('BARGE', 'MBC'):
        raise ValueError('unknown trip kind: %r' % kind)
    trip_id = _int(form.get('trip_id'))
    if trip_id is None:
        raise ValueError('missing trip id')
    berth_id = _int(form.get('berth_id'))
    route_id = _int(form.get('route_id'))
    priority = _int(form.get('priority'))
    planned = _dt(form.get('planned_at'))
    remarks = (form.get('remarks') or '').strip() or None

    conn = get_db()
    cur = get_cursor(conn)
    if not any((berth_id, route_id, priority, planned, remarks)):
        cur.execute('DELETE FROM rp03_arrival_plan WHERE trip_kind = %s AND trip_id = %s',
                    [kind, trip_id])
    else:
        cur.execute('''
            INSERT INTO rp03_arrival_plan
                (trip_kind, trip_id, berth_id, route_id, priority, planned_at,
                 remarks, updated_by, updated_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
            ON CONFLICT (trip_kind, trip_id) DO UPDATE SET
                berth_id   = EXCLUDED.berth_id,
                route_id   = EXCLUDED.route_id,
                priority   = EXCLUDED.priority,
                planned_at = EXCLUDED.planned_at,
                remarks    = EXCLUDED.remarks,
                updated_by = EXCLUDED.updated_by,
                updated_at = EXCLUDED.updated_at
        ''', [kind, trip_id, berth_id, route_id, priority, planned, remarks,
              user, datetime.now()])
    conn.commit()
    cur.close()
    conn.close()


def _fetch():
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT b.id, b.barge_name, b.trip_number, b.cargo_name, b.discharge_quantity,
               b.contractor_name, h.doc_num AS doc_num, h.vessel_name, h.operation_type,
               b.trip_start, b.along_side_vessel, b.commenced_loading,
               b.completed_loading, b.cast_off_mv, b.anchored_gull_island,
               b.aweigh_gull_island, b.amf_at_port,
               b.along_side_berth, b.commence_discharge_berth, b.completed_discharge_berth
        FROM ldud_barge_lines b
        JOIN ldud_header h ON h.id = b.ldud_id
        WHERE COALESCE(b.barge_name, '') <> ''
    ''')
    barges = cur.fetchall()
    cur.execute('''
        SELECT m.id, m.mbc_name, m.doc_num, m.cargo_name, m.bl_quantity,
               m.quantity_uom, m.load_port, m.operation_type,
               lp.arrived_load_port,
               lp.alongside_berth      AS lp_alongside_berth,
               lp.loading_commenced    AS lp_loading_commenced,
               lp.loading_completed    AS lp_loading_completed,
               lp.cast_off_load_port,
               dp.arrival_gull_island, dp.departure_gull_island,
               dp.vessel_arrival_port, dp.unloading_commenced, dp.unloading_completed,
               dp.vessel_unloading_berth
        FROM mbc_header m
        LEFT JOIN LATERAL (SELECT * FROM mbc_load_port_lines      x WHERE x.mbc_id = m.id ORDER BY x.id LIMIT 1) lp ON TRUE
        LEFT JOIN LATERAL (SELECT * FROM mbc_discharge_port_lines x WHERE x.mbc_id = m.id ORDER BY x.id LIMIT 1) dp ON TRUE
        WHERE COALESCE(m.mbc_name, '') <> ''
    ''')
    mbcs = cur.fetchall()
    cur.close()
    conn.close()
    return barges, mbcs


def _leg_table(rows, stages, arrival_col):
    """{stage: (p10, p50, p90, n)} hours from that milestone to arrival."""
    samples = defaultdict(list)
    for r in rows:
        arrival = _dt(r[arrival_col])
        if not arrival:
            continue
        for col, _label in stages:
            t = _dt(r.get(col))
            if not t:
                continue
            hrs = (arrival - t).total_seconds() / 3600
            if 0 <= hrs <= MAX_H:
                samples[col].append(hrs)
    table = {}
    for col, vals in samples.items():
        if len(vals) < MIN_N:
            continue
        deciles = quantiles(vals, n=10)
        table[col] = (round(deciles[0], 1), round(median(vals), 1),
                      round(deciles[8], 1), len(vals))
    return table


def _model_etas(barges, mbcs, now):
    """{(kind, trip_id): v6 prediction}, or {} — the board must still open on
    plain medians if scikit-learn or the model file is missing on this server."""
    try:
        from .eta_predict import predict_board
        return predict_board(barges, mbcs, now=now)
    except Exception:
        log.exception('RP03: v6 ETA model unavailable, using median legs')
        return {}


def _learning(barges, mbcs, now):
    """Track record of the logged ETAs, or None before eta_service has run."""
    try:
        from .eta_learn import status
        return status(barges, mbcs, now)
    except Exception:
        log.exception('RP03: ETA track record unavailable')
        return None


def _open_trips(rows, stages, table, arrival_col, done_col, spec, now, plans, etas):
    """Live trips, each tagged 'inbound' / 'arrived' / 'unknown'.

    One flat list rather than three, so the caller can pick the latest trip per
    vessel ACROSS buckets — otherwise a barge on a fresh run shows up again as
    an older, never-discharged trip sitting in 'at port'."""
    labels = dict(stages)
    trips = []
    for r in rows:
        arrival = _dt(r[arrival_col])
        done = _dt(r[done_col])
        if arrival and done:
            continue                        # discharged and gone
        last = None
        for col, _label in stages:
            t = _dt(r.get(col))
            if t:
                last = (col, t)
        seen_at = arrival or (last[1] if last else None)
        if not seen_at or seen_at < now - timedelta(days=STALE_DAYS):
            continue                        # abandoned document, not a voyage
        trip = {
            'id':         r['id'],
            'kind':       spec['kind'],
            'name':       (r[spec['name']] or '').strip(),
            'doc':        r.get('doc_num') or '',
            'cargo':      (r.get('cargo_name') or '').strip(),
            'qty':        r.get(spec['qty']),
            'uom':        (r.get('quantity_uom') or 'MT'),
            'via':        (r.get(spec['via']) or '').strip(),
            'trip_label': spec['trip_label'](r),
            'last_stage': labels.get(last[0], '') if last else '',
            'last_at':    last[1] if last else None,
            'sort_at':    seen_at,
        }
        plan = plans.get((spec['kind'], r['id'])) or {}
        trip.update({
            'berth_id': plan.get('berth_id'), 'berth': plan.get('berth_name'),
            'route_id': plan.get('route_id'), 'route': plan.get('route_name'),
            'priority': plan.get('priority'), 'remarks': plan.get('remarks'),
            'planned_by': plan.get('updated_by'),
        })
        if arrival:
            trip.update({'bucket': 'arrived', 'arrived_at': arrival})
            trips.append(trip)
            continue
        band = table.get(last[0]) if last else None
        if band:
            p10, p50, p90, n = band
            trip.update({
                'eta':     last[1] + timedelta(hours=p50),
                'eta_p10': last[1] + timedelta(hours=p10),
                'eta_p90': last[1] + timedelta(hours=p90),
                'leg_p50': p50,
                'leg_n':   n,
            })
        pred = etas.get((spec['kind'], r['id']))
        if pred:
            # A median-only stage returns no band of its own; keep the table's.
            trip.update({'eta': pred['eta'], 'eta_source': 'model',
                         'missed_tide': pred['missed_tide']})
            if pred['eta_p10'] and pred['eta_p90']:
                trip.update({'eta_p10': pred['eta_p10'], 'eta_p90': pred['eta_p90']})
        pinned = _dt(plan.get('planned_at'))
        if pinned:
            # The planner's slot wins; the derived ETA stays visible beside it so
            # the plan drifting away from reality is something you can see.
            trip['derived_eta'] = trip.get('eta')
            if trip.get('eta'):
                trip['drift_min'] = round((trip['eta'] - pinned).total_seconds() / 60)
            trip.update({'eta': pinned, 'pinned': True})
        # No band and no pin means either no milestone at all, or one with too
        # few past samples to average — show it, don't guess. Hiding it loses
        # the cargo.
        trip['bucket'] = 'inbound' if trip.get('eta') else 'unknown'
        trips.append(trip)
    return trips


def _flag_berth_clashes(trips):
    """Mark trips sharing a berth within BERTH_GAP_H of each other."""
    by_berth = defaultdict(list)
    for t in trips:
        when = t.get('eta') or t.get('arrived_at')
        if t.get('berth_id') and when:
            by_berth[t['berth_id']].append((when, t))
    for group in by_berth.values():
        group.sort(key=lambda p: p[0])
        for (t1, a), (t2, b) in zip(group, group[1:]):
            if (t2 - t1).total_seconds() / 3600 < BERTH_GAP_H:
                a['berth_clash'] = b['berth_clash'] = True


def _latest_per_vessel(trips):
    """One row per barge / MBC — the trip with the most recent milestone."""
    best = {}
    for t in trips:
        key = (t['kind'], t['name'].upper())
        if key not in best or t['sort_at'] > best[key]['sort_at']:
            best[key] = t
    return list(best.values())


def _slot_time(ops_start, idx):
    return (ops_start + timedelta(minutes=SLOT_MIN * idx)).strftime('%H:%M')


def _merge_rows(slots, ops_start, live_idx):
    """One row per slot, except that runs of slots empty on EVERY day collapse
    into a single 'no arrivals' row. The live slot never merges, so the now-line
    stays on the board even in the middle of a quiet stretch."""
    def quiet(i):
        return not any(slots[i]) and i != live_idx

    rows, i, n = [], 0, len(slots)
    while i < n:
        if not quiet(i):
            rows.append({'label': '%s–%s' % (_slot_time(ops_start, i),
                                             _slot_time(ops_start, i + 1)),
                         'now': i == live_idx, 'empty': False, 'span': 1,
                         'cells': slots[i]})
            i += 1
            continue
        j = i
        while j < n and quiet(j):
            j += 1
        rows.append({'label': '%s – %s' % (_slot_time(ops_start, i),
                                           _slot_time(ops_start, j)),
                     'now': False, 'empty': True, 'span': j - i, 'cells': None})
        i = j
    return rows


def _slot_order(t):
    """Inside a slot: planner priority first, then ETA. Unranked sorts last."""
    return (t['priority'] if t.get('priority') else NO_PRIORITY, t['eta'])


def board(now=None):
    """Two operational days of half-hourly arrival slots, plus what falls outside."""
    now = now or datetime.now()
    barges, mbcs = _fetch()
    plans = _plans()

    etas = _model_etas(barges, mbcs, now)
    barge_table = _leg_table(barges, BARGE_STAGES, BARGE_ARRIVAL)
    mbc_table   = _leg_table(mbcs,   MBC_STAGES,   MBC_ARRIVAL)

    trips = _latest_per_vessel(
        _open_trips(barges, BARGE_STAGES, barge_table, BARGE_ARRIVAL, BARGE_DONE,
                    {'kind': 'BARGE', 'name': 'barge_name', 'qty': 'discharge_quantity',
                     'via': 'vessel_name',
                     'trip_label': lambda r: 'Trip %s' % r['trip_number'] if r.get('trip_number') else ''},
                    now, plans, etas)
        + _open_trips(mbcs, MBC_STAGES, mbc_table, MBC_ARRIVAL, MBC_DONE,
                      {'kind': 'MBC', 'name': 'mbc_name', 'qty': 'bl_quantity',
                       'via': 'load_port', 'trip_label': lambda r: r.get('doc_num') or ''},
                      now, plans, etas))
    inbound = [t for t in trips if t['bucket'] == 'inbound']
    at_port = [t for t in trips if t['bucket'] == 'arrived']
    unknown = [t for t in trips if t['bucket'] == 'unknown']
    _flag_berth_clashes(inbound + at_port)
    missed = sum(1 for t in inbound if t.get('missed_tide'))

    # Operational day: before 07:00 we are still inside yesterday's ops day.
    ops_start = now.replace(hour=OPS_HOUR, minute=0, second=0, microsecond=0)
    if now.hour < OPS_HOUR:
        ops_start -= timedelta(days=1)
    slot_secs = SLOT_MIN * 60

    slots = [[[] for _ in range(HORIZON_DAYS)] for _ in range(SLOTS_PER_DAY)]
    # Arrived trips show at the hour they arrived, today only. Older ones are not
    # listed, but still count for clashes because they are holding the berth.
    arrived = [t for t in sorted(at_port, key=lambda x: x['arrived_at'])
               if ops_start <= t['arrived_at'] < ops_start + timedelta(days=1)]
    for t in arrived:
        slots[int((t['arrived_at'] - ops_start).total_seconds() // slot_secs)][0].append(t)
    clashes = sum(1 for t in inbound + arrived if t.get('berth_clash'))

    overdue, beyond = [], []
    for t in sorted(inbound, key=_slot_order):
        idx = int((t['eta'] - ops_start).total_seconds() // slot_secs)
        if idx < 0:
            overdue.append(t)               # ETA has passed, arrival not stamped
        elif idx >= SLOTS_PER_DAY * HORIZON_DAYS:
            beyond.append(t)
        else:
            slots[idx % SLOTS_PER_DAY][idx // SLOTS_PER_DAY].append(t)

    days = [{'start': ops_start + timedelta(days=d),
             'end':   ops_start + timedelta(days=d + 1),
             'label': ('Today', 'Tomorrow')[d] if d < 2 else 'Day %d' % (d + 1)}
            for d in range(HORIZON_DAYS)]
    live_idx = int((now - ops_start).total_seconds() // slot_secs)
    berths, routes = _options()

    basis = ([{'kind': 'BARGE', 'stage': lbl, 'p50': barge_table[c][1], 'n': barge_table[c][3]}
              for c, lbl in BARGE_STAGES if c in barge_table] +
             [{'kind': 'MBC', 'stage': lbl, 'p50': mbc_table[c][1], 'n': mbc_table[c][3]}
              for c, lbl in MBC_STAGES if c in mbc_table])

    return {
        'now': now, 'ops_start': ops_start, 'days': days,
        'rows': _merge_rows(slots, ops_start, live_idx),
        'slots': slots, 'overdue': overdue, 'beyond': beyond,
        'arrived': arrived,
        'unknown': sorted(unknown, key=lambda x: x['sort_at'], reverse=True),
        'inbound_count': sum(len(c) for row in slots for c in row) - len(arrived),
        'basis': basis, 'berths': berths, 'routes': routes, 'clashes': clashes,
        'missed': missed, 'modelled': sum(1 for t in inbound if t.get('eta_source') == 'model'),
        'slot_secs': slot_secs, 'learning': _learning(barges, mbcs, now),
    }


def _busy_stamp():
    """Where the data actually ends — so the selfcheck has live rows to assert
    on even against a restored dump whose data stopped months ago. The 95th
    percentile, not the max: a handful of mistyped future dates would otherwise
    park 'now' in a gap where every real trip is already stale."""
    barges, mbcs = _fetch()
    stamps = [_dt(r.get(c)) for r in barges for c, _l in BARGE_STAGES]
    stamps += [_dt(r.get(c)) for r in mbcs for c, _l in MBC_STAGES]
    epochs = sorted(s.timestamp() for s in stamps if s)
    if not epochs:
        return datetime.now()
    return datetime.fromtimestamp(epochs[int(len(epochs) * 0.95)])


def _check_merge():
    """Pure check on the row merging — no DB, so it runs on any dataset."""
    start = datetime(2026, 1, 1, OPS_HOUR)
    empty = [[[] for _ in range(HORIZON_DAYS)] for _ in range(SLOTS_PER_DAY)]

    rows = _merge_rows(empty, start, -1)         # nothing anywhere, no live slot
    assert len(rows) == 1 and rows[0]['empty'] and rows[0]['span'] == SLOTS_PER_DAY

    busy = [[[] for _ in range(HORIZON_DAYS)] for _ in range(SLOTS_PER_DAY)]
    busy[5][1].append({'name': 'X'})             # occupied on the second day only
    rows = _merge_rows(busy, start, 9)
    assert sum(r['span'] for r in rows) == SLOTS_PER_DAY, 'rows must cover the day'
    assert [r['span'] for r in rows] == [5, 1, 3, 1, SLOTS_PER_DAY - 10]
    assert not rows[1]['empty'] and rows[1]['cells'] is busy[5]
    assert rows[3]['now'] and not rows[3]['empty'], 'the live slot must never merge'
    assert all(r['cells'] is None for r in rows if r['empty'])

    a = {'berth_id': 1, 'eta': start, 'name': 'A'}
    b = {'berth_id': 1, 'eta': start + timedelta(hours=BERTH_GAP_H - 1), 'name': 'B'}
    c = {'berth_id': 1, 'eta': start + timedelta(hours=BERTH_GAP_H * 3), 'name': 'C'}
    _flag_berth_clashes([a, b, c])
    assert a.get('berth_clash') and b.get('berth_clash') and not c.get('berth_clash')


def _demo(now=None):
    """Self-check: run against the live DB and assert the board holds together."""
    _check_merge()
    b = board(now=now)
    if not b['inbound_count'] and now is None:
        # Every trip is older than STALE_DAYS: rewind to where the data ends so
        # the assertions below actually exercise a populated board.
        return _demo(now=_busy_stamp())
    assert len(b['slots']) == SLOTS_PER_DAY
    assert all(len(r) == HORIZON_DAYS for r in b['slots'])
    assert b['ops_start'].hour == OPS_HOUR
    assert b['ops_start'] <= b['now'] < b['ops_start'] + timedelta(days=1)
    assert sum(r['span'] for r in b['rows']) == SLOTS_PER_DAY
    assert not any(r['empty'] and r['cells'] for r in b['rows'])
    on_board = [t for row in b['slots'] for cell in row for t in cell]
    for t in b['arrived']:
        i = int((t['arrived_at'] - b['ops_start']).total_seconds() // b['slot_secs'])
        assert t in b['slots'][i][0], 'arrival not in its own slot today'
    placed = [t for t in on_board if t['bucket'] == 'inbound']
    assert len(on_board) == len(placed) + len(b['arrived'])
    assert len(placed) == b['inbound_count']
    for t in placed:
        assert b['ops_start'] <= t['eta'] < b['ops_start'] + timedelta(days=HORIZON_DAYS)
        if t.get('pinned'):
            if t.get('derived_eta'):
                assert t['drift_min'] == round(
                    (t['derived_eta'] - t['eta']).total_seconds() / 60), 'drift mismatch'
        else:
            assert t['eta_p10'] <= t['eta'] <= t['eta_p90'], 'band inverted for %s' % t['name']
            if t.get('eta_source') == 'model':
                assert t['eta'] >= b['now'].replace(second=0, microsecond=0), 'model ETA in the past'
            # a stamp mistyped into the future is not a milestone to the model
            assert t['eta'] >= min(t['last_at'], b['now'].replace(second=0, microsecond=0)), 'ETA before its own milestone'
    for cell in (c for row in b['slots'] for c in row):
        ranks = [_slot_order(t) for t in cell if t['bucket'] == 'inbound']
        assert ranks == sorted(ranks), 'priority order broken inside a slot'
    for t in b['overdue']:
        assert t['eta'] < b['ops_start']
    names = [(t['kind'], t['name'].upper())
             for t in placed + b['overdue'] + b['beyond'] + b['arrived'] + b['unknown']]
    assert len(names) == len(set(names)), 'a vessel appears on the board twice'
    for row in b['basis']:
        assert row['n'] >= MIN_N and 0 <= row['p50'] <= MAX_H
    planned = [t for t in placed + b['overdue'] + b['arrived'] + b['unknown']
               if t.get('berth_id') or t.get('route_id') or t.get('priority')]
    print('%d of the inbound trips use the v6 model, %d missed a tide' % (b['modelled'], b['missed']))
    print('RP03 selfcheck OK @ %s — %d inbound in %d h, %d overdue, %d arrived today, %d no-ETA, '
          'basis %d stages, %d rows (%d merged), %d planned, %d berths, %d routes'
          % (b['now'].strftime('%Y-%m-%d %H:%M'), b['inbound_count'],
             24 * HORIZON_DAYS, len(b['overdue']), len(b['arrived']),
             len(b['unknown']), len(b['basis']), len(b['rows']),
             sum(1 for r in b['rows'] if r['empty']), len(planned),
             len(b['berths']), len(b['routes'])))


if __name__ == '__main__':
    _demo()
