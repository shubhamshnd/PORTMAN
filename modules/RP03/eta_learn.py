"""RP03 self-learning: log every ETA, compare with the real arrival, retrain weekly.

Runs from eta_service.py, never inside a web request. The web app only reads
eta_live.joblib, which is replaced atomically and only when a retrained model
passes the checks in eta_fit.retrain(). eta_prev.joblib is the one before it:
delete eta_live.joblib (or copy eta_prev over it) to roll back.
"""
import json
import logging
import os
from datetime import datetime, timedelta
from statistics import median

import joblib
import pandas as pd

from database import get_db, get_cursor
from . import eta_fit, model
from .eta_features import RP03_COLUMNS, hours
from .eta_predict import HERE, load_bundle, predict_board

log = logging.getLogger(__name__)

LIVE, PREV = HERE / 'eta_live.joblib', HERE / 'eta_prev.joblib'
RETRAIN_DAYS = 7
REPORT_DAYS = 30


def _rows(sql, args=()):
    conn = get_db()
    try:
        cur = get_cursor(conn)
        cur.execute(sql, args)
        return [dict(r) for r in cur.fetchall()]
    finally:
        conn.close()


def log_predictions(now=None):
    """Store the first ETA given for each trip at each milestone. Returns rows added."""
    now = now or datetime.now()
    barges, mbcs = model._fetch()
    etas = predict_board(barges, mbcs, now=now)
    legs = {'BARGE': model._leg_table(barges, model.BARGE_STAGES, model.BARGE_ARRIVAL),
            'MBC':   model._leg_table(mbcs,   model.MBC_STAGES,   model.MBC_ARRIVAL)}
    column = {k.upper(): {v: c for c, v in m.items()} for k, m in RP03_COLUMNS.items()}
    trained = load_bundle().get('trained_until')
    conn = get_db()
    try:
        cur = get_cursor(conn)
        added = 0
        for (kind, trip_id), e in etas.items():
            leg = legs[kind].get(column[kind][e['stage']])
            cur.execute('''
                INSERT INTO rp03_eta_log (trip_kind, trip_id, stage, stage_at, predicted_at, eta,
                                          eta_p10, eta_p90, median_eta, model_trained_until)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (trip_kind, trip_id, stage) DO NOTHING
            ''', [kind, trip_id, e['stage'], e['last_at'], now, e['eta'], e['eta_p10'], e['eta_p90'],
                  e['last_at'] + timedelta(hours=leg[1]) if leg else None,
                  float(trained) if trained is not None else None])
            added += cur.rowcount
        conn.commit()
        return added
    finally:
        conn.close()


def accuracy(barges, mbcs, now, days=REPORT_DAYS):
    """How the logged ETAs turned out: typical error of v6 and of the median rule, in hours."""
    arrived = {('BARGE', r['id']): model._dt(r[model.BARGE_ARRIVAL]) for r in barges}
    arrived.update({('MBC', r['id']): model._dt(r[model.MBC_ARRIVAL]) for r in mbcs})
    out = {}
    for r in _rows('SELECT * FROM rp03_eta_log WHERE predicted_at >= %s', [now - timedelta(days=days)]):
        actual = arrived.get((r['trip_kind'], r['trip_id']))
        if not actual or actual < r['predicted_at']:
            continue        # still on the way, or the arrival was entered for a time before we spoke
        hrs = lambda a: abs((a - actual).total_seconds()) / 3600
        s = out.setdefault(r['trip_kind'], {'model': [], 'median': [], 'in_band': []})
        s['model'].append(hrs(r['eta']))
        if r['median_eta']:
            s['median'].append(hrs(r['median_eta']))
        if r['eta_p10'] and r['eta_p90']:
            s['in_band'].append(r['eta_p10'] <= actual <= r['eta_p90'])
    return {k: {'n': len(s['model']), 'typical_h': round(median(s['model']), 1),
                'median_rule_h': round(median(s['median']), 1) if s['median'] else None,
                'within_2h_pct': round(100 * sum(x <= 2 for x in s['model']) / len(s['model'])),
                'in_band_pct': round(100 * sum(s['in_band']) / len(s['in_band'])) if s['in_band'] else None}
            for k, s in out.items()}


def status(barges, mbcs, now):
    """For the board. None until the migration has run."""
    if not _rows("SELECT to_regclass('rp03_eta_log') AS t")[0]['t']:
        return None
    last = _rows('SELECT ran_at, promoted, reason FROM rp03_eta_run ORDER BY id DESC LIMIT 1')
    return {'accuracy': accuracy(barges, mbcs, now), 'last_run': last[0] if last else None}


def retrain_due(now):
    last = _rows('SELECT MAX(ran_at) AS t FROM rp03_eta_run')[0]['t']
    return last is None or last <= now - timedelta(days=RETRAIN_DAYS)


def learn(now=None):
    """One guarded retrain. Always records the outcome; swaps the model only on a pass."""
    now = now or datetime.now()
    barges, mbcs = model._fetch()
    # On Export, along_side_berth is the loading berth: not an arrival to learn from.
    frames = {k: pd.DataFrame([dict(r) for r in rows if r['operation_type'] != 'Export']).rename(columns=RP03_COLUMNS[k])
              for k, rows in (('barge', barges), ('mbc', mbcs))}
    bundle, detail = eta_fit.retrain(frames, _rows('SELECT * FROM tide_master'), load_bundle(), hours([now])[0])
    detail['live_accuracy'] = accuracy(barges, mbcs, now)
    if bundle is not None:
        tmp = LIVE.with_suffix('.tmp')
        joblib.dump(bundle, tmp, compress=3)
        if LIVE.exists():
            os.replace(LIVE, PREV)
        os.replace(tmp, LIVE)
    conn = get_db()
    try:
        cur = get_cursor(conn)
        cur.execute('INSERT INTO rp03_eta_run (ran_at, promoted, reason, detail) VALUES (%s, %s, %s, %s)',
                    [now, bundle is not None, detail['reason'], json.dumps(detail)])
        conn.commit()
    finally:
        conn.close()
    log.info('RP03 retrain: %s', detail)
    return bundle is not None, detail
