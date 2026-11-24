"""Live v6 ETAs from RP03's own rows.

    etas = predict_board(barges, mbcs, now)   # {('BARGE', id): {...}, ('MBC', id): {...}}

Each trip is predicted from its latest stamped milestone. Pass ALL rows of RP03 _fetch(), not just
open trips: completed ones are the history and open ones are the queue the model reads.
This file is copied verbatim to PORTMAN modules/RP03/; edit it here and copy it over.
"""

from __future__ import annotations

from datetime import datetime
from functools import lru_cache
from pathlib import Path

import joblib
import numpy as np
import pandas as pd

if __package__:  # inside PORTMAN: modules.RP03
    from .eta_features import EPOCH, KINDS, RP03_COLUMNS, design, high_waters, hours, predict, predict_band, queue_state, tide_features, trip_features
else:
    from eta_features import EPOCH, KINDS, RP03_COLUMNS, design, high_waters, hours, predict, predict_band, queue_state, tide_features, trip_features

HERE = Path(__file__).resolve().parent
STALE_HOURS = 7 * 24  # same cut-off as RP03: an untouched trip is an abandoned document
HISTORY_HOURS = 12 * 24


@lru_cache(maxsize=1)
def _load(path: str, mtime: float) -> dict:
    return joblib.load(path)


def load_bundle() -> dict:
    """The retrained model when there is one, else the shipped one. Picks up a swap without a restart."""
    for path in (HERE / "eta_live.joblib", HERE / "eta_v6.joblib", HERE / "eta_artifacts_v6" / "eta_v6.joblib"):
        if path.exists():
            return _load(str(path), path.stat().st_mtime)
    raise FileNotFoundError("no v6 ETA model file")


def _frame(kind: str, rows) -> pd.DataFrame:
    frame = rows.copy() if isinstance(rows, pd.DataFrame) else pd.DataFrame([dict(r) for r in rows])
    return frame.rename(columns=RP03_COLUMNS[kind]).reset_index(drop=True)


def _stamp(h: float) -> datetime:
    return (EPOCH + pd.Timedelta(hours=float(h))).round("min").to_pydatetime()


def _next_tide_shift(tide_coef: np.ndarray, eta: float, now: float) -> float:
    """Hours to add so a trip that missed its tide lands in the next one after now."""
    p, _ = high_waters(tide_coef, eta - 24, now + 48)
    own = p[np.argmin(np.abs(p - eta))]
    return next(w - own for w in p if eta + (w - own) >= now)


def predict_board(barges, mbcs, now: datetime | None = None, bundle: dict | None = None) -> dict:
    bundle = bundle or load_bundle()
    now_h = hours([now or datetime.now()])[0]
    out: dict = {}
    for kind, rows in (("barge", barges), ("mbc", mbcs)):
        if not len(rows):
            continue
        frame, stages = _frame(kind, rows), KINDS[kind]["stages"]
        t = np.column_stack([hours(frame[s]) for s in stages])
        t[t > now_h + 1] = np.nan                       # mistyped future stamps are not milestones
        arrival = hours(frame["arrival_ts"])
        stamped = np.isfinite(t)
        last = np.where(stamped.any(1), t.shape[1] - 1 - stamped[:, ::-1].argmax(1), -1)
        last_at = np.where(last >= 0, t[np.arange(len(t)), last], -np.inf)
        live = np.isnan(arrival) & (last >= 0) & (last_at >= now_h - STALE_HOURS)
        # Recent history is all the queue features look at; trimming keeps a page load fast.
        newest = np.nan_to_num(np.column_stack([t, arrival]), nan=-np.inf).max(axis=1)
        history = frame[newest >= now_h - STALE_HOURS - HISTORY_HOURS]
        for i, stage in enumerate(stages):
            pick = np.flatnonzero(live & (last == i))
            model = bundle["stages"].get((kind, stage))
            if not len(pick) or model is None:
                continue
            todo, at = frame.iloc[pick].reset_index(drop=True), last_at[pick]
            lo, hi = np.full(len(pick), np.nan), np.full(len(pick), np.nan)
            if model["selected"] == "median":
                point = np.full(len(pick), model["median_hours"])
            else:
                x, base = design(trip_features(kind, stage, todo, model["codes"]), queue_state(kind, stage, todo, history),
                                 tide_features(at, bundle["tide"], model["lead"]), model["selected"])
                point = predict(model["models"], x, base)
                lo, hi = predict_band(model["models"], x, base, point, model["widen"])
            for j in range(len(pick)):
                eta = at[j] + point[j]
                shift = _next_tide_shift(bundle["tide"], eta, now_h) if eta < now_h else 0.0
                early, late = at[j] + lo[j] + shift, at[j] + hi[j] + shift
                if shift:
                    early = max(early, now_h)
                out[(kind.upper(), int(todo["id"].iloc[j]))] = {
                    "eta": _stamp(eta + shift),
                    "eta_p10": _stamp(early) if np.isfinite(early) else None,
                    "eta_p90": _stamp(late) if np.isfinite(late) else None,
                    "stage": stage, "last_at": _stamp(at[j]),
                    "remaining_hours": round(eta + shift - now_h, 2), "missed_tide": bool(shift),
                }
    return out
