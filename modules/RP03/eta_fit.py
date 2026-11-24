"""v6 fitting, and the guarded retrain used in production.

This file is copied verbatim to PORTMAN modules/RP03/; edit it here and copy it over.

What was measured on 19 Sep - 22 Nov 2026 (barge cast-off MAE 212.7 min with a fixed model):
weekly retrain on all data 210.3; weighting recent trips more 211.6-214.4; shifting predictions by
the recent error 229.0. So learning means refitting on everything, and no error-shift correction.
"""

from __future__ import annotations

import numpy as np
import pandas as pd
from sklearn.ensemble import HistGradientBoostingRegressor

if __package__:
    from .eta_features import KINDS, MAX_HOURS, category_codes, design, fit_tide, high_waters, hours, predict, queue_state, tide_features, trip_features
else:
    from eta_features import KINDS, MAX_HOURS, category_codes, design, fit_tide, high_waters, hours, predict, queue_state, tide_features, trip_features

SEEDS = 3
WINDOW_DAYS = 14       # newest trips held back to judge a retrained model
MIN_WINDOW_TRIPS = 20  # fewer than this in a stage proves nothing, so it has no vote
WORST_ERROR_RATIO = 2  # typical error may not exceed this multiple of the shipped model's


def estimator(seed: int, loss: str = "absolute_error", quantile: float | None = None) -> HistGradientBoostingRegressor:
    return HistGradientBoostingRegressor(
        loss=loss, quantile=quantile, max_iter=300, learning_rate=0.04, max_leaf_nodes=12, min_samples_leaf=40,
        l2_regularization=3.0, max_features=0.5, early_stopping=False, random_state=seed,
    )


def lead_hours(y: np.ndarray) -> float:
    """High waters closer than this cannot be made from here (fastest 5% less 2 h slack)."""
    return float(max(0.5, np.quantile(y, 0.05) - 2))


def fit(x: pd.DataFrame, target: np.ndarray, with_band: bool = False) -> dict:
    columns = [c for c in x if x[c].nunique() > 1]
    x = x[columns]
    out = {"columns": columns, "point": [estimator(s).fit(x, target) for s in range(SEEDS)]}
    if with_band:
        out["p10"] = estimator(0, "quantile", 0.1).fit(x, target)
        out["p90"] = estimator(0, "quantile", 0.9).fit(x, target)
    return out


def stage_rows(kind: str, stage: str, history: pd.DataFrame) -> tuple[pd.DataFrame, np.ndarray, np.ndarray, np.ndarray, dict]:
    """Usable trips for one milestone in time order: rows, stage time, arrival time, duration, exclusions."""
    t, arrival = hours(history[stage]), hours(history["arrival_ts"])
    y = arrival - t
    order = np.flatnonzero(np.isfinite(y) & (y >= 0) & (y <= MAX_HOURS))
    order = order[np.argsort(t[order], kind="stable")]
    excluded = {"rows_total": len(y), "missing_timestamp": int((~np.isfinite(y)).sum()), "negative": int((y < 0).sum()), "over_120_hours": int((y > MAX_HOURS).sum())}
    return history.iloc[order].reset_index(drop=True), t[order], arrival[order], y[order], excluded


def fit_bundle(kind, stage, rows, t, y, queue, tide_coef, selected, widen, with_band=True) -> dict:
    bundle = {"selected": selected, "lead": lead_hours(y), "median_hours": float(np.median(y)), "widen": widen, "training_rows": len(y)}
    if selected != "median":
        bundle["codes"] = category_codes(kind, rows)
        x, base = design(trip_features(kind, stage, rows, bundle["codes"]), queue, tide_features(t, tide_coef, bundle["lead"]), selected)
        bundle["models"] = fit(x, y - base, with_band)
    return bundle


def apply_bundle(kind, stage, model, rows, t, queue, tide_coef) -> np.ndarray:
    if model["selected"] == "median":
        return np.full(len(t), model["median_hours"])
    x, base = design(trip_features(kind, stage, rows, model["codes"]), queue, tide_features(t, tide_coef, model["lead"]), model["selected"])
    return predict(model["models"], x, base)


def refit_tide(tide_rows, champion: dict, now: float) -> tuple[np.ndarray, str]:
    """Tide from the current TM01 table, unless that table is thin or disagrees with the known tide."""
    try:
        coef, rejected = fit_tide(tide_rows)
    except (ValueError, KeyError) as exc:
        return champion["tide"], f"kept the previous tide: {exc}"
    old, new = high_waters(champion["tide"], now, now + 24 * 45)[0], high_waters(coef, now, now + 24 * 45)[0]
    drift = float(np.median([np.abs(new - x).min() for x in old])) * 60 if len(new) else np.inf
    if drift > 30:
        return champion["tide"], f"kept the previous tide: new table moves high water by {drift:.0f} min"
    return coef, f"tide refitted, {len(rejected)} rows ignored"


def retrain(frames: dict, tide_rows, champion: dict, now: float) -> tuple[dict | None, dict]:
    """(new bundle or None, detail). None means: keep the current model, detail['reason'] says why.

    The recipe is judged on the newest WINDOW_DAYS with those trips held back. It must beat the median
    rule for each kind and stay within WORST_ERROR_RATIO of the shipped model's typical error. Only
    then is it refitted on everything. Candidate choice per stage stays as shipped.
    """
    tide, note = refit_tide(tide_rows, champion, now)
    detail: dict = {"tide": note, "window_days": WINDOW_DAYS, "kinds": {}}
    cutoff, data = now - 24 * WINDOW_DAYS, {}
    score = {k: {"model": [], "median": [], "ratio": []} for k in KINDS}
    for (kind, stage), old in champion["stages"].items():
        rows, t, arrival, y, _ = stage_rows(kind, stage, frames[kind])
        past = arrival <= now
        rows, t, arrival, y = rows[past].reset_index(drop=True), t[past], arrival[past], y[past]
        if len(y) < 0.9 * old["training_rows"]:
            detail["reason"] = f"{kind} {stage}: only {len(y)} usable trips, current model had {old['training_rows']}"
            return None, detail
        queue = queue_state(kind, stage, rows, frames[kind])
        data[(kind, stage)] = rows, t, y, queue, old
        tr, te = np.flatnonzero(arrival < cutoff), np.flatnonzero(t >= cutoff)
        if len(te) < MIN_WINDOW_TRIPS:
            continue
        pick = lambda f, i: f.iloc[i].reset_index(drop=True)
        trial = fit_bundle(kind, stage, pick(rows, tr), t[tr], y[tr], pick(queue, tr), tide, old["selected"], 0.0, with_band=False)
        error = np.abs(apply_bundle(kind, stage, trial, pick(rows, te), t[te], pick(queue, te), tide) - y[te])
        score[kind]["model"].append(error)
        score[kind]["median"].append(np.abs(np.median(y[tr]) - y[te]))
        score[kind]["ratio"].append(error * 60 / champion["reference"][(kind, stage)])
    for kind, s in score.items():
        if not s["model"]:
            detail["reason"] = f"{kind}: fewer than {MIN_WINDOW_TRIPS} trips per stage in the last {WINDOW_DAYS} days, nothing to judge on"
            return None, detail
        s = {k: np.concatenate(v) for k, v in s.items()}
        detail["kinds"][kind] = {"predictions": len(s["model"]), "mae_h": round(float(s["model"].mean()), 2), "median_rule_mae_h": round(float(s["median"].mean()), 2), "typical_error_ratio": round(float(np.median(s["ratio"])), 2)}
        if s["model"].mean() > s["median"].mean():
            detail["reason"] = f"{kind}: no better than the plain median on the last {WINDOW_DAYS} days"
            return None, detail
        if np.median(s["ratio"]) > WORST_ERROR_RATIO:
            detail["reason"] = f"{kind}: typical error is {np.median(s['ratio']):.1f}x the shipped model; check recent data entry"
            return None, detail
    stages = {(kind, stage): fit_bundle(kind, stage, rows, t, y, queue, tide, old["selected"], old["widen"]) for (kind, stage), (rows, t, y, queue, old) in data.items()}
    detail["reason"] = "promoted"
    return {**champion, "tide": tide, "stages": stages, "trained_until": now}, detail
