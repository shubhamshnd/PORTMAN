"""v6 features shared by training and live prediction (one code path, so no train/serve skew).

Why v6: Dharamtar is tide-gated. Barges and MBCs arrive in a ~5 h window of every 12.42 h
tidal cycle, so duration from any upstream event is a sawtooth of the tide phase at that event.
v5 only had clock hour and lunar day, from which a tree cannot rebuild tide phase.

Tide comes from PORTMAN TM01 (tide_master). Every other input exists in RP03's _fetch().
This file is copied verbatim to PORTMAN modules/RP03/; edit it here and copy it over.
"""

from __future__ import annotations

import numpy as np
import pandas as pd

EPOCH = pd.Timestamp("2026-01-01")
MAX_HOURS = 120.0

# Milestones earliest first, same chains as RP03 BARGE_STAGES / MBC_STAGES.
# RP03's three Gull/AMF barge stages are never stamped in the data, so they have no model:
# a trip there falls back to its last modelled milestone.
KINDS = {
    "barge": {
        "stages": ["trip_start_ts", "alongside_mother_vessel_ts", "loading_commenced_ts", "loading_completed_ts", "cast_off_mother_vessel_ts"],
        "name": "barge_name", "qty": "discharge_quantity",
        "categories": ["contractor_name", "cargo_name"],
        "leg_key": "mother_vessel",                                   # who queues together before arrival
        "total_keys": ["mother_vessel", "cargo_name", "contractor_name"],
        "done": "discharge_completed_ts",                             # optional: not in the first export
    },
    "mbc": {
        "stages": ["arrived_load_port_ts", "alongside_load_berth_ts", "loading_commenced_ts", "loading_completed_ts", "cast_off_load_port_ts", "arrival_gull_island_ts", "departure_gull_island_ts"],
        "name": "mbc_name", "qty": "bl_quantity",
        "categories": ["mbc_name", "cargo_name", "load_port"],
        "leg_key": "load_port",
        "total_keys": ["load_port", "cargo_name"],
        "done": "unloading_completed_ts",
    },
}
RP03_COLUMNS = {
    "barge": {
        "trip_start": "trip_start_ts", "along_side_vessel": "alongside_mother_vessel_ts",
        "commenced_loading": "loading_commenced_ts", "completed_loading": "loading_completed_ts",
        "cast_off_mv": "cast_off_mother_vessel_ts", "along_side_berth": "arrival_ts",
        "completed_discharge_berth": "discharge_completed_ts", "vessel_name": "mother_vessel",
    },
    "mbc": {
        "arrived_load_port": "arrived_load_port_ts", "lp_alongside_berth": "alongside_load_berth_ts",
        "lp_loading_commenced": "loading_commenced_ts", "lp_loading_completed": "loading_completed_ts",
        "cast_off_load_port": "cast_off_load_port_ts", "arrival_gull_island": "arrival_gull_island_ts",
        "departure_gull_island": "departure_gull_island_ts", "vessel_arrival_port": "arrival_ts",
        "unloading_completed": "unloading_completed_ts",
    },
}

# Tidal constituent speeds, degrees/hour. Standard astronomical constants.
CONSTITUENTS = {
    "m2": 28.9841042, "s2": 30.0, "n2": 28.4397295, "k2": 30.0821373, "k1": 15.0410686, "o1": 13.9430356, "p1": 14.9589314,
    "q1": 13.3986609, "j1": 15.5854433, "l2": 29.5284789, "2n2": 27.8953548, "mu2": 27.9682084, "nu2": 28.5125831,
    "m4": 57.9682084, "ms4": 58.9841042, "mn4": 57.4238337, "mk3": 44.0251729, "m3": 43.4761563, "m6": 86.9523127,
    "2ms6": 87.9682084, "msf": 1.0158958, "mm": 0.5443747, "sa": 0.0410686,
}
OMEGA = np.deg2rad(np.array(list(CONSTITUENTS.values())))


def hours(values) -> np.ndarray:
    """Timestamps -> float hours since EPOCH (NaN when missing or invalid)."""
    parsed = pd.to_datetime(pd.Series(list(values), dtype=object), format="mixed", errors="coerce")
    return ((parsed - EPOCH).dt.total_seconds() / 3600).to_numpy(dtype=float)


def clean_key(values) -> np.ndarray:
    out = pd.Series(list(values), dtype="string").str.strip().str.casefold().fillna("")
    return out.replace("jsw jaigad port", "jsw jaigad").to_numpy(dtype=object)  # same port, two spellings


# ---------------------------------------------------------------- tide
def _height_basis(t: np.ndarray) -> np.ndarray:
    return np.column_stack([np.ones_like(t), np.cos(np.outer(t, OMEGA)), np.sin(np.outer(t, OMEGA))])


def _rate_basis(t: np.ndarray) -> np.ndarray:
    return np.column_stack([np.zeros_like(t), -OMEGA * np.sin(np.outer(t, OMEGA)), OMEGA * np.cos(np.outer(t, OMEGA))])


def fit_tide(tide_rows) -> tuple[np.ndarray, pd.DataFrame]:
    """TM01 tide_master (HW/LW times and heights) -> (harmonic coefficients, rejected rows).

    A row is a turning point: height h and zero rate. Fitting both gives a tide for any time, so the
    table may have holes or start after the training data. Rows far off the fit are typos and dropped.
    """
    frame = tide_rows.copy() if isinstance(tide_rows, pd.DataFrame) else pd.DataFrame([dict(r) for r in tide_rows])
    t, h = hours(frame["tide_datetime"]), pd.to_numeric(frame["tide_meters"], errors="coerce").to_numpy(dtype=float)
    usable = np.isfinite(t) & np.isfinite(h)
    if usable.sum() < 200:
        raise ValueError(f"need about 2 months of tide_master rows, got {usable.sum()}")
    keep, safe_t = usable.copy(), np.where(usable, t, 0.0)
    for _ in range(6):
        x = np.vstack([_height_basis(t[keep]), 3 * _rate_basis(t[keep])])
        y = np.r_[h[keep], np.zeros(keep.sum())]
        coef = np.linalg.solve(x.T @ x + np.eye(x.shape[1]), x.T @ y)
        keep = usable & (np.abs(_height_basis(safe_t) @ coef - h) < 0.5) & (np.abs(_rate_basis(safe_t) @ coef) < 0.35)
    if keep.sum() < 0.8 * usable.sum():
        raise ValueError(f"only {keep.sum()} of {usable.sum()} tide_master rows fit one tide")
    return coef, frame.loc[~keep]


def high_waters(coef: np.ndarray, lo: float, hi: float) -> tuple[np.ndarray, np.ndarray]:
    """High-water times (to the minute) and heights in [lo, hi]."""
    grid = np.arange(np.floor(lo), hi, 1 / 60)
    height = _height_basis(grid) @ coef
    i = np.flatnonzero((height[1:-1] > height[:-2]) & (height[1:-1] >= height[2:])) + 1
    return grid[i], height[i]


def tide_features(t: np.ndarray, coef: np.ndarray, lead: float) -> pd.DataFrame:
    """`tide_base` = hours to the first high water the trip can make (>= lead hours away)."""
    out = {"tide_height": _height_basis(t) @ coef, "tide_rate": _rate_basis(t) @ coef}
    names = ("prev", "base", "next2", "next3")
    gap, level = np.empty((len(t), 4)), np.empty((len(t), 4))
    for week in np.unique(np.floor(t / 168)):  # weekly chunks: no giant minute grid when old rows are mixed in
        rows = np.flatnonzero(np.floor(t / 168) == week)
        p, h = high_waters(coef, week * 168 - 30, (week + 1) * 168 + lead + 60)
        j = np.searchsorted(p, t[rows] + lead)
        for k in range(4):
            gap[rows, k], level[rows, k] = p[j + k - 1] - t[rows], h[j + k - 1]
    for k, name in enumerate(names):
        out[f"tide_{name}"], out[f"tide_{name}_height"] = gap[:, k], level[:, k]
    return pd.DataFrame(out)


# ---------------------------------------------------------------- queue state
QUEUE_STATS = ["open_all", "open_key", "open_age_key", "recent_all", "recent_key", "recent_max_key", "done24_all", "done24_key"]


def queue_features(prefix, t, key, start, end, log_key) -> pd.DataFrame:
    """State of one leg (start -> end event) as known strictly before each t.

    Legs with no end yet count as in progress, so live history must include open trips.
    """
    duration = end - start
    ok = np.isfinite(start) & (np.isnan(end) | ((duration >= 0) & (duration <= MAX_HOURS)))
    start, log_key = start[ok], log_key[ok]
    end = np.where(np.isnan(end[ok]), np.inf, end[ok])
    duration = end - start
    values = np.full((len(t), len(QUEUE_STATS)), np.nan)
    for i, now in enumerate(t):
        age = now - start
        same = log_key == key[i]
        opened = (age > 0) & (age <= MAX_HOURS) & (end >= now)
        mine = opened & same
        done = (end < now) & (end > now - 96)
        day = done & (end > now - 24)
        row = [opened.sum(), mine.sum(), age[mine].max() if mine.any() else 0.0, np.nan, np.nan, np.nan, day.sum(), (day & same).sum()]
        for col, mask in ((3, done), (4, done & same)):
            idx = np.flatnonzero(mask)
            if len(idx):
                idx = idx[np.lexsort((start[idx], end[idx]))[-6:]]
                row[col] = duration[idx].mean()
                if col == 4:
                    row[5] = duration[idx].max()
        values[i] = row
    return pd.DataFrame(values, columns=[f"{prefix}_{n}" for n in QUEUE_STATS])


# ---------------------------------------------------------------- model frame
def category_codes(kind: str, rows: pd.DataFrame) -> dict:
    """Stable ordinal per category, rarest first; unseen values become missing."""
    return {
        col: {v: i for i, v in enumerate(pd.Series(clean_key(rows[col])).value_counts().sort_index(kind="stable").sort_values(kind="stable").index)}
        for col in KINDS[kind]["categories"]
    }


def trip_features(kind: str, stage: str, rows: pd.DataFrame, codes: dict) -> pd.DataFrame:
    """The trip's own facts, all known at its `stage` timestamp."""
    cfg = KINDS[kind]
    t = hours(rows[stage])
    trip = {"qty": pd.to_numeric(rows[cfg["qty"]], errors="coerce").to_numpy(dtype=float)}
    if kind == "barge":
        trip["trip_number"] = pd.to_numeric(rows["trip_number"], errors="coerce").to_numpy(dtype=float)
    for earlier in cfg["stages"][:cfg["stages"].index(stage)]:
        gap = t - hours(rows[earlier])
        trip[f"since_{earlier}"] = np.where((gap >= 0) & (gap <= MAX_HOURS), gap, np.nan)
    year = 2 * np.pi * t / (24 * 365.25)
    trip.update({"hour_of_day": t % 24, "weekday": np.floor(t / 24) % 7, "year_sin": np.sin(year), "year_cos": np.cos(year)})
    for col in cfg["categories"]:
        trip[f"{col}_code"] = np.array([codes[col].get(v, np.nan) for v in clean_key(rows[col])], dtype=float)
    return pd.DataFrame(trip)


def queue_state(kind: str, stage: str, rows: pd.DataFrame, history: pd.DataFrame) -> pd.DataFrame:
    """Who else is on the water, as known at each row's `stage` timestamp."""
    cfg = KINDS[kind]
    t = hours(rows[stage])
    chain = [*cfg["stages"][cfg["stages"].index(stage):], "arrival_ts"]
    log = {c: hours(history[c]) for c in chain}
    legs = [(f"leg{i}", a, b, cfg["leg_key"]) for i, (a, b) in enumerate(zip(chain, chain[1:])) if len(chain) > 2]
    legs += [(f"total_{k}", stage, "arrival_ts", k) for k in cfg["total_keys"]]
    parts = [queue_features(p, t, clean_key(rows[k]), log[a], log[b], clean_key(history[k])) for p, a, b, k in legs]
    if cfg["done"] in history:
        parts.append(queue_features("port", t, clean_key(rows["cargo_name"]), log["arrival_ts"], hours(history[cfg["done"]]), clean_key(history["cargo_name"])))
    return pd.concat(parts, axis=1)


# ---------------------------------------------------------------- model application
# candidate -> (uses tide, uses queue, predicts offset from the high water)
CANDIDATES = {
    "median": None,
    "no_tide": (False, True, False),
    "tide": (True, False, False),
    "tide_offset": (True, False, True),
    "tide_queue": (True, True, False),
    "tide_queue_offset": (True, True, True),
}


def design(trip: pd.DataFrame, queue: pd.DataFrame, tide: pd.DataFrame, candidate: str) -> tuple[pd.DataFrame, np.ndarray]:
    use_tide, use_queue, offset = CANDIDATES[candidate]
    parts = [trip] + ([tide] if use_tide else []) + ([queue] if use_queue else [])
    return pd.concat(parts, axis=1), tide["tide_base"].to_numpy() if offset else np.zeros(len(trip))


def predict(models: dict, x: pd.DataFrame, base: np.ndarray) -> np.ndarray:
    x = x[models["columns"]]
    return np.maximum(0, base + np.median([m.predict(x) for m in models["point"]], axis=0))


def predict_band(models: dict, x: pd.DataFrame, base: np.ndarray, point: np.ndarray, widen: float = 0.0) -> tuple[np.ndarray, np.ndarray]:
    x = x[models["columns"]]
    return np.clip(base + models["p10"].predict(x) - widen, 0, point), np.maximum(base + models["p90"].predict(x) + widen, point)
