from flask import render_template, session, jsonify, request
from datetime import datetime, timedelta, date
import json

from database import get_db, get_cursor

from .. import bp
from ..daily_ops.model import fy_label

# Reuse the existing Port Overview helpers.
from ..port_overview.views import (
    login_required,
    _days_left_in_month,
    _cargo_by_type,
    _current_fy_by_type,
    _load_fy_targets,
)


# ============================================================================
# CARGO STATISTICS
# ============================================================================

CARGO_CONFIG = [
    ("IBRM", ["IBRM"]),
    ("CBRM", ["CBRM"]),
    ("Fluxes", ["FLUXES", "FLUX"]),
    ("Clinker/ Slag", ["CLINKER", "SLAG"]),
    ("FG Goods", ["FINISH GOODS", "FINISHED GOODS", "FG GOODS"]),
]


def _norm(value):
    return str(value or "").strip().upper()


def _qty(by_type, names):
    wanted = {_norm(x) for x in names}

    return round(
        sum(
            float(value or 0)
            for key, value in (by_type or {}).items()
            if _norm(key) in wanted
        ),
        2,
    )


def _get_month_target(financial_year, report_date):
    """
    Get the financial-target record for the report month.

    Financial year is Apr-Mar:
        Apr -> index 0
        May -> index 1
        ...
        Mar -> index 11

    This returns ONLY the target for the selected report month,
    not the sum of all 12 months.
    """
    targets = _load_fy_targets(financial_year) or []

    fiscal_month_index = (
        report_date.month - 4
        if report_date.month >= 4
        else report_date.month + 8
    )

    # Preferred: use the matching month/date information if available.
    month_number = report_date.month
    month_name = report_date.strftime("%B").upper()

    for month in targets:
        if not isinstance(month, dict):
            continue

        # Support common month fields used by financial_target.
        candidates = [
            month.get("month"),
            month.get("month_number"),
            month.get("month_no"),
            month.get("month_name"),
            month.get("name"),
            month.get("date"),
        ]

        matched = False

        for candidate in candidates:
            if candidate is None:
                continue

            candidate_text = _norm(candidate)

            if candidate_text == month_name:
                matched = True
                break

            try:
                if int(float(candidate)) == month_number:
                    matched = True
                    break
            except (TypeError, ValueError):
                pass

        if matched:
            return month

    # Fallback for _load_fy_targets implementations that return
    # Apr-Mar in a fixed 12-item list without month metadata.
    if len(targets) >= 12 and 0 <= fiscal_month_index < len(targets):
        return targets[fiscal_month_index]

    return {}


def _get_targets(financial_year, report_date):
    """
    Read monthly ABP and SCM values from financial_year_targets.

    ABP = Base Target
    SCM = Outlook Target; if Outlook is blank, Base Target is used.

    IMPORTANT:
    Only the selected report month's financial_target is used.
    """
    month_target = _get_month_target(financial_year, report_date)
    month_categories = month_target.get("categories") or {}

    result = {}

    for label, categories in CARGO_CONFIG:
        wanted = {_norm(x) for x in categories}

        abp = 0.0
        scm = 0.0

        for category, values in month_categories.items():
            if _norm(category) not in wanted:
                continue

            values = values or {}

            base = float(values.get("base") or 0)

            outlook = values.get("outlook")
            if outlook in (None, ""):
                outlook = base
            else:
                outlook = float(outlook or 0)

            abp += base
            scm += outlook

        result[label] = {
            "abp": round(abp, 2),
            "scm": round(scm, 2),
        }

    return result


def _stats_cargo_consumption(start_date, end_date):
    """
    Read Consumption from stats_cargo for the requested date range.

    stats_cargo stores the consumption section as JSON in `data`.
    The consumption fields are:
        ibrm
        cbrm
        fluxes
        clinker_slag
        fg_goods
        total

    Returns category totals in the same names used by CARGO_CONFIG.
    """
    result = {
        "IBRM": 0.0,
        "CBRM": 0.0,
        "FLUXES": 0.0,
        "CLINKER": 0.0,
        "SLAG": 0.0,
        "FINISH GOODS": 0.0,
        "FINISHED GOODS": 0.0,
        "FG GOODS": 0.0,
    }

    conn = get_db()
    cur = get_cursor(conn)

    try:
        cur.execute(
            """
            SELECT entry_date, data
            FROM stats_cargo
            WHERE section = %s
              AND entry_date >= %s
              AND entry_date <= %s
            ORDER BY entry_date ASC, id ASC;
            """,
            (
                "consumption",
                start_date,
                end_date,
            ),
        )

        rows = cur.fetchall()

        for row in rows:
            raw_data = row["data"]

            if isinstance(raw_data, str):
                try:
                    raw_data = json.loads(raw_data)
                except (TypeError, ValueError):
                    raw_data = {}

            data = raw_data or {}

            def value(field):
                try:
                    return float(data.get(field) or 0)
                except (TypeError, ValueError):
                    return 0.0

            result["IBRM"] += value("ibrm")
            result["CBRM"] += value("cbrm")
            result["FLUXES"] += value("fluxes")
            result["CLINKER"] += value("clinker_slag")
            result["FG GOODS"] += value("fg_goods")

        return {
            key: round(value, 2)
            for key, value in result.items()
        }

    finally:
        conn.close()



# ============================================================================
# STATS_CARGO PRODUCTION
# ============================================================================

PRODUCTION_FIELDS = [
    # Support both the DB field names used by CARGO_STATS and the
    # underscore aliases used by the Production API/frontend.
    ("BF-1", ["bf1", "bf_1"]),
    ("BF-2", ["bf2", "bf_2"]),
    ("PP-1", ["pp1", "pp_1"]),
    ("PP-2", ["pp2", "pp_2"]),
    ("SP-1", ["sp1", "sp_1"]),
    ("SP-2", ["sp2", "sp_2"]),
    ("CO-1 (PUSHINGS)", ["co1_pushings"]),
    ("CO-2 (A&B) (PUSHINGS)", ["co2ab_pushings", "co2_ab_pushings"]),
    ("CO-2 (C&D) (PUSHINGS)", ["co2cd_pushings", "co2_cd_pushings"]),
]


def _stats_cargo_production(start_date, end_date):
    """
    Read Production DIRECTLY from stats_cargo.

    CARGO_STATS stores Production as:
        section = 'production'
        data = {
            'bf1': ...,
            'bf2': ...,
            'pp1': ...,
            'pp2': ...,
            'sp1': ...,
            'sp2': ...,
            'co1_pushings': ...,
            'co2ab_pushings': ...,
            'co2cd_pushings': ...,
            'total': ...
        }

    Return one record per unit/date.  The Jinja calculates:
      Actual Production = selected date
      MTD = month start through selected date
      YTD = 1-Apr through selected date
    """
    conn = get_db()
    cur = get_cursor(conn)

    try:
        cur.execute(
            """
            SELECT entry_date, data
            FROM stats_cargo
            WHERE LOWER(TRIM(section)) = 'production'
              AND entry_date >= %s
              AND entry_date <= %s
            ORDER BY entry_date ASC, id ASC;
            """,
            (start_date, end_date),
        )

        result = []

        for db_row in cur.fetchall():
            entry_date = db_row["entry_date"]
            raw_data = db_row["data"]

            if isinstance(raw_data, str):
                try:
                    raw_data = json.loads(raw_data)
                except (TypeError, ValueError):
                    raw_data = {}

            data = raw_data if isinstance(raw_data, dict) else {}

            date_text = (
                entry_date.strftime("%Y-%m-%d")
                if hasattr(entry_date, "strftime")
                else str(entry_date)[:10]
            )

            for unit, fields in PRODUCTION_FIELDS:
                # Some stats_cargo records use bf_1 / bf_2 etc., while older
                # records use bf1 / bf2.  Try every supported alias.
                raw_value = None
                for field in fields:
                    if field in data and data.get(field) not in (None, ""):
                        raw_value = data.get(field)
                        break

                # Handle numeric strings, blanks, None, etc.
                try:
                    actual = float(raw_value) if raw_value not in (None, "") else 0.0
                except (TypeError, ValueError):
                    actual = 0.0

                result.append({
                    "date": date_text,
                    "unit": unit,
                    "actual_production": round(actual, 2),
                    "target_production": None,
                })

        return result

    finally:
        conn.close()



# ============================================================================
# STATS_CARGO STOCKS
# ============================================================================

STOCK_FIELDS = {
    "IBRM": "ibrm",
    "CBRM": "cbrm",
    "Fluxes": "fluxes",
    "Clinker/ Slag": "clinker_slag",
    "FG Goods": "fg_goods",
}


def _stats_cargo_stock(section, report_date):
    """
    Read the stock snapshot from stats_cargo.

    CARGO_STATS uses:
        section = 'stock_rmhs'  -> RMHS / JSW Steel stock
        section = 'stock_pnp'   -> PNP stock

    Stock is a point-in-time value, NOT an MTD/YTD sum.

    The record for the selected report date is preferred. If that date does
    not exist, the latest stock record before the report date is used.
    """

    result = {
        "IBRM": None,
        "CBRM": None,
        "Fluxes": None,
        "Clinker/ Slag": None,
        "FG Goods": None,
        "TOTAL": None,
    }

    conn = get_db()
    cur = get_cursor(conn)

    try:
        # First try the exact report date.
        cur.execute(
            """
            SELECT entry_date, data
            FROM stats_cargo
            WHERE LOWER(TRIM(section)) = %s
              AND entry_date = %s
            ORDER BY id DESC
            LIMIT 1;
            """,
            (section.lower(), report_date),
        )

        row = cur.fetchone()

        # If there is no exact snapshot, use the latest available snapshot
        # before the selected report date.
        if not row:
            cur.execute(
                """
                SELECT entry_date, data
                FROM stats_cargo
                WHERE LOWER(TRIM(section)) = %s
                  AND entry_date <= %s
                ORDER BY entry_date DESC, id DESC
                LIMIT 1;
                """,
                (section.lower(), report_date),
            )
            row = cur.fetchone()

        if not row:
            return result

        raw_data = row["data"]

        if isinstance(raw_data, str):
            try:
                raw_data = json.loads(raw_data)
            except (TypeError, ValueError):
                raw_data = {}

        data = raw_data if isinstance(raw_data, dict) else {}

        def numeric(field):
            value = data.get(field)

            if value in (None, ""):
                return None

            try:
                return round(float(value), 2)
            except (TypeError, ValueError):
                return None

        for label, field in STOCK_FIELDS.items():
            result[label] = numeric(field)

        # Prefer the stored total. If it is absent, calculate it from the
        # five cargo categories.
        stored_total = numeric("total")

        if stored_total is not None:
            result["TOTAL"] = stored_total
        else:
            values = [
                result[label]
                for label in STOCK_FIELDS
                if result[label] is not None
            ]

            result["TOTAL"] = (
                round(sum(values), 2)
                if values
                else None
            )

        return result

    finally:
        conn.close()


def _build_report(selected_date):
    """
    User-selected date is the reporting cut-off date + 1 day.

    Example:
        selected date = 28-Aug-2026
        report date   = 27-Aug-2026
    """
    report_date = selected_date - timedelta(days=1)
    report_date_s = report_date.strftime("%Y-%m-%d")

    month_start = report_date.replace(day=1)
    month_start_s = month_start.strftime("%Y-%m-%d")

    fy_start_year = (
        report_date.year
        if report_date.month >= 4
        else report_date.year - 1
    )

    financial_year = fy_label(fy_start_year)

    # ========================================================================
    # ACTUAL CARGO HANDLED
    # ========================================================================
    # Use EXACTLY the same calculations as Port Overview.
    #
    # If today is 03-Sep-2026 and the report is generated for that date:
    #
    #   Actual Day = Port Overview "Yesterday" = 02-Sep-2026
    #   Actual MTD = Port Overview "Sep 2026" through 02-Sep-2026
    #   Actual YTD = Port Overview "FY 2026-2027" through 02-Sep-2026
    #
    # Do NOT read Actual Cargo Handled from stats_cargo consumption.
    # Port Overview is the single source for these actuals.
    #
    # _cargo_by_type() gives the same daily/monthly live cargo values used by
    # the Port Overview cards.
    day_by_type = _cargo_by_type(
        report_date_s,
        report_date_s,
    )

    mtd_by_type = _cargo_by_type(
        month_start_s,
        report_date_s,
    )

    # This is the SAME FY calculation used by the Port Overview FY card:
    # historical April + live May onward.
    ytd_by_type = _current_fy_by_type(
        report_date_s,
    )

    # IMPORTANT:
    # Targets are for the selected report month only.
    targets = _get_targets(
        financial_year,
        report_date,
    )

    # Asking rate is based on remaining days in the report month.
    balance_days = _days_left_in_month(report_date)

    # Stock is a point-in-time snapshot for the report date.
    # JSW Steel column uses RMHS stock; PNP column uses PNP stock.
    rmhs_stock = _stats_cargo_stock(
        "stock_rmhs",
        report_date,
    )

    pnp_stock = _stats_cargo_stock(
        "stock_pnp",
        report_date,
    )

    rows = []

    for label, categories in CARGO_CONFIG:
        target = targets.get(
            label,
            {"abp": 0.0, "scm": 0.0},
        )

        abp = target["abp"]
        scm = target["scm"]

        actual_day = _qty(day_by_type, categories)
        actual_mtd = _qty(mtd_by_type, categories)
        actual_ytd = _qty(ytd_by_type, categories)

        # GAP is monthly target/outlook minus MTD actual.
        #
        # As per ABP:
        #     GAP = ABP in Outlook - MTD
        #
        # As per SCM:
        #     GAP = SCM/Outlook in Outlook - MTD
        #
        # Do not allow a negative remaining quantity.
        gap_abp = max(abp - actual_mtd, 0.0)
        gap_scm = max(scm - actual_mtd, 0.0)

        asking_abp = (
            gap_abp / balance_days
            if balance_days > 0
            else 0.0
        )

        asking_scm = (
            gap_scm / balance_days
            if balance_days > 0
            else 0.0
        )

        rows.append({
            "key": label,
            "cargo": label,
            "abp": round(abp, 2),
            "scm": round(scm, 2),
            "actual_day": round(actual_day, 2),
            "actual_mtd": round(actual_mtd, 2),
            "actual_ytd": round(actual_ytd, 2),
            "gap_abp": round(gap_abp, 2),
            "gap_scm": round(gap_scm, 2),
            "asking_abp": round(asking_abp, 2),
            "asking_scm": round(asking_scm, 2),

            # RMHS = JSW Steel stock
            # PNP  = PNP stock
            "stock_jsw": rmhs_stock.get(label),
            "stock_pnp": pnp_stock.get(label),
        })

    total = {
        "key": "TOTAL",
        "cargo": "Total",
        "abp": round(sum(r["abp"] for r in rows), 2),
        "scm": round(sum(r["scm"] for r in rows), 2),
        "actual_day": round(sum(r["actual_day"] for r in rows), 2),
        "actual_mtd": round(sum(r["actual_mtd"] for r in rows), 2),
        "actual_ytd": round(sum(r["actual_ytd"] for r in rows), 2),
        "gap_abp": round(sum(r["gap_abp"] for r in rows), 2),
        "gap_scm": round(sum(r["gap_scm"] for r in rows), 2),
        "asking_abp": round(sum(r["asking_abp"] for r in rows), 2),
        "asking_scm": round(sum(r["asking_scm"] for r in rows), 2),
        # Stock totals come directly from the stock snapshots.
        "stock_jsw": rmhs_stock.get("TOTAL"),
        "stock_pnp": pnp_stock.get("TOTAL"),
    }

    rows.append(total)

    # This is the value your Jinja header should use:
    #
    # As per ABP -> financial_target.abp
    # As per SCM -> financial_target.scm
    #
    # These are MONTHLY values for the selected report month.
    financial_target = {
        "abp": total["abp"],
        "scm": total["scm"],
    }

    return {
        "selected_date": selected_date.strftime("%Y-%m-%d"),
        "report_date": report_date_s,
        "financial_year": financial_year,
        "financial_target": financial_target,
        "balance_days": balance_days,
        "stock_rmhs": rmhs_stock,
        "stock_pnp": pnp_stock,
        "rows": rows,
    }



# ============================================================================
# DAILY PORT OPERATIONS SUMMARY
# ============================================================================

def _safe_float(value):
    try:
        if value is None or value == "":
            return 0.0
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _fetch_daily_port_operations(selected_date):
    """
    Build one row for every calendar day from the 1st of the selected month
    through the selected date.

    Sources:
      - Jetty Discharge       = Jetty Handling / lueu_lines
      - Steel Plant Discharge = Steel cargo from the same 24-hour SMS data
      - Cement Plant Discharge= Clinker + Slag from the same 24-hour SMS data
      - MBC Discharge         = MBC discharge in the last 24-hour window
      - MV Discharge          = MV discharge in the last 24-hour window
      - Daily Consumption     = stats_cargo / section=consumption

    The row DATE is the report date.  For MV/MBC, the 24-hour window is
    08:00 of the previous day through 08:00 of that report date, matching
    the existing 24-hours report convention.
    """

    month_start = selected_date.replace(day=1)

    conn = get_db()
    cur = get_cursor(conn)

    try:
        # ------------------------------------------------------------
        # 1) Jetty / Steel / Cement cargo from 24-hour SMS data.
        # ------------------------------------------------------------
        #
        # lueu_lines is already the source used by the existing Jetty
        # Handling and Steel/Cement cargo calculations.
        #
        # We fetch all dates in one query rather than one query per day.
        cur.execute(
            """
            SELECT
                TO_DATE(l.entry_date, 'YYYY-MM-DD') AS report_day,
                COALESCE(SUM(l.quantity), 0) AS jetty_qty,
                COALESCE(
                    SUM(
                        CASE
                            WHEN UPPER(TRIM(COALESCE(vc.cargo_type, '')))
                                 NOT IN ('CLINKER', 'SLAG')
                            THEN COALESCE(l.quantity, 0)
                            ELSE 0
                        END
                    ),
                    0
                ) AS steel_qty,
                COALESCE(
                    SUM(
                        CASE
                            WHEN UPPER(TRIM(COALESCE(vc.cargo_type, '')))
                                 IN ('CLINKER', 'SLAG')
                            THEN COALESCE(l.quantity, 0)
                            ELSE 0
                        END
                    ),
                    0
                ) AS cement_qty
            FROM lueu_lines l
            LEFT JOIN vessel_cargo vc
                ON UPPER(TRIM(vc.cargo_name))
                 = UPPER(TRIM(l.cargo_name))
            WHERE l.is_deleted = false
              AND l.cargo_name IS NOT NULL
              AND TRIM(l.cargo_name) <> ''
              AND TO_DATE(l.entry_date, 'YYYY-MM-DD')
                    BETWEEN %s AND %s
            GROUP BY TO_DATE(l.entry_date, 'YYYY-MM-DD')
            ORDER BY report_day
            """,
            (month_start, selected_date),
        )

        jetty_rows = cur.fetchall()

        jetty_by_day = {
            row["report_day"]: {
                "jetty": round(_safe_float(row["jetty_qty"]), 2),
                "steel": round(_safe_float(row["steel_qty"]), 2),
                "cement": round(_safe_float(row["cement_qty"]), 2),
            }
            for row in jetty_rows
        }

        # ------------------------------------------------------------
        # 2) MV discharge — 24 hrs for every day.
        # ------------------------------------------------------------
        #
        # Existing report uses start_time's date for the 24-hour quantity.
        # Keep exactly that data convention for this table.
        cur.execute(
            """
            SELECT
                TO_DATE(start_time, 'YYYY-MM-DD') AS report_day,
                COALESCE(SUM(quantity), 0) AS qty
            FROM ldud_vessel_operations
            WHERE TO_DATE(start_time, 'YYYY-MM-DD')
                    BETWEEN %s AND %s
            GROUP BY TO_DATE(start_time, 'YYYY-MM-DD')
            ORDER BY report_day
            """,
            (month_start, selected_date),
        )

        mv_rows = cur.fetchall()

        mv_by_day = {
            row["report_day"]: round(_safe_float(row["qty"]), 2)
            for row in mv_rows
        }

        # ------------------------------------------------------------
        # 3) MBC discharge — last 24 hrs for every day.
        # ------------------------------------------------------------
        #
        # For a report date D:
        #     08:00 on D-1 -> 08:00 on D
        #
        # lueu_lines stores the MBC SMS quantities by entry_date, so the
        # first/last calendar dates are handled separately to preserve the
        # 24-hour cutoff.
        mbc_by_day = {}

        day_cursor = month_start
        while day_cursor <= selected_date:
            window_end = datetime.combine(
                day_cursor,
                datetime.min.time(),
            ).replace(hour=8)

            window_start = window_end - timedelta(hours=24)

            # The current MBC SMS table stores entry_date as a date string.
            # Sum the SMS date belonging to the 24-hour reporting day.
            #
            # For the normal 08:00 -> 08:00 reporting convention this is
            # the previous calendar date.
            mbc_target_date = day_cursor - timedelta(days=1)

            cur.execute(
                """
                SELECT
                    COALESCE(SUM(COALESCE(quantity, 0)), 0) AS qty
                FROM lueu_lines
                WHERE source_type = 'MBC'
                  AND TO_DATE(entry_date, 'YYYY-MM-DD') = %s
                """,
                (mbc_target_date,),
            )

            mbc_row = cur.fetchone()
            mbc_by_day[day_cursor] = round(
                _safe_float(mbc_row["qty"] if mbc_row else 0),
                2,
            )

            day_cursor += timedelta(days=1)

        # ------------------------------------------------------------
        # 4) Daily Consumption from stats_cargo / consumption.
        # ------------------------------------------------------------
        #
        # stats_cargo stores the section data as JSON.  The five consumption
        # fields are summed exactly like the existing Cargo Statistics page.
        cur.execute(
            """
            SELECT
                entry_date,
                data
            FROM stats_cargo
            WHERE LOWER(TRIM(section)) = 'consumption'
              AND entry_date >= %s
              AND entry_date <= %s
            ORDER BY entry_date, id
            """,
            (month_start, selected_date),
        )

        consumption_rows = cur.fetchall()

        consumption_by_day = {}

        for row in consumption_rows:
            raw_data = row["data"]

            if isinstance(raw_data, str):
                try:
                    raw_data = json.loads(raw_data)
                except (TypeError, ValueError):
                    raw_data = {}

            if not isinstance(raw_data, dict):
                raw_data = {}

            # Prefer the stored Total.  If Total is not present, calculate
            # it from the five standard consumption fields.
            total_value = raw_data.get("total")

            if total_value not in (None, ""):
                total = _safe_float(total_value)
            else:
                total = sum(
                    _safe_float(raw_data.get(field))
                    for field in (
                        "ibrm",
                        "cbrm",
                        "fluxes",
                        "clinker_slag",
                        "fg_goods",
                    )
                )

            # entry_date can be returned as a date or string depending on
            # the database driver/schema.
            raw_date = row["entry_date"]

            if isinstance(raw_date, datetime):
                day = raw_date.date()
            elif isinstance(raw_date, date):
                day = raw_date
            else:
                try:
                    day = datetime.strptime(
                        str(raw_date)[:10],
                        "%Y-%m-%d",
                    ).date()
                except (TypeError, ValueError):
                    continue

            consumption_by_day[day] = round(
                consumption_by_day.get(day, 0.0) + total,
                2,
            )

        # ------------------------------------------------------------
        # 5) Build the complete 1..selected-date matrix.
        # ------------------------------------------------------------
        rows = []

        day_cursor = month_start

        while day_cursor <= selected_date:
            jetty = jetty_by_day.get(
                day_cursor,
                {"jetty": 0.0, "steel": 0.0, "cement": 0.0},
            )

            rows.append({
                "date": day_cursor.strftime("%Y-%m-%d"),
                "jetty_discharge": round(jetty["jetty"], 2),
                "steel_plant_discharge": round(jetty["steel"], 2),
                "cement_plant_discharge": round(jetty["cement"], 2),
                "mbc_discharge": round(
                    mbc_by_day.get(day_cursor, 0.0),
                    2,
                ),
                "mv_discharge": round(
                    mv_by_day.get(day_cursor, 0.0),
                    2,
                ),
                "daily_consumption": round(
                    consumption_by_day.get(day_cursor, 0.0),
                    2,
                ),
            })

            day_cursor += timedelta(days=1)

        # ------------------------------------------------------------
        # 6) Header "Avg" values = average of the displayed days.
        # ------------------------------------------------------------
        count = len(rows)

        def avg(field):
            if count <= 0:
                return 0.0

            return round(
                sum(_safe_float(row[field]) for row in rows) / count,
                2,
            )

        averages = {
            "jetty_discharge": avg("jetty_discharge"),
            "steel_plant_discharge": avg("steel_plant_discharge"),
            "cement_plant_discharge": avg("cement_plant_discharge"),
            "mbc_discharge": avg("mbc_discharge"),
            "mv_discharge": avg("mv_discharge"),
            "daily_consumption": avg("daily_consumption"),
        }

        return {
            "rows": rows,
            "averages": averages,
            "from_date": month_start.strftime("%Y-%m-%d"),
            "to_date": selected_date.strftime("%Y-%m-%d"),
        }

    finally:
        conn.close()



# ============================================================================
# DAILY PORT OPERATIONS SUMMARY API
# ============================================================================


# ============================================================================
# JETTY & MV DISCHARGE BAR CHART API
# Uses the same month-to-selected-date Daily Port Operations data.
# ============================================================================
@bp.route("/api/module/RP01/cargo_statistics/discharge-chart")
@login_required
def cargo_statistics_discharge_chart():
    selected_date_s = (request.args.get("date") or "").strip()

    if selected_date_s:
        try:
            selected_date = datetime.strptime(
                selected_date_s,
                "%Y-%m-%d",
            ).date()
        except ValueError:
            return jsonify({
                "status": "error",
                "error": "Invalid date. Expected YYYY-MM-DD.",
            }), 400
    else:
        selected_date = datetime.now().date()

    try:
        result = _fetch_daily_port_operations(selected_date)

        # Keep the chart API deliberately small.  It returns one point for
        # every calendar day from the 1st through the selected date.
        chart_rows = []
        for row in result.get("rows", []):
            chart_rows.append({
                "date": row.get("date"),
                "steel_plant_discharge": row.get("steel_plant_discharge", 0),
                "daily_consumption": row.get("daily_consumption", 0),
                "mv_discharge": row.get("mv_discharge", 0),
                "mbc_discharge": row.get("mbc_discharge", 0),
            })

        return jsonify({
            "status": "ok",
            "selected_date": selected_date.strftime("%Y-%m-%d"),
            "from_date": result.get("from_date"),
            "to_date": result.get("to_date"),
            "rows": chart_rows,
        })

    except Exception as exc:
        print(
            "DISCHARGE CHART ERROR:",
            repr(exc),
        )
        return jsonify({
            "status": "error",
            "error": str(exc),
        }), 500


@bp.route("/api/module/RP01/cargo_statistics/daily-operations")
@login_required
def cargo_statistics_daily_operations():
    selected_date_s = (request.args.get("date") or "").strip()

    if selected_date_s:
        try:
            selected_date = datetime.strptime(
                selected_date_s,
                "%Y-%m-%d",
            ).date()
        except ValueError:
            return jsonify({
                "status": "error",
                "error": "Invalid date. Expected YYYY-MM-DD.",
            }), 400
    else:
        selected_date = datetime.now().date()

    try:
        result = _fetch_daily_port_operations(selected_date)

        return jsonify({
            "status": "ok",
            "selected_date": selected_date.strftime("%Y-%m-%d"),
            **result,
        })

    except Exception as exc:
        print(
            "DAILY PORT OPERATIONS ERROR:",
            repr(exc),
        )

        return jsonify({
            "status": "error",
            "error": str(exc),
        }), 500



# ============================================================================
# RECEIPT / SUPPLY OF RM APIs
# Source: stats_cargo
# ============================================================================
RECEIPT_CONFIG = {
    "pnp": ("pnp_receipt", [
        ("Through Rake", "Jabalpur Fines", "M.P", ["rake_jabalpur_fines", "jabalpur_fines"]),
        ("Through Rake", "MEL", "Karnataka", ["rake_mel", "mel"]),
        ("Through Rake", "Surajgad Fines", "Vidharbha", ["rake_surajgad_fines", "surajgad_fines"]),
        ("Through Rake", "Corex Coal", "Vijaynagar", ["rake_corex_coal", "corex_coal", "corex"]),
        ("Through Rake", "BMM Pellets", "", ["rake_bmm_pellets", "bmm_pellets", "bmm"]),
        ("Through MV", "Limestone", "Norway Port", ["mv_limestone", "limestone_mv", "limestone"]),
    ]),
    "karanja": ("karanja_receipt", [
        ("Through MV", "OLIVINE", "Norway Port", ["olivine", "mv_olivine", "karanja_olivine"]),
        ("Receipt at Plant", "", "By Road", ["by_road", "road", "karanja_by_road"]),
        ("Receipt at Plant", "", "By Barge", ["by_barge", "barge", "karanja_by_barge"]),
    ]),
    "roha": ("roha_receipt", [
        ("Through Rake", "MEL", "Karnataka", ["rake_mel", "mel", "roha_mel"]),
    ]),
    "dolvi": ("dolvi_receipt", [
        ("Through Rake", "MEL", "Karnataka", ["rake_mel", "mel"]),
        ("Through Rake", "MMEC Vedanta Fines", "Karnataka", ["rake_mmec_vedanta_fines", "mmec_vedanta_fines"]),
        ("Through Rake", "BHQ Flux", "Karnataka", ["rake_bhq_flux", "bhq_flux"]),
        ("Through Rake", "Jabalpur Fines", "M.P", ["rake_jabalpur_fines", "jabalpur_fines"]),
        ("Through Rake", "Jamshedpur Coking Coal", "Jharkhand", ["rake_jamshedpur_coking_coal", "jamshedpur_coking_coal"]),
        ("Through Rake", "Surajgad Fines", "-", ["rake_surajgad_fines", "surajgad_fines"]),
        ("Through Rake", "ZTC", "Karnataka", ["rake_ztc", "ztc"]),
        ("Through Rake", "Limestone", "Rajasthan", ["rake_limestone", "limestone"]),
        ("Through Rake", "Dolomite", "Rajasthan", ["rake_dolomite", "dolomite"]),
        ("Through Rake", "Pellets", "Others", ["rake_pellets", "pellets"]),
        ("Through Rake", "Orissa Fines", "Orissa", ["rake_orissa_fines", "orissa_fines"]),
        ("Through Road", "Dolomite", "Rajasthan", ["road_dolomite", "dolomite_road"]),
        ("Through Road", "Limestone 40'80", "Rajasthan", ["road_limestone_40_80", "limestone_40_80"]),
    ]),
}

def _receipt_value(data, fields):
    data = data if isinstance(data, dict) else {}
    for field in fields:
        value = data.get(field)
        if value in (None, ""):
            continue
        try:
            return float(value)
        except (TypeError, ValueError):
            continue
    return 0.0

def _stats_cargo_receipt(section, start_date, end_date):
    conn = get_db()
    cur = get_cursor(conn)
    try:
        cur.execute(
            """
            SELECT entry_date, data
            FROM stats_cargo
            WHERE LOWER(TRIM(section)) = %s
              AND entry_date >= %s
              AND entry_date <= %s
            ORDER BY entry_date ASC, id ASC;
            """,
            (section.lower().strip(), start_date, end_date),
        )
        result = []
        for row in cur.fetchall():
            raw = row["data"]
            if isinstance(raw, str):
                try:
                    raw = json.loads(raw)
                except (TypeError, ValueError):
                    raw = {}
            result.append({
                "date": row["entry_date"].strftime("%Y-%m-%d") if hasattr(row["entry_date"], "strftime") else str(row["entry_date"])[:10],
                "data": raw if isinstance(raw, dict) else {},
            })
        return result
    finally:
        conn.close()

def _receipt_api_data(key, selected_date):
    section, config_rows = RECEIPT_CONFIG[key]
    fy_start_year = selected_date.year if selected_date.month >= 4 else selected_date.year - 1
    fy_start = datetime(fy_start_year, 4, 1).date()
    records = _stats_cargo_receipt(section, fy_start, selected_date)
    result = []
    for group, cargo, source, fields in config_rows:
        result.append({
            "group": group,
            "cargo": cargo,
            "source": source,
            "fields": fields,
            "daily": [
                {"date": r["date"], "value": round(_receipt_value(r["data"], fields), 2)}
                for r in records
            ],
        })
    return result

def _receipt_selected_date():
    value = (request.args.get("date") or "").strip()
    if not value:
        return datetime.now().date()
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError:
        return None

def _receipt_endpoint(key):
    selected_date = _receipt_selected_date()
    if selected_date is None:
        return jsonify({"status": "error", "error": "Invalid date. Expected YYYY-MM-DD."}), 400
    rows = _receipt_api_data(key, selected_date)
    return jsonify({
        "status": "ok",
        "section": RECEIPT_CONFIG[key][0],
        "selected_date": selected_date.strftime("%Y-%m-%d"),
        "rows": rows,
        "data": rows,
    })

@bp.route("/api/pnp_receipt")
@login_required
def cargo_statistics_pnp_receipt_data():
    return _receipt_endpoint("pnp")

@bp.route("/api/karanja_receipt")
@login_required
def cargo_statistics_karanja_receipt_data():
    return _receipt_endpoint("karanja")

@bp.route("/api/roha_receipt")
@login_required
def cargo_statistics_roha_receipt_data():
    return _receipt_endpoint("roha")

@bp.route("/api/dolvi_receipt")
@login_required
def cargo_statistics_dolvi_receipt_data():
    return _receipt_endpoint("dolvi")

# ============================================================================
# PAGE ROUTE
# ============================================================================

@bp.route("/module/RP01/cargo_statistics/")
@login_required
def cargo_statistics_index():
    return render_template(
        "cargo_statistics.html",
        username=session.get("username"),
    )


# ============================================================================
# PRODUCTION DATA API
# ============================================================================

@bp.route("/api/module/RP01/cargo_statistics/production")
@bp.route("/api/production")
@login_required
def cargo_statistics_production_data():
    """
    Return Production records from stats_cargo for the selected date's
    financial year up to the selected date.

    The Jinja page uses these records to calculate:
        Actual Production (TPD) = selected report date
        MTD                     = month start through report date
        YTD                     = 1-Apr through report date
    """
    selected_date_s = (request.args.get("date") or "").strip()

    if selected_date_s:
        try:
            selected_date = datetime.strptime(
                selected_date_s,
                "%Y-%m-%d",
            ).date()
        except ValueError:
            return jsonify({
                "status": "error",
                "error": "Invalid date. Expected YYYY-MM-DD.",
            }), 400
    else:
        selected_date = datetime.now().date()

    fy_start_year = (
        selected_date.year
        if selected_date.month >= 4
        else selected_date.year - 1
    )

    fy_start_date = datetime(
        fy_start_year,
        4,
        1,
    ).date()

    production = _stats_cargo_production(
        fy_start_date,
        selected_date,
    )

    return jsonify({
        "status": "ok",
        "selected_date": selected_date.strftime("%Y-%m-%d"),
        "production": production,
        "rows": production,
    })


# ============================================================================
# DATA API
# ============================================================================

@bp.route("/api/module/RP01/cargo_statistics/data")
@login_required
def cargo_statistics_data():
    now = datetime.now()

    selected_date_s = (request.args.get("date") or "").strip()

    if selected_date_s:
        try:
            selected_date = datetime.strptime(
                selected_date_s,
                "%Y-%m-%d",
            ).date()
        except ValueError:
            return jsonify({
                "status": "error",
                "error": "Invalid date. Expected YYYY-MM-DD.",
            }), 400
    else:
        selected_date = now.date()

    report = _build_report(selected_date)

    return jsonify({
        "status": "ok",
        **report,
        "as_of": now.strftime("%Y-%m-%d %H:%M:%S"),
    })
