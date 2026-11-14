from datetime import date, datetime
import json
from typing import Optional, List, Dict, Any

from flask import (
    Blueprint,
    render_template,
    request,
    jsonify,
    session,
    redirect,
    url_for
)

from database import get_db, get_cursor


# ============================================================
# BLUEPRINT
# ============================================================

bp = Blueprint(
    "cargo_stats",
    __name__,
    url_prefix="/module/CARGO_STATS",
    template_folder=".",
)


# ============================================================
# MODULE INFORMATION
# ============================================================

MODULE_INFO = {
    "code": "CARGO_STATS",
    "name": "Cargo Utilization",
    "description": "Daily consumption, production and stock (RMHS / PNP) tracking",
    "icon": "bar-chart-2",
}


# ============================================================
# DATABASE
# ============================================================

TABLE_NAME = "stats_cargo"

VALID_SECTIONS = {
    "consumption",
    "production",
    "stock_rmhs",
    "stock_pnp",
    "pnp_receipt",
    "karanja_receipt",
    "roha_receipt",
    "dolvi_receipt",
}

RECEIPT_SECTIONS = {
    "pnp_receipt",
    "karanja_receipt",
    "roha_receipt",
    "dolvi_receipt",
}


# ============================================================
# SECTION VALIDATION
# ============================================================

def _validate_section(section: str):

    if section not in VALID_SECTIONS:

        raise ValueError(
            f"Invalid section '{section}'. "
            f"Must be one of {sorted(VALID_SECTIONS)}"
        )

    return section


# ============================================================
# LOGIN
# ============================================================

def _require_login():

    return "user_id" in session


# ============================================================
# UPDATE EXISTING RECORD
#
# IMPORTANT:
# This function NEVER creates a new ID.
# It updates the existing row using ID.
# ============================================================

def update_entry(
    entry_id: int,
    entry_date: date,
    section: str,
    data: Dict[str, Any]
) -> int:

    section = _validate_section(section)

    conn = get_db()
    cur = get_cursor(conn)

    try:

        # ----------------------------------------------------
        # First make sure the ID belongs to this section
        # ----------------------------------------------------

        cur.execute(
            f"""
            SELECT id
            FROM {TABLE_NAME}
            WHERE id = %s
              AND section = %s;
            """,
            (
                entry_id,
                section
            )
        )

        existing = cur.fetchone()

        if not existing:

            raise ValueError(
                f"Record with ID {entry_id} "
                f"does not exist."
            )


        # ----------------------------------------------------
        # UPDATE SAME ROW
        # ----------------------------------------------------

        cur.execute(
            f"""
            UPDATE {TABLE_NAME}
            SET
                entry_date = %s,
                data = %s,
                updated_at = now()
            WHERE id = %s
              AND section = %s
            RETURNING id;
            """,
            (
                entry_date,
                json.dumps(data),
                entry_id,
                section
            )
        )

        row = cur.fetchone()

        if not row:

            raise ValueError(
                f"Unable to update record "
                f"{entry_id}."
            )

        conn.commit()

        return row["id"]

    except Exception:

        conn.rollback()

        raise

    finally:

        conn.close()


# ============================================================
# INSERT NEW RECORD
#
# This is ONLY called for a genuinely new row.
# ============================================================

def insert_entry(
    entry_date: date,
    section: str,
    data: Dict[str, Any]
) -> int:

    section = _validate_section(section)

    conn = get_db()
    cur = get_cursor(conn)

    try:

        # ----------------------------------------------------
        # Original daily sections allow only one row per date.
        # Receipt sections can contain multiple rows on the same day.
        # ----------------------------------------------------
        if section not in RECEIPT_SECTIONS:
            cur.execute(
                f"""
                SELECT id
                FROM {TABLE_NAME}
                WHERE entry_date = %s
                  AND section = %s
                LIMIT 1;
                """,
                (
                    entry_date,
                    section
                )
            )

            existing = cur.fetchone()

            if existing:
                raise ValueError(
                    f"A {section} record already "
                    f"exists for {entry_date}."
                )


        # ----------------------------------------------------
        # INSERT
        # ----------------------------------------------------

        cur.execute(
            f"""
            INSERT INTO {TABLE_NAME}
            (
                entry_date,
                section,
                data,
                updated_at
            )
            VALUES
            (
                %s,
                %s,
                %s,
                now()
            )
            RETURNING id;
            """,
            (
                entry_date,
                section,
                json.dumps(data)
            )
        )

        row = cur.fetchone()

        conn.commit()

        return row["id"]

    except Exception:

        conn.rollback()

        raise

    finally:

        conn.close()


# ============================================================
# GET RECORD BY ID
# ============================================================

def get_entry_by_id(
    entry_id: int,
    section: str
):

    section = _validate_section(section)

    conn = get_db()
    cur = get_cursor(conn)

    try:

        cur.execute(
            f"""
            SELECT
                id,
                entry_date,
                section,
                data,
                updated_at
            FROM {TABLE_NAME}
            WHERE id = %s
              AND section = %s;
            """,
            (
                entry_id,
                section
            )
        )

        return cur.fetchone()

    finally:

        conn.close()


# ============================================================
# GET HISTORY
# ============================================================

def get_history(
    section: str,
    start_date: Optional[date] = None,
    end_date: Optional[date] = None
) -> List[Dict[str, Any]]:

    section = _validate_section(section)

    where_parts = [
        "section = %s"
    ]

    params: List[Any] = [
        section
    ]

    if start_date:

        where_parts.append(
            "entry_date >= %s"
        )

        params.append(
            start_date
        )

    if end_date:

        where_parts.append(
            "entry_date <= %s"
        )

        params.append(
            end_date
        )

    where_clause = (
        "WHERE " +
        " AND ".join(where_parts)
    )

    conn = get_db()
    cur = get_cursor(conn)

    try:

        cur.execute(
            f"""
            SELECT
                id,
                entry_date,
                section,
                data,
                updated_at
            FROM {TABLE_NAME}
            {where_clause}
            ORDER BY entry_date DESC, id DESC;
            """,
            params
        )

        return cur.fetchall()

    finally:

        conn.close()


# ============================================================
# DELETE BY ID
# ============================================================

def delete_entry_by_id(
    entry_id: int,
    section: str
):

    section = _validate_section(section)

    conn = get_db()
    cur = get_cursor(conn)

    try:

        cur.execute(
            f"""
            DELETE FROM {TABLE_NAME}
            WHERE id = %s
              AND section = %s;
            """,
            (
                entry_id,
                section
            )
        )

        conn.commit()

    except Exception:

        conn.rollback()

        raise

    finally:

        conn.close()


# ============================================================
# SERIALIZE
# ============================================================

def serialize_row(row):

    if row is None:

        return None

    raw_data = row["data"]

    if isinstance(raw_data, str):

        raw_data = json.loads(
            raw_data
        )

    out = dict(
        raw_data or {}
    )

    # --------------------------------------------------------
    # IMPORTANT:
    # Always return database ID.
    # JavaScript uses this ID for updates.
    # --------------------------------------------------------

    out["id"] = row["id"]


    # --------------------------------------------------------
    # Date
    # --------------------------------------------------------

    ed = row["entry_date"]

    if isinstance(ed, date):

        out["date"] = ed.isoformat()

    else:

        out["date"] = ed


    # --------------------------------------------------------
    # Section
    # --------------------------------------------------------

    out["section"] = row["section"]


    # --------------------------------------------------------
    # Updated
    # --------------------------------------------------------

    ua = row["updated_at"]

    if isinstance(ua, datetime):

        out["updated_at"] = ua.isoformat()

    else:

        out["updated_at"] = ua


    return out


# ============================================================
# PAGE
# ============================================================

@bp.route("/")
def index():

    if not _require_login():

        return redirect(
            url_for("login")
        )

    return render_template(
        "stats.html"
    )


# ============================================================
# GET DATA
# ============================================================

@bp.route(
    "/api/<section>",
    methods=["GET"]
)
def api_list(section):

    if not _require_login():

        return jsonify({
            "error": "Not logged in"
        }), 401

    try:

        rows = get_history(
            section
        )

    except ValueError as e:

        return jsonify({
            "error": str(e)
        }), 400

    return jsonify([
        serialize_row(row)
        for row in rows
    ])


# ============================================================
# SAVE
#
# EXISTING ROW:
#     ID exists -> UPDATE
#
# NEW ROW:
#     ID doesn't exist -> INSERT
# ============================================================

@bp.route(
    "/api/<section>",
    methods=["POST"]
)
def api_upsert(section):

    if not _require_login():

        return jsonify({
            "error": "Not logged in"
        }), 401


    # --------------------------------------------------------
    # Validate section
    # --------------------------------------------------------

    try:

        _validate_section(
            section
        )

    except ValueError as e:

        return jsonify({
            "error": str(e)
        }), 400


    # --------------------------------------------------------
    # Read JSON
    # --------------------------------------------------------

    payload = request.get_json(
        silent=True
    ) or {}


    entry_id = payload.get(
        "id"
    )

    entry_date_str = payload.get(
        "date"
    )

    entry_data = payload.get(
        "data"
    )


    # --------------------------------------------------------
    # Validate date
    # --------------------------------------------------------

    # Receipt tables do not display a Date column.
    # Their database date is automatically set to today when omitted.
    if not entry_date_str and section in RECEIPT_SECTIONS:
        entry_date_str = date.today().isoformat()

    if not entry_date_str:
        return jsonify({
            "error": "Date is required."
        }), 400

    try:
        entry_date = date.fromisoformat(entry_date_str)

    except ValueError:

        return jsonify({
            "error": (
                "Invalid date. "
                "Expected YYYY-MM-DD."
            )
        }), 400


    # --------------------------------------------------------
    # Validate data
    # --------------------------------------------------------

    if not isinstance(
        entry_data,
        dict
    ):

        return jsonify({
            "error": "Data must be an object."
        }), 400


    # ========================================================
    # EXISTING ROW
    #
    # ID IS PRESENT
    #
    # UPDATE ONLY
    # ========================================================

    if entry_id is not None:

        try:

            entry_id = int(
                entry_id
            )

        except (
            ValueError,
            TypeError
        ):

            return jsonify({
                "error": "Invalid record ID."
            }), 400


        try:

            updated_id = update_entry(
                entry_id=entry_id,
                entry_date=entry_date,
                section=section,
                data=entry_data
            )


        except ValueError as e:

            return jsonify({
                "error": str(e)
            }), 404


        except Exception as e:

            error_text = str(e)

            if (
                "uq_stats_cargo_entry_date_section"
                in error_text
            ):

                return jsonify({
                    "error": (
                        f"A {section} record already "
                        f"exists for {entry_date}."
                    )
                }), 409

            return jsonify({
                "error": error_text
            }), 500


        # ----------------------------------------------------
        # Get SAME ROW after update
        # ----------------------------------------------------

        saved = get_entry_by_id(
            updated_id,
            section
        )


        return jsonify({
            "success": True,
            "created": False,
            "updated": True,
            "id": updated_id,
            "data": serialize_row(saved)
        }), 200


    # ========================================================
    # NEW ROW
    #
    # NO ID
    #
    # ONLY HERE DO WE INSERT
    # ========================================================

    try:

        new_id = insert_entry(
            entry_date=entry_date,
            section=section,
            data=entry_data
        )

    except ValueError as e:

        return jsonify({
            "error": str(e)
        }), 409

    except Exception as e:

        error_text = str(e)

        if (
            "uq_stats_cargo_entry_date_section"
            in error_text
        ):

            return jsonify({
                "error": (
                    f"A {section} record already "
                    f"exists for {entry_date}."
                )
            }), 409

        return jsonify({
            "error": error_text
        }), 500


    saved = get_entry_by_id(
        new_id,
        section
    )


    return jsonify({
        "success": True,
        "created": True,
        "updated": False,
        "id": new_id,
        "data": serialize_row(saved)
    }), 200


# ============================================================
# DELETE DYNAMIC COLUMN
# ============================================================

@bp.route(
    "/api/<section>/column/<field>",
    methods=["DELETE"]
)
def api_delete_column(section, field):

    if not _require_login():
        return jsonify({
            "success": False,
            "error": "Not logged in"
        }), 401

    conn = None

    try:
        section = _validate_section(section)

        field = (field or "").strip()

        if not field:
            return jsonify({
                "success": False,
                "error": "Column field is required."
            }), 400

        conn = get_db()
        cur = get_cursor(conn)

        # --------------------------------------------------------
        # Get every record in this section.
        # --------------------------------------------------------

        cur.execute(
            f"""
            SELECT id, data
            FROM {TABLE_NAME}
            WHERE section = %s
            ORDER BY id;
            """,
            (section,)
        )

        rows = cur.fetchall()

        affected_rows = 0

        def json_dict(raw):
            """
            Convert PostgreSQL JSON / string / bytes into a dict.
            """
            if raw is None:
                return {}

            if isinstance(raw, dict):
                return dict(raw)

            if isinstance(raw, bytes):
                try:
                    raw = raw.decode("utf-8")
                except Exception:
                    return {}

            if isinstance(raw, str):
                try:
                    value = json.loads(raw)
                except Exception:
                    return {}

                return value if isinstance(value, dict) else {}

            try:
                value = dict(raw)
                return value if isinstance(value, dict) else {}
            except Exception:
                return {}

        # --------------------------------------------------------
        # Remove the column from EVERY database row.
        # --------------------------------------------------------

        for row in rows:

            row_id = row["id"]
            data = json_dict(row["data"])

            if not data:
                continue

            changed = False

            # Exact key from the URL.
            if field in data:
                del data[field]
                changed = True

            # Also match normalized keys.
            normalized_field = (
                field.strip()
                .lower()
                .replace(" ", "_")
                .replace("-", "_")
            )

            for key in list(data.keys()):

                if str(key).startswith("__"):
                    continue

                normalized_key = (
                    str(key)
                    .strip()
                    .lower()
                    .replace(" ", "_")
                    .replace("-", "_")
                )

                if normalized_key == normalized_field:
                    del data[key]
                    changed = True

            # ----------------------------------------------------
            # Remove metadata entry.
            # ----------------------------------------------------

            dynamic_columns = data.get(
                "__dynamic_columns"
            )

            if isinstance(dynamic_columns, dict):

                for metadata_key in list(
                    dynamic_columns.keys()
                ):

                    normalized_metadata_key = (
                        str(metadata_key)
                        .strip()
                        .lower()
                        .replace(" ", "_")
                        .replace("-", "_")
                    )

                    if (
                        metadata_key == field
                        or
                        normalized_metadata_key
                        == normalized_field
                    ):
                        del dynamic_columns[
                            metadata_key
                        ]
                        changed = True

                if dynamic_columns:

                    data["__dynamic_columns"] = (
                        dynamic_columns
                    )

                else:

                    data.pop(
                        "__dynamic_columns",
                        None
                    )

            # ----------------------------------------------------
            # IMPORTANT:
            # Write the COMPLETE cleaned JSON back.
            # This is a PostgreSQL JSON column, so we send a
            # JSON string, NOT a jsonb expression or SQL record.
            # ----------------------------------------------------

            if changed:

                cleaned_json = json.dumps(
                    data,
                    ensure_ascii=False
                )

                cur.execute(
                    f"""
                    UPDATE {TABLE_NAME}
                    SET
                        data = %s,
                        updated_at = now()
                    WHERE id = %s
                      AND section = %s;
                    """,
                    (
                        cleaned_json,
                        row_id,
                        section
                    )
                )

                affected_rows += 1

        conn.commit()

        # --------------------------------------------------------
        # HARD verification.
        # --------------------------------------------------------

        cur.execute(
            f"""
            SELECT id, data
            FROM {TABLE_NAME}
            WHERE section = %s
            ORDER BY id;
            """,
            (section,)
        )

        remaining = []

        for row in cur.fetchall():

            data = json_dict(row["data"])

            # Actual field still present.
            if field in data:
                remaining.append(row["id"])
                continue

            # Normalized field still present.
            normalized_field = (
                field.strip()
                .lower()
                .replace(" ", "_")
                .replace("-", "_")
            )

            found = False

            for key in data.keys():

                if str(key).startswith("__"):
                    continue

                normalized_key = (
                    str(key)
                    .strip()
                    .lower()
                    .replace(" ", "_")
                    .replace("-", "_")
                )

                if normalized_key == normalized_field:
                    found = True
                    break

            if found:
                remaining.append(row["id"])
                continue

            # Metadata still present.
            dynamic_columns = data.get(
                "__dynamic_columns"
            )

            if isinstance(dynamic_columns, dict):

                for metadata_key in dynamic_columns.keys():

                    normalized_metadata_key = (
                        str(metadata_key)
                        .strip()
                        .lower()
                        .replace(" ", "_")
                        .replace("-", "_")
                    )

                    if (
                        metadata_key == field
                        or
                        normalized_metadata_key
                        == normalized_field
                    ):
                        remaining.append(row["id"])
                        break

        if remaining:

            conn.rollback()

            return jsonify({
                "success": False,
                "deleted": False,
                "field": field,
                "affected_rows": affected_rows,
                "remaining_rows": remaining,
                "error": (
                    f'Value for column "{field}" still '
                    f"exists in database rows: {remaining}"
                )
            }), 500

        return jsonify({
            "success": True,
            "deleted": True,
            "field": field,
            "affected_rows": affected_rows
        }), 200

    except Exception as e:

        if conn:
            conn.rollback()

        return jsonify({
            "success": False,
            "deleted": False,
            "error": str(e)
        }), 500

    finally:

        if conn:
            conn.close()


# ============================================================
# DELETE
# ============================================================

@bp.route(
    "/api/<section>/id/<int:entry_id>",
    methods=["DELETE"]
)
def api_delete_by_id(
    section,
    entry_id
):

    if not _require_login():

        return jsonify({
            "error": "Not logged in"
        }), 401

    try:

        delete_entry_by_id(
            entry_id,
            section
        )

    except ValueError as e:

        return jsonify({
            "error": str(e)
        }), 400

    return jsonify({
        "success": True,
        "deleted": True
    }), 200