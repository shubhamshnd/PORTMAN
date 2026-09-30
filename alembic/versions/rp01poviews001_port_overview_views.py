"""RP01 Port Overview: read-only reporting views (vw_port_overview_*)

Exposes every dataset behind /module/RP01/port-overview/ as a plain view so a
reporting user can be granted SELECT on them and build their own dashboards.
Views only — no table, column or row is touched. Each view carries ALL rows
(no "today" filter); the consumer filters by date/shift/FY.

Logic mirrors modules/RP01/RP01/port_overview/views.py and the helpers it
calls (Barge_Position_Report._fetch_all_barges/_fetch_tide_data,
shift_report._fetch_shift_pivot/_fetch_delays). Keep in sync if those change.

Grant access with e.g.:
    GRANT SELECT ON ALL <view names> TO <reporting_user>;

Revision ID: rp01poviews001
Revises: rp02cargobd001
Create Date: 2026-09-30
"""
from typing import Sequence, Union
from alembic import op

revision: str = 'rp01poviews001'
down_revision: Union[str, None] = 'rp02cargobd001'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


# Text timestamps are only compared/cast when they look ISO, so one bad
# value in a source row can never make a whole view error out.
_ISO = r"'^\d{4}-\d{2}-\d{2}'"
_NUM = r"'^\s*-?\d+(\.\d+)?\s*$'"

VIEWS = {
    # --- Berth map: drawn berth boxes on Clean_berths.png (941x1672) ---------
    'vw_port_overview_berth_layout': """
        SELECT b.id AS berth_id,
               b.berth_name,
               b.berth_sequence,
               (b.image_position->>'cx')::numeric                AS cx,
               (b.image_position->>'cy')::numeric                AS cy,
               (b.image_position->>'w')::numeric                 AS w,
               (b.image_position->>'h')::numeric                 AS h,
               COALESCE((b.image_position->>'angle')::numeric, 0) AS angle
        FROM port_berth_master b
        WHERE b.image_position IS NOT NULL
          AND jsonb_typeof(b.image_position) = 'object'
    """,

    # --- Live barges/MBCs at port (_fetch_all_barges) ------------------------
    'vw_port_overview_live_assets': f"""
        WITH discharge_sums AS (
            SELECT TRIM(UPPER(ll.barge_name)) AS barge_name, ll.source_id,
                   SUM(COALESCE(ll.quantity, 0)) AS discharged_qty
            FROM lueu_lines ll
            WHERE ll.is_deleted IS NOT TRUE AND ll.source_type = 'VCN'
            GROUP BY TRIM(UPPER(ll.barge_name)), ll.source_id
        ),
        barges AS (
            SELECT l.id, 'BARGE'::text AS asset_type, l.barge_name AS name,
                   l.trip_number, l.cargo_name,
                   COALESCE(l.discharge_quantity, 0)::numeric AS total_qty,
                   GREATEST(COALESCE(l.discharge_quantity, 0) - COALESCE(ds.discharged_qty, 0), 0)::numeric AS balance_qty,
                   COALESCE(NULLIF(TRIM(l.commence_discharge_berth), ''), NULLIF(TRIM(l.along_side_berth), ''), '') AS berth,
                   l.along_side_berth AS arrived_at,
                   l.commence_discharge_berth AS commenced_at,
                   l.completed_discharge_berth AS completed_at,
                   CASE WHEN NULLIF(TRIM(l.completed_discharge_berth), '') IS NOT NULL THEN 'Discharge Completed'
                        WHEN NULLIF(TRIM(l.commence_discharge_berth), '') IS NOT NULL THEN 'Under Discharge'
                        ELSE 'Waiting' END AS status
            FROM ldud_barge_lines l
            LEFT JOIN ldud_header h ON h.id = l.ldud_id
            LEFT JOIN discharge_sums ds
                   ON ds.barge_name = TRIM(UPPER(CONCAT(l.barge_name, ' / ', COALESCE(l.trip_number::text, '1'))))
                  AND ds.source_id = h.vcn_id
            WHERE COALESCE(TRIM(l.barge_name), '') <> ''
              AND NOT (TRIM(l.along_side_berth) ~ {_ISO} AND TRIM(l.along_side_berth) < '2026-05-01')
              AND NULLIF(TRIM(l.cast_off_port), '') IS NULL
              AND COALESCE(NULLIF(TRIM(l.completed_discharge_berth), ''), NULLIF(TRIM(l.commence_discharge_berth), ''),
                           NULLIF(TRIM(l.along_side_berth), '')) IS NOT NULL
        ),
        latest_mbc AS (
            SELECT h.id, h.mbc_name, h.cargo_name,
                   COALESCE(h.bl_quantity, 0) AS bl_qty,
                   COALESCE(p.vessel_unloading_berth, e.berth_master)       AS berth,
                   COALESCE(p.vessel_arrival_port,   e.arrived_at_port)     AS arrival_port,
                   COALESCE(p.unloading_commenced,   e.loading_commenced)   AS unloading_commenced,
                   COALESCE(p.unloading_completed,   e.loading_completed)   AS unloading_completed,
                   COALESCE(p.vessel_cast_off,       e.cast_off_from_berth) AS mbc_cast_off,
                   ROW_NUMBER() OVER (PARTITION BY h.id ORDER BY COALESCE(p.id, e.id) DESC) AS rn
            FROM mbc_header h
            LEFT JOIN mbc_discharge_port_lines p ON p.mbc_id = h.id AND h.operation_type ILIKE 'Import'
            LEFT JOIN mbc_export_load_port_lines e ON e.mbc_id = h.id AND h.operation_type ILIKE 'Export'
            WHERE p.mbc_id IS NOT NULL OR e.mbc_id IS NOT NULL
        ),
        mbcs AS (
            SELECT m.id, 'MBC'::text, m.mbc_name, NULL::integer, m.cargo_name,
                   m.bl_qty::numeric,
                   GREATEST(m.bl_qty - COALESCE(q.actual_qty, 0), 0)::numeric,
                   UPPER(COALESCE(m.berth, '')),
                   m.arrival_port, m.unloading_commenced, m.unloading_completed,
                   CASE WHEN NULLIF(TRIM(m.unloading_completed), '') IS NOT NULL THEN 'Discharge Completed'
                        WHEN NULLIF(TRIM(m.unloading_commenced), '') IS NOT NULL THEN 'Under Discharge'
                        ELSE 'Waiting' END
            FROM latest_mbc m
            LEFT JOIN (SELECT source_id, SUM(COALESCE(quantity, 0)) AS actual_qty
                       FROM lueu_lines
                       WHERE source_type = 'MBC' AND is_deleted IS NOT TRUE
                       GROUP BY source_id) q ON q.source_id = m.id
            WHERE m.rn = 1
              AND NULLIF(TRIM(m.arrival_port), '') IS NOT NULL
              AND NOT (TRIM(m.arrival_port) ~ {_ISO} AND TRIM(m.arrival_port) < '2026-05-01')
              AND NULLIF(TRIM(m.mbc_cast_off), '') IS NULL
        )
        SELECT * FROM barges
        UNION ALL
        SELECT * FROM mbcs
    """,

    # --- Latest LUEU logged interval per barge/MBC ---------------------------
    'vw_port_overview_last_lueu_completion': """
        SELECT DISTINCT ON (UPPER(TRIM(barge_name)))
               UPPER(TRIM(barge_name)) AS barge_key,
               entry_date, from_time, to_time,
               TRIM(entry_date || ' ' || to_time) AS last_completed
        FROM lueu_lines
        WHERE is_deleted IS NOT TRUE
          AND COALESCE(TRIM(barge_name), '') <> ''
          AND COALESCE(TRIM(to_time), '') <> ''
          AND COALESCE(TRIM(entry_date), '') <> ''
        ORDER BY UPPER(TRIM(barge_name)), entry_date DESC, id DESC
    """,

    # --- Berth occupancy: every saved Barge Position layout item, refreshed
    #     with live balances. is_latest_layout = what the dashboard shows. ---
    'vw_port_overview_berth_occupancy': f"""
        WITH reports AS (
            SELECT r.*,
                   ROW_NUMBER() OVER (ORDER BY r.report_date DESC,
                                      CASE r.shift WHEN 'C' THEN 3 WHEN 'B' THEN 2 WHEN 'A' THEN 1 ELSE 0 END DESC,
                                      r.updated_at DESC) = 1 AS is_latest_layout
            FROM barge_position_report r
            WHERE jsonb_typeof(r.berth_layout) = 'array'
              AND jsonb_array_length(r.berth_layout) > 0
        ),
        items AS (
            SELECT r.id AS report_id, r.report_date, r.shift, r.is_latest_layout,
                   UPPER(TRIM(i->>'berth')) AS berth,
                   TRIM(i->>'name') AS name,
                   UPPER(COALESCE(NULLIF(i->>'type', ''), 'BARGE')) AS asset_type,
                   UPPER(COALESCE(NULLIF(i->>'position', ''), 'A/S')) AS position,
                   NULLIF(i->>'cargo', '') AS saved_cargo,
                   CASE WHEN COALESCE(NULLIF(i->>'total', ''), i->>'qty') ~ {_NUM}
                        THEN COALESCE(NULLIF(i->>'total', ''), i->>'qty')::numeric END AS saved_total_qty,
                   CASE WHEN i->>'balance' ~ {_NUM} THEN (i->>'balance')::numeric END AS saved_balance_qty,
                   NULLIF(TRIM(COALESCE(NULLIF(TRIM(i->>'unloading_commenced'), ''), i->>'commence_discharge_berth')), '') AS saved_commenced
            FROM reports r
            CROSS JOIN LATERAL jsonb_array_elements(r.berth_layout) i
        ),
        live AS (
            SELECT DISTINCT ON (UPPER(TRIM(name))) *
            FROM vw_port_overview_live_assets
            ORDER BY UPPER(TRIM(name)), id DESC
        ),
        merged AS (
            SELECT it.*,
                   COALESCE(NULLIF(lv.cargo_name, ''), it.saved_cargo, '') AS cargo,
                   COALESCE(NULLIF(lv.total_qty, 0), it.saved_total_qty, 0) AS total_qty,
                   COALESCE(NULLIF(lv.balance_qty, 0), it.saved_balance_qty, 0) AS balance_qty,
                   COALESCE(it.saved_commenced, NULLIF(TRIM(lv.commenced_at), '')) AS commenced,
                   lv.id IS NOT NULL AS matched_live,
                   lc.last_completed
            FROM items it
            LEFT JOIN live lv ON UPPER(TRIM(lv.name)) = UPPER(it.name)
            LEFT JOIN vw_port_overview_last_lueu_completion lc ON lc.barge_key = UPPER(it.name)
            WHERE COALESCE(it.berth, '') NOT IN ('', 'WAITING')
              AND COALESCE(it.name, '') <> ''
        )
        SELECT report_id, report_date, shift, is_latest_layout,
               berth, name, asset_type, cargo, position,
               CASE position WHEN 'A/S' THEN 0 WHEN 'D/B' THEN 1 WHEN 'T/B' THEN 2
                             WHEN 'F/B' THEN 3 WHEN 'S/B' THEN 4 ELSE 0 END AS tier,
               total_qty, balance_qty,
               CASE WHEN commenced IS NULL THEN 'Waiting'
                    WHEN balance_qty <= 0 THEN 'Discharge Completed'
                    ELSE 'Under Discharge' END AS status,
               commenced, last_completed, matched_live
        FROM merged
    """,

    # --- Upcoming arrivals: left Gull Island, nothing logged at port yet ----
    'vw_port_overview_upcoming_arrivals': """
        WITH b AS (
            SELECT DISTINCT ON (UPPER(TRIM(l.barge_name)))
                   TRIM(l.barge_name) AS name, l.cargo_name,
                   COALESCE(l.discharge_quantity, 0)::numeric AS qty,
                   COALESCE(NULLIF(TRIM(l.aweigh_gull_island_empty::text), ''), NULLIF(TRIM(l.aweigh_gull_island::text), '')) AS since,
                   COALESCE(NULLIF(TRIM(l.completed_loading::text), ''), NULLIF(TRIM(l.cast_off_mv::text), '')) AS loaded,
                   COALESCE(NULLIF(TRIM(l.amf_at_port::text), ''), NULLIF(TRIM(l.along_side_berth::text), ''),
                            NULLIF(TRIM(l.commence_discharge_berth::text), ''), NULLIF(TRIM(l.completed_discharge_berth::text), ''),
                            NULLIF(TRIM(l.cast_off_berth::text), ''), NULLIF(TRIM(l.cast_off_port::text), '')) AS at_port
            FROM ldud_barge_lines l
            JOIN ldud_header h ON h.id = l.ldud_id
            WHERE COALESCE(TRIM(l.barge_name), '') <> ''
              AND COALESCE(h.operation_type, '') <> 'Export'
            ORDER BY UPPER(TRIM(l.barge_name)), l.id DESC
        ),
        m AS (
            SELECT DISTINCT ON (UPPER(TRIM(h.mbc_name)))
                   TRIM(h.mbc_name) AS name, h.cargo_name,
                   COALESCE(h.bl_quantity, 0)::numeric AS qty,
                   NULLIF(TRIM(d.departure_gull_island::text), '') AS since,
                   COALESCE(NULLIF(TRIM(d.arrived_yellow_crane::text), ''), NULLIF(TRIM(d.vessel_arrival_port::text), ''),
                            NULLIF(TRIM(d.vessel_all_made_fast::text), ''), NULLIF(TRIM(d.unloading_commenced::text), ''),
                            NULLIF(TRIM(d.cleaning_commenced::text), ''), NULLIF(TRIM(d.cleaning_completed::text), ''),
                            NULLIF(TRIM(d.unloading_completed::text), ''), NULLIF(TRIM(d.vessel_cast_off::text), ''),
                            NULLIF(TRIM(d.sailed_out_load_port::text), '')) AS at_port
            FROM mbc_discharge_port_lines d
            JOIN mbc_header h ON h.id = d.mbc_id
            WHERE COALESCE(TRIM(h.mbc_name), '') <> ''
            ORDER BY UPPER(TRIM(h.mbc_name)), d.id DESC
        )
        SELECT 'BARGE'::text AS asset_type, name, COALESCE(cargo_name, '') AS cargo, qty, since AS departed_gull_island
        FROM b WHERE since IS NOT NULL AND loaded IS NOT NULL AND at_port IS NULL
        UNION ALL
        SELECT 'MBC', name, COALESCE(cargo_name, ''), qty, since
        FROM m WHERE since IS NOT NULL AND at_port IS NULL
    """,

    # --- Shift notes + in-charge (one row per note; reports without notes kept) ---
    'vw_port_overview_shift_notes': """
        SELECT r.report_date, r.shift, TRIM(COALESCE(r.shift_incharge, '')) AS shift_incharge,
               n.ord AS note_no, n.note
        FROM barge_position_report r
        LEFT JOIN LATERAL jsonb_array_elements_text(
                CASE WHEN jsonb_typeof(r.notes) = 'array' THEN r.notes ELSE '[]'::jsonb END
             ) WITH ORDINALITY AS n(note, ord) ON TRUE
    """,

    # --- Cargo throughput: one row per LUEU quantity line (live + historical).
    #     Today/Yesterday/Month cards = LIVE rows by entry_date.
    #     FY card = rows where counts_toward_fy (historical April + live May-Mar).
    #     All Time = FY card + vw_port_overview_cargo_baseline. -------------
    'vw_port_overview_cargo_throughput': f"""
        WITH vc AS (
            SELECT DISTINCT ON (UPPER(TRIM(cargo_name))) UPPER(TRIM(cargo_name)) AS k, cargo_type
            FROM vessel_cargo ORDER BY UPPER(TRIM(cargo_name)), id
        ),
        rows AS (
            SELECT 'LIVE'::text AS source, l.id AS line_id,
                   TO_DATE(BTRIM(l.entry_date), 'YYYY-MM-DD') AS entry_date,
                   NULLIF(TRIM(l.shift), '') AS shift, l.cargo_name, l.barge_name,
                   l.equipment_name, l.route_name, l.berth_name,
                   l.quantity::numeric AS quantity
            FROM lueu_lines l
            WHERE l.is_deleted IS NOT TRUE
              AND l.quantity IS NOT NULL
              AND BTRIM(l.entry_date) ~ {_ISO}
            UNION ALL
            SELECT 'HISTORICAL', h.id, h.entry_date,
                   NULLIF(TRIM(h.shift), ''), h.cargo_name, h.barge_name,
                   h.equipment_name, h.route_name, h.berth_name,
                   h.quantity
            FROM rp01_historical_lueu h
            WHERE h.quantity IS NOT NULL AND h.entry_date IS NOT NULL
        )
        SELECT r.source, r.line_id, r.entry_date, r.shift,
               r.cargo_name, COALESCE(vc.cargo_type, 'OTHERS') AS cargo_type,
               r.barge_name, r.equipment_name, r.route_name, r.berth_name, r.quantity,
               EXTRACT(YEAR FROM r.entry_date - INTERVAL '3 months')::int AS fy_start_year,
               (EXTRACT(YEAR FROM r.entry_date - INTERVAL '3 months')::int || '-' ||
                EXTRACT(YEAR FROM r.entry_date - INTERVAL '3 months')::int + 1) AS financial_year,
               DATE_TRUNC('month', r.entry_date)::date AS month_start,
               (r.source = 'HISTORICAL') = (EXTRACT(MONTH FROM r.entry_date) = 4) AS counts_toward_fy
        FROM rows r
        LEFT JOIN vc ON vc.k = UPPER(TRIM(r.cargo_name))
    """,

    # --- All-time baseline per cargo_type, FY2012-13 .. FY2025-26 ------------
    'vw_port_overview_cargo_baseline': """
        SELECT * FROM (VALUES
            ('IBRM', 116229319::numeric), ('FLUXES', 24038328), ('CBRM', 52438887),
            ('CLINKER', 3453541), ('SLAG', 1439086), ('FINISH GOODS', 157549),
            ('OTHERS', 354512)
        ) AS t(cargo_type, qty_till_mar_2026)
    """,

    # --- FY targets: month level (effective = Outlook if set, else Base) ----
    'vw_port_overview_fy_targets': """
        SELECT f.financial_year,
               (t->>'month_num')::int AS month_num,
               t->>'month' AS month_name,
               COALESCE(NULLIF(t->>'base_target', '')::numeric, 0) AS base_target,
               NULLIF(t->>'outlook', '')::numeric AS outlook,
               COALESCE(NULLIF(t->>'outlook', '')::numeric, NULLIF(t->>'base_target', '')::numeric, 0) AS effective_target,
               f.updated_at
        FROM financial_year_targets f
        CROSS JOIN LATERAL jsonb_array_elements(
                CASE WHEN jsonb_typeof(f.targets) = 'array' THEN f.targets ELSE '[]'::jsonb END) t
        WHERE jsonb_typeof(t) = 'object'
    """,

    # --- FY targets: category level (IBRM/CBRM/FLUXES/CLINKER/SLAG) ---------
    'vw_port_overview_fy_category_targets': """
        SELECT f.financial_year,
               (t->>'month_num')::int AS month_num,
               t->>'month' AS month_name,
               c.key AS category,
               COALESCE(NULLIF(c.value->>'base', '')::numeric, 0) AS base_target,
               NULLIF(c.value->>'outlook', '')::numeric AS outlook,
               COALESCE(NULLIF(c.value->>'outlook', '')::numeric, NULLIF(c.value->>'base', '')::numeric, 0) AS effective_target
        FROM financial_year_targets f
        CROSS JOIN LATERAL jsonb_array_elements(
                CASE WHEN jsonb_typeof(f.targets) = 'array' THEN f.targets ELSE '[]'::jsonb END) t
        CROSS JOIN LATERAL jsonb_each(
                CASE WHEN jsonb_typeof(t->'categories') = 'object' THEN t->'categories' ELSE '{}'::jsonb END) c
        WHERE jsonb_typeof(t) = 'object'
    """,

    # --- Delays per LUEU line. Dashboard's "top delays" = entry_date = today,
    #     excluding in_top_delays = false, summed by delay_label. -----------
    'vw_port_overview_delays': r"""
        WITH d AS (
            SELECT l.id AS line_id, l.entry_date, l.shift, l.delay_name,
                   COALESCE(t.type, 'Other') AS delay_type,
                   l.equipment_name, l.system_name, l.route_name,
                   TRIM(l.from_time) AS from_time, TRIM(l.to_time) AS to_time,
                   CASE WHEN TRIM(l.from_time) ~ '^\d{1,2}:\d{2}$' AND TRIM(l.to_time) ~ '^\d{1,2}:\d{2}$'
                             AND split_part(TRIM(l.from_time), ':', 1)::int < 24 AND split_part(TRIM(l.from_time), ':', 2)::int < 60
                             AND split_part(TRIM(l.to_time), ':', 1)::int < 24 AND split_part(TRIM(l.to_time), ':', 2)::int < 60
                        THEN (((split_part(TRIM(l.to_time), ':', 1)::int * 60 + split_part(TRIM(l.to_time), ':', 2)::int)
                             - (split_part(TRIM(l.from_time), ':', 1)::int * 60 + split_part(TRIM(l.from_time), ':', 2)::int)
                             + 1440) % 1440)
                        ELSE 0 END AS total_minutes
            FROM lueu_lines l
            LEFT JOIN port_delay_types t ON t.name = l.delay_name
            WHERE l.delay_name IS NOT NULL AND l.delay_name <> ''
              AND l.is_deleted IS NOT TRUE
        )
        SELECT d.*,
               CASE WHEN LOWER(REPLACE(d.delay_type, ' ', '')) = 'maintenancedelays'
                    THEN NULLIF(TRIM(d.system_name), '')
                    ELSE NULLIF(TRIM(d.equipment_name), '') END AS delay_source,
               COALESCE(
                   (CASE WHEN LOWER(REPLACE(d.delay_type, ' ', '')) = 'maintenancedelays'
                         THEN NULLIF(TRIM(d.system_name), '')
                         ELSE NULLIF(TRIM(d.equipment_name), '') END) || ' — ', ''
               ) || TRIM(d.delay_name) AS delay_label,
               LOWER(TRIM(d.delay_name)) NOT IN ('idle', 'unloading') AS in_top_delays
        FROM d
    """,

    # --- MBC voyage-stage status (JSW Infra / JSW Shipping fleet) ------------
    'vw_port_overview_mbc_status': """
        SELECT TRIM(m.mbc_name) AS mbc_name,
               CASE
                   WHEN h.id IS NULL THEN NULL
                   WHEN NULLIF(TRIM(l.arrived_load_port), '') IS NOT NULL
                        AND NULLIF(TRIM(l.loading_commenced), '') IS NULL THEN NULL
                   WHEN NULLIF(TRIM(d.unloading_completed), '') IS NOT NULL THEN NULL
                   WHEN COALESCE(NULLIF(TRIM(h.cargo_name), ''), '') <> '' THEN TRIM(h.cargo_name)
                   ELSE NULL
               END AS cargo_name,
               CASE
                   WHEN h.id IS NULL THEN 'EMPTY : WAITING AT JAIGAD'
                   WHEN NULLIF(TRIM(l.arrived_load_port), '') IS NOT NULL
                        AND NULLIF(TRIM(l.loading_commenced), '') IS NULL THEN 'EMPTY : WAITING AT LOAD PORT'
                   WHEN NULLIF(TRIM(l.loading_commenced), '') IS NOT NULL
                        AND NULLIF(TRIM(l.loading_completed), '') IS NULL THEN 'UNDER LOADING'
                   WHEN NULLIF(TRIM(l.loading_completed), '') IS NOT NULL
                        AND NULLIF(TRIM(l.cast_off_load_port), '') IS NULL THEN 'LOADED : WAITING AT LOAD PORT'
                   WHEN NULLIF(TRIM(l.cast_off_load_port), '') IS NOT NULL
                        AND NULLIF(TRIM(d.arrival_gull_island), '') IS NULL THEN 'LOADED : ON THE WAY TO GULL'
                   WHEN NULLIF(TRIM(d.arrival_gull_island), '') IS NOT NULL
                        AND NULLIF(TRIM(d.departure_gull_island), '') IS NULL THEN 'LOADED : WAITING AT GULL'
                   WHEN NULLIF(TRIM(d.departure_gull_island), '') IS NOT NULL
                        AND NULLIF(TRIM(d.vessel_arrival_port), '') IS NULL THEN 'LOADED : ON THE WAY TO DHARAMTAR'
                   WHEN NULLIF(TRIM(d.vessel_arrival_port), '') IS NOT NULL
                        AND NULLIF(TRIM(d.unloading_commenced), '') IS NULL THEN 'LOADED : WAITING AT DHARAMTAR'
                   WHEN NULLIF(TRIM(d.unloading_commenced), '') IS NOT NULL
                        AND NULLIF(TRIM(d.unloading_completed), '') IS NULL THEN 'UNDER DISCHARGE AT DHARAMTAR'
                   WHEN NULLIF(TRIM(d.unloading_completed), '') IS NOT NULL
                        AND d.sailed_out_load_port IS NOT NULL
                        AND d.reached_load_port IS NOT NULL THEN 'EMPTY : WAITING AT JAIGAD'
                   WHEN NULLIF(TRIM(d.unloading_completed), '') IS NOT NULL
                        AND d.sailed_out_load_port IS NOT NULL THEN 'EMPTY : ON THE WAY TO JAIGAD'
                   WHEN NULLIF(TRIM(d.unloading_completed), '') IS NOT NULL THEN 'EMPTY : WAITING AT DHARAMTAR'
                   ELSE 'NA'
               END AS mbc_status
        FROM mbc_master m
        LEFT JOIN LATERAL (
            SELECT h.* FROM mbc_header h
            WHERE TRIM(h.mbc_name) = TRIM(m.mbc_name)
            ORDER BY h.id DESC LIMIT 1
        ) h ON TRUE
        LEFT JOIN mbc_load_port_lines l ON l.mbc_id = h.id
        LEFT JOIN mbc_discharge_port_lines d ON d.mbc_id = h.id
        WHERE UPPER(TRIM(COALESCE(m.mbc_owner_name, ''))) IN ('JSW INFRA', 'JSW SHIPPING')
    """,

    # --- Tide table. HW/LW = local peak vs neighbours (dashboard alternates
    #     from the first of the next 6 readings, which is equivalent). -------
    'vw_port_overview_tide': f"""
        WITH t AS (
            SELECT id, TRIM(tide_datetime)::timestamp AS tide_at, tide_meters
            FROM tide_master
            WHERE TRIM(tide_datetime) ~ {_ISO}
        )
        SELECT id, tide_at, tide_meters,
               CASE WHEN COALESCE(tide_meters > LEAD(tide_meters) OVER w,
                                  tide_meters > LAG(tide_meters) OVER w)
                    THEN 'HW' ELSE 'LW' END AS tide_type
        FROM t
        WINDOW w AS (ORDER BY tide_at)
    """,

    # --- Cached WeatherAPI payload (current conditions + raw forecast JSON) --
    'vw_port_overview_weather': """
        SELECT w.cache_key, w.fetched_at,
               (w.data->'current'->>'temp_c')::numeric       AS temp_c,
               (w.data->'current'->>'feelslike_c')::numeric  AS feelslike_c,
               w.data->'current'->'condition'->>'text'       AS condition_text,
               w.data->'current'->'condition'->>'icon'       AS condition_icon,
               (w.data->'current'->>'wind_kph')::numeric     AS wind_kph,
               w.data->'current'->>'wind_dir'                AS wind_dir,
               (w.data->'current'->>'gust_kph')::numeric     AS gust_kph,
               (w.data->'current'->>'humidity')::numeric     AS humidity,
               (w.data->'current'->>'pressure_mb')::numeric  AS pressure_mb,
               (w.data->'current'->>'vis_km')::numeric       AS vis_km,
               w.data->'current'->>'last_updated'            AS last_updated,
               w.data->'forecast'->'forecastday'             AS forecast_days
        FROM weather_cache w
    """,
}


def upgrade() -> None:
    # Dict order matters: berth_occupancy reads live_assets + last_lueu_completion.
    for name, sql in VIEWS.items():
        op.execute(f"CREATE OR REPLACE VIEW {name} AS {sql}")


def downgrade() -> None:
    for name in reversed(list(VIEWS)):
        op.execute(f"DROP VIEW IF EXISTS {name}")
