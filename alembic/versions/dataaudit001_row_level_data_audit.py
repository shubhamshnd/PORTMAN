"""Row-level data audit: every INSERT / UPDATE / DELETE on module tables.

One generic trigger records the change into data_audit, whatever wrote it —
module screens, Admin tools, scripts or a manual SQL fix. The app user comes
from the `app.user` setting that database.get_db() stamps on each connection;
writes without it are recorded as 'db:<role>'.

Binary and secret columns (bytea, password/secret/token/otp/api_key) are
stripped per table when the trigger is attached, so they never reach the log.

A table created later is not audited until a migration runs
    SELECT data_audit_attach('<table>');

Revision ID: dataaudit001
Revises: mbct01merge002
Create Date: 2026-11-20
"""
from alembic import op

revision = 'dataaudit001'
down_revision = 'mbct01merge002'
branch_labels = None
depends_on = None

# Logs, queues, tokens, caches and per-user UI state: either already an audit
# trail, high-churn machine data, or secrets end to end.
EXCLUDED = (
    'alembic_version', 'data_audit', 'approval_log', 'cutover_audit',
    'integration_logs', 'mail_queue', 'sap_outbound_queue', 'sap_inbound_tokens',
    'password_reset_tokens', 'weather_cache', 'saved_filters',
    'saved_pivot_reports', 'user_delay_preferences',
)


def upgrade():
    op.execute("""
        CREATE TABLE IF NOT EXISTS data_audit (
            id          BIGSERIAL PRIMARY KEY,
            table_name  TEXT NOT NULL,
            row_id      TEXT,
            op          CHAR(1) NOT NULL,          -- I / U / D
            old_data    JSONB,                     -- full row before (U, D)
            new_data    JSONB,                     -- full row (I) or changed fields only (U)
            changed_by  TEXT,
            changed_at  TIMESTAMP NOT NULL DEFAULT now()
        )
    """)
    op.execute("CREATE INDEX IF NOT EXISTS ix_data_audit_changed_at ON data_audit (changed_at)")
    op.execute("CREATE INDEX IF NOT EXISTS ix_data_audit_row ON data_audit (table_name, row_id)")

    op.execute("""
        CREATE OR REPLACE FUNCTION data_audit_trg() RETURNS trigger
        LANGUAGE plpgsql AS $$
        DECLARE
            o jsonb; n jsonb; diff jsonb; rid text; k text;
        BEGIN
            IF TG_OP <> 'INSERT' THEN o := to_jsonb(OLD); END IF;
            IF TG_OP <> 'DELETE' THEN n := to_jsonb(NEW); END IF;
            -- TG_ARGV = columns never copied (binaries, secrets), fixed at attach time
            IF TG_NARGS > 0 THEN
                FOREACH k IN ARRAY TG_ARGV LOOP
                    o := o - k;
                    n := n - k;
                END LOOP;
            END IF;
            rid := COALESCE(n->>'id', o->>'id');

            IF TG_OP = 'UPDATE' THEN
                SELECT jsonb_object_agg(e.key, e.value) INTO diff
                FROM jsonb_each(n) e
                WHERE e.value IS DISTINCT FROM o->e.key;
                IF diff IS NULL THEN RETURN NULL; END IF;   -- no-op save
                n := diff;
            END IF;

            INSERT INTO data_audit (table_name, row_id, op, old_data, new_data, changed_by)
            VALUES (TG_TABLE_NAME, rid, left(TG_OP, 1), o, n,
                    COALESCE(NULLIF(current_setting('app.user', true), ''), 'db:' || session_user));
            RETURN NULL;
        END $$
    """)

    op.execute("""
        CREATE OR REPLACE FUNCTION data_audit_attach(tbl text) RETURNS void
        LANGUAGE plpgsql AS $$
        DECLARE strip text;
        BEGIN
            SELECT string_agg(quote_literal(column_name), ', ') INTO strip
            FROM information_schema.columns
            WHERE table_schema = 'public' AND table_name = tbl
              AND (data_type = 'bytea'
                   OR column_name ~* '(password|secret|token|otp|api_key)');
            EXECUTE format('DROP TRIGGER IF EXISTS data_audit ON %I', tbl);
            EXECUTE format('CREATE TRIGGER data_audit AFTER INSERT OR UPDATE OR DELETE ON %I
                            FOR EACH ROW EXECUTE FUNCTION data_audit_trg(%s)', tbl, COALESCE(strip, ''));
        END $$
    """)

    excluded = ', '.join(f"'{t}'" for t in EXCLUDED)
    op.execute(f"""
        SELECT data_audit_attach(table_name)
        FROM information_schema.tables
        WHERE table_schema = 'public' AND table_type = 'BASE TABLE'
          AND table_name NOT IN ({excluded})
    """)


def downgrade():
    op.execute("""
        DO $$
        DECLARE t text;
        BEGIN
            FOR t IN SELECT event_object_table FROM information_schema.triggers
                     WHERE trigger_name = 'data_audit' GROUP BY 1 LOOP
                EXECUTE format('DROP TRIGGER IF EXISTS data_audit ON %I', t);
            END LOOP;
        END $$
    """)
    op.execute("DROP FUNCTION IF EXISTS data_audit_attach(text)")
    op.execute("DROP FUNCTION IF EXISTS data_audit_trg()")
    op.execute("DROP TABLE IF EXISTS data_audit")
