"""RP03 Arrival Planner: ETA log and retrain history

Two tables, owned by RP03 and written only by eta_service.py. rp03_eta_log keeps
the first ETA given for each trip at each milestone, so it can be compared with
the arrival that followed. rp03_eta_run records every retrain and whether the
new model was promoted. Dropping both only loses that history.

trip_id has no foreign key for the same reason as rp03_arrival_plan.

Revision ID: rp03eta001
Revises: dataaudit001
Create Date: 2026-11-24
"""
from typing import Sequence, Union

from alembic import op

revision: str = 'rp03eta001'
down_revision: Union[str, None] = 'dataaudit001'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.execute('''
        CREATE TABLE IF NOT EXISTS rp03_eta_log (
            id                  SERIAL PRIMARY KEY,
            trip_kind           TEXT NOT NULL CHECK (trip_kind IN ('BARGE', 'MBC')),
            trip_id             INTEGER NOT NULL,
            stage               TEXT NOT NULL,
            stage_at            TIMESTAMP NOT NULL,
            predicted_at        TIMESTAMP NOT NULL,
            eta                 TIMESTAMP NOT NULL,
            eta_p10             TIMESTAMP,
            eta_p90             TIMESTAMP,
            median_eta          TIMESTAMP,
            model_trained_until DOUBLE PRECISION,
            UNIQUE (trip_kind, trip_id, stage)
        )
    ''')
    op.execute('CREATE INDEX IF NOT EXISTS ix_rp03_eta_log_predicted_at ON rp03_eta_log (predicted_at)')
    op.execute('''
        CREATE TABLE IF NOT EXISTS rp03_eta_run (
            id       SERIAL PRIMARY KEY,
            ran_at   TIMESTAMP NOT NULL,
            promoted BOOLEAN NOT NULL,
            reason   TEXT,
            detail   JSONB
        )
    ''')


def downgrade() -> None:
    op.execute('DROP TABLE IF EXISTS rp03_eta_run')
    op.execute('DROP TABLE IF EXISTS rp03_eta_log')
