"""create TGU01 daily log table

Revision ID: tgu01merge001
Revises: bgt01merge001
Create Date: 2026-11-16

Creates:
- tgu01_daily_log
"""

from typing import Sequence, Union
from alembic import op

revision: str = "tgu01merge001"
down_revision: Union[str, None] = "bgt01merge001"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.execute("""
        CREATE TABLE IF NOT EXISTS tgu01_daily_log (
            id SERIAL PRIMARY KEY,
            tug_name TEXT NOT NULL,
            entry_date TEXT NOT NULL,

            -- [1] HSD FUEL SECTION (LTR)
            hsd_opening_rob TEXT DEFAULT '0',
            hsd_received TEXT DEFAULT '0',
            hsd_me_hours TEXT DEFAULT '0',
            hsd_dg_hours TEXT DEFAULT '0',
            hsd_me_consumption TEXT DEFAULT '0',
            hsd_dg_consumption TEXT DEFAULT '0',
            hsd_total_hours TEXT DEFAULT '0',
            hsd_total_consumption TEXT DEFAULT '0',
            hsd_closing_rob TEXT DEFAULT '0',

            -- [2] FRESH WATER SECTION (MT)
            fw_opening_rob TEXT DEFAULT '0',
            fw_received TEXT DEFAULT '0',
            fw_consumption TEXT DEFAULT '0',
            fw_closing_rob TEXT DEFAULT '0',

            -- [3] ENGINE METER READINGS (HRS)
            port_initial TEXT DEFAULT '0',
            port_final TEXT DEFAULT '0',
            port_diff TEXT DEFAULT '0',
            stbd_initial TEXT DEFAULT '0',
            stbd_final TEXT DEFAULT '0',
            stbd_diff TEXT DEFAULT '0',

            -- [4] SHORE POWER & AUDIT
            shore_power_hours TEXT DEFAULT '0',
            remarks TEXT,
            created_by TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE (tug_name, entry_date)
        )
    """)
    op.execute("""
        CREATE INDEX IF NOT EXISTS idx_tgu01_tug_date ON tgu01_daily_log(tug_name, entry_date DESC)
    """)


def downgrade() -> None:
    op.execute("DROP INDEX IF EXISTS idx_tgu01_tug_date")
    op.execute("DROP TABLE IF EXISTS tgu01_daily_log")

