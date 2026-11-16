"""create all BGT01 tables

Revision ID: bgt01merge001
Revises: mbct01merge001
Create Date: 2026-11-16

Creates:
- bgt01_header
- bgt01_movement_activity
"""

from typing import Sequence, Union
from alembic import op

revision: str = "bgt01merge001"
down_revision: Union[str, None] = "mbct01merge001"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    # 1. BGT01 HEADER
    op.execute("""
        CREATE TABLE IF NOT EXISTS bgt01_header (
            id SERIAL PRIMARY KEY,
            vcn_id INTEGER REFERENCES vcn_header(id) ON DELETE SET NULL,
            vcn_doc_num TEXT,
            vessel_name TEXT,
            barge_name TEXT,
            trip_number TEXT,
            doc_date TEXT,
            cargo_name TEXT,
            doc_status TEXT DEFAULT 'Draft',
            created_by TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # 2. BGT01 MOVEMENT ACTIVITY (One row per activity, with tugs_data JSONB)
    op.execute("""
        CREATE TABLE IF NOT EXISTS bgt01_movement_activity (
            id SERIAL PRIMARY KEY,
            header_id INTEGER NOT NULL REFERENCES bgt01_header(id) ON DELETE CASCADE,
            movement_date TEXT,
            activity TEXT,
            from_berth TEXT,
            to_berth TEXT,
            marine_officer TEXT,
            remark TEXT,
            tugs_data JSONB DEFAULT '[]'::jsonb,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)


def downgrade() -> None:
    op.execute("""DROP TABLE IF EXISTS bgt01_movement_activity""")
    op.execute("""DROP TABLE IF EXISTS bgt01_header""")

