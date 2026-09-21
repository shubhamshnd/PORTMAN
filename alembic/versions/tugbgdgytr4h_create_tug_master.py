"""Marine: create Tug master table

Revision ID: tugbgdgytr4h
Revises: bpobgdgytr4h
Create Date: 2026-09-21
"""

from typing import Sequence, Union
from alembic import op

revision: str = 'tugbgdgytr4h'
down_revision: Union[str, None] = 'bpobgdgytr4h'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None

def upgrade() -> None:
    op.execute("""
        CREATE TABLE IF NOT EXISTS tug_master (
            id SERIAL PRIMARY KEY,
            tug_name VARCHAR(150) NOT NULL,
            company_name VARCHAR(150),
            tug_capacity VARCHAR(100),
            bhp VARCHAR(100),
            contact_no VARCHAR(20)
        )
    """)

def downgrade() -> None:
    op.execute("DROP TABLE IF EXISTS tug_master")
