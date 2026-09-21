"""Marine: create Berthing Officer master table

Revision ID: bomx8y3p2v9q
Revises: tugbgdgytr4h
Create Date: 2026-09-21
"""

from typing import Sequence, Union
from alembic import op

revision: str = 'bomx8y3p2v9q'
down_revision: Union[str, None] = 'tugbgdgytr4h'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None

def upgrade() -> None:
    op.execute("""
        CREATE TABLE IF NOT EXISTS berthing_officer_master (
            id SERIAL PRIMARY KEY,
            name VARCHAR(150) NOT NULL,
            company_name VARCHAR(150),
            contact_no VARCHAR(20)
        )
    """)

def downgrade() -> None:
    op.execute("DROP TABLE IF EXISTS berthing_officer_master")
