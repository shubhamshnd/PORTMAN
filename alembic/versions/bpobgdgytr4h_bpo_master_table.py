"""Marine: create BPO master table

Revision ID: bpobgdgytr4h

Revises: rp03plan001

Create Date: 2026-09-21

"""

from typing import Sequence, Union

from alembic import op


revision: str = 'bpobgdgytr4h'

down_revision: Union[str, None] = 'rp03plan001'

branch_labels: Union[str, Sequence[str], None] = None

depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:

    op.execute("""
        CREATE TABLE IF NOT EXISTS bpo_master (
            id SERIAL PRIMARY KEY,
            name VARCHAR(150) NOT NULL,
            contact_no VARCHAR(20)
        )
    """)


def downgrade() -> None:

    op.execute("DROP TABLE IF EXISTS bpo_master")
