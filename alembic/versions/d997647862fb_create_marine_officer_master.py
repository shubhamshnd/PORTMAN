"""Marine: create marine officer master table

Revision ID: d997647862fb
Revises: bomx8y3p2v9q
Create Date: 2026-09-22
"""
from typing import Sequence, Union
from alembic import op


# revision identifiers, used by Alembic.
revision: str = 'd997647862fb'
down_revision: Union[str, None] = 'bomx8y3p2v9q'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.execute("""
        CREATE TABLE IF NOT EXISTS marine_officer_master (
            id SERIAL PRIMARY KEY,
            name VARCHAR(150) NOT NULL,
            contact_no VARCHAR(20)
        )
    """)


def downgrade() -> None:
    op.execute("DROP TABLE IF EXISTS marine_officer_master")
