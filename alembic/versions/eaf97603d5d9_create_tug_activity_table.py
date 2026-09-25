"""create_tug_activity_master_table

Revision ID: eaf97603d5d9
Revises: f7529a9c6b51
Create Date: 2026-09-25 10:21:12.543694

"""
from typing import Sequence, Union

from alembic import op


# revision identifiers, used by Alembic.
revision: str = 'eaf97603d5d9'
down_revision: Union[str, None] = 'f7529a9c6b51'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.execute("""
        CREATE TABLE IF NOT EXISTS activity_master (
            id SERIAL PRIMARY KEY,
            name VARCHAR(150) NOT NULL
        )
    """)


def downgrade() -> None:
    op.execute("DROP TABLE IF EXISTS activity_master")
