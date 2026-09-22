"""create_conveyor_operator_master

Revision ID: f7529a9c6b51
Revises: d997647862fb
Create Date: 2026-09-22 11:33:55.584398

"""
from typing import Sequence, Union

from alembic import op


# revision identifiers, used by Alembic.
revision: str = 'f7529a9c6b51'
down_revision: Union[str, None] = 'd997647862fb'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.execute("""
        CREATE TABLE IF NOT EXISTS conveyor_operator_master (
            id SERIAL PRIMARY KEY,
            name VARCHAR(150) NOT NULL,
            contact_no VARCHAR(20)
        )
    """)


def downgrade() -> None:
    op.execute("DROP TABLE IF EXISTS conveyor_operator_master")
