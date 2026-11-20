"""add created_by and created_date to mbct01_header

Revision ID: mbct01merge002
Revises: tgu01merge001
Create Date: 2026-11-19

Changes:
- Add created_by to mbct01_header
- Add created_date to mbct01_header
"""

from typing import Sequence, Union
from alembic import op


# revision identifiers, used by Alembic.
revision: str = "mbct01merge002"
down_revision: Union[str, None] = "tgu01merge001"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.execute("""
        ALTER TABLE mbct01_header 
        ADD COLUMN IF NOT EXISTS created_by JSONB,
        ADD COLUMN IF NOT EXISTS created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP;

        DO $$
        BEGIN
            IF EXISTS (
                SELECT 1 FROM information_schema.columns 
                WHERE table_name = 'mbct01_header' 
                AND column_name = 'created_by' 
                AND data_type <> 'jsonb'
            ) THEN
                ALTER TABLE mbct01_header 
                ALTER COLUMN created_by TYPE JSONB USING 
                    CASE 
                        WHEN created_by IS NULL THEN NULL
                        WHEN created_by ~ '^\\{.*\\}$' THEN created_by::jsonb
                        ELSE jsonb_build_object('username', created_by)
                    END;
            END IF;
        END $$;
    """)


def downgrade() -> None:
    op.execute("""
        ALTER TABLE mbct01_header 
        DROP COLUMN IF EXISTS created_by,
        DROP COLUMN IF EXISTS created_date;
    """)
