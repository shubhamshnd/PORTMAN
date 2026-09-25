"""RP02 Backdated Cargo Handling: rp02_cargo_handling_backdated table

Revision ID: rp02cargobd001
Revises: eaf97603d5d9
Create Date: 2026-09-25
"""
from typing import Sequence, Union
from alembic import op

revision: str = 'rp02cargobd001'
down_revision: Union[str, None] = 'eaf97603d5d9'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.execute("""
        CREATE TABLE IF NOT EXISTS rp02_cargo_handling_backdated (
            id                  SERIAL PRIMARY KEY,
            vessel_name         VARCHAR(255) NOT NULL,
            vessel_type         VARCHAR(50),
            material_po         VARCHAR(100),
            cargo_type          VARCHAR(100),
            cargo_name          VARCHAR(255),
            bl_qty_mt           VARCHAR(100),
            actual_discharge    VARCHAR(100),
            load_port           VARCHAR(100),
            discharge_commenced VARCHAR(100),
            discharge_completed VARCHAR(100),
            consignee           VARCHAR(255),
            flag                VARCHAR(100),
            invoice_number      VARCHAR(100),
            status              VARCHAR(50),
            uploaded_by         TEXT,
            uploaded_at         TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE INDEX IF NOT EXISTS ix_rp02_cargo_bd_vessel ON rp02_cargo_handling_backdated (vessel_name);
        CREATE INDEX IF NOT EXISTS ix_rp02_cargo_bd_commenced ON rp02_cargo_handling_backdated (discharge_commenced);
        CREATE INDEX IF NOT EXISTS ix_rp02_cargo_bd_completed ON rp02_cargo_handling_backdated (discharge_completed);
    """)


def downgrade() -> None:
    op.execute("DROP TABLE IF EXISTS rp02_cargo_handling_backdated")


