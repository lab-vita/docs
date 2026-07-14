"""initial

Revision ID: 0001
Revises: 
Create Date: 2026-01-01 00:00:00.000000

"""
from typing import Sequence, Union
from alembic import op
import sqlalchemy as sa
from sqlalchemy.dialects import postgresql

revision: str = "0001"
down_revision: Union[str, None] = None
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.create_table(
        "processes",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("key", sa.String(64), nullable=False),
        sa.Column("title", sa.String(255), nullable=False),
        sa.Column("description", sa.Text(), nullable=True, server_default=""),
        sa.Column("icon", sa.String(16), nullable=True, server_default="📄"),
        sa.Column("iblock_id", sa.String(64), nullable=True, server_default=""),
        sa.Column("bp_template_id", sa.String(64), nullable=True, server_default=""),
        sa.Column("template_file_id", sa.String(64), nullable=True, server_default=""),
        sa.Column("output_folder_id", sa.String(64), nullable=True, server_default=""),
        sa.Column("iblock_fields", postgresql.JSONB(), nullable=True),
        sa.Column("is_active", sa.Boolean(), nullable=True, server_default="true"),
        sa.Column("sort_order", sa.Integer(), nullable=True, server_default="0"),
        sa.Column("created_at", sa.DateTime(), nullable=True),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint("key"),
    )

    op.create_table(
        "form_fields",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("process_id", sa.Integer(), nullable=False),
        sa.Column("name", sa.String(128), nullable=False),
        sa.Column("label", sa.String(255), nullable=False),
        sa.Column("field_type", sa.String(32), nullable=True, server_default="text"),
        sa.Column("required", sa.Boolean(), nullable=True, server_default="true"),
        sa.Column("options", postgresql.JSONB(), nullable=True),
        sa.Column("placeholder", sa.String(255), nullable=True, server_default=""),
        sa.Column("sort_order", sa.Integer(), nullable=True, server_default="0"),
        sa.ForeignKeyConstraint(["process_id"], ["processes.id"], ondelete="CASCADE"),
        sa.PrimaryKeyConstraint("id"),
    )

    op.create_table(
        "document_signatures",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("process_id", sa.Integer(), nullable=False),
        sa.Column("placeholder", sa.String(128), nullable=False),
        sa.Column("label", sa.String(255), nullable=True, server_default=""),
        sa.Column("stage", sa.String(32), nullable=True, server_default="initial"),
        sa.Column("source", sa.String(32), nullable=True, server_default="employee_profile"),
        sa.ForeignKeyConstraint(["process_id"], ["processes.id"], ondelete="CASCADE"),
        sa.PrimaryKeyConstraint("id"),
    )


def downgrade() -> None:
    op.drop_table("document_signatures")
    op.drop_table("form_fields")
    op.drop_table("processes")
