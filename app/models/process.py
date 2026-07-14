from datetime import datetime
from typing import List, Optional
from sqlalchemy import Integer, String, Boolean, DateTime, ForeignKey, Text
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy.orm import Mapped, mapped_column, relationship
from app.database import Base


class Process(Base):
    """Бизнес-процесс — единица каталога заявок."""
    __tablename__ = "processes"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    key: Mapped[str] = mapped_column(String(64), unique=True, nullable=False)
    title: Mapped[str] = mapped_column(String(255), nullable=False)
    description: Mapped[str] = mapped_column(Text, default="")
    icon: Mapped[str] = mapped_column(String(16), default="📄")

    # Bitrix24 IDs
    iblock_id: Mapped[str] = mapped_column(String(64), default="")
    bp_template_id: Mapped[str] = mapped_column(String(64), default="")
    template_file_id: Mapped[str] = mapped_column(String(64), default="")
    output_folder_id: Mapped[str] = mapped_column(String(64), default="")

    # Поля инфоблока для хранения данных заявки (KEY|BITRIX_FIELD)
    # Пример: {"SUMMA": "PROPERTY_516", "NAZNACHENIE": "PROPERTY_518"}
    iblock_fields: Mapped[Optional[dict]] = mapped_column(JSONB, default=dict)

    is_active: Mapped[bool] = mapped_column(Boolean, default=True)
    sort_order: Mapped[int] = mapped_column(Integer, default=0)
    created_at: Mapped[datetime] = mapped_column(
        DateTime, default=datetime.utcnow
    )

    fields: Mapped[List["FormField"]] = relationship(
        "FormField",
        back_populates="process",
        cascade="all, delete-orphan",
        order_by="FormField.sort_order",
    )
    signatures: Mapped[List["DocumentSignature"]] = relationship(
        "DocumentSignature",
        back_populates="process",
        cascade="all, delete-orphan",
    )


class FormField(Base):
    """Поле формы заявки конкретного процесса."""
    __tablename__ = "form_fields"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    process_id: Mapped[int] = mapped_column(
        Integer, ForeignKey("processes.id", ondelete="CASCADE"), nullable=False
    )
    name: Mapped[str] = mapped_column(String(128), nullable=False)
    label: Mapped[str] = mapped_column(String(255), nullable=False)
    # text | number | date | date_range | select | textarea | positions
    field_type: Mapped[str] = mapped_column(String(32), default="text")
    required: Mapped[bool] = mapped_column(Boolean, default=True)
    # Варианты для type=select: [{"value": "x", "label": "Y"}]
    options: Mapped[Optional[list]] = mapped_column(JSONB, default=list)
    placeholder: Mapped[str] = mapped_column(String(255), default="")
    sort_order: Mapped[int] = mapped_column(Integer, default=0)

    process: Mapped["Process"] = relationship("Process", back_populates="fields")


class DocumentSignature(Base):
    """Настройка подписи в шаблоне документа."""
    __tablename__ = "document_signatures"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    process_id: Mapped[int] = mapped_column(
        Integer, ForeignKey("processes.id", ondelete="CASCADE"), nullable=False
    )
    # Placeholder в Alt Text шаблона: SIGN_EMPLOYEE, SIGN_APPROVER
    placeholder: Mapped[str] = mapped_column(String(128), nullable=False)
    label: Mapped[str] = mapped_column(String(255), default="")
    # initial — при первой генерации, approval — после согласования
    stage: Mapped[str] = mapped_column(String(32), default="initial")
    # employee_profile — из поля UF_ профиля пользователя
    # bp_variable — file_id приходит из переменной БП
    source: Mapped[str] = mapped_column(String(32), default="employee_profile")

    process: Mapped["Process"] = relationship(
        "Process", back_populates="signatures"
    )
