"""
Заполняет БД начальными данными при первом запуске.
Запускается из entrypoint.sh после alembic upgrade head.
"""
import asyncio
import os
from sqlalchemy import select
from app.database import SessionLocal
from app.models.process import Process, FormField, DocumentSignature
from datetime import datetime


CASH_PROCESS = {
    "key": "cash",
    "title": "Выдача наличных",
    "description": "Заявка на выдачу денежных средств под отчёт",
    "icon": "💵",
    "iblock_id": os.getenv("CASH_REQUEST_IBLOCK_ID", ""),
    "bp_template_id": os.getenv("CASH_REQUEST_BP_ID", ""),
    "template_file_id": "",   # заполнить через /admin
    "output_folder_id": "",   # заполнить через /admin
    "iblock_fields": {
        "SUMMA": os.getenv("FIELD_SUMMA", "PROPERTY_516"),
        "NAZNACHENIE": os.getenv("FIELD_NAZNACHENIE", "PROPERTY_518"),
    },
    "fields": [
        {"name": "TITLE", "label": "Название заявки",
         "field_type": "text", "required": True,
         "placeholder": "Например: Ремонт забора", "sort_order": 0},
        {"name": "POSITIONS", "label": "Позиции",
         "field_type": "positions", "required": True,
         "placeholder": "", "sort_order": 1},
    ],
    "signatures": [
        {"placeholder": "SIGN_EMPLOYEE", "label": "Подпись сотрудника",
         "stage": "initial", "source": "employee_profile"},
        {"placeholder": "SIGN_APPROVER", "label": "Подпись руководителя",
         "stage": "approval", "source": "bp_variable"},
    ],
}


async def seed():
    async with SessionLocal() as db:
        # Проверяем — уже есть или нет
        result = await db.execute(
            select(Process).where(Process.key == "cash")
        )
        existing = result.scalar_one_or_none()
        if existing:
            print("Seed: процесс 'cash' уже существует, пропускаем")
            return

        process = Process(
            key=CASH_PROCESS["key"],
            title=CASH_PROCESS["title"],
            description=CASH_PROCESS["description"],
            icon=CASH_PROCESS["icon"],
            iblock_id=CASH_PROCESS["iblock_id"],
            bp_template_id=CASH_PROCESS["bp_template_id"],
            template_file_id=CASH_PROCESS["template_file_id"],
            output_folder_id=CASH_PROCESS["output_folder_id"],
            iblock_fields=CASH_PROCESS["iblock_fields"],
            is_active=True,
            sort_order=0,
            created_at=datetime.utcnow(),
        )

        for f in CASH_PROCESS["fields"]:
            process.fields.append(FormField(**f))

        for s in CASH_PROCESS["signatures"]:
            process.signatures.append(DocumentSignature(**s))

        db.add(process)
        await db.commit()
        print("Seed: процесс 'cash' создан")


if __name__ == "__main__":
    asyncio.run(seed())
