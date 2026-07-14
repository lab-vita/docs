"""
/api/process/{key} — отдаёт конфиг процесса в JSON для построения формы.
"""
from fastapi import APIRouter, HTTPException
from sqlalchemy import select
from sqlalchemy.orm import selectinload

from app.database import SessionLocal
from app.models.process import Process

router = APIRouter(prefix="/api")


@router.get("/process/{key}")
async def get_process_config(key: str):
    """
    Возвращает конфигурацию формы для указанного процесса.
    Используется универсальным form.html для построения формы на клиенте.
    """
    async with SessionLocal() as db:
        result = await db.execute(
            select(Process)
            .options(selectinload(Process.fields))
            .where(Process.key == key, Process.is_active == True)
        )
        process = result.scalar_one_or_none()

    if not process:
        raise HTTPException(status_code=404, detail=f"Процесс '{key}' не найден")

    return {
        "key": process.key,
        "title": process.title,
        "icon": process.icon,
        "fields": [
            {
                "name": f.name,
                "label": f.label,
                "type": f.field_type,
                "required": f.required,
                "options": f.options or [],
                "placeholder": f.placeholder or "",
            }
            for f in sorted(process.fields, key=lambda x: x.sort_order)
        ],
    }
