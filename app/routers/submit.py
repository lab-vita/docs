"""
/submit — универсальный обработчик отправки заявки.
Работает для любого процесса из каталога.
"""
import logging
from datetime import datetime
from typing import Any, Dict

from fastapi import APIRouter
from pydantic import BaseModel
from sqlalchemy import select
from sqlalchemy.orm import selectinload

from app.config import settings
from app.database import SessionLocal
from app.models.process import Process
from app.services.bitrix import BitrixClient

logger = logging.getLogger(__name__)
router = APIRouter()


class SubmitRequest(BaseModel):
    process_key: str
    user_auth_id: str          # AUTH_ID пользователя — токен для создания от его имени
    fields: Dict[str, Any]     # значения полей формы


@router.post("/submit")
async def submit(req: SubmitRequest):
    """
    Создаёт элемент в инфоблоке Bitrix24 и запускает БП.
    Подпись сотрудника подтягивается автоматически из профиля пользователя.
    """
    try:
        async with SessionLocal() as db:
            result = await db.execute(
                select(Process)
                .options(selectinload(Process.fields), selectinload(Process.signatures))
                .where(Process.key == req.process_key, Process.is_active == True)
            )
            process = result.scalar_one_or_none()

        if not process:
            return {"success": False, "error": f"Процесс '{req.process_key}' не найден"}

        async with BitrixClient() as bitrix:
            # Получаем профиль пользователя для подтягивания подписи
            profile = await bitrix.get_user_profile(req.user_auth_id)
            user_id = str(profile.get("ID", ""))
            sign_file_id = profile.get(settings.sign_field, "")

            logger.info(f"Submit: user_id={user_id}, process={req.process_key}")

            # Собираем поля для инфоблока согласно маппингу процесса
            iblock_fields = _build_iblock_fields(process, req.fields)

            element_code = f"{req.process_key}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
            element_id = await bitrix.create_list_element(
                iblock_id=process.iblock_id,
                fields=iblock_fields,
                element_code=element_code,
                auth=req.user_auth_id,
            )

            if not element_id:
                return {"success": False, "error": "Не удалось создать заявку в Bitrix24"}

            logger.info(f"Создан элемент ID={element_id}")

            # Запускаем БП
            await bitrix.start_workflow(
                template_id=process.bp_template_id,
                document_id=["lists", "BizprocDocument", str(element_id)],
                auth=req.user_auth_id,
            )

        return {"success": True, "element_id": element_id}

    except Exception as e:
        logger.error(f"Ошибка submit [{req.process_key}]: {e}", exc_info=True)
        return {"success": False, "error": str(e)}


def _build_iblock_fields(process: Process, form_fields: dict) -> dict:
    """
    Формирует словарь полей для lists.element.add.
    Использует iblock_fields (маппинг form_key → PROPERTY_XXX) из конфига процесса.
    Поле POSITIONS обрабатывается особо — разбивается на сумму и строки позиций.
    """
    mapping: dict = process.iblock_fields or {}
    result = {}

    # Название заявки всегда в NAME
    if "TITLE" in form_fields:
        result["NAME"] = form_fields["TITLE"]

    positions = form_fields.get("POSITIONS", [])

    for form_key, bitrix_field in mapping.items():
        if form_key == "SUMMA" and positions:
            total = sum(float(p.get("amount", 0)) for p in positions)
            result[bitrix_field] = f"{int(total)}|RUB"

        elif form_key == "NAZNACHENIE" and positions:
            lines = "\n".join(
                f"{p.get('name', '')}|{int(p.get('amount', 0))}"
                for p in positions
            )
            result[bitrix_field] = lines

        elif form_key in form_fields:
            result[bitrix_field] = form_fields[form_key]

    return result
