"""
/activity-handler — обработчик кастомной активити Bitrix24.
Вызывается из дизайнера БП при достижении шага "Генерация документа".
"""
import json
import logging

from fastapi import APIRouter, Request

from app.services.bitrix import BitrixClient, _load_tokens, _save_tokens
from app.services.document import (
    generate_document,
    parse_variables,
    parse_signatures,
    format_request_goal,
)

logger = logging.getLogger(__name__)
router = APIRouter()


def _extract_props(data: dict) -> dict:
    """
    Парсит свойства активити из form-encoded формата.
    properties[key] и properties[key][N] → словарь {key: value | [values]}
    """
    props = {}
    for key, value in data.items():
        if not key.startswith("properties["):
            continue
        rest = key[11:]
        if "][" in rest:
            base_key = rest[: rest.index("][")]
            if base_key not in props:
                props[base_key] = []
            if isinstance(props[base_key], list):
                props[base_key].append(value)
        elif rest.endswith("]"):
            props[rest[:-1]] = value
    return props


@router.post("/activity-handler")
async def activity_handler(request: Request):
    """
    Обработчик активити — генерирует .docx документ из шаблона
    и возвращает ID готового файла обратно в бизнес-процесс.
    """
    body = await request.body()
    logger.info(f"Вызов активити: {body[:300]}")

    # Пробуем form data, потом JSON
    form_data = getattr(request.state, "form_data", None)
    if form_data is None:
        try:
            form_data = dict(await request.form())
        except Exception:
            form_data = {}

    if not form_data:
        try:
            form_data = json.loads(body)
        except Exception:
            form_data = {}

    props = _extract_props(form_data)

    event_token = form_data.get("event_token")
    template_id = props.get("template_id", "")
    folder_id = props.get("folder_id", "")
    filename = props.get("filename", "document.docx")
    source_file_id = props.get("source_file_id", "")
    doc_data = parse_variables(props.get("variables", []))
    signatures = parse_signatures(props.get("signatures", []))

    logger.info(f"Переменные: {doc_data}")
    logger.info(f"Подписи: {signatures}")

    # Обновляем токен из данных запроса — он всегда свежий
    request_access_token = form_data.get("auth[access_token]")
    if request_access_token:
        tokens = _load_tokens()
        tokens["access_token"] = request_access_token
        tokens["refresh_token"] = form_data.get("auth[refresh_token]", tokens.get("refresh_token", ""))
        tokens["client_endpoint"] = form_data.get("auth[client_endpoint]", tokens.get("client_endpoint", ""))
        _save_tokens(tokens)

    # Форматируем REQUEST_GOAL если есть
    if "REQUEST_GOAL" in doc_data:
        doc_data["REQUEST_GOAL"] = format_request_goal(doc_data["REQUEST_GOAL"])

    async with BitrixClient() as bitrix:
        file_id = await generate_document(
            bitrix=bitrix,
            template_id=template_id,
            folder_id=folder_id,
            filename=filename,
            data=doc_data,
            signatures=signatures,
            source_file_id=source_file_id,
        )

        # Отправляем результат обратно в БП
        await bitrix.send_bp_event(
            event_token=event_token,
            return_values={"file_id": file_id},
        )

    logger.info(f"Активити завершено, file_id={file_id}")
    return {"status": "ok", "file_id": file_id}
