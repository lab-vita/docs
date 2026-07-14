"""
/app — главная точка входа приложения из Bitrix24.
Без параметра process= → меню процессов.
С параметром process=<key> → форма конкретного процесса.
"""
import logging
from pathlib import Path

from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse
from sqlalchemy import select

from app.config import settings
from app.database import SessionLocal
from app.models.process import Process
from app.services.bitrix import BitrixClient

logger = logging.getLogger(__name__)
router = APIRouter()

_TEMPLATES_DIR = Path(__file__).parent.parent / "templates"


async def _resolve_user(request: Request, form_data: dict) -> tuple[str, str]:
    """Получает user_id и auth_id из Bitrix24. Возвращает (user_id, auth_id)."""
    auth_id = (
        request.query_params.get("AUTH_ID")
        or request.query_params.get("auth_id")
        or form_data.get("AUTH_ID")
        or form_data.get("auth_id")
        or ""
    )
    domain = (
        request.query_params.get("DOMAIN")
        or form_data.get("DOMAIN")
        or ""
    )
    user_id = ""
    if auth_id and domain:
        try:
            async with BitrixClient() as bitrix:
                profile = await bitrix.get_user_profile(auth_id)
            user_id = str(profile.get("ID", ""))
        except Exception as e:
            logger.error(f"Ошибка получения профиля: {e}")
    return user_id, auth_id


def _inject_user(html: str, user_id: str, auth_id: str) -> str:
    """Подставляет user_id и auth_id в HTML-шаблон."""
    html = html.replace(
        "const userId = urlParams.get('user_id') || '';",
        f"const userId = urlParams.get('user_id') || '{user_id}';",
    )
    html = html.replace(
        "const authToken = urlParams.get('auth_token') || '';",
        f"const authToken = urlParams.get('auth_token') || '{auth_id}';",
    )
    return html


@router.api_route("/app", methods=["GET", "POST"], response_class=HTMLResponse)
async def app_handler(request: Request):
    """
    Главный обработчик приложения.
    process= не задан → каталог процессов.
    process=<key>     → универсальная форма для этого процесса.
    """
    query = dict(request.query_params)
    form_data = getattr(request.state, "form_data", None) or {}
    if not form_data and request.method == "POST":
        try:
            form_data = dict(await request.form())
        except Exception:
            form_data = {}

    user_id, auth_id = await _resolve_user(request, form_data)
    process_key = query.get("process", "")

    if process_key:
        # Проверяем что процесс существует и активен
        async with SessionLocal() as db:
            result = await db.execute(
                select(Process).where(Process.key == process_key, Process.is_active == True)
            )
            process = result.scalar_one_or_none()

        if not process:
            return HTMLResponse(
                f"<p style='font-family:sans-serif;padding:20px'>Процесс не найден: {process_key}</p>",
                status_code=404,
            )

        # Отдаём универсальную форму
        html = (_TEMPLATES_DIR / "form.html").read_text(encoding="utf-8")
        html = html.replace("__PROCESS_KEY__", process_key)
        html = html.replace("__PROCESS_TITLE__", process.title)
        html = _inject_user(html, user_id, auth_id)
        return HTMLResponse(html)

    # Меню процессов
    async with SessionLocal() as db:
        result = await db.execute(
            select(Process)
            .where(Process.is_active == True)
            .order_by(Process.sort_order)
        )
        processes = result.scalars().all()

    import json
    processes_json = json.dumps(
        {p.key: {"title": p.title, "description": p.description, "icon": p.icon}
         for p in processes},
        ensure_ascii=False,
    )

    html = (_TEMPLATES_DIR / "menu.html").read_text(encoding="utf-8")
    html = html.replace("__APP_TITLE__", settings.app_title)
    html = html.replace("__PROCESSES_JSON__", processes_json)
    html = _inject_user(html, user_id, auth_id)
    return HTMLResponse(html)
