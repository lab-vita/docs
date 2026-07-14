"""
Middleware безопасности для Bitrix24-приложения.

BitrixSourceMiddleware  — гарантирует что запрос пришёл из Bitrix24 iframe.
AdminAccessMiddleware   — проверяет что пользователь есть в ADMIN_USER_IDS.
"""
import logging
from pathlib import Path
from typing import Optional

from fastapi import Request
from fastapi.responses import HTMLResponse
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import Response

from app.config import settings

logger = logging.getLogger(__name__)

# Маршруты, которые не требуют проверки источника
_PUBLIC_PATHS = {"/health", "/install"}

_BLOCKED_HTML = """<!DOCTYPE html>
<html lang="ru"><head><meta charset="UTF-8">
<style>
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;
display:flex;align-items:center;justify-content:center;height:100vh;margin:0;background:#f5f5f5;}
.card{background:#fff;border-radius:8px;padding:32px 40px;
box-shadow:0 2px 8px rgba(0,0,0,.1);text-align:center;max-width:420px;}
h2{color:#1a1a1a;margin-bottom:8px;font-size:18px;}
p{color:#666;font-size:13px;line-height:1.5;}
</style></head><body>
<div class="card">
  <h2>Доступ ограничен</h2>
  <p>Это приложение доступно только внутри портала Bitrix24.<br>
  Откройте его через левое меню.</p>
</div>
</body></html>"""

_ACCESS_DENIED_HTML = """<!DOCTYPE html>
<html lang="ru"><head><meta charset="UTF-8">
<style>
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;
display:flex;align-items:center;justify-content:center;height:100vh;margin:0;background:#f5f5f5;}
.card{background:#fff;border-radius:8px;padding:32px 40px;
box-shadow:0 2px 8px rgba(0,0,0,.1);text-align:center;max-width:420px;}
h2{color:#1a1a1a;margin-bottom:8px;font-size:18px;}
p{color:#666;font-size:13px;line-height:1.5;}
</style></head><body>
<div class="card">
  <h2>Нет доступа</h2>
  <p>У вас недостаточно прав для просмотра этого раздела.<br>
  Обратитесь к администратору.</p>
</div>
</body></html>"""


def _extract_auth_id(request: Request, form_data: Optional[dict] = None) -> Optional[str]:
    """Ищет AUTH_ID в query params и form data."""
    auth_id = request.query_params.get("AUTH_ID") or request.query_params.get("auth_id")
    if auth_id:
        return auth_id
    if form_data:
        return form_data.get("AUTH_ID") or form_data.get("auth_id")
    return None


def _has_bitrix_context(request: Request, form_data: Optional[dict] = None) -> bool:
    """
    Проверяет наличие признаков Bitrix24-контекста в запросе.
    Bitrix24 всегда передаёт PLACEMENT или AUTH_ID при открытии iframe.
    """
    qp = request.query_params
    # GET-запрос с PLACEMENT или AUTH_ID → из Bitrix24
    if qp.get("PLACEMENT") or qp.get("AUTH_ID") or qp.get("auth_id"):
        return True
    # POST-запрос: ищем в form data
    if form_data:
        if form_data.get("PLACEMENT") or form_data.get("AUTH_ID") or form_data.get("auth_id"):
            return True
        # Bitrix24 activity handler передаёт event_token
        if form_data.get("event_token"):
            return True
    return False


class BitrixSourceMiddleware(BaseHTTPMiddleware):
    """
    Блокирует прямые запросы извне — без Bitrix24-контекста.
    Исключения: /health, /install (нужны для деплоя и OAuth).
    """

    async def dispatch(self, request: Request, call_next) -> Response:
        path = request.url.path

        # Публичные маршруты пропускаем
        if path in _PUBLIC_PATHS:
            return await call_next(request)

        # Для /admin проверка выполняется в AdminAccessMiddleware
        # Здесь только проверяем контекст
        form_data = None
        if request.method == "POST":
            try:
                form_data = dict(await request.form())
                # Сохраняем в state чтобы роутер мог переиспользовать
                request.state.form_data = form_data
            except Exception:
                pass

        if not _has_bitrix_context(request, form_data):
            logger.warning(f"Заблокирован запрос без Bitrix24-контекста: {path} от {request.client}")
            return HTMLResponse(_BLOCKED_HTML, status_code=403)

        return await call_next(request)


class AdminAccessMiddleware(BaseHTTPMiddleware):
    """
    Проверяет доступ к /admin/*.
    Получает user_id из профиля Bitrix24 и сверяет с ADMIN_USER_IDS.
    """

    async def dispatch(self, request: Request, call_next) -> Response:
        if not request.url.path.startswith("/admin"):
            return await call_next(request)

        if not settings.admin_user_ids:
            logger.warning("ADMIN_USER_IDS не настроен — доступ к /admin закрыт")
            return HTMLResponse(_ACCESS_DENIED_HTML, status_code=403)

        # Собираем AUTH_ID
        form_data = getattr(request.state, "form_data", None)
        if form_data is None and request.method == "POST":
            try:
                form_data = dict(await request.form())
                request.state.form_data = form_data
            except Exception:
                form_data = {}

        auth_id = _extract_auth_id(request, form_data)
        if not auth_id:
            return HTMLResponse(_ACCESS_DENIED_HTML, status_code=403)

        # Получаем профиль пользователя
        try:
            from app.services.bitrix import BitrixClient
            async with BitrixClient() as bitrix:
                profile = await bitrix.get_user_profile(auth_id)
            user_id = int(profile.get("ID", 0))
        except Exception as e:
            logger.error(f"Ошибка получения профиля для /admin: {e}")
            return HTMLResponse(_ACCESS_DENIED_HTML, status_code=403)

        if user_id not in settings.admin_user_ids:
            logger.warning(f"Доступ к /admin запрещён для user_id={user_id}")
            return HTMLResponse(_ACCESS_DENIED_HTML, status_code=403)

        request.state.admin_user_id = user_id
        return await call_next(request)
