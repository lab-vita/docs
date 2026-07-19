"""
/install — точка входа при установке приложения Bitrix24.
Сохраняет OAuth-токены, регистрирует активити и левое меню.
При открытии не из страницы настроек — показывает меню процессов.
"""
import json
import logging
from pathlib import Path

from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse

from app.config import settings
from app.services.bitrix import BitrixClient, _save_tokens

logger = logging.getLogger(__name__)
router = APIRouter()


def _render_installed_page() -> str:
    return f"""<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>body{{font-family:sans-serif;display:flex;align-items:center;justify-content:center;
height:100vh;margin:0;background:#f5f5f5;}}
.card{{background:#fff;border-radius:8px;padding:32px 40px;
box-shadow:0 2px 8px rgba(0,0,0,.1);text-align:center;}}
h2{{color:#1a1a1a;margin-bottom:8px;}} p{{color:#666;font-size:14px;}}
</style></head><body>
<div class="card"><h2>Приложение установлено</h2>
<p>Откройте «{settings.app_title}» в левом меню Bitrix24.</p></div>
</body></html>"""


@router.api_route("/install", methods=["GET", "POST"])
async def install(request: Request):
    """
    Вызывается Bitrix24 при установке/переустановке приложения.
    Сохраняет токены и регистрирует активити + пункт меню.
    При открытии через LEFT_MENU — перенаправляет на /app.
    """
    query = dict(request.query_params)
    form_data = getattr(request.state, "form_data", None)
    if form_data is None:
        try:
            form_data = dict(await request.form())
        except Exception:
            form_data = {}

    params = {**query, **form_data}
    logger.info(f"Установка приложения: placement={params.get('PLACEMENT')}, domain={params.get('DOMAIN')}")

    # Извлекаем данные авторизации
    domain = params.get("DOMAIN") or params.get("auth[domain]", "")
    client_endpoint = (
        params.get("auth[client_endpoint]")
        or (f"https://{domain}/rest/" if domain else "")
    )
    tokens = {
        "access_token": params.get("AUTH_ID") or params.get("auth[access_token]", ""),
        "refresh_token": params.get("AUTH_REFRESH_ID") or params.get("auth[refresh_token]", ""),
        "domain": domain,
        "member_id": params.get("member_id") or params.get("auth[member_id]", ""),
        "client_endpoint": client_endpoint,
    }
    _save_tokens(tokens)

    placement = params.get("PLACEMENT", "DEFAULT")
    placement_options = params.get("PLACEMENT_OPTIONS", "")

    # Страница настроек разработчика — выполняем регистрацию
    is_devops = "devops" in placement_options and "edit" in placement_options

    if placement != "DEFAULT" or not is_devops:
        # Открыто через меню или iframe — показываем приложение
        from app.routers.app import app_handler
        return await app_handler(request)

    # Страница devops — регистрируем и возвращаем страницу успеха
    try:
        async with BitrixClient() as bitrix:
            await bitrix.register_activity()
            await bitrix.bind_menu()
        logger.info("Регистрация завершена успешно")
    except Exception as e:
        logger.error(f"Ошибка регистрации: {e}")

    return HTMLResponse(_render_installed_page())
