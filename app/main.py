"""
Умные бизнес-процессы — точка входа FastAPI приложения.
"""
import logging

from fastapi import FastAPI
from fastapi.responses import JSONResponse

from app.config import settings
from app.middleware.bitrix_auth import BitrixSourceMiddleware, AdminAccessMiddleware
from app.routers import install, app as app_router, api, submit, activity, admin

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-7s %(name)s — %(message)s",
)
logger = logging.getLogger(__name__)

application = FastAPI(
    title=settings.app_title,
    docs_url=None,   # скрываем /docs в продакшне
    redoc_url=None,
)

# ── Middleware (порядок важен — выполняются снизу вверх) ──────────────────────
application.add_middleware(AdminAccessMiddleware)
application.add_middleware(BitrixSourceMiddleware)

# ── Роутеры ───────────────────────────────────────────────────────────────────
application.include_router(install.router)
application.include_router(app_router.router)
application.include_router(api.router)
application.include_router(submit.router)
application.include_router(activity.router)
application.include_router(admin.router)


# ── Health check ──────────────────────────────────────────────────────────────
@application.get("/health")
async def health():
    return {"status": "ok"}


# ── Legacy redirect: /form → /app?process=cash ────────────────────────────────
@application.api_route("/form", methods=["GET", "POST"])
async def form_legacy():
    from fastapi.responses import RedirectResponse
    return RedirectResponse("/app?process=cash", status_code=302)
