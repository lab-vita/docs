"""
/admin — панель управления процессами.
Доступна только пользователям из ADMIN_USER_IDS (проверяется в middleware).
"""
import json
import logging
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse, RedirectResponse
from sqlalchemy import select
from sqlalchemy.orm import selectinload

from app.database import SessionLocal
from app.models.process import Process, FormField, DocumentSignature

logger = logging.getLogger(__name__)
router = APIRouter(prefix="/admin")

_TEMPLATES_DIR = Path(__file__).parent.parent / "templates" / "admin"


def _render(template: str, **ctx) -> str:
    html = (_TEMPLATES_DIR / template).read_text(encoding="utf-8")
    for key, value in ctx.items():
        html = html.replace(f"__{key.upper()}__", str(value))
    return html


def _get_auth_id(request: Request) -> str:
    """Извлекает AUTH_ID из query params или form_data."""
    auth_id = (
        request.query_params.get("AUTH_ID")
        or request.query_params.get("auth_id")
        or ""
    )
    if not auth_id:
        form_data = getattr(request.state, "form_data", None) or {}
        auth_id = form_data.get("AUTH_ID") or form_data.get("auth_id") or ""
    return auth_id


# ── Список процессов ──────────────────────────────────────────────────────────

@router.get("", response_class=HTMLResponse)
@router.get("/", response_class=HTMLResponse)
async def admin_index(request: Request):
    async with SessionLocal() as db:
        result = await db.execute(
            select(Process).order_by(Process.sort_order, Process.id)
        )
        processes = result.scalars().all()

    processes_json = json.dumps(
        [
            {
                "id": p.id,
                "key": p.key,
                "title": p.title,
                "description": p.description,
                "icon": p.icon,
                "is_active": p.is_active,
                "sort_order": p.sort_order,
            }
            for p in processes
        ],
        ensure_ascii=False,
    )
    auth_id = _get_auth_id(request)
    html = (_TEMPLATES_DIR / "index.html").read_text(encoding="utf-8")
    html = html.replace("__PROCESSES_JSON__", processes_json)
    html = html.replace("__AUTH_ID__", auth_id)
    return HTMLResponse(html)


# ── Создание процесса ─────────────────────────────────────────────────────────

@router.get("/process/new", response_class=HTMLResponse)
async def admin_process_new_form(request: Request):
    auth_id = _get_auth_id(request)
    html = (_TEMPLATES_DIR / "process_edit.html").read_text(encoding="utf-8")
    html = html.replace("__PROCESS_JSON__", "null")
    html = html.replace("__TITLE__", "Новый процесс")
    html = html.replace("__FORM_ACTION__", f"/admin/process/new?AUTH_ID={auth_id}")
    html = html.replace("__AUTH_ID__", auth_id)
    return HTMLResponse(html)


@router.post("/process/new")
async def admin_process_new_save(request: Request):
    auth_id = _get_auth_id(request)
    form_data = getattr(request.state, "form_data", None)
    if form_data is None:
        try:
            form_data = dict(await request.form())
        except Exception:
            form_data = {}

    process = Process(
        key=form_data.get("key", "").strip(),
        title=form_data.get("title", "").strip(),
        description=form_data.get("description", "").strip(),
        icon=form_data.get("icon", "📄").strip(),
        iblock_id=form_data.get("iblock_id", "").strip(),
        bp_template_id=form_data.get("bp_template_id", "").strip(),
        template_file_id=form_data.get("template_file_id", "").strip(),
        output_folder_id=form_data.get("output_folder_id", "").strip(),
        iblock_fields=_parse_json_field(form_data.get("iblock_fields", "{}")),
        is_active=form_data.get("is_active") == "on",
        sort_order=int(form_data.get("sort_order", 0) or 0),
    )

    _apply_fields(process, form_data)
    _apply_signatures(process, form_data)

    async with SessionLocal() as db:
        db.add(process)
        await db.commit()
        await db.refresh(process)

    return RedirectResponse(f"/admin/process/{process.id}?AUTH_ID={auth_id}", status_code=303)


# ── Редактирование процесса ───────────────────────────────────────────────────

@router.get("/process/{process_id}", response_class=HTMLResponse)
async def admin_process_edit_form(process_id: int, request: Request):
    async with SessionLocal() as db:
        result = await db.execute(
            select(Process)
            .options(selectinload(Process.fields), selectinload(Process.signatures))
            .where(Process.id == process_id)
        )
        process = result.scalar_one_or_none()

    if not process:
        return HTMLResponse("<p>Процесс не найден</p>", status_code=404)

    auth_id = _get_auth_id(request)

    process_json = json.dumps(
        {
            "id": process.id,
            "key": process.key,
            "title": process.title,
            "description": process.description,
            "icon": process.icon,
            "iblock_id": process.iblock_id,
            "bp_template_id": process.bp_template_id,
            "template_file_id": process.template_file_id,
            "output_folder_id": process.output_folder_id,
            "iblock_fields": process.iblock_fields or {},
            "is_active": process.is_active,
            "sort_order": process.sort_order,
            "fields": [
                {
                    "id": f.id,
                    "name": f.name,
                    "label": f.label,
                    "field_type": f.field_type,
                    "required": f.required,
                    "options": f.options or [],
                    "placeholder": f.placeholder or "",
                    "sort_order": f.sort_order,
                }
                for f in sorted(process.fields, key=lambda x: x.sort_order)
            ],
            "signatures": [
                {
                    "id": s.id,
                    "placeholder": s.placeholder,
                    "label": s.label,
                    "stage": s.stage,
                    "source": s.source,
                }
                for s in process.signatures
            ],
        },
        ensure_ascii=False,
    )

    html = (_TEMPLATES_DIR / "process_edit.html").read_text(encoding="utf-8")
    html = html.replace("__PROCESS_JSON__", process_json)
    html = html.replace("__TITLE__", process.title)
    html = html.replace("__FORM_ACTION__", f"/admin/process/{process_id}?AUTH_ID={auth_id}")
    html = html.replace("__AUTH_ID__", auth_id)
    return HTMLResponse(html)


@router.post("/process/{process_id}")
async def admin_process_edit_save(process_id: int, request: Request):
    auth_id = _get_auth_id(request)
    form_data = getattr(request.state, "form_data", None)
    if form_data is None:
        try:
            form_data = dict(await request.form())
        except Exception:
            form_data = {}

    async with SessionLocal() as db:
        result = await db.execute(
            select(Process)
            .options(selectinload(Process.fields), selectinload(Process.signatures))
            .where(Process.id == process_id)
        )
        process = result.scalar_one_or_none()
        if not process:
            return HTMLResponse("<p>Процесс не найден</p>", status_code=404)

        process.key = form_data.get("key", process.key).strip()
        process.title = form_data.get("title", process.title).strip()
        process.description = form_data.get("description", process.description).strip()
        process.icon = form_data.get("icon", process.icon).strip()
        process.iblock_id = form_data.get("iblock_id", process.iblock_id).strip()
        process.bp_template_id = form_data.get("bp_template_id", process.bp_template_id).strip()
        process.template_file_id = form_data.get("template_file_id", process.template_file_id).strip()
        process.output_folder_id = form_data.get("output_folder_id", process.output_folder_id).strip()
        process.iblock_fields = _parse_json_field(
            form_data.get("iblock_fields", json.dumps(process.iblock_fields or {}))
        )
        process.is_active = form_data.get("is_active") == "on"
        process.sort_order = int(form_data.get("sort_order", process.sort_order) or 0)

        # Пересоздаём поля и подписи
        for f in list(process.fields):
            await db.delete(f)
        for s in list(process.signatures):
            await db.delete(s)
        await db.flush()

        _apply_fields(process, form_data)
        _apply_signatures(process, form_data)

        await db.commit()

    return RedirectResponse(f"/admin/process/{process_id}?AUTH_ID={auth_id}", status_code=303)


# ── Вкл/Выкл процесса ─────────────────────────────────────────────────────────

@router.post("/process/{process_id}/toggle")
async def admin_process_toggle(process_id: int):
    async with SessionLocal() as db:
        result = await db.execute(select(Process).where(Process.id == process_id))
        process = result.scalar_one_or_none()
        if not process:
            return {"success": False, "error": "Процесс не найден"}
        process.is_active = not process.is_active
        await db.commit()
        return {"success": True, "is_active": process.is_active}


# ── Удаление процесса ─────────────────────────────────────────────────────────

@router.post("/process/{process_id}/delete")
async def admin_process_delete(process_id: int):
    async with SessionLocal() as db:
        result = await db.execute(select(Process).where(Process.id == process_id))
        process = result.scalar_one_or_none()
        if not process:
            return {"success": False, "error": "Процесс не найден"}
        await db.delete(process)
        await db.commit()
    return RedirectResponse("/admin", status_code=303)


# ── Вспомогательные функции ───────────────────────────────────────────────────

def _parse_json_field(value: str) -> dict:
    try:
        result = json.loads(value)
        return result if isinstance(result, dict) else {}
    except Exception:
        return {}


def _apply_fields(process: Process, form_data: dict) -> None:
    """
    Читает поля формы из form_data в формате:
    fields[0][name], fields[0][label], fields[0][field_type], ...
    """
    fields_map: dict[int, dict] = {}
    for key, value in form_data.items():
        if not key.startswith("fields["):
            continue
        rest = key[7:]
        idx_end = rest.index("]")
        idx = int(rest[:idx_end])
        field_key = rest[idx_end + 2:-1]
        if idx not in fields_map:
            fields_map[idx] = {}
        fields_map[idx][field_key] = value

    for idx in sorted(fields_map.keys()):
        f = fields_map[idx]
        if not f.get("name") or not f.get("label"):
            continue
        options_raw = f.get("options", "[]")
        try:
            options = json.loads(options_raw)
        except Exception:
            options = []
        process.fields.append(
            FormField(
                name=f["name"].strip(),
                label=f["label"].strip(),
                field_type=f.get("field_type", "text"),
                required=f.get("required") == "on",
                options=options,
                placeholder=f.get("placeholder", "").strip(),
                sort_order=idx,
            )
        )


def _apply_signatures(process: Process, form_data: dict) -> None:
    """
    Читает подписи из form_data в формате:
    signatures[0][placeholder], signatures[0][label], ...
    """
    sigs_map: dict[int, dict] = {}
    for key, value in form_data.items():
        if not key.startswith("signatures["):
            continue
        rest = key[11:]
        idx_end = rest.index("]")
        idx = int(rest[:idx_end])
        field_key = rest[idx_end + 2:-1]
        if idx not in sigs_map:
            sigs_map[idx] = {}
        sigs_map[idx][field_key] = value

    for idx in sorted(sigs_map.keys()):
        s = sigs_map[idx]
        if not s.get("placeholder"):
            continue
        process.signatures.append(
            DocumentSignature(
                placeholder=s["placeholder"].strip(),
                label=s.get("label", "").strip(),
                stage=s.get("stage", "initial"),
                source=s.get("source", "employee_profile"),
            )
        )
