import os
import io
import json
import requests
from fastapi import FastAPI, HTTPException, Request
from pydantic import BaseModel
from typing import List, Optional
from docx import Document
from docx.oxml.ns import qn
from datetime import datetime
import logging

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="Document Generator for Bitrix24")

# Константы приложения из переменных окружения
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
BITRIX_WEBHOOK = os.getenv("BITRIX_WEBHOOK_URL")
APP_URL = os.getenv("APP_URL", "https://docs.lab-vita.ru")
TOKEN_FILE = "/app/tokens.json"


# === Модели данных ===

class SignatureEntry(BaseModel):
    placeholder: str  # Замещающий текст картинки-заглушки в шаблоне
    signature_id: str  # ID файла подписи на Диске Битрикс24


class GenerateRequest(BaseModel):
    template_id: str
    folder_id: str
    filename: str
    data: dict
    signatures: Optional[List[SignatureEntry]] = []


# === Работа с токенами ===

def save_tokens(tokens: dict):
    """Сохраняет токены OAuth в файл"""
    with open(TOKEN_FILE, "w") as f:
        json.dump(tokens, f)
    logger.info("Токены сохранены")


def load_tokens() -> dict:
    """Загружает токены OAuth из файла"""
    try:
        with open(TOKEN_FILE, "r") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}


def refresh_access_token() -> str:
    """Обновляет access_token через refresh_token"""
    tokens = load_tokens()
    if not tokens.get("refresh_token"):
        raise HTTPException(status_code=401, detail="Нет refresh_token — переустановите приложение")

    resp = requests.post("https://oauth.bitrix.info/oauth/token/", params={
        "grant_type": "refresh_token",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "refresh_token": tokens["refresh_token"],
    })
    resp.raise_for_status()
    new_tokens = resp.json()
    save_tokens(new_tokens)
    logger.info("Токен обновлён")
    return new_tokens["access_token"]


def get_access_token() -> str:
    """Возвращает актуальный access_token, при необходимости обновляет"""
    tokens = load_tokens()
    if not tokens.get("access_token"):
        return refresh_access_token()
    return tokens["access_token"]


# === Вызовы REST API Битрикс24 ===

def b24_call(method: str, params: dict, use_webhook: bool = False) -> dict:
    """Универсальный вызов REST API Битрикс24"""
    if use_webhook and BITRIX_WEBHOOK:
        url = f"{BITRIX_WEBHOOK}{method}.json"
        resp = requests.post(url, json=params)
    else:
        access_token = get_access_token()
        tokens = load_tokens()
        # Используем client_endpoint если есть, иначе формируем из domain
        client_endpoint = tokens.get("client_endpoint", "")
        domain = tokens.get("domain", "")
        if client_endpoint:
            url = f"{client_endpoint}{method}.json"
        else:
            url = f"https://{domain}/rest/{method}.json"
        logger.info(f"B24 call: {url}")
        resp = requests.post(url, json={**params, "auth": access_token})

    resp.raise_for_status()
    result = resp.json()

    # Если токен протух — обновляем и повторяем
    if isinstance(result, dict) and result.get("error") == "expired_token":
        access_token = refresh_access_token()
        tokens = load_tokens()
        domain = tokens.get("domain", "")
        url = f"https://{domain}/rest/{method}.json"
        resp = requests.post(url, json={**params, "auth": access_token})
        resp.raise_for_status()
        result = resp.json()

    return result.get("result", result)


# === Интеграция с Диском Битрикс24 ===

def b24_download_file(file_id: str, use_webhook: bool = True) -> bytes:
    """Скачивает файл с Диска Битрикс24 по ID"""
    result = b24_call("disk.file.get", {"id": file_id}, use_webhook=use_webhook)
    download_url = result.get("DOWNLOAD_URL")
    if not download_url:
        raise HTTPException(status_code=404, detail=f"Файл {file_id} не найден на Диске")
    file_resp = requests.get(download_url)
    file_resp.raise_for_status()
    return file_resp.content


def b24_upload_file(folder_id: str, filename: str, content: bytes) -> str:
    """Загружает файл на Диск Битрикс24 и возвращает ID"""
    url = f"{BITRIX_WEBHOOK}disk.folder.uploadfile.json"
    resp = requests.post(url, params={"id": folder_id, "data[NAME]": filename})
    resp.raise_for_status()
    upload_url = resp.json().get("result", {}).get("uploadUrl")
    if not upload_url:
        raise HTTPException(status_code=500, detail="Не получен upload URL")

    files = {"file": (filename, content, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
    upload_resp = requests.post(upload_url, files=files)
    upload_resp.raise_for_status()
    file_id = upload_resp.json().get("result", {}).get("ID")
    if not file_id:
        raise HTTPException(status_code=500, detail="Не получен ID файла")
    return str(file_id)


# === Обработка документа ===

def replace_paragraph_text(paragraph, data: dict):
    """
    Заменяет переменные вида ${KEY} в одном параграфе.
    Собирает полный текст из всех run-ов чтобы обойти разбивку Word.
    """
    for key, value in data.items():
        placeholder = f"${{{key}}}"
        if placeholder not in paragraph.text:
            continue
        full_text = "".join(run.text for run in paragraph.runs)
        if placeholder not in full_text:
            continue
        new_text = full_text.replace(placeholder, str(value))
        if paragraph.runs:
            paragraph.runs[0].text = new_text
            for run in paragraph.runs[1:]:
                run.text = ""


def replace_text(doc: Document, data: dict):
    """Заменяет текстовые переменные во всём документе — в параграфах и таблицах"""
    for paragraph in doc.paragraphs:
        replace_paragraph_text(paragraph, data)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_paragraph_text(paragraph, data)


def replace_image(doc: Document, placeholder_desc: str, image_bytes: bytes):
    """
    Заменяет картинку-заглушку на реальное изображение подписи.
    Поиск по полю "Описание" замещающего текста (Alt Text).
    """

    def process_paragraphs(paragraphs):
        for p in paragraphs:
            for run in p.runs:
                drawings = (
                        run._element.findall('.//' + qn('wp:inline')) +
                        run._element.findall('.//' + qn('wp:anchor'))
                )
                for drawing in drawings:
                    docPr = drawing.find('.//' + qn('wp:docPr'))
                    if docPr is None:
                        continue
                    if docPr.get('descr', '') != placeholder_desc:
                        continue
                    blip = drawing.find('.//' + qn('a:blip'))
                    if blip is None:
                        continue
                    r_embed = blip.get(qn('r:embed'))
                    img_part = doc.part.related_parts[r_embed]
                    img_part._blob = image_bytes
                    logger.info(f"Заменена картинка: {placeholder_desc}")
                    return True
        return False

    process_paragraphs(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                process_paragraphs(cell.paragraphs)


def parse_variables(variables_list) -> dict:
    """
    Парсит переменные из формата KEY|VALUE в словарь.
    Поддерживает строку и список строк.
    Пример: ["EMPLOYEE_NAME|Егошин Алексей", "AMOUNT|20000"] -> {"EMPLOYEE_NAME": "Егошин Алексей", ...}
    """
    result = {}
    if isinstance(variables_list, str):
        variables_list = [variables_list]
    for item in variables_list or []:
        if "|" in item:
            key, _, value = item.partition("|")
            result[key.strip()] = value.strip()
    return result


def parse_signatures(signatures_list) -> list:
    """
    Парсит подписи из формата PLACEHOLDER|FILE_ID в список словарей.
    Пример: ["SIGN_EMPLOYEE|837426"] -> [{"placeholder": "SIGN_EMPLOYEE", "signature_id": "837426"}]
    """
    result = []
    if isinstance(signatures_list, str):
        signatures_list = [signatures_list]
    for item in signatures_list or []:
        if "|" in item:
            placeholder, _, file_id = item.partition("|")
            result.append({"placeholder": placeholder.strip(), "signature_id": file_id.strip()})
    return result


# === Регистрация активити в Битрикс24 ===

def register_activity(domain: str, access_token: str):
    """Регистрирует кастомное активити в дизайнере БП"""
    params = {
        "auth": access_token,
        "CODE": "generate_document",
        "HANDLER": f"{APP_URL}/activity-handler",
        "AUTH_USER_ID": 1,
        "USE_SUBSCRIPTION": "Y",
        "NAME": {"ru": "Генерация документа", "en": "Generate Document"},
        "DESCRIPTION": {
            "ru": "Генерирует документ из шаблона .docx с подстановкой переменных и подписей",
            "en": "Generates a document from a .docx template"
        },
        "PROPERTIES": {
            "template_id": {
                "Name": {"ru": "ID шаблона на Диске", "en": "Template ID"},
                "Type": "string",
                "Required": "Y",
                "Multiple": "N",
            },
            "folder_id": {
                "Name": {"ru": "ID папки для сохранения", "en": "Folder ID"},
                "Type": "string",
                "Required": "Y",
                "Multiple": "N",
            },
            "filename": {
                "Name": {"ru": "Имя файла", "en": "Filename"},
                "Type": "string",
                "Required": "Y",
                "Multiple": "N",
            },
            "variables": {
                # Множественное поле — пользователь добавляет строки KEY|VALUE
                "Name": {"ru": "Переменные шаблона (KEY|VALUE)", "en": "Template variables (KEY|VALUE)"},
                "Type": "string",
                "Required": "N",
                "Multiple": "Y",
            },
            "signatures": {
                # Множественное поле — пользователь добавляет строки PLACEHOLDER|FILE_ID
                "Name": {"ru": "Подписи (PLACEHOLDER|FILE_ID)", "en": "Signatures (PLACEHOLDER|FILE_ID)"},
                "Type": "string",
                "Required": "N",
                "Multiple": "Y",
            },
        },
        "RETURN_PROPERTIES": {
            "file_id": {
                "Name": {"ru": "ID сгенерированного файла", "en": "Generated file ID"},
                "Type": "string",
                "Multiple": "N",
                "Default": None,
            }
        },
    }

    tokens = load_tokens()
    client_endpoint = tokens.get("client_endpoint") or f"https://{domain}/rest/"
    url = f"{client_endpoint}bizproc.activity.add.json"
    resp = requests.post(url, json=params)
    resp.raise_for_status()
    result = resp.json()
    logger.info(f"Результат регистрации активити: {result}")
    return result


# === Эндпоинты ===

@app.api_route("/install", methods=["GET", "POST"])
async def install(request: Request):
    """
    Вызывается Битрикс24 при установке приложения.
    Сохраняет токены и регистрирует активити в дизайнере БП.
    """
    # Битрикс24 может передавать параметры как в query string так и в теле POST
    query_params = dict(request.query_params)
    try:
        form_params = dict(await request.form())
    except Exception:
        form_params = {}

    # Объединяем параметры, form имеет приоритет
    params = {**query_params, **form_params}
    logger.info(f"Установка приложения: {params}")

    tokens = {
        "access_token": params.get("AUTH_ID") or params.get("auth_id") or params.get("auth[access_token]"),
        "refresh_token": params.get("AUTH_REFRESH_ID") or params.get("auth_refresh_id") or params.get(
            "auth[refresh_token]"),
        "domain": params.get("DOMAIN") or params.get("domain") or params.get("auth[domain]"),
        "member_id": params.get("member_id") or params.get("auth[member_id]"),
        "client_endpoint": params.get("auth[client_endpoint]"),
    }
    save_tokens(tokens)

    try:
        register_activity(tokens["domain"], tokens["access_token"])
        logger.info("Активити зарегистрировано")
    except Exception as e:
        logger.error(f"Ошибка регистрации активити: {e}")

    return {"status": "ok", "message": "Приложение установлено, активити зарегистрировано"}


@app.post("/app")
async def app_handler(request: Request):
    """Обработчик событий приложения от Битрикс24"""
    body = await request.body()
    logger.info(f"Событие приложения: {body[:200]}")
    return {"status": "ok"}


@app.post("/activity-handler")
async def activity_handler(request: Request):
    """
    Обработчик активити — вызывается Битрикс24 когда БП
    доходит до действия 'Генерация документа'.

    Получает параметры в формате:
    - variables: список строк KEY|VALUE (переменные шаблона)
    - signatures: список строк PLACEHOLDER|FILE_ID (подписи)
    - template_id, folder_id, filename — обязательные поля
    """
    body = await request.body()
    logger.info(f"Вызов активити: {body[:500]}")

    try:
        data = json.loads(body)
    except Exception:
        form = await request.form()
        data = dict(form)

    props = data.get("properties", {})
    event_token = data.get("event_token")
    auth = data.get("auth", {})

    # Основные параметры
    template_id = props.get("template_id", "")
    folder_id = props.get("folder_id", "")
    filename = props.get("filename", "document.docx")

    # Парсим переменные и подписи из формата KEY|VALUE
    doc_data = parse_variables(props.get("variables", []))
    signatures = parse_signatures(props.get("signatures", []))

    logger.info(f"Переменные: {doc_data}")
    logger.info(f"Подписи: {signatures}")

    # Генерируем документ — используем токен приложения для скачивания
    template_bytes = b24_download_file(template_id, use_webhook=False)
    doc = Document(io.BytesIO(template_bytes))
    replace_text(doc, doc_data)

    for sig in signatures:
        sign_bytes = b24_download_file(sig["signature_id"], use_webhook=False)
        replace_image(doc, sig["placeholder"], sign_bytes)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    name, ext = filename.rsplit(".", 1) if "." in filename else (filename, "docx")
    unique_filename = f"{name}_{timestamp}.{ext}"

    file_id = b24_upload_file(folder_id, unique_filename, output.read())
    logger.info(f"Документ сгенерирован, ID: {file_id}")

    # Возвращаем результат в БП через bizproc.event.send
    tokens = load_tokens()
    domain = tokens.get("domain", auth.get("domain", ""))
    access_token = auth.get("access_token") or get_access_token()

    return_url = f"https://{domain}/rest/bizproc.event.send.json"
    return_resp = requests.post(return_url, json={
        "auth": access_token,
        "event_token": event_token,
        "return_values": {"file_id": file_id},
        "log_message": f"Документ сгенерирован: {unique_filename}",
    })
    logger.info(f"Результат отправки в БП: {return_resp.text}")

    return {"status": "ok", "file_id": file_id}


@app.post("/generate")
async def generate_document(req: GenerateRequest):
    """Ручная генерация документа через POST запрос (для тестирования)"""
    if not BITRIX_WEBHOOK:
        raise HTTPException(status_code=500, detail="BITRIX_WEBHOOK_URL не настроен")

    logger.info(f"Генерация документа: {req.filename}")

    template_bytes = b24_download_file(req.template_id)
    doc = Document(io.BytesIO(template_bytes))
    replace_text(doc, req.data)

    for sig in req.signatures:
        logger.info(f"Скачиваем подпись ID={sig.signature_id} для {sig.placeholder}")
        sign_bytes = b24_download_file(sig.signature_id)
        replace_image(doc, sig.placeholder, sign_bytes)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    name, ext = req.filename.rsplit(".", 1) if "." in req.filename else (req.filename, "docx")
    unique_filename = f"{name}_{timestamp}.{ext}"

    logger.info(f"Загружаем в папку ID={req.folder_id} как {unique_filename}")
    file_id = b24_upload_file(req.folder_id, unique_filename, output.read())

    logger.info(f"Готово! ID файла: {file_id}")
    return {"file_id": file_id, "filename": unique_filename}


@app.get("/health")
async def health():
    return {"status": "ok"}
