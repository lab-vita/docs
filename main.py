import os
import requests
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import List, Optional
import logging

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="Bitrix24 Business Process Document Generator")

# URL вебхука Битрикс24 берётся из переменной окружения
BITRIX_WEBHOOK = os.getenv("BITRIX_WEBHOOK_URL")


# === Модели данных ===

class SignatureEntry(BaseModel):
    """Одна подпись в документе"""
    placeholder: str  # Замещающий текст картинки-заглушки в шаблоне, например SIGN_EMPLOYEE
    signature_id: str  # ID файла подписи на Диске Битрикс24


class GenerateRequest(BaseModel):
    """Запрос на генерацию документа"""
    template_id: str  # ID шаблона .docx на Диске Битрикс24
    folder_id: str  # ID папки для сохранения результата
    filename: str  # Базовое имя итогового файла
    data: dict  # Текстовые переменные для подстановки
    signatures: Optional[List[SignatureEntry]] = []  # Список подписей


# === Интеграция с Диском Битрикс24 ===

def b24_download_file(file_id: str) -> bytes:
    """Скачивает файл с Диска Битрикс24 по ID и возвращает содержимое в байтах"""
    # Получаем информацию о файле, в том числе DOWNLOAD_URL
    url = f"{BITRIX_WEBHOOK}disk.file.get.json"
    resp = requests.post(url, json={"id": file_id})
    resp.raise_for_status()

    result = resp.json().get("result", {})
    download_url = result.get("DOWNLOAD_URL")
    if not download_url:
        raise HTTPException(status_code=404, detail=f"Файл {file_id} не найден на Диске")

    # Скачиваем файл по полученному URL
    file_resp = requests.get(download_url)
    file_resp.raise_for_status()
    return file_resp.content


def b24_upload_file(folder_id: str, filename: str, content: bytes) -> str:
    """Загружает файл на Диск Битрикс24 и возвращает ID загруженного файла"""

    # Шаг 1: запрашиваем у Битрикс24 одноразовый URL для загрузки файла
    url = f"{BITRIX_WEBHOOK}disk.folder.uploadfile.json"
    resp = requests.post(url, params={
        "id": folder_id,
        "data[NAME]": filename,
    })
    resp.raise_for_status()

    upload_url = resp.json().get("result", {}).get("uploadUrl")
    if not upload_url:
        raise HTTPException(status_code=500, detail=f"Не получен upload URL: {resp.json()}")

    # Шаг 2: загружаем файл по полученному URL через multipart/form-data
    files = {"file": (filename, content, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
    upload_resp = requests.post(upload_url, files=files)
    logger.info(f"Ответ загрузки: {upload_resp.status_code} {upload_resp.text[:200]}")
    upload_resp.raise_for_status()

    file_id = upload_resp.json().get("result", {}).get("ID")
    if not file_id:
        raise HTTPException(status_code=500, detail=f"Не получен ID файла: {upload_resp.json()}")
    return str(file_id)


# === Эндпоинты ===

@app.get("/health")
async def health():
    return {"status": "ok"}
