import os
import requests
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import List, Optional
from docx import Document
from docx.oxml.ns import qn
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
    url = f"{BITRIX_WEBHOOK}disk.file.get.json"
    resp = requests.post(url, json={"id": file_id})
    resp.raise_for_status()

    result = resp.json().get("result", {})
    download_url = result.get("DOWNLOAD_URL")
    if not download_url:
        raise HTTPException(status_code=404, detail=f"Файл {file_id} не найден на Диске")

    file_resp = requests.get(download_url)
    file_resp.raise_for_status()
    return file_resp.content


def b24_upload_file(folder_id: str, filename: str, content: bytes) -> str:
    """Загружает файл на Диск Битрикс24 и возвращает ID загруженного файла"""
    url = f"{BITRIX_WEBHOOK}disk.folder.uploadfile.json"
    resp = requests.post(url, params={
        "id": folder_id,
        "data[NAME]": filename,
    })
    resp.raise_for_status()

    upload_url = resp.json().get("result", {}).get("uploadUrl")
    if not upload_url:
        raise HTTPException(status_code=500, detail=f"Не получен upload URL: {resp.json()}")

    files = {"file": (filename, content, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
    upload_resp = requests.post(upload_url, files=files)
    logger.info(f"Ответ загрузки: {upload_resp.status_code} {upload_resp.text[:200]}")
    upload_resp.raise_for_status()

    file_id = upload_resp.json().get("result", {}).get("ID")
    if not file_id:
        raise HTTPException(status_code=500, detail=f"Не получен ID файла: {upload_resp.json()}")
    return str(file_id)


# === Обработка документа ===

def replace_paragraph_text(paragraph, data: dict):
    """
    Заменяет переменные вида ${KEY} в одном параграфе.

    Проблема: Word может разбивать текст на несколько run-ов,
    из-за чего ${KEY} оказывается разбит между ними.
    Решение: собираем полный текст параграфа, заменяем,
    записываем результат в первый run, остальные очищаем.
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
    Заменяет картинку-заглушку в документе на реальное изображение подписи.

    Поиск картинки ведётся по полю "Описание" замещающего текста (Alt Text).
    В шаблоне Word: правая кнопка на картинке → Изменить замещающий текст →
    поле "Описание" = placeholder_desc (например SIGN_EMPLOYEE).

    Замена происходит на уровне бинарных данных — blob изображения
    заменяется напрямую в part документа без изменения размеров и положения.
    """

    def process_paragraphs(paragraphs):
        for p in paragraphs:
            for run in p.runs:
                # Ищем все inline и anchor изображения в run-е
                drawings = (
                        run._element.findall('.//' + qn('wp:inline')) +
                        run._element.findall('.//' + qn('wp:anchor'))
                )
                for drawing in drawings:
                    # Получаем docPr — блок с метаданными изображения
                    docPr = drawing.find('.//' + qn('wp:docPr'))
                    if docPr is None:
                        continue

                    # Проверяем совпадение замещающего текста
                    if docPr.get('descr', '') != placeholder_desc:
                        continue

                    # Находим blip — ссылку на файл изображения внутри документа
                    blip = drawing.find('.//' + qn('a:blip'))
                    if blip is None:
                        continue

                    # Получаем relationship ID и заменяем blob
                    r_embed = blip.get(qn('r:embed'))
                    img_part = doc.part.related_parts[r_embed]
                    img_part._blob = image_bytes
                    logger.info(f"Заменена картинка: {placeholder_desc}")
                    return True
        return False

    # Ищем в основных параграфах
    process_paragraphs(doc.paragraphs)

    # Ищем в таблицах
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                process_paragraphs(cell.paragraphs)


# === Эндпоинты ===

@app.get("/health")
async def health():
    return {"status": "ok"}
