import os
import io
import requests
from fastapi import FastAPI, HTTPException
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


# === Эндпоинты ===

@app.post("/generate")
async def generate_document(req: GenerateRequest):
    """
    Генерирует документ из шаблона и загружает на Диск Битрикс24.

    Порядок действий:
    1. Скачиваем шаблон .docx с Диска Б24
    2. Заменяем текстовые переменные ${KEY} значениями из data
    3. Заменяем картинки-заглушки реальными подписями
    4. Сохраняем документ в буфер
    5. Загружаем на Диск Б24 с уникальным именем (добавляем временную метку)
    6. Возвращаем ID и имя загруженного файла
    """
    if not BITRIX_WEBHOOK:
        raise HTTPException(status_code=500, detail="BITRIX_WEBHOOK_URL не настроен")

    logger.info(f"Генерация документа: {req.filename}")

    # Шаг 1: скачиваем шаблон
    template_bytes = b24_download_file(req.template_id)
    logger.info("Шаблон скачан")

    # Шаг 2: открываем шаблон и заменяем текстовые переменные
    doc = Document(io.BytesIO(template_bytes))
    replace_text(doc, req.data)

    # Шаг 3: заменяем картинки подписей
    for sig in req.signatures:
        logger.info(f"Скачиваем подпись ID={sig.signature_id} для {sig.placeholder}")
        sign_bytes = b24_download_file(sig.signature_id)
        replace_image(doc, sig.placeholder, sign_bytes)

    # Шаг 4: сохраняем документ в буфер
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    # Шаг 5: формируем уникальное имя файла с временной меткой
    # чтобы избежать ошибки "файл с таким именем уже существует"
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
