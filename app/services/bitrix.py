"""
Async клиент для Bitrix24 REST API.
Все HTTP-вызовы только здесь — остальной код не знает об httpx.
"""
import json
import logging
from pathlib import Path
from typing import Optional
import httpx
from app.config import settings

logger = logging.getLogger(__name__)


def _load_tokens() -> dict:
    try:
        return json.loads(Path(settings.token_file).read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _save_tokens(tokens: dict) -> None:
    Path(settings.token_file).write_text(
        json.dumps(tokens, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    logger.info("Токены сохранены")


class BitrixClient:
    """Async клиент. Создавать через контекстный менеджер или один раз на запрос."""

    def __init__(self):
        self._client = httpx.AsyncClient(timeout=30.0)

    async def aclose(self):
        await self._client.aclose()

    async def __aenter__(self):
        return self

    async def __aexit__(self, *args):
        await self._client.aclose()

    # ── Токены ────────────────────────────────────────────────────────────────

    def get_client_endpoint(self) -> str:
        return _load_tokens().get("client_endpoint", "")

    def get_access_token(self) -> str:
        return _load_tokens().get("access_token", "")

    async def refresh_token(self) -> str:
        """Обновляет access_token через refresh_token, сохраняет и возвращает новый."""
        tokens = _load_tokens()
        if not tokens.get("refresh_token"):
            raise RuntimeError("Нет refresh_token — переустановите приложение")

        resp = await self._client.post(
            "https://oauth.bitrix.info/oauth/token/",
            params={
                "grant_type": "refresh_token",
                "client_id": settings.client_id,
                "client_secret": settings.client_secret,
                "refresh_token": tokens["refresh_token"],
            },
        )
        resp.raise_for_status()
        new_tokens = resp.json()
        _save_tokens(new_tokens)
        logger.info("Токен обновлён")
        return new_tokens["access_token"]

    async def _app_token(self) -> str:
        """Возвращает актуальный токен приложения (не пользователя)."""
        token = self.get_access_token()
        if not token:
            token = await self.refresh_token()
        return token

    async def _call(
        self,
        method: str,
        params: dict,
        auth: Optional[str] = None,
        retry: bool = True,
    ) -> dict:
        """
        Вызывает REST-метод Bitrix24.
        auth — токен пользователя (если None — используем токен приложения).
        При 401 однократно обновляет токен и повторяет запрос.
        """
        endpoint = self.get_client_endpoint()
        if not endpoint:
            raise RuntimeError("client_endpoint не найден — переустановите приложение")

        token = auth or await self._app_token()
        resp = await self._client.post(
            f"{endpoint}{method}",
            params={"auth": token},
            json=params,
        )

        # Bitrix24 возвращает 200 даже при ошибке авторизации
        if resp.status_code == 401 or (
            resp.status_code == 200
            and resp.json().get("error") in ("expired_token", "INVALID_TOKEN")
            and retry
            and not auth  # обновляем только токен приложения
        ):
            logger.info("Токен истёк, обновляем...")
            token = await self.refresh_token()
            resp = await self._client.post(
                f"{endpoint}{method}",
                params={"auth": token},
                json=params,
            )

        resp.raise_for_status()
        data = resp.json()
        if "error" in data:
            raise RuntimeError(f"Bitrix24 API error: {data['error']} — {data.get('error_description', '')}")
        return data.get("result", data)

    # ── Пользователи ─────────────────────────────────────────────────────────

    async def get_user_profile(self, auth_id: str) -> dict:
        """
        Возвращает профиль текущего пользователя по его AUTH_ID.
        Включает пользовательские поля (UF_*).
        """
        endpoint = self.get_client_endpoint()
        if not endpoint:
            raise RuntimeError("client_endpoint не найден — переустановите приложение")
        resp = await self._client.post(
            f"{endpoint}profile.json",
            params={"auth": auth_id},
        )
        resp.raise_for_status()
        data = resp.json()
        if "error" in data:
            raise RuntimeError(f"Ошибка профиля: {data['error']}")
        return data.get("result", {})

    async def get_user_by_id(self, user_id: int) -> dict:
        """Возвращает данные пользователя по ID через токен приложения."""
        results = await self._call("user.get.json", {"ID": user_id})
        if isinstance(results, list) and results:
            return results[0]
        return {}

    # ── Диск ──────────────────────────────────────────────────────────────────

    async def download_file(self, file_id: str) -> bytes:
        """Скачивает файл с Диска Bitrix24 по ID."""
        result = await self._call("disk.file.get.json", {"id": file_id})
        download_url = result.get("DOWNLOAD_URL")
        if not download_url:
            raise RuntimeError(f"Файл {file_id} не найден на Диске")
        file_resp = await self._client.get(download_url)
        file_resp.raise_for_status()
        return file_resp.content

    async def upload_file(self, folder_id: str, filename: str, content: bytes) -> str:
        """Загружает файл в папку на Диске, возвращает ID файла."""
        # Получаем upload URL через вебхук (он всегда актуален)
        webhook = settings.bitrix_webhook_url.rstrip("/")
        resp = await self._client.post(
            f"{webhook}/disk.folder.uploadfile.json",
            params={"id": folder_id, "data[NAME]": filename},
        )
        resp.raise_for_status()
        upload_url = resp.json().get("result", {}).get("uploadUrl")
        if not upload_url:
            raise RuntimeError("Не получен uploadUrl от Bitrix24")

        upload_resp = await self._client.post(
            upload_url,
            files={"file": (filename, content,
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
        )
        upload_resp.raise_for_status()
        file_id = upload_resp.json().get("result", {}).get("ID")
        if not file_id:
            raise RuntimeError("Не получен ID загруженного файла")
        return str(file_id)

    # ── Списки (инфоблоки) ────────────────────────────────────────────────────

    async def create_list_element(
        self,
        iblock_id: str,
        fields: dict,
        element_code: str,
        auth: str,
    ) -> str:
        """
        Создаёт элемент списка (инфоблок) от имени пользователя.
        Возвращает ID элемента.
        """
        result = await self._call(
            "lists.element.add.json",
            {
                "IBLOCK_TYPE_ID": "bitrix_processes",
                "IBLOCK_ID": iblock_id,
                "ELEMENT_CODE": element_code,
                "FIELDS": fields,
            },
            auth=auth,
        )
        return str(result)

    # ── Бизнес-процессы ───────────────────────────────────────────────────────

    async def start_workflow(
        self,
        template_id: str,
        document_id: list,
        auth: str,
    ) -> dict:
        """Запускает шаблон БП на документ от имени пользователя."""
        return await self._call(
            "bizproc.workflow.start.json",
            {
                "TEMPLATE_ID": template_id,
                "DOCUMENT_ID": document_id,
                "PARAMETERS": {},
            },
            auth=auth,
        )

    async def send_bp_event(self, event_token: str, return_values: dict) -> dict:
        """Отправляет результат выполнения активити обратно в БП."""
        return await self._call(
            "bizproc.event.send.json",
            {
                "event_token": event_token,
                "return_values": return_values,
            },
        )

    # ── Установка ─────────────────────────────────────────────────────────────

    async def register_activity(self) -> None:
        """Регистрирует/перерегистрирует кастомную активити в Bitrix24."""
        endpoint = self.get_client_endpoint()
        token = await self._app_token()

        # Удаляем старую версию
        await self._client.post(
            f"{endpoint}bizproc.activity.delete.json",
            params={"auth": token},
            json={"CODE": "generate_document"},
        )

        params = {
            "CODE": "generate_document",
            "HANDLER": f"{settings.app_url}/activity-handler",
            "AUTH_USER_ID": 1,
            "USE_SUBSCRIPTION": "Y",
            "NAME": {"ru": "Генерация документа", "en": "Generate Document"},
            "DESCRIPTION": {
                "ru": "Генерирует документ из шаблона .docx с подстановкой переменных и подписей",
                "en": "Generates a document from a .docx template",
            },
            "PROPERTIES": {
                "template_id": {
                    "Name": {"ru": "ID шаблона на Диске", "en": "Template ID"},
                    "Type": "string", "Required": "Y", "Multiple": "N",
                },
                "folder_id": {
                    "Name": {"ru": "ID папки для сохранения", "en": "Folder ID"},
                    "Type": "string", "Required": "Y", "Multiple": "N",
                },
                "filename": {
                    "Name": {"ru": "Имя файла", "en": "Filename"},
                    "Type": "string", "Required": "Y", "Multiple": "N",
                },
                "variables": {
                    "Name": {"ru": "Переменные шаблона (KEY|VALUE)", "en": "Template variables"},
                    "Type": "string", "Required": "N", "Multiple": "Y",
                },
                "signatures": {
                    "Name": {"ru": "Подписи (PLACEHOLDER|FILE_ID)", "en": "Signatures"},
                    "Type": "string", "Required": "N", "Multiple": "Y",
                },
                "source_file_id": {
                    "Name": {"ru": "ID существующего файла", "en": "Source file ID"},
                    "Type": "string", "Required": "N", "Multiple": "N",
                },
            },
            "RETURN_PROPERTIES": {
                "file_id": {
                    "Name": {"ru": "ID сгенерированного файла", "en": "Generated file ID"},
                    "Type": "string", "Multiple": "N", "Default": None,
                }
            },
        }

        resp = await self._client.post(
            f"{endpoint}bizproc.activity.add.json",
            params={"auth": token},
            json=params,
        )
        resp.raise_for_status()
        logger.info(f"Активити зарегистрировано: {resp.json().get('result')}")

    async def bind_menu(self) -> None:
        """
        Регистрирует пункт левого меню глобально.
        Примечание: Bitrix24 Cloud не поддерживает принудительную привязку
        к конкретным пользователям (ERROR_PLACEMENT_USER_MODE) для локальных приложений.
        Администратор портала должен добавить приложение в меню вручную через
        настройки Bitrix24 или пользователи добавят его самостоятельно.
        """
        endpoint = self.get_client_endpoint()
        token = await self._app_token()

        # Убираем старые привязки
        for old_handler in [f"{settings.app_url}/form", f"{settings.app_url}/app"]:
            await self._client.post(
                f"{endpoint}placement.unbind.json",
                params={"auth": token},
                json={"PLACEMENT": "LEFT_MENU", "HANDLER": old_handler},
            )

        # Глобальная привязка
        resp = await self._client.post(
            f"{endpoint}placement.bind.json",
            params={"auth": token},
            json={
                "PLACEMENT": "LEFT_MENU",
                "HANDLER": f"{settings.app_url}/app",
                "TITLE": settings.app_title,
            },
        )
        logger.info(f"Глобальная привязка меню: {resp.json()}")

