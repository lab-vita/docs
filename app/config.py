import os
from pydantic_settings import BaseSettings
from pydantic import field_validator
from typing import List


class Settings(BaseSettings):
    # Bitrix24
    client_id: str = ""
    client_secret: str = ""
    bitrix_webhook_url: str = ""
    app_url: str = "https://docs.lab-vita.ru"

    # Database
    database_url: str = "postgresql+asyncpg://labvita:labvita@postgres:5432/labvita"

    # Security
    admin_user_ids: List[int] = []
    sign_field: str = "UF_USR_1784019697611"

    # App
    app_title: str = "Умные бизнес-процессы"
    token_file: str = "/app/tokens.json"

    @field_validator("admin_user_ids", mode="before")
    @classmethod
    def parse_admin_ids(cls, v):
        if isinstance(v, str):
            return [int(x.strip()) for x in v.split(",") if x.strip()]
        return v

    model_config = {"env_file": ".env", "extra": "ignore"}


settings = Settings()
