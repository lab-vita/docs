FROM python:3.13-slim

WORKDIR /app

# Зависимости
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Исходный код
COPY app/ ./app/
COPY migrations/ ./migrations/
COPY alembic.ini .
COPY entrypoint.sh .

RUN sed -i 's/\r//' entrypoint.sh && chmod +x entrypoint.sh

# Создаём файл токенов если не существует
RUN touch /app/tokens.json

EXPOSE 5000

ENTRYPOINT ["./entrypoint.sh"]

