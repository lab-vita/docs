#!/bin/sh
set -e

echo "Running database migrations..."
cd /app
alembic upgrade head

echo "Running seed..."
python -m app.seed

echo "Starting application..."
exec uvicorn app.main:application --host 0.0.0.0 --port 5000
