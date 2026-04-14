#!/bin/bash
set -e

echo "=== ASELBY — Démarrage ==="

echo "1. Attente de la base de données..."
python manage.py wait_for_db

echo "2. Application des migrations..."
python manage.py migrate --no-input

echo "3. Collecte des fichiers statiques..."
python manage.py collectstatic --no-input --clear 2>/dev/null || python manage.py collectstatic --no-input

echo "4. Démarrage de Gunicorn..."
exec gunicorn config.wsgi:application \
    --bind 0.0.0.0:8000 \
    --workers 2 \
    --timeout 120 \
    --access-logfile - \
    --error-logfile -
