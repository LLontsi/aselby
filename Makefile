.PHONY: up down build dev migrate import import-test import-reset shell logs collectstatic createsuperuser

# ---- Production ----
up:
	docker-compose up -d

down:
	docker-compose down

build:
	docker-compose build --no-cache

logs:
	docker-compose logs -f web

# ---- Import des données Excel ----
# Le dossier ASELBYTONTINE_2026 est monté dans /data/excel
import:
	docker-compose exec web python manage.py import_aselby_2026 --fichiers /data/excel

import-test:
	docker-compose exec web python manage.py import_aselby_2026 --fichiers /data/excel --dry-run

import-reset:
	docker-compose exec web python manage.py import_aselby_2026 --fichiers /data/excel --reset

# ---- Autres commandes ----
shell:
	docker-compose exec web python manage.py shell

migrate:
	docker-compose exec web python manage.py migrate

collectstatic:
	docker-compose exec web python manage.py collectstatic --noinput

createsuperuser:
	docker-compose exec web python manage.py createsuperuser
