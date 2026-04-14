# TontineCM — ASELBY

Application web de gestion de tontine pour l'Association des Élites Bangou de Yaoundé (ASELBY).

## Stack
- **Backend** : Django 5 + PostgreSQL
- **Frontend** : Django Templates + Tailwind CSS
- **Déploiement** : Docker + Nginx

## Démarrage rapide

```bash
# 1. Copier et configurer l'environnement
cp .env.example .env
# Éditez .env avec vos valeurs

# 2. Lancer Docker
docker-compose up -d

# 3. Migrations
docker-compose exec web python manage.py migrate

# 4. Importer les données Excel 2026
docker-compose exec web python manage.py import_aselby_2026 \
    --fichiers /chemin/vers/fichiers/excel

# 5. Accéder à l'application
# http://localhost
```

## Comptes par défaut (après import)
| Rôle | Username | Mot de passe |
|------|----------|--------------|
| Bureau | `tontine_2026` | `tontine_2026` |
| Membre | `<prénom>` | numéro de téléphone ou `aselby_2026` |

## Structure
```
├── backend/
│   ├── apps/         # 15 applications Django
│   ├── config/       # Settings, URLs, WSGI
│   ├── templates/    # Templates HTML (Django)
│   ├── static/       # CSS, JS, images
│   └── manage.py
└── docker/           # Nginx, Postgres
```

## Développeur
LONTSI LAMBOU RONALDINO — PropentaTech — 2026
