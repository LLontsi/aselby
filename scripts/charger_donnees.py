#!/usr/bin/env python3
"""
Script d'import ASELBY 2026 — exécution locale sans Docker.

Usage:
    cd aselby_final/backend
    pip install -r requirements/base.txt
    python ../scripts/charger_donnees.py --fichiers ../ASELBYTONTINE_2026

Prérequis:
    - PostgreSQL installé et accessible
    - Base de données créée : createdb aselby_db
    - Variables d'environnement configurées (ou modifier les valeurs ci-dessous)
"""

import os
import sys
import argparse

# ============================================================
# CONFIGURATION — modifie ces valeurs si besoin
# ============================================================
DB_NAME     = os.getenv('DB_NAME',     'aselby_db')
DB_USER     = os.getenv('DB_USER',     'aselby_user')
DB_PASSWORD = os.getenv('DB_PASSWORD', 'aselby_pass')
DB_HOST     = os.getenv('DB_HOST',     'localhost')
DB_PORT     = os.getenv('DB_PORT',     '5432')
SECRET_KEY  = os.getenv('SECRET_KEY',  'local-dev-secret-key-aselby-2026')
# ============================================================

def configurer_django(backend_dir):
    """Configure Django sans docker-compose."""
    sys.path.insert(0, backend_dir)
    os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings.dev')

    # Injecter les variables de base de données
    os.environ['DB_NAME']     = DB_NAME
    os.environ['DB_USER']     = DB_USER
    os.environ['DB_PASSWORD'] = DB_PASSWORD
    os.environ['DB_HOST']     = DB_HOST
    os.environ['DB_PORT']     = DB_PORT
    os.environ['SECRET_KEY']  = SECRET_KEY
    os.environ['DEBUG']       = 'True'

    import django
    django.setup()
    print("✓ Django configuré")


def verifier_connexion_db():
    """Vérifie que la base de données est accessible."""
    from django.db import connection
    try:
        connection.ensure_connection()
        print(f"✓ Base de données connectée ({DB_HOST}:{DB_PORT}/{DB_NAME})")
        return True
    except Exception as e:
        print(f"✗ Connexion à la base impossible : {e}")
        print(f"\n  Vérifie que PostgreSQL tourne et que la base existe :")
        print(f"  createdb {DB_NAME}")
        print(f"  createuser {DB_USER} avec le mot de passe {DB_PASSWORD}")
        return False


def appliquer_migrations():
    """Applique les migrations Django."""
    print("\n→ Application des migrations...")
    from django.core.management import call_command
    call_command('migrate', verbosity=1, interactive=False)
    print("✓ Migrations appliquées")


def lancer_import(dossier_excel, dry_run=False, reset=False):
    """Lance le management command d'import."""
    from django.core.management import call_command

    mode = "[DRY-RUN] " if dry_run else ""
    print(f"\n→ {mode}Import des données depuis : {dossier_excel}")

    kwargs = {'fichiers': dossier_excel, 'dry_run': dry_run, 'reset': reset}
    call_command('import_aselby_2026', **kwargs)


def main():
    parser = argparse.ArgumentParser(description='Import ASELBY 2026 en local')
    parser.add_argument('--fichiers', required=True,
                        help='Chemin vers le dossier ASELBYTONTINE_2026')
    parser.add_argument('--dry-run', action='store_true',
                        help='Simuler sans écrire en base')
    parser.add_argument('--reset', action='store_true',
                        help='Supprimer les données existantes avant import')
    parser.add_argument('--skip-migrate', action='store_true',
                        help='Ne pas relancer migrate')
    args = parser.parse_args()

    # Trouver le dossier backend
    script_dir  = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(script_dir)
    backend_dir = os.path.join(project_dir, 'backend')

    if not os.path.isdir(backend_dir):
        print(f"✗ Dossier backend introuvable : {backend_dir}")
        sys.exit(1)

    if not os.path.isdir(args.fichiers):
        print(f"✗ Dossier Excel introuvable : {args.fichiers}")
        sys.exit(1)

    print("=" * 60)
    print("  ASELBY — Import des données 2026")
    print("=" * 60)

    # 1. Configurer Django
    configurer_django(backend_dir)

    # 2. Vérifier la connexion
    if not verifier_connexion_db():
        sys.exit(1)

    # 3. Migrations
    if not args.skip_migrate and not args.dry_run:
        appliquer_migrations()

    # 4. Import
    lancer_import(args.fichiers, dry_run=args.dry_run, reset=args.reset)

    print("\n" + "=" * 60)
    if not args.dry_run:
        print("  Import terminé avec succès !")
        print("  Accès : http://localhost:8000")
        print("  Bureau  : tontine_2026 / tontine_2026")
        print("  Membre  : bakop / 655592631")
    else:
        print("  [DRY-RUN] Simulation terminée — aucune donnée écrite")
    print("=" * 60)


if __name__ == '__main__':
    main()
