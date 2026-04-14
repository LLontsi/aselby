import time
from django.core.management.base import BaseCommand
from django.db import connections
from django.db.utils import OperationalError


class Command(BaseCommand):
    help = "Attend que la base de données soit disponible"

    def handle(self, *args, **options):
        self.stdout.write("Attente de la base de données...")
        db_conn = None
        attempts = 0
        while not db_conn:
            try:
                db_conn = connections['default']
                db_conn.ensure_connection()
            except OperationalError:
                attempts += 1
                self.stdout.write(f"  Base non disponible, tentative {attempts}... (attente 2s)")
                time.sleep(2)
        self.stdout.write(self.style.SUCCESS("Base de données disponible !"))
