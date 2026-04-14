#!/bin/bash
# Exécuté UNE SEULE FOIS par postgres au 1er démarrage
# Les relances suivantes → volume persistant → ce script est ignoré
set -e

echo "======================================"
echo "ASELBY — Initialisation base de données"
echo "======================================"

DUMP_FILE="/docker/dump/aselby_dump.sql"

if [ -f "$DUMP_FILE" ]; then
    echo "Dump trouvé : chargement en cours..."
    psql -v ON_ERROR_STOP=1 \
         --username "$POSTGRES_USER" \
         --dbname   "$POSTGRES_DB" \
         < "$DUMP_FILE"
    echo "✓ Dump chargé — BD initialisée"
else
    echo "Aucun dump trouvé dans /docker/dump/"
    echo "La BD démarre vide. Lancer les imports manuellement."
fi

echo "======================================"
