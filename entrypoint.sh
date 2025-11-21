#!/bin/bash
set -e

# Lancer l'API FastAPI avec Uvicorn
exec uvicorn main:app \
    --host 0.0.0.0 \
    --port 8000 \
    --workers ${WORKERS:-2} \
    --timeout-keep-alive 75
