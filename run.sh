#!/usr/bin/env bash
# IRONMAN Takip — kurulum + çalıştırma (macOS / Linux)
set -e
cd "$(dirname "$0")"
if [ ! -d ".venv" ]; then
  echo "→ Sanal ortam oluşturuluyor..."
  python3 -m venv .venv
  ./.venv/bin/pip install --quiet --upgrade pip
  ./.venv/bin/pip install --quiet -r requirements.txt
  echo "→ Kurulum tamam."
fi
exec ./.venv/bin/python app.py "$@"
