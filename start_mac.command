#!/bin/zsh
# Arrencada local per a Mac Intel (no canvia el codi del projecte)
cd "$(dirname "$0")"
if [ -d ".venv" ]; then
  source .venv/bin/activate
else
  echo "⚠️ No s'ha trobat l'entorn virtual .venv. Executa ./setup_mac_intel.command primer."
  exit 1
fi
# Alliberar el port 5051 si està en ús
lsof -ti:5051 | xargs kill -9 2>/dev/null
# Arrencar aplicació
python container/app.py

