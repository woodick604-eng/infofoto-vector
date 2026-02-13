#!/usr/bin/env zsh
# Configuració automàtica per a Mac Intel (sense canviar el codi del projecte)
set -euo pipefail
cd "$(dirname "$0")"

echo "🔧 Comprovant Python 3..."
if ! command -v python3 >/dev/null 2>&1; then
  echo "❌ No s'ha trobat python3 al sistema."
  echo "   Instal·la'l (p.ex. amb Homebrew: brew install python) i torna-ho a provar."
  exit 1
fi

# Crear entorn virtual si no existeix
if [ ! -d ".venv" ]; then
  echo "📦 Creant entorn virtual .venv..."
  python3 -m venv .venv
else
  echo "ℹ️ Ja existeix .venv; s'utilitzarà l'existent."
fi

echo "➡️ Activant entorn virtual..."
source .venv/bin/activate

echo "⬆️ Actualitzant pip i wheel..."
python -m pip install --upgrade pip wheel

# Detectar fitxer de requeriments
REQS=""
if [ -f "requirements.txt" ]; then
  REQS="requirements.txt"
elif [ -f "container/requirements.txt" ]; then
  REQS="container/requirements.txt"
elif [ -f "requeriments.txt" ]; then
  REQS="requeriments.txt"
fi


if [ -n "$REQS" ]; then
  echo "📥 Instal·lant dependències des de $REQS ..."
  pip install -r "$REQS"
else
  echo "⚠️ No s'ha trobat ni requirements.txt ni requeriments.txt."
  echo "   Pots instal·lar manualment les dependències necessàries quan convingui."
fi

# Missatge final
cat <<'EOF'
✅ Entorn preparat correctament.

Per arrencar l'aplicació:
  ./start_mac.command

Si macOS bloqueja l'script:
  - Clic dret > Obre, o
  - chmod +x setup_mac_intel.command start_mac.command

Si el port 5051 està ocupat, tanca processos previs o reinicia l'script d'arrencada.
EOF
