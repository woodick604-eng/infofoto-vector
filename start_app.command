#!/bin/zsh
set -euo pipefail
cd "$(dirname "$0")"

PORT=5051

pids=$(lsof -tiTCP:$PORT 2>/dev/null || true)
[ -n "${pids:-}" ] && kill -9 $pids || true

xattr -dr com.apple.quarantine . 2>/dev/null || true

if [ ! -d .venv ]; then
  python3 -m venv .venv
fi
source .venv/bin/activate
python -m pip install --upgrade pip >/dev/null
pip install -q flask python-docx pillow pillow-heif

cat > run_local.py <<PY
import threading, time, webbrowser
from app import app
def op():
    time.sleep(1)
    webbrowser.open("http://127.0.0.1:${PORT}/")
thread = threading.Thread(target=op, daemon=True)
thread.start()
app.run(host="127.0.0.1", port=${PORT}, debug=False)
PY

python run_local.py
