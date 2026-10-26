#!/usr/bin/env bash
# BOQ Agent - start the web app (Mac/Linux). Run:  bash run-app.sh
cd "$(dirname "$0")"
if [ ! -f .venv/bin/activate ]; then
  echo "Run 'bash setup.sh' first (one time only)."
  exit 1
fi
source .venv/bin/activate
( sleep 2 && open "http://127.0.0.1:8000" 2>/dev/null || xdg-open "http://127.0.0.1:8000" 2>/dev/null ) &
python -m uvicorn boq.webapp:app --port 8000
