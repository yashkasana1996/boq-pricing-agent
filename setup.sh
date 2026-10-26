#!/usr/bin/env bash
# BOQ Agent - one-time setup (Mac/Linux). Run:  bash setup.sh
set -e
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null; then
  echo "Python 3 not found. Install from https://www.python.org/downloads/ then rerun."
  exit 1
fi

echo "Creating isolated environment (.venv)..."
python3 -m venv .venv
source .venv/bin/activate

echo "Installing dependencies (needs internet)..."
pip install --quiet -r requirements.txt

echo "Running the full self-test against the sample tender files..."
python -m pytest tests -q

echo
echo "If the line above says '8 passed', setup is DONE."
echo "Next:  bash run-app.sh"
