#!/bin/bash
# Double-click this file (or run in Terminal) to open the CCS Worksheet Update tool.
cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
  echo "Creating virtual environment..."
  python3 -m venv .venv
  .venv/bin/pip install -r requirements.txt
fi

echo "Opening CCS Worksheet Update in your browser..."
.venv/bin/streamlit run app.py
