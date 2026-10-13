#!/bin/bash
set -e
cd "$(dirname "$0")"
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
echo 'Installation complete. Run Run Web Mac.command each day.'
read -r -p 'Press Enter to close.'
