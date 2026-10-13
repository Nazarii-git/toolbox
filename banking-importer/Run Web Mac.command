#!/bin/bash
set -e
cd "$(dirname "$0")"
if [ ! -x .venv/bin/python ]; then
 echo 'Run Install Mac.command first.'
 read -r -p 'Press Enter to close.'
 exit 1
fi
.venv/bin/python web_app.py
