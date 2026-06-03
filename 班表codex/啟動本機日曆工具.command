#!/bin/zsh
cd "$(dirname "$0")" || exit 1

python3 -c "import openpyxl" >/dev/null 2>&1
if [ $? -ne 0 ]; then
  python3 -m pip install -r requirements.txt
fi

python3 local_calendar_app.py
