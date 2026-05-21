@echo off
setlocal EnableExtensions

set "ROOT=%~dp0"
cd /d "%ROOT%"

py -3 -m pip install -r requirements.txt
if errorlevel 1 (
  python -m pip install -r requirements.txt
)

py -3 gui_app.py
if errorlevel 1 (
  python gui_app.py
)
