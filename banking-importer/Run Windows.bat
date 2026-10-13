@echo off
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe (
 echo Run Install Windows.bat first.
 pause
 exit /b 1
)
.venv\Scripts\python.exe windows_app.py
if errorlevel 1 pause
