@echo off
cd /d "%~dp0"
py -3 -m venv .venv
if errorlevel 1 goto failed
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto failed
echo Installation complete. Double-click Run Windows.bat each day.
pause
exit /b 0
:failed
echo Installation failed. Install Python 3 from python.org, including the Python launcher, and retry.
pause
exit /b 1
