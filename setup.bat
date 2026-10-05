@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    python -m venv .venv
    if errorlevel 1 goto failed
)
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed
echo Setup complete. Double-click run.bat to launch.
pause
exit /b 0
:failed
echo Setup failed. Check Python 3.11+ and your network connection.
pause
exit /b 1
