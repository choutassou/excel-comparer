@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo First run setup.bat to install dependencies.
    pause
    exit /b 1
)
".venv\Scripts\python.exe" -m source %*
if errorlevel 1 pause
