@echo off
rem WIN5 launcher (double-click to start)
cd /d "%~dp0"

where pythonw >nul 2>&1
if %errorlevel%==0 (
    start "" pythonw launcher.py
    exit /b
)

python launcher.py
if errorlevel 1 pause
