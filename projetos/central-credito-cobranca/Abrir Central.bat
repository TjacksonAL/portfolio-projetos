@echo off
cd /d "%~dp0"
where python >nul 2>nul
if %errorlevel%==0 (
    python central.py
) else (
    py central.py
)
pause
