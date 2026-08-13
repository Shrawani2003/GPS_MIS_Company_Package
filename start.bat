@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ============================================
echo  GPS vs MIS Fleet Dashboard - Setup and Start
echo ============================================
echo.

REM --- Check Python is installed ---
where python >nul 2>nul
if errorlevel 1 (
    echo Python is not installed on this computer, or not on PATH.
    echo Please install it first from https://python.org
    echo ^(check "Add Python to PATH" during install^), then run this file again.
    pause
    exit /b 1
)

REM --- Install dependencies (safe to run every time - skips if already present) ---
echo Checking required packages...
python -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo.
    echo Something went wrong installing packages. See the message above.
    pause
    exit /b 1
)

REM --- Set up admin login if not already done ---
if not exist ".streamlit\credentials.json" (
    echo.
    echo ---------------------------------------------
    echo  First-time setup: choose your admin login
    echo ---------------------------------------------
    set /p ADMIN_USER=Choose an admin username:
    set /p ADMIN_PASS=Choose an admin password:
    if not exist ".streamlit" mkdir ".streamlit"
    > ".streamlit\credentials.json" (
        echo {
        echo   "username": "!ADMIN_USER!",
        echo   "password": "!ADMIN_PASS!"
        echo }
    )
    echo.
    echo Saved. You will use this username/password to log in from now on.
    echo ^(To change it later, use "Change password" in the app sidebar after logging in.^)
    echo.
)

echo.
echo Starting the dashboard...
echo.
python -m streamlit run app.py

pause
