@echo off
setlocal
cd /d "%~dp0"

set "PY_CMD=python"
if exist "C:\Python314\python.exe" set "PY_CMD=C:\Python314\python.exe"

"%PY_CMD%" -c "import xlsxwriter, python_calamine, pymssql" >nul 2>&1
if not errorlevel 1 goto run_app

echo [INFO] Detected missing dependencies. Installing from requirements.txt...
"%PY_CMD%" -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [ERROR] Failed to install packages. Please run: pip install -r requirements.txt
    pause
    exit /b 1
)

:run_app
"%PY_CMD%" main.py

if errorlevel 1 (
    echo.
    echo [ERROR] Application exited with error code %errorlevel%.
    pause
)
