@echo off
setlocal
cd /d "%~dp0"
set "PY_CMD=python"
python --version >nul 2>&1
if %errorlevel% neq 0 (
    if exist "C:\Python314\python.exe" (
        set "PY_CMD=C:\Python314\python.exe"
    ) else (
        echo [ERROR] Python not found!
        pause
        exit /b 1
    )
)
%PY_CMD% build_exe.py
pause
