@echo off
cd /d "%~dp0"

where python >nul 2>&1
if %errorlevel% equ 0 (
    python main.py
    goto end
)

if exist "C:\Python314\python.exe" (
    "C:\Python314\python.exe" main.py
    goto end
)

echo [ERROR] Python not found in PATH or C:\Python314\python.exe!
pause

:end
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Application exited with error.
    pause
)
