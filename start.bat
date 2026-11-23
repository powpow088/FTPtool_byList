@echo off
cd /d "%~dp0"

if exist "C:\Python314\python.exe" (
    "C:\Python314\python.exe" main.py
    goto end
)

python main.py

:end
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Application exited with error code %errorlevel%.
    pause
)
