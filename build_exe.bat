@echo off
setlocal
title SFTP Batch Downloader - Build EXE

cd /d "%~dp0"

echo ==================================================
echo   Building SFTP Batch Downloader Standalone EXE
echo ==================================================
echo.

set "PY_CMD=python"
python --version >nul 2>&1
if %errorlevel% neq 0 (
    if exist "C:\Python314\python.exe" (
        set "PY_CMD=C:\Python314\python.exe"
    ) else (
        echo [ERROR] Python not found! Please check PATH.
        pause
        exit /b 1
    )
)

echo [1/4] Checking PyInstaller...
%PY_CMD% -m pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing pyinstaller...
    %PY_CMD% -m pip install pyinstaller
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to install pyinstaller.
        pause
        exit /b 1
    )
)

echo [2/4] Cleaning previous build cache...
if exist "build" rd /s /q "build"
if exist "dist\SFTP批次下載工具.exe" del /f /q "dist\SFTP批次下載工具.exe"

echo [3/4] Compiling with PyInstaller...
%PY_CMD% -m PyInstaller --clean SFTPTool.spec

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Build failed! Check messages above.
    pause
    exit /b 1
)

echo.
echo ==================================================
echo   [SUCCESS] Standalone EXE created successfully!
echo   Location: dist\SFTP批次下載工具.exe
echo ==================================================
echo.
pause
