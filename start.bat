@echo off
chcp 65001 >nul
cd /d "%~dp0"

:: 1. 優先使用專案指定之 C:\Python314\python.exe (具備完整相依套件且免疫 Windows 11 SAC 限制)
set "PY_CMD="
if exist "C:\Python314\python.exe" (
    set "PY_CMD=C:\Python314\python.exe"
) else (
    where python >nul 2>&1
    if %errorlevel% equ 0 (
        set "PY_CMD=python"
    )
)

if "%PY_CMD%"=="" (
    echo [ERROR] 找不到 Python 執行環境！請確認已安裝 Python 3.10+。
    pause
    exit /b 1
)

:: 2. 自動檢測必要相依套件，若缺少則自動補齊
"%PY_CMD%" -c "import xlsxwriter, openpyxl, customtkinter, paramiko" >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] 偵測到缺少相依套件，正在自動安裝 (請稍候)...
    "%PY_CMD%" -m pip install -r requirements.txt
)

:: 3. 啟動主程式
"%PY_CMD%" main.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] 程式異常結束 (錯誤代碼: %errorlevel%)。
    pause
)
