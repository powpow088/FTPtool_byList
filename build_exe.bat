@echo off
setlocal
cd /d "%~dp0"

if exist "C:\Python314\python.exe" (
    "C:\Python314\python.exe" build_exe.py
    goto end
)

python build_exe.py

:end
pause
