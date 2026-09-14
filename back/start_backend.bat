@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Ambiente virtual nao encontrado.
    echo Execute primeiro: python -m venv .venv
    echo Depois: .venv\Scripts\pip install -r requirements.txt
    pause
    exit /b 1
)
.venv\Scripts\python.exe main.py
pause
