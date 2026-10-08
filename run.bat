@echo off
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe (
    echo Creando entorno de Python...
    python -m venv .venv || (echo Instale Python 3.10+ desde python.org & pause & exit /b 1)
)
.venv\Scripts\python.exe -m pip install -q -r requirements.txt
start "" .venv\Scripts\pythonw.exe app.py
