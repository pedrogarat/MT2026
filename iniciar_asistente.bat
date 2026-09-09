@echo off
title Asistente EBSS - Estudo Basico de Seguridade e Saude
cd /d "%~dp0"

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERRO] Non se atopou Python instalado neste ordenador.
    echo Por favor, instala Python 3 dende https://www.python.org/downloads/
    pause
    exit /b 1
)

set "PYTHON_EXE=python"

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -c "import sys" >nul 2>nul
    if %errorlevel% equ 0 (
        set "PYTHON_EXE=.venv\Scripts\python.exe"
    ) else (
        echo [INFO] Detectouse unha contorna .venv de outro ordenador. Recreando...
        rmdir /s /q .venv >nul 2>nul
    )
)

if not exist ".venv\Scripts\python.exe" (
    echo [INFO] Configurando contorna de Python para este ordenador...
    where uv >nul 2>nul
    if %errorlevel% equ 0 (
        uv venv .venv >nul 2>nul
        uv pip install --system-certs --python ".venv\Scripts\python.exe" -r requirements.txt >nul 2>nul
    ) else (
        python -m venv .venv
        ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    )
    if exist ".venv\Scripts\python.exe" (
        set "PYTHON_EXE=.venv\Scripts\python.exe"
    )
)

echo [INFO] Executando servidor web do Asistente EBSS...
"%PYTHON_EXE%" main.py
pause
