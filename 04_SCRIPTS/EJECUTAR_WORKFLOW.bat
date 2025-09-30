@echo off
title Equality Momentum - Workflow Automatizado
color 0A

REM Cambiar al directorio del script
cd /d "%~dp0"

REM Verificar si existe el entorno virtual
if exist "..\EM\Scripts\python.exe" (
    echo Usando entorno virtual EM...
    "..\EM\Scripts\python.exe" ejecutar_workflow.py
) else (
    REM Verificar si Python está disponible globalmente
    python --version >nul 2>&1
    if errorlevel 1 (
        echo.
        echo ==========================================
        echo   ERROR: Python no encontrado
        echo ==========================================
        echo.
        echo Por favor instale Python desde:
        echo https://python.org
        echo.
        echo Asegurese de marcar "Add Python to PATH"
        echo durante la instalacion.
        echo.
        echo O active el entorno virtual EM si existe.
        echo.
        pause
        exit /b 1
    )
    
    REM Ejecutar workflow con Python global
    python ejecutar_workflow.py
)
