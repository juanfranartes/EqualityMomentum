@echo off
title Verificación del Sistema
color 0E

echo.
echo ==========================================
echo       VERIFICACIÓN DEL SISTEMA
echo ==========================================
echo.

REM Verificar Python
echo Verificando Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python NO encontrado
    echo    Instalar desde: https://python.org
) else (
    python --version
    echo ✅ Python OK
)

REM Verificar pip
echo.
echo Verificando pip...
pip --version >nul 2>&1
if errorlevel 1 (
    echo ❌ pip NO encontrado
) else (
    echo ✅ pip OK
)

REM Verificar dependencias
echo.
echo Verificando dependencias...
python -c "import pandas; print('✅ pandas OK')" 2>nul || echo "❌ pandas falta"
python -c "import numpy; print('✅ numpy OK')" 2>nul || echo "❌ numpy falta"  
python -c "import openpyxl; print('✅ openpyxl OK')" 2>nul || echo "❌ openpyxl falta"
python -c "import tkinter; print('✅ tkinter OK')" 2>nul || echo "❌ tkinter falta"

REM Verificar estructura de carpetas
echo.
echo Verificando estructura de carpetas...
if exist "..\01_DATOS_SIN_PROCESAR" (echo ✅ 01_DATOS_SIN_PROCESAR) else (echo ❌ 01_DATOS_SIN_PROCESAR falta)
if exist "..\02_RESULTADOS" (echo ✅ 02_RESULTADOS) else (echo ❌ 02_RESULTADOS falta)
if exist "..\03_LOGS" (echo ✅ 03_LOGS) else (echo ❌ 03_LOGS falta)

REM Verificar archivos principales
echo.
echo Verificando archivos principales...
if exist "procesador_registro_retributivo.py" (echo ✅ Script principal) else (echo ❌ Script principal falta)
if exist "requirements.txt" (echo ✅ requirements.txt) else (echo ❌ requirements.txt falta)
if exist "..\PROCESADOR_REGISTROS.exe" (echo ✅ Ejecutable) else (echo ❌ Ejecutable falta)

echo.
echo ==========================================
echo        VERIFICACIÓN COMPLETADA
echo ==========================================
echo.
pause