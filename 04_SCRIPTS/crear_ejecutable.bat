@echo off
title Equality Momentum - Creador de Ejecutable
color 0A

echo.
echo ==========================================
echo   CREANDO EJECUTABLE - EQUALITY MOMENTUM
echo ==========================================
echo.

REM Verificar si Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no encontrado
    echo Por favor instale Python desde https://python.org
    pause
    exit /b 1
)

echo ✓ Python encontrado

REM Instalar pyinstaller si no está instalado
echo.
echo Instalando/actualizando PyInstaller...
pip install --upgrade pyinstaller

REM Verificar si el script existe
if not exist "procesador_registro_retributivo.py" (
    echo.
    echo ERROR: No se encontró procesador_registro_retributivo.py
    echo Asegúrese de que el archivo está en esta carpeta
    pause
    exit /b 1
)

echo ✓ Script encontrado

REM Crear ejecutable
echo.
echo Creando ejecutable...
echo (Esto puede tardar varios minutos...)

pyinstaller --onefile ^
    --windowed ^
    --name="PROCESADOR_REGISTROS" ^
    --icon=logo.ico ^
    --distpath=../ ^
    --workpath=build ^
    --specpath=specs ^
    --clean ^
    --noconfirm ^
    procesador_registro_retributivo.py

REM Verificar si se creó correctamente
if exist "..\PROCESADOR_REGISTROS.exe" (
    echo.
    echo ==========================================
    echo   ¡EJECUTABLE CREADO EXITOSAMENTE! ✓
    echo ==========================================
    echo.
    echo Ubicación: PROCESADOR_REGISTROS.exe
    echo.
    echo El sistema está listo para usar:
    echo 1. Poner archivos Excel en: 01_DATOS_SIN_PROCESAR
    echo 2. Ejecutar: PROCESADOR_REGISTROS.exe
    echo 3. Ver resultados en: 02_RESULTADOS
    echo.
) else (
    echo.
    echo ERROR: No se pudo crear el ejecutable
    echo Revisar los mensajes de error arriba
)

echo.
pause