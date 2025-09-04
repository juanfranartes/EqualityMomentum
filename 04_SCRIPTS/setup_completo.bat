@echo off
title Setup Completo - Equality Momentum
color 0B

echo.
echo ==========================================
echo     CONFIGURACIÓN COMPLETA DEL SISTEMA
echo          EQUALITY MOMENTUM
echo ==========================================
echo.

REM Crear estructura de carpetas
echo Creando estructura de carpetas...

if not exist "..\01_DATOS_SIN_PROCESAR" mkdir "..\01_DATOS_SIN_PROCESAR"
if not exist "..\02_RESULTADOS" mkdir "..\02_RESULTADOS"
if not exist "..\03_LOGS" mkdir "..\03_LOGS"

echo ✓ Carpetas creadas

REM Instalar dependencias
echo.
echo Instalando dependencias de Python...
pip install -r requirements.txt

echo ✓ Dependencias instaladas

REM Crear ejecutable
echo.
echo Creando ejecutable...
call crear_ejecutable.bat

echo.
echo ==========================================
echo        ¡CONFIGURACIÓN COMPLETADA!
echo ==========================================
echo.
echo Estructura final:
echo   📁 01_DATOS_SIN_PROCESAR  (Archivos Excel de entrada)
echo   📁 02_RESULTADOS          (Reportes generados)
echo   📁 03_LOGS               (Logs de procesamiento)
echo   📁 03_SCRIPTS            (Scripts del sistema)
echo   🔧 PROCESADOR_REGISTROS.exe (Ejecutable principal)
echo.
echo Para usar el sistema:
echo 1. Copiar archivos Excel en 01_DATOS_SIN_PROCESAR
echo 2. Ejecutar PROCESADOR_REGISTROS.exe
echo 3. Revisar resultados en 02_RESULTADOS
echo.
pause