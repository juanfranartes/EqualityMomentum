@echo off
REM Lanzador de EqualityMomentum
REM Ejecuta la aplicación principal

echo ========================================
echo      EQUALITYMOMENTUM
echo ========================================
echo.

REM Verificar que estamos en el directorio correcto
if not exist "04_SCRIPTS\app_principal.py" (
    echo ERROR: No se encuentra app_principal.py
    echo Por favor, ejecute este script desde la raiz del proyecto
    pause
    exit /b 1
)

REM Activar entorno virtual si existe
if exist "EM\Scripts\activate.bat" (
    echo Activando entorno virtual...
    call EM\Scripts\activate.bat
) else (
    echo ADVERTENCIA: No se encontro el entorno virtual EM
    echo Usando Python del sistema
    echo.
)

REM Ejecutar aplicación
echo Iniciando EqualityMomentum...
echo.

cd 04_SCRIPTS
python app_principal.py

REM Si hubo error
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: La aplicacion termino con errores
    echo Revise los logs en 03_LOGS
    pause
)
