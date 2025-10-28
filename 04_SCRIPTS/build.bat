@echo off
REM Script de compilación para EqualityMomentum
REM Compila la aplicación usando PyInstaller

echo ========================================
echo  COMPILACION DE EQUALITYMOMENTUM
echo ========================================
echo.

REM Verificar que estamos en el directorio correcto
if not exist "app_principal.py" (
    echo ERROR: No se encuentra app_principal.py
    echo Por favor, ejecute este script desde la carpeta 04_SCRIPTS
    pause
    exit /b 1
)

REM Activar entorno virtual si existe
if exist "..\EM\Scripts\activate.bat" (
    echo Activando entorno virtual...
    call ..\EM\Scripts\activate.bat
) else (
    echo ADVERTENCIA: No se encontro el entorno virtual EM
    echo Asegurese de tener PyInstaller instalado
    echo.
)

REM Limpiar builds anteriores
echo Limpiando builds anteriores...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "__pycache__" rmdir /s /q "__pycache__"
echo.

REM Compilar con PyInstaller
echo Compilando aplicacion con PyInstaller...
echo Esto puede tomar varios minutos...
echo.

python -m PyInstaller EqualityMomentum.spec --clean

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: La compilacion fallo
    pause
    exit /b 1
)

echo.
echo ========================================
echo  COMPILACION EXITOSA
echo ========================================
echo.
echo El ejecutable se encuentra en: dist\EqualityMomentum\
echo.

REM Crear estructura de carpetas necesarias
echo Creando estructura de carpetas...
if not exist "dist\EqualityMomentum\01_DATOS_SIN_PROCESAR" mkdir "dist\EqualityMomentum\01_DATOS_SIN_PROCESAR"
if not exist "dist\EqualityMomentum\02_RESULTADOS" mkdir "dist\EqualityMomentum\02_RESULTADOS"
if not exist "dist\EqualityMomentum\03_LOGS" mkdir "dist\EqualityMomentum\03_LOGS"
if not exist "dist\EqualityMomentum\05_INFORMES" mkdir "dist\EqualityMomentum\05_INFORMES"

REM Copiar archivos adicionales
echo Copiando archivos adicionales...
copy "config.json" "dist\EqualityMomentum\" > nul 2>&1
copy "..\version.json" "dist\EqualityMomentum\" > nul 2>&1

echo.
echo Carpetas creadas correctamente
echo.
echo Para probar la aplicacion, ejecute:
echo   dist\EqualityMomentum\EqualityMomentum.exe
echo.
pause
