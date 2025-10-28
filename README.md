# EqualityMomentum

Sistema profesional de gestión de registros retributivos con análisis de brechas salariales e informes automatizados.

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![Python](https://img.shields.io/badge/python-3.8+-green)
![License](https://img.shields.io/badge/license-MIT-orange)

## 📋 Descripción

**EqualityMomentum** es una aplicación de escritorio que permite:

- ✅ Procesar datos de registros retributivos (formatos estándar y Triodos Bank)
- 📊 Generar informes profesionales en Word y PDF con análisis estadístico
- 📈 Calcular automáticamente brechas salariales por sexo
- 🔒 Aplicar filtros de privacidad LOPD/RGPD
- 🎨 Interfaz corporativa intuitiva con identidad visual personalizada
- 🔄 Sistema de actualizaciones automáticas

## 🎨 Identidad Corporativa

- **Colores principales:**
  - Azul: `#1f3c89`
  - Naranja: `#ff5c39`
  - Blanco: `#ffffff`

- **Tipografías:**
  - Títulos: Lusitana (48px)
  - Texto: Work Sans (18px)

## 💻 Requisitos del Sistema

- **Sistema Operativo:** Windows 10 o superior
- **Espacio en disco:** 500 MB mínimo
- **RAM:** 4 GB mínimo (8 GB recomendado)
- **Resolución de pantalla:** 1280x720 mínimo

## 🚀 Instalación

### Opción 1: Instalador (Recomendado)

1. Descargue el instalador más reciente desde [Releases](https://github.com/TU_USUARIO/EqualityMomentum/releases)
2. Ejecute `EqualityMomentum_Setup_vX.X.X.exe`
3. Siga las instrucciones del asistente de instalación
4. La aplicación se instalará en `C:\Program Files\EqualityMomentum`
5. Se crearán carpetas de trabajo en `Documentos\EqualityMomentum`

### Opción 2: Desde el Código Fuente (Desarrolladores)

```bash
# Clonar el repositorio
git clone https://github.com/TU_USUARIO/EqualityMomentum.git
cd EqualityMomentum

# Crear entorno virtual
python -m venv EM

# Activar entorno virtual
EM\Scripts\activate

# Instalar dependencias
pip install -r 04_SCRIPTS\requirements.txt

# Ejecutar aplicación
cd 04_SCRIPTS
python app_principal.py
```

## 📖 Uso

### 1. Procesar Datos

1. Abra la aplicación **EqualityMomentum**
2. Haga clic en **"PROCESAR DATOS"**
3. Seleccione el archivo Excel con los datos sin procesar
4. Elija el tipo de procesamiento:
   - **Estándar**: Para archivos con estructura maestro
   - **Triodos**: Para archivos de Triodos Bank (protegidos con contraseña)
5. Seleccione la carpeta donde guardar los resultados
6. Haga clic en **"Procesar"**

**Resultado:** Se generará un archivo Excel procesado en la carpeta de resultados.

### 2. Generar Informe

1. Haga clic en **"GENERAR INFORME"**
2. Seleccione el archivo Excel procesado
3. Elija el tipo de informe:
   - **CONSOLIDADO**: Informe completo con todos los análisis
   - **PROMEDIO**: Solo análisis con promedios
   - **MEDIANA**: Solo análisis con medianas
   - **COMPLEMENTOS**: Solo análisis de complementos
4. Seleccione la carpeta donde guardar el informe
5. Haga clic en **"Generar Informe"**

**Resultado:** Se generarán dos archivos (`.docx` y `.pdf`) en la carpeta de informes.

## 🔄 Actualizaciones

La aplicación verifica automáticamente si hay actualizaciones disponibles al iniciarse.

**Para buscar actualizaciones manualmente:**

1. Abra la aplicación
2. Haga clic en **"Buscar actualizaciones"**
3. Si hay una actualización disponible, siga las instrucciones

**Para desarrolladores:**

```bash
# Crear una nueva versión
cd 04_SCRIPTS
python build_release.py
```

## 📁 Estructura de Carpetas

```
EqualityMomentum/
├── 00_DOCUMENTACION/      # Documentación y recursos gráficos
│   ├── isotipo.jpg        # Logo corporativo
│   └── ...
├── 01_DATOS_SIN_PROCESAR/ # Archivos Excel de entrada
├── 02_RESULTADOS/         # Archivos Excel procesados
├── 03_LOGS/               # Archivos de log
├── 04_SCRIPTS/            # Código fuente
│   ├── app_principal.py           # Aplicación principal
│   ├── procesar_datos.py          # Procesador estándar
│   ├── procesar_datos_triodos.py  # Procesador Triodos
│   ├── generar_informe_optimizado.py  # Generador de informes
│   ├── logger_manager.py          # Sistema de logging
│   ├── updater.py                 # Sistema de actualizaciones
│   ├── config.json                # Configuración
│   └── requirements.txt           # Dependencias
├── 05_INFORMES/           # Informes generados (.docx y .pdf)
└── README.md              # Este archivo
```

## 🛠️ Desarrollo

### Compilar la Aplicación

```bash
cd 04_SCRIPTS

# Compilar con PyInstaller
build.bat

# O crear un release completo (incrementa versión, compila y crea instalador)
python build_release.py
```

### Crear Instalador

```bash
# Requisito: Inno Setup instalado
# Ejecutar desde 04_SCRIPTS:

# Opción 1: Script automático (recomendado)
python build_release.py

# Opción 2: Manual
# 1. Compilar con build.bat
# 2. Abrir installer.iss con Inno Setup
# 3. Compilar
```

### Estructura del Código

- **app_principal.py**: Interfaz principal con navegación
- **logger_manager.py**: Sistema centralizado de logging con captura de errores
- **updater.py**: Sistema de actualización automática desde GitHub
- **procesar_datos.py**: Lógica de procesamiento de datos estándar
- **procesar_datos_triodos.py**: Lógica de procesamiento para Triodos Bank
- **generar_informe_optimizado.py**: Generación de informes con análisis estadístico

## 🐛 Solución de Problemas

### La aplicación no inicia

1. Verifique que su sistema cumple con los requisitos mínimos
2. Reinstale la aplicación
3. Revise los logs en `Documentos\EqualityMomentum\Logs`

### Errores durante el procesamiento

1. Verifique que el archivo Excel tiene el formato correcto
2. Revise los logs para ver el error específico
3. Envíe el reporte de error al equipo de desarrollo

### No se pueden generar informes PDF

1. Verifique que se genera correctamente el archivo .docx
2. El error puede estar relacionado con la conversión a PDF
3. Revise los logs para más detalles

### ¿Dónde encuentro los logs?

**Desde la aplicación:**
- Haga clic en "Abrir carpeta de logs"

**Manualmente:**
- `C:\Users\[TU_USUARIO]\Documents\EqualityMomentum\Logs`
- O en la carpeta del proyecto: `03_LOGS`

## 📞 Soporte

Si encuentra problemas:

1. Revise la carpeta de logs
2. Busque archivos `ERROR_REPORT_*.txt` y `ERROR_REPORT_*.json`
3. Envíe estos archivos al equipo de desarrollo

## 📝 Changelog

### v1.0.0 (2025-10-28)
- ✨ Versión inicial del sistema
- ✅ Procesamiento de datos estándar y Triodos
- 📊 Generación de informes en Word y PDF
- 🎨 Interfaz corporativa con identidad visual
- 🔄 Sistema de actualizaciones automáticas
- 📝 Sistema de logging centralizado
- 🔒 Filtros de privacidad LOPD/RGPD

## 📄 Licencia

Este proyecto está bajo la Licencia MIT. Ver archivo `LICENSE` para más detalles.

## 👥 Autores

**EqualityMomentum Team**

---

**Desarrollado con ❤️ para promover la igualdad salarial**
