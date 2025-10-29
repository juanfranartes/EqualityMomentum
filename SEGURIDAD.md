# 🔒 CONFIGURACIÓN DE SEGURIDAD - NO SUBIR DATOS

# IMPORTANTE: Este archivo asegura que ningún dato sensible se sube a GitHub

# ====================================================================
# CARPETAS CON DATOS SENSIBLES (IGNORADAS POR GIT)
# ====================================================================

01_DATOS_SIN_PROCESAR/     # Archivos Excel originales del cliente
02_RESULTADOS/             # Archivos Excel procesados
03_LOGS/                   # Logs (no contienen datos personales, pero se ignoran)
05_INFORMES/               # Informes Word generados

# ====================================================================
# COMPORTAMIENTO EN PRODUCCIÓN (Streamlit Cloud)
# ====================================================================

# Cuando la aplicación se despliega en Streamlit Cloud u otro servidor:
# 
# 1. Los archivos se borran INMEDIATAMENTE después de procesarse
# 2. Solo se mantienen en memoria RAM durante la sesión del usuario
# 3. Al cerrar el navegador, TODO se elimina automáticamente
# 4. No se crean carpetas de resultados en el servidor

# ====================================================================
# COMPORTAMIENTO EN LOCAL (Tu ordenador)
# ====================================================================

# Cuando ejecutas localmente (streamlit run streamlit_app.py):
#
# 1. Se crean las carpetas temporalmente
# 2. Puedes borrarlas manualmente cuando quieras
# 3. Git NO las sube a GitHub (están en .gitignore)

# ====================================================================
# VERIFICAR QUE TODO ESTÁ SEGURO
# ====================================================================

# Ejecuta este comando para ver qué archivos se subirían a GitHub:
#   git status
#
# NUNCA deberías ver archivos .xlsx, .docx o .log en la lista

# ====================================================================
# BORRAR DATOS LOCALES MANUALMENTE
# ====================================================================

# En Windows PowerShell:
#   Remove-Item 02_RESULTADOS\* -Force
#   Remove-Item 03_LOGS\* -Force
#   Remove-Item 05_INFORMES\* -Force

# En Linux/Mac:
#   rm -f 02_RESULTADOS/*
#   rm -f 03_LOGS/*
#   rm -f 05_INFORMES/*

# ====================================================================
# GARANTÍAS DE SEGURIDAD
# ====================================================================

✅ Archivos sensibles en .gitignore
✅ Borrado automático en producción
✅ Procesamiento en memoria RAM
✅ Sin base de datos
✅ Sin almacenamiento permanente en servidor
✅ Logs sin datos personales

# ====================================================================
# SI DESPLIEGAS EN STREAMLIT CLOUD
# ====================================================================

# Los archivos se procesarán completamente en memoria y se borrarán
# inmediatamente. No quedarán en el servidor.

# El usuario:
# 1. Sube su archivo → Se procesa en RAM
# 2. Descarga el resultado → El archivo va a SU ordenador
# 3. Cierra el navegador → Todo se borra de la sesión

# TÚ (dueño del servidor):
# - NO ves los archivos
# - NO quedan guardados
# - NO hay acceso a datos sensibles

# ====================================================================
