# 🚀 GUÍA DE INSTALACIÓN Y USO - EqualityMomentum Web

## 📋 Requisitos Previos

- Python 3.8 o superior
- Entorno virtual `EM` (ya existe en el proyecto)

---

## ⚙️ INSTALACIÓN

### 1. Activar el Entorno Virtual

```powershell
# En Windows PowerShell
.\EM\Scripts\Activate.ps1

# En Windows CMD
.\EM\Scripts\activate.bat
```

### 2. Instalar Streamlit (si no está instalado)

```powershell
pip install streamlit
```

O instalar todas las dependencias:

```powershell
pip install -r requirements.txt
```

### 3. Verificar Instalación

```powershell
python test_streamlit.py
```

Deberías ver:
```
✅ Python 3.x
✅ streamlit x.x.x
✅ pandas x.x.x
✅ numpy x.x.x
✅ python-docx instalado
✅ matplotlib x.x.x
✅ Todos los scripts de 04_SCRIPTS
```

---

## 🎯 EJECUCIÓN

### Iniciar la Aplicación Web

```powershell
streamlit run streamlit_app.py
```

La aplicación se abrirá automáticamente en tu navegador en:
```
http://localhost:8501
```

---

## 📖 CÓMO USAR LA APLICACIÓN

### Paso 1: Seleccionar Tipo de Archivo

Elige el tipo de archivo que vas a procesar:

- **General**: Para archivos Excel estándar con hoja "BASE GENERAL"
- **Triodos**: Para archivos de Triodos Bank (protegidos con contraseña)

### Paso 2: Seleccionar Acción

Elige qué quieres hacer:

- **Procesar Datos**: Solo genera el Excel procesado
- **Generar Informe**: Solo genera el informe Word (requiere archivo ya procesado)
- **Ambas**: Procesa datos Y genera informe (opción más común)

### Paso 3: Configurar Contraseña (opcional)

Si tu archivo está protegido con contraseña:
1. Marca la casilla "¿El archivo tiene contraseña?"
2. Ingresa la contraseña (por defecto para Triodos: "Triodos2025")

### Paso 4: Subir Archivo

- Arrastra tu archivo Excel a la zona de carga
- O haz clic para seleccionarlo desde tu computadora
- Tamaño máximo: 50MB

### Paso 5: Procesar

Haz clic en el botón correspondiente:
- 🚀 "Procesar y Generar Informe" (si seleccionaste "Ambas")
- 📊 "Procesar Datos" (si solo quieres el Excel)
- 📄 "Generar Informe" (si solo quieres el Word)

### Paso 6: Descargar Resultados

Una vez completado el procesamiento:

- **Excel Procesado**: Contiene los datos con columnas equiparadas
  - Nombre: `REPORTE_[nombre_archivo]_[fecha].xlsx`
  
- **Informe Word**: Informe completo con gráficos y análisis
  - Nombre: `INFORME_[nombre_archivo]_[fecha].docx`

---

## 📂 ESTRUCTURA DE ARCHIVOS

Después del procesamiento, encontrarás:

```
EqualityMomentum/
├── 02_RESULTADOS/          # Excel procesados
│   └── REPORTE_*.xlsx
│
├── 03_LOGS/                # Logs de procesamiento
│   ├── procesamiento_YYYYMMDD.log
│   └── informe_YYYYMMDD.log
│
└── 05_INFORMES/            # Informes Word
    └── registro_retributivo_*.docx
```

---

## 🔧 SCRIPTS UTILIZADOS

La aplicación web utiliza los siguientes scripts:

### Para Procesamiento General
- **Script**: `04_SCRIPTS\procesar_datos.py`
- **Clase**: `ProcesadorRegistroRetributivo`
- **Uso**: Archivos Excel estándar

### Para Procesamiento Triodos
- **Script**: `04_SCRIPTS\procesar_datos_triodos.py`
- **Clase**: `ProcesadorTriodos`
- **Uso**: Archivos de Triodos Bank (protegidos)

### Para Generar Informes
- **Script**: `04_SCRIPTS\generar_informe_optimizado.py`
- **Clase**: `GeneradorInformeOptimizado`
- **Uso**: Generación de informes Word con gráficos

---

## 🛡️ SEGURIDAD Y PRIVACIDAD

La aplicación está diseñada para máxima privacidad:

✅ **Sin Base de Datos**: No se almacena información permanentemente

✅ **Procesamiento en Memoria**: Los datos solo existen en RAM durante la sesión

✅ **Limpieza Automática**: 
- Los archivos temporales se eliminan automáticamente
- Los datos en sesión se limpian al cerrar el navegador
- Puedes limpiar manualmente con el botón "🗑️ Limpiar Sesión"

✅ **Logs Sin Datos Personales**: Los logs no contienen información sensible

---

## 🐛 SOLUCIÓN DE PROBLEMAS

### Error: "streamlit no instalado"

```powershell
pip install streamlit
```

### Error: "Import could not be resolved"

- **Causa**: Advertencias de Pylance
- **Solución**: Ignorar, los imports funcionan en ejecución

### Error al Procesar Archivo

1. Verificar formato del archivo Excel
2. Comprobar que existe la hoja requerida:
   - Archivo General: "BASE GENERAL"
   - Archivo Triodos: formato específico de Triodos
3. Revisar logs en `03_LOGS/`

### Error al Generar Informe

1. Verificar que existe archivo procesado en `02_RESULTADOS/`
2. Comprobar que el DataFrame tiene las columnas requeridas
3. Revisar que existe la plantilla en `00_DOCUMENTACION/Registro retributivo/`

### La Aplicación No Se Abre

```powershell
# Verificar que streamlit está instalado
streamlit --version

# Reinstalar si es necesario
pip install --upgrade streamlit

# Ejecutar con puerto específico
streamlit run streamlit_app.py --server.port 8502
```

---

## 📊 FUNCIONALIDADES

### Procesamiento de Datos

- ✅ Equiparación por grupos profesionales
- ✅ Cálculo de complementos salariales
- ✅ Análisis de jornadas parciales
- ✅ Estadísticas por sexo
- ✅ Validación de datos

### Generación de Informes

- ✅ Análisis por promedio
- ✅ Análisis por mediana
- ✅ Análisis de complementos
- ✅ Gráficos de barras horizontales
- ✅ Tablas comparativas por sexo
- ✅ Formato profesional Word

---

## 💡 CONSEJOS DE USO

### Para Mejores Resultados

1. **Formato del Archivo**: Asegúrate de que tu Excel tiene el formato correcto
2. **Datos Completos**: Verifica que no falten columnas importantes
3. **Contraseña Correcta**: Si el archivo está protegido, usa la contraseña correcta
4. **Internet Activo**: Necesario para cargar las fuentes de Google Fonts

### Flujo de Trabajo Recomendado

1. **Primera vez**: Usar opción "Ambas" para generar Excel e informe
2. **Ajustes**: Si necesitas regenerar el informe, usa "Generar Informe"
3. **Múltiples archivos**: Procesar uno por uno, descargando resultados antes de continuar

---

## 🔄 ACTUALIZACIÓN

Para actualizar la aplicación:

```powershell
# Actualizar repositorio
git pull

# Actualizar dependencias
pip install -r requirements.txt --upgrade

# Verificar instalación
python test_streamlit.py
```

---

## 📞 SOPORTE

### Documentación

- **Cambios realizados**: `CAMBIOS_STREAMLIT.md`
- **Resumen**: `RESUMEN_MODIFICACIONES.md`
- **Este archivo**: `GUIA_USO.md`

### Logs

Los logs de procesamiento están en:
- `03_LOGS/procesamiento_YYYYMMDD.log`
- `03_LOGS/procesamiento_triodos_YYYYMMDD.log`
- `03_LOGS/informe_YYYYMMDD.log`

### Scripts de Verificación

```powershell
# Verificar imports
python verificar_imports.py

# Verificar compatibilidad
python test_streamlit.py
```

---

## 🎉 ¡LISTO!

Tu aplicación EqualityMomentum está lista para usar. 

**Comando rápido para empezar:**

```powershell
.\EM\Scripts\Activate.ps1
streamlit run streamlit_app.py
```

---

*Versión: 2.0 - Última actualización: 29 de octubre de 2025*
