# EqualityMomentum Web 🌐

Aplicación web para procesamiento de Registros Retributivos, desarrollada con Streamlit. **Sin almacenamiento de datos** - Todo el procesamiento se realiza en memoria.

## 🌟 Características

- ✅ **Procesamiento en memoria**: Sin base de datos, sin almacenamiento persistente
- ✅ **Privacidad total**: Los archivos se eliminan automáticamente tras el procesamiento
- ✅ **Interfaz simple**: Drag & drop para subir archivos
- ✅ **Dos formatos**: Soporte para formato General y Triodos
- ✅ **Descarga directa**: Excel procesado e Informe Word con gráficos
- ✅ **Auto-actualización**: Push a GitHub → Deploy automático

## 📋 Requisitos

- Python 3.8+
- Dependencias listadas en `requirements.txt`

## 🚀 Instalación Local

### 1. Clonar el repositorio

```bash
git clone https://github.com/TU_USUARIO/EqualityMomentum.git
cd EqualityMomentum
```

### 2. Crear entorno virtual

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Linux/Mac
source venv/bin/activate
```

### 3. Instalar dependencias

```bash
pip install -r requirements.txt
```

### 4. Ejecutar la aplicación

```bash
streamlit run streamlit_app.py
```

La aplicación estará disponible en: `http://localhost:8501`

## ☁️ Despliegue en Streamlit Cloud (GRATIS)

### Paso 1: Preparar el repositorio

1. Asegúrate de que todos los cambios estén commiteados:

```bash
git add .
git commit -m "Release: EqualityMomentum Web v2.0"
git push origin main
```

### Paso 2: Configurar Streamlit Cloud

1. Ve a [share.streamlit.io](https://share.streamlit.io)
2. Inicia sesión con tu cuenta de GitHub
3. Haz clic en **"New app"**
4. Selecciona:
   - **Repository**: `TU_USUARIO/EqualityMomentum`
   - **Branch**: `main`
   - **Main file path**: `streamlit_app.py`
5. Haz clic en **"Deploy"**

### Paso 3: Esperar el despliegue

- El proceso tarda 2-3 minutos
- Streamlit Cloud instalará automáticamente las dependencias de `requirements.txt`
- Una vez completado, recibirás una URL pública (ej: `https://equalitymomentum.streamlit.app`)

### Auto-actualización

Cada vez que hagas `git push` a la rama `main`, Streamlit Cloud **actualizará automáticamente** la aplicación web. No necesitas hacer nada más.

```bash
# Hacer cambios en el código
git add .
git commit -m "Mejora: descripción del cambio"
git push origin main

# ✅ La web se actualiza automáticamente en 1-2 minutos
```

## 🎯 Uso de la Aplicación

### 1. Seleccionar tipo de archivo

- **General**: Para archivos Excel estándar con hoja "BASE GENERAL"
- **Triodos**: Para archivos de Triodos Bank (protegidos con contraseña)

### 2. Subir archivo

- Arrastra el archivo Excel o haz clic para seleccionar
- Tamaño máximo: **50 MB**
- Formatos: `.xlsx`, `.xls`

### 3. Procesar datos

- Haz clic en **"Procesar Archivo"**
- Espera 10-30 segundos (dependiendo del tamaño)

### 4. Descargar resultados

- **Excel procesado**: Datos con columnas equiparadas
- **Informe Word**: Informe completo con gráficos y tablas

## 🔒 Privacidad y Seguridad

### ¿Qué NO se guarda?

- ❌ Archivos subidos
- ❌ Archivos procesados
- ❌ Datos de empleados
- ❌ Logs con información sensible
- ❌ Cookies de tracking

### ¿Cómo funciona?

1. El usuario sube un archivo → se carga en **memoria RAM**
2. Se procesa en **memoria** (sin escribir a disco)
3. Se generan los resultados en **memoria**
4. El usuario descarga los archivos
5. Al cerrar la ventana → **todos los datos se eliminan**

### Cumplimiento GDPR/LOPD

- ✅ No almacenamiento de datos personales
- ✅ Procesamiento en memoria temporal
- ✅ Limpieza automática de sesiones
- ✅ Sin transferencia a terceros
- ✅ Filtros de privacidad en informes (oculta datos cuando hay 1 solo empleado)

## 🛠️ Arquitectura Técnica

```
EqualityMomentum/
├── streamlit_app.py          # Interfaz web principal
├── core/
│   ├── __init__.py
│   ├── procesador.py         # Procesamiento de datos (BytesIO)
│   ├── generador.py          # Generación de informes Word (BytesIO)
│   └── utils.py              # Utilidades comunes
├── templates/                 # Plantillas Word (opcional)
├── .streamlit/
│   └── config.toml           # Configuración de Streamlit
├── requirements.txt          # Dependencias Python
└── README_WEB.md             # Esta documentación
```

## 📊 Tecnologías Utilizadas

- **Streamlit**: Framework web
- **Pandas**: Procesamiento de datos
- **OpenPyXL**: Lectura/escritura de Excel
- **Python-docx**: Generación de documentos Word
- **Matplotlib/Seaborn**: Generación de gráficos
- **msoffcrypto-tool**: Desencriptación de archivos protegidos

## 🐛 Solución de Problemas

### La app no inicia en local

```bash
# Verificar instalación de Streamlit
streamlit --version

# Reinstalar dependencias
pip install -r requirements.txt --upgrade
```

### Error al procesar archivo Triodos

- Verifica que la contraseña sea correcta (por defecto: `Triodos2025`)
- Asegúrate de que el archivo tenga la hoja "BASE GENERAL"

### El despliegue en Streamlit Cloud falla

1. Verifica que `requirements.txt` esté en la raíz del repositorio
2. Comprueba que `streamlit_app.py` esté en la raíz
3. Revisa los logs en el panel de Streamlit Cloud

### Archivo demasiado grande

- Límite en Streamlit Cloud (gratuito): **50 MB**
- Si necesitas procesar archivos más grandes, considera:
  - Dividir el archivo en partes
  - Usar Streamlit Cloud Team (plan de pago)
  - Desplegar en otro servicio (Render, Hugging Face)

## 🔄 Alternativas de Despliegue

### Opción 2: Hugging Face Spaces

1. Crea un Space en [huggingface.co/spaces](https://huggingface.co/spaces)
2. Selecciona **Streamlit** como SDK
3. Conecta tu repositorio de GitHub
4. Auto-deploy activado

### Opción 3: Render.com

1. Crea una cuenta en [render.com](https://render.com)
2. Nuevo **Web Service**
3. Conecta GitHub
4. Comando de inicio: `streamlit run streamlit_app.py --server.port=$PORT --server.address=0.0.0.0`

## 📝 Changelog

### v2.0.0 (2025-01-XX)
- ✨ Versión web completa
- ✅ Procesamiento en memoria (sin almacenamiento)
- ✅ Interfaz Streamlit moderna
- ✅ Soporte para archivos General y Triodos
- ✅ Generación de informes Word con gráficos
- ✅ Auto-despliegue desde GitHub

### v1.0.1 (Anterior)
- Aplicación de escritorio (PyInstaller)
- Interfaz tkinter

## 🤝 Contribuir

1. Fork el proyecto
2. Crea una rama para tu feature (`git checkout -b feature/nueva-funcionalidad`)
3. Commit tus cambios (`git commit -m 'Añade nueva funcionalidad'`)
4. Push a la rama (`git push origin feature/nueva-funcionalidad`)
5. Abre un Pull Request

## 📄 Licencia

Este proyecto es privado y de uso interno.

## 📧 Soporte

Para preguntas o problemas, contacta con el equipo de desarrollo.

---

**EqualityMomentum Web v2.0** | Registro Retributivo | 🔒 Sin almacenamiento de datos
