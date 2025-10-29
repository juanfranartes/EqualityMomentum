# 🚀 Guía Rápida de Despliegue

## ✅ Checklist Pre-Despliegue

Antes de desplegar, verifica que tienes:

- [x] Código completo en la rama `main`
- [x] `streamlit_app.py` en la raíz del repositorio
- [x] `requirements.txt` actualizado en la raíz
- [x] Carpeta `core/` con módulos de procesamiento
- [x] Carpeta `.streamlit/` con `config.toml`
- [x] `.gitignore` configurado correctamente

## 📦 Paso 1: Preparar el Repositorio

```bash
# 1. Verificar que todos los archivos nuevos estén añadidos
git status

# 2. Añadir archivos nuevos
git add streamlit_app.py
git add core/
git add requirements.txt
git add .streamlit/
git add README_WEB.md
git add DESPLIEGUE.md

# 3. Commit
git commit -m "Release: EqualityMomentum Web v2.0 - Versión Streamlit"

# 4. Push a GitHub
git push origin main
```

## ☁️ Paso 2: Desplegar en Streamlit Cloud

### Opción A: Desde la Web (Recomendado)

1. **Ir a Streamlit Cloud**
   - URL: https://share.streamlit.io
   - Haz clic en "Sign in" → Selecciona GitHub

2. **Crear Nueva App**
   - Haz clic en **"New app"** (botón arriba a la derecha)

3. **Configurar el Despliegue**
   - **Repository**: Selecciona `TU_USUARIO/EqualityMomentum`
   - **Branch**: `main`
   - **Main file path**: `streamlit_app.py`
   - **App URL** (opcional): Puedes personalizar la URL

4. **Deploy**
   - Haz clic en **"Deploy!"**
   - Espera 2-3 minutos mientras se instalan las dependencias

5. **¡Listo!**
   - URL de tu app: `https://equalitymomentum.streamlit.app` (o la que personalizaste)

### Opción B: Con Streamlit CLI (Avanzado)

```bash
# Instalar Streamlit CLI
pip install streamlit

# Desplegar desde terminal
streamlit deploy streamlit_app.py
```

## 🔄 Auto-Actualización

Una vez desplegado, **cada `git push` actualizará automáticamente la web**:

```bash
# 1. Hacer cambios en el código
# ... editar archivos ...

# 2. Commit y push
git add .
git commit -m "Fix: corrección de bug en procesamiento"
git push origin main

# 3. ✅ Streamlit Cloud detecta el push y redespliegue automáticamente
# Tiempo de actualización: ~1-2 minutos
```

## 🎛️ Configuración Avanzada (Opcional)

### Variables de Entorno / Secrets

Si necesitas configurar contraseñas u otros secretos:

1. En el panel de Streamlit Cloud, ve a **Settings** → **Secrets**

2. Añade tus secretos en formato TOML:

```toml
[triodos]
password = "Triodos2025"

[general]
max_upload_size_mb = 50
```

3. En tu código, accede a ellos:

```python
import streamlit as st

password = st.secrets["triodos"]["password"]
```

### Recursos de la App

Por defecto (plan gratuito):
- **RAM**: 1 GB
- **CPU**: 1 core compartido
- **Apps**: 3 simultáneas
- **Inactividad**: App se suspende tras 7 días sin uso

Si necesitas más recursos:
- Upgrade a **Streamlit Cloud Team** ($20/mes por usuario)
- O desplegar en Render/Railway/Hugging Face

## 🧪 Paso 3: Probar la Aplicación

### Test Manual

1. **Abre la URL de tu app**

2. **Prueba con archivo General**:
   - Selecciona "General"
   - Sube un archivo Excel de prueba
   - Haz clic en "Procesar Archivo"
   - Verifica que se genere el Excel y Word
   - Descarga ambos archivos y ábrelos

3. **Prueba con archivo Triodos** (si aplica):
   - Selecciona "Triodos"
   - Introduce la contraseña
   - Sube el archivo
   - Procesa y descarga

4. **Verifica privacidad**:
   - Cierra la ventana
   - Abre de nuevo la app
   - Verifica que no hay datos previos (sesión limpia)

### Verificación de Logs

En el panel de Streamlit Cloud:
- Ve a **Manage app** → **Logs**
- Verifica que no hay errores críticos
- Comprueba que el procesamiento funciona correctamente

## ❌ Solución de Problemas Comunes

### Error: "Module not found"

**Causa**: Falta una dependencia en `requirements.txt`

**Solución**:
```bash
# Añadir la dependencia faltante
echo "nombre-paquete>=version" >> requirements.txt
git add requirements.txt
git commit -m "Fix: añadir dependencia faltante"
git push origin main
```

### Error: "File not found: streamlit_app.py"

**Causa**: El archivo no está en la raíz del repositorio

**Solución**:
- Verifica que `streamlit_app.py` esté en la raíz (no en subcarpetas)
- En Streamlit Cloud, verifica que el **Main file path** sea correcto

### La app se queda "Loading" indefinidamente

**Causa**: Error en el código que impide la carga

**Solución**:
1. Revisa los logs en Streamlit Cloud
2. Prueba localmente: `streamlit run streamlit_app.py`
3. Corrige el error y haz push

### "Memory limit exceeded"

**Causa**: Archivo demasiado grande o uso excesivo de RAM

**Solución**:
- Reduce el tamaño del archivo de entrada
- Optimiza el código para usar menos memoria
- Considera upgrade a plan de pago

## 🔐 Seguridad y Privacidad

### Verificación de Privacidad

Confirma que tu despliegue cumple con:

1. **No almacenamiento de archivos**:
   - Los archivos solo existen en `st.session_state` (memoria)
   - Se eliminan al cerrar la sesión

2. **Sin logs sensibles**:
   - Revisa los logs de Streamlit Cloud
   - No debe haber datos de empleados en los logs

3. **HTTPS habilitado**:
   - Streamlit Cloud usa HTTPS automáticamente
   - Verifica que la URL empieza con `https://`

4. **Sin tracking**:
   - En `.streamlit/config.toml` está `gatherUsageStats = false`

## 📊 Monitoreo Post-Despliegue

### Métricas a Vigilar

En el panel de Streamlit Cloud:

1. **Uptime**: Debe ser >99%
2. **Response time**: <5 segundos para carga inicial
3. **Errores**: Revisar semanalmente los logs
4. **Uso de recursos**: RAM y CPU no deben estar al 100% constantemente

### Notificaciones

Configura notificaciones en Streamlit Cloud:
- Ve a **Settings** → **Notifications**
- Activa alertas para:
  - App down
  - Deploy failed
  - High error rate

## 🎉 ¡Despliegue Completado!

Tu aplicación web está ahora:

- ✅ **Online 24/7** en una URL pública
- ✅ **Auto-actualizada** con cada git push
- ✅ **Segura** con HTTPS
- ✅ **Sin almacenamiento** de datos sensibles
- ✅ **Lista para usar** por tu equipo

### Próximos Pasos

1. **Comparte la URL** con tu equipo
2. **Crea un bookmark** en el navegador
3. **Documenta el proceso** para usuarios finales
4. **Configura monitoreo** si es crítico

---

## 📞 Soporte

Si tienes problemas durante el despliegue:

1. Revisa esta guía completa
2. Consulta los logs en Streamlit Cloud
3. Verifica el [README_WEB.md](README_WEB.md)
4. Contacta con el equipo de desarrollo

**¡Éxito con tu despliegue! 🚀**
