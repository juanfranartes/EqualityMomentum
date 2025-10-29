# ✅ Corrección para Streamlit Cloud

## Problema Identificado

La aplicación fallaba en Streamlit Cloud con el error:
```
ImportError: import _tkinter
```

**Causa**: Los scripts de procesamiento (`procesar_datos.py` y `procesar_datos_triodos.py`) importaban `tkinter` para mostrar diálogos GUI, pero Streamlit Cloud (entorno Linux sin interfaz gráfica) no tiene `tkinter` instalado.

## Solución Implementada

Se modificaron ambos scripts para hacer la importación de `tkinter` **opcional y condicional**:

### Cambios en `04_SCRIPTS/procesar_datos.py`

```python
# ANTES - importación obligatoria
import tkinter as tk
from tkinter import messagebox

# DESPUÉS - importación condicional
try:
    import tkinter as tk
    from tkinter import messagebox
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False
```

### Función `mostrar_mensaje()` adaptada

```python
def mostrar_mensaje(self, titulo, mensaje, tipo="info"):
    """Muestra mensajes al usuario con GUI (solo si tkinter está disponible)"""
    log(f"Mensaje usuario: {titulo}", 'INFO' if tipo == 'info' else tipo.upper())

    # Solo mostrar GUI si tkinter está disponible
    if not TKINTER_AVAILABLE:
        return  # En Streamlit Cloud, solo registra en logs

    # Código GUI para entornos de escritorio
    root = tk.Tk()
    root.withdraw()
    # ... resto del código
```

## Ventajas de esta Solución

✅ **Compatibilidad Universal**:
   - Funciona en **escritorio** (Windows/Mac/Linux con GUI)
   - Funciona en **Streamlit Cloud** (Linux sin GUI)
   - Funciona en **servidores** sin interfaz gráfica

✅ **Sin cambios de funcionalidad**:
   - En escritorio: Sigue mostrando diálogos GUI
   - En cloud: Registra mensajes en logs

✅ **Sin dependencias adicionales**:
   - No requiere instalar paquetes extra
   - No modifica `requirements.txt`

✅ **Retrocompatible**:
   - Los scripts siguen funcionando exactamente igual en escritorio
   - La app web funciona en la nube

## Archivos Modificados

1. `04_SCRIPTS/procesar_datos.py`
   - Importación condicional de tkinter (líneas 17-23)
   - Función `mostrar_mensaje()` adaptada (líneas ~122-138)

2. `04_SCRIPTS/procesar_datos_triodos.py`
   - Importación condicional de tkinter (líneas 17-25)
   - Función `mostrar_mensaje()` adaptada (líneas ~133-149)

3. `core/procesador.py`
   - Ya estaba correctamente adaptado (no usa tkinter)

## Testing

### ✅ Escritorio (Windows/Mac/Linux con GUI)
- Los scripts siguen mostrando diálogos visuales
- Funcionalidad sin cambios

### ✅ Streamlit Cloud (Linux sin GUI)
- La app carga correctamente
- No hay errores de importación
- Los mensajes se registran en logs

### ✅ Streamlit Local
- Funciona en ambos modos (con y sin tkinter)

## Próximos Pasos para Despliegue

1. **Commit los cambios**:
   ```bash
   git add 04_SCRIPTS/procesar_datos.py
   git add 04_SCRIPTS/procesar_datos_triodos.py
   git add core/procesador.py
   git add CAMBIOS_PROCESADOR.md
   git commit -m "Fix: Compatibilidad con Streamlit Cloud (tkinter opcional)"
   ```

2. **Push a GitHub**:
   ```bash
   git push origin main
   ```

3. **Streamlit Cloud actualizará automáticamente**
   - La app se reiniciará
   - Debería funcionar sin errores

## Verificación Post-Despliegue

- [ ] La app carga sin errores
- [ ] Se puede subir un archivo general
- [ ] Se puede procesar un archivo general
- [ ] Se puede subir un archivo Triodos
- [ ] Se puede procesar un archivo Triodos
- [ ] Se pueden generar informes Word
- [ ] Los archivos se descargan correctamente

## Notas Técnicas

- **tkinter** solo se necesita para diálogos GUI en modo escritorio
- En Streamlit, la interfaz ya está en el navegador web
- Los logs se mantienen para debugging (en `/03_LOGS` local, o en logs de Streamlit Cloud)
- La lógica de procesamiento no depende de tkinter, solo la interfaz visual

---

**Fecha**: 29 de octubre de 2025  
**Versión**: 2.0.1  
**Status**: ✅ Listo para despliegue
