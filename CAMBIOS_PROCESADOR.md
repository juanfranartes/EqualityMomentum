# Cambios en el Procesador de Datos

## Fecha: 29 de octubre de 2025

## Resumen
Se ha refactorizado el módulo `core/procesador.py` para que la aplicación web utilice directamente los scripts de procesamiento de `04_SCRIPTS` en lugar de mantener lógica duplicada.

## Cambios Realizados

### 1. Simplificación de `core/procesador.py`

**Antes:**
- El archivo contenía ~650 líneas con lógica completa de procesamiento duplicada
- Mantenía su propia implementación de equiparación, cálculos y complementos
- Riesgo de inconsistencias entre la versión web y los scripts de escritorio

**Después:**
- El archivo tiene ~130 líneas actuando como adaptador
- Delega todo el procesamiento a los scripts originales en `04_SCRIPTS`
- Garantiza que web y escritorio usan exactamente la misma lógica

### 2. Funcionamiento Actual

La clase `ProcesadorRegistroRetributivo` en `core/procesador.py` ahora:

1. **Inicializa** dos procesadores:
   - `ProcesadorGeneral` desde `04_SCRIPTS/procesar_datos.py`
   - `ProcesadorTriodos` desde `04_SCRIPTS/procesar_datos_triodos.py`

2. **Método `procesar_excel_general(archivo_bytes)`**:
   - Recibe archivo en memoria (BytesIO) desde Streamlit
   - Crea archivo temporal en disco
   - Llama a `procesador_general.leer_y_procesar_excel()`
   - Llama a `procesador_general.crear_reporte_excel()`
   - Convierte resultado a BytesIO
   - Limpia archivos temporales
   - Devuelve BytesIO a Streamlit

3. **Método `procesar_excel_triodos(archivo_bytes, password)`**:
   - Mismo flujo pero usando `procesador_triodos.leer_y_procesar_triodos()`
   - Maneja archivos protegidos con contraseña

### 3. Scripts Utilizados

#### Archivo General
- **Script**: `04_SCRIPTS/procesar_datos.py`
- **Clase**: `ProcesadorRegistroRetributivo`
- **Métodos llamados**:
  - `leer_y_procesar_excel(ruta_archivo)`
  - `crear_reporte_excel(archivo_original, df_procesado)`

#### Archivo Triodos
- **Script**: `04_SCRIPTS/procesar_datos_triodos.py`
- **Clase**: `ProcesadorTriodos`
- **Métodos llamados**:
  - `leer_y_procesar_triodos(ruta_archivo)`
  - `crear_reporte_excel(archivo_original, df_procesado)`

### 4. Ventajas de este Enfoque

✅ **Mantenimiento unificado**: Un solo lugar donde actualizar la lógica de procesamiento
✅ **Consistencia**: Web y escritorio usan exactamente el mismo código
✅ **Simplicidad**: `core/procesador.py` es ahora solo un adaptador ligero
✅ **Menos bugs**: Eliminamos el riesgo de divergencia entre implementaciones

### 5. Flujo en Streamlit

```
Usuario sube archivo → streamlit_app.py
                              ↓
                   Detecta tipo (General/Triodos)
                              ↓
                   core/procesador.py (adaptador)
                              ↓
            Archivo temporal en disco
                              ↓
        04_SCRIPTS/procesar_datos[_triodos].py
                              ↓
              Procesamiento completo
                              ↓
            Archivo temporal de salida
                              ↓
         core/procesador.py (lee y convierte)
                              ↓
          BytesIO de vuelta a Streamlit
                              ↓
           Usuario descarga resultado
```

## Archivos Modificados

- `core/procesador.py` - Completamente refactorizado (650 → 130 líneas)
- `04_SCRIPTS/procesar_datos.py` - Añadida compatibilidad con entornos sin GUI
- `04_SCRIPTS/procesar_datos_triodos.py` - Añadida compatibilidad con entornos sin GUI

### Compatibilidad con Streamlit Cloud

Se han modificado los scripts para que funcionen en entornos sin interfaz gráfica:

- **Importación condicional de tkinter**: Los scripts detectan automáticamente si tkinter está disponible
- **Función `mostrar_mensaje` adaptativa**: Solo muestra diálogos GUI si tkinter está presente
- **Compatible con Linux/Cloud**: Los scripts funcionan perfectamente en Streamlit Cloud sin tkinter

```python
# Importación segura
try:
    import tkinter as tk
    from tkinter import messagebox
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False
```

## Archivos No Modificados (funcionan igual)

- `streamlit_app.py` - No requiere cambios, la interfaz es la misma
- `04_SCRIPTS/procesar_datos.py` - Script original sin cambios
- `04_SCRIPTS/procesar_datos_triodos.py` - Script original sin cambios
- `core/generador.py` - Generador de informes sin cambios

## Testing Recomendado

1. ✅ Probar carga de archivo general en la app web
2. ✅ Probar carga de archivo Triodos en la app web
3. ✅ Verificar que los archivos procesados son idénticos a los generados por scripts de escritorio
4. ✅ Comprobar limpieza de archivos temporales
5. ✅ Verificar manejo de errores y excepciones

## Notas Técnicas

- Los archivos temporales se crean con `tempfile.NamedTemporaryFile`
- Se usa `delete=False` para mantener control manual de la limpieza
- Los archivos temporales se eliminan siempre, incluso en caso de error
- La ruta `04_SCRIPTS` se agrega dinámicamente a `sys.path` para importación
