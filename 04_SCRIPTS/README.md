# Equality Momentum - Sistema Automatizado de Registros Retributivos

## 🚀 Uso Rápido (1 Click)

**Para ejecutar todo el proceso completo:**

1. **Colocar datos**: Ponga sus archivos Excel en la carpeta `01_DATOS_SIN_PROCESAR`
2. **Ejecutar**: Haga doble click en `EJECUTAR_WORKFLOW.bat`
3. **Ver resultados**: 
   - Datos procesados → `02_RESULTADOS`
   - Informes generados → `05_INFORMES`

## 📁 Estructura del Sistema

```
📁 01_DATOS_SIN_PROCESAR/  ← Poner aquí los archivos Excel originales
📁 02_RESULTADOS/          ← Datos procesados y limpios
📁 03_LOGS/               ← Logs de procesamiento
📁 04_SCRIPTS/            ← Scripts del sistema (este directorio)
📁 05_INFORMES/           ← Informes Word finales generados
```

## 🔧 Archivos Principales

- **`EJECUTAR_WORKFLOW.bat`** - ⭐ ARCHIVO PRINCIPAL - Ejecuta todo con 1 click
- **`ejecutar_workflow.py`** - Script maestro que coordina todo el proceso
- **`procesar_datos.py`** - Procesa y limpia los datos del Excel
- **`generar_informe.py`** - Genera el informe Word con gráficos
- **`report_config.yaml`** - Configuración de los gráficos e informe
- **`requirements.txt`** - Dependencias de Python
- **`limpiar_sistema.py`** - Limpia archivos temporales

## ⚙️ ¿Qué hace el workflow automático?

1. **Verifica el sistema** - Comprueba Python y dependencias
2. **Instala dependencias** - Si faltan, las instala automáticamente
3. **Procesa datos** - Limpia y calcula valores del Excel
4. **Genera informe** - Crea documento Word con gráficos y tablas
5. **Muestra resultados** - Informa dónde están los archivos generados

## 🛠️ Otros Comandos Útiles

### Ejecutar solo procesamiento de datos:
```bash
python procesar_datos.py
```

### Ejecutar solo generación de informe:
```bash
python generar_informe.py
```

### Verificar que todo funciona:
```bash
python verificar_resultados.py
```

### Limpiar archivos temporales:
```bash
python limpiar_sistema.py
```

## 🐛 Solución de Problemas

### Error: "Python no encontrado"
- Instalar Python desde https://python.org
- ✅ Marcar "Add Python to PATH" durante la instalación

### Error: "No se encontraron archivos Excel"
- Verificar que los archivos .xlsx están en `01_DATOS_SIN_PROCESAR`
- Los archivos no deben estar abiertos en Excel

### Error: "Dependencias faltantes"
- El sistema las instala automáticamente
- Si persiste, ejecutar manualmente: `pip install -r requirements.txt`

### Error de codificación UTF-8
- El sistema maneja automáticamente la codificación
- Si persiste, verificar que los archivos Excel no tienen caracteres especiales

## 📊 Archivos que se Eliminaron

Los siguientes archivos eran innecesarios y se eliminaron para simplificar:
- `crear_ejecutable.bat` (funcionalidad duplicada)
- `setup_completo.bat` (muy complejo para el uso requerido)
- `verificar_sistema.bat` (incluido en el workflow principal)

## 📞 Soporte

Si tiene problemas, revisar los logs en la carpeta `03_LOGS` para más detalles del error.
