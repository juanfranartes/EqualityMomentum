# INSTRUCCIONES: Uso de Plantilla en Streamlit

## 📋 Resumen del cambio

El sistema ha sido actualizado para usar una plantilla de Word personalizada en la generación de informes, compatible tanto con el uso local como con Streamlit Cloud.

## 📁 Ubicación de la plantilla

El sistema buscará la plantilla en el siguiente orden:

1. **`templates/plantilla_informe.docx`** (Recomendado para Streamlit Cloud)
2. **`00_DOCUMENTACION/Registro retributivo/Reg Retributivo  NUEVA PLANTILLA.docx`** (Uso local)

## ⚙️ Configuración requerida

### Para uso en Streamlit Cloud:

1. Coloca tu archivo de plantilla Word en la carpeta `templates/`
2. Renómbralo como `plantilla_informe.docx`
3. Asegúrate de incluirlo en el commit y push a GitHub
4. Verifica que el archivo esté presente en el repositorio antes de desplegar

### Para uso local:

Puedes usar cualquiera de las dos ubicaciones mencionadas arriba.

## ✅ Verificación

Para verificar que la plantilla está siendo utilizada:

1. Revisa los logs generados durante la creación del informe
2. Busca el mensaje: `"Usando plantilla: [ruta]"`
3. Si ves: `"ADVERTENCIA: No se encontró la plantilla..."` significa que la plantilla no está disponible

## 🔄 Comportamiento sin plantilla

Si no se encuentra la plantilla en ninguna ubicación:
- El sistema creará un documento Word desde cero
- Se aplicará formato básico
- Se incluirá todo el contenido (tablas, gráficos, análisis)
- NO se aplicarán estilos corporativos predefinidos de la plantilla

## 📝 Contenido de la plantilla

La plantilla debe incluir:
- Logo y encabezado corporativo de EqualityMomentum
- Estilos predefinidos para títulos, texto y tablas
- Formato y diseño corporativo
- Pie de página con información legal

## 🚀 Despliegue en Streamlit Cloud

**IMPORTANTE**: Para que la plantilla funcione en Streamlit Cloud:

```bash
# 1. Añadir la plantilla al repositorio
git add templates/plantilla_informe.docx

# 2. Commit
git commit -m "Añadir plantilla de informe para Streamlit"

# 3. Push a GitHub
git push origin main
```

## 🔍 Modificaciones realizadas

### Archivo: `04_SCRIPTS/generar_informe_optimizado.py`

**Antes:**
```python
plantilla_path = Path(__file__).parent.parent / '00_DOCUMENTACION' / 'Registro retributivo' / 'Reg Retributivo  NUEVA PLANTILLA.docx'
```

**Después:**
```python
plantilla_paths = [
    Path(__file__).parent.parent / 'templates' / 'plantilla_informe.docx',
    Path(__file__).parent.parent / '00_DOCUMENTACION' / 'Registro retributivo' / 'Reg Retributivo  NUEVA PLANTILLA.docx'
]

plantilla_path = None
for path in plantilla_paths:
    if path.exists():
        plantilla_path = path
        break
```

## 📞 Soporte

Si encuentras problemas con la plantilla:
1. Verifica que el archivo existe en la ubicación correcta
2. Revisa los logs en `03_LOGS/` para mensajes de error
3. Asegúrate de que el archivo no esté corrupto
4. Verifica que tenga permisos de lectura

---

**Fecha de actualización**: 29 de octubre de 2025
**Versión**: 2.0
