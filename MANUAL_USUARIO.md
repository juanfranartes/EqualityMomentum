# Manual de Usuario - EqualityMomentum

**Sistema de Gestión de Registros Retributivos**

Versión 1.0.0

---

## 📑 Índice

1. [Introducción](#introducción)
2. [Instalación](#instalación)
3. [Primer Uso](#primer-uso)
4. [Procesamiento de Datos](#procesamiento-de-datos)
5. [Generación de Informes](#generación-de-informes)
6. [Actualizaciones](#actualizaciones)
7. [Solución de Problemas](#solución-de-problemas)
8. [Preguntas Frecuentes](#preguntas-frecuentes)
9. [Soporte Técnico](#soporte-técnico)

---

## 🎯 Introducción

**EqualityMomentum** es un sistema profesional que permite procesar datos de registros retributivos y generar informes automáticos con análisis de brechas salariales.

### Características principales:

- ✅ **Procesamiento automático** de datos en Excel
- 📊 **Informes profesionales** en Word y PDF
- 🔒 **Privacidad garantizada** con filtros LOPD/RGPD
- 🎨 **Interfaz intuitiva** con identidad corporativa
- 🔄 **Actualizaciones automáticas**

### ¿Qué hace la aplicación?

1. **Procesa datos**: Toma archivos Excel con información salarial y los normaliza/estandariza
2. **Calcula equiparaciones**: Normaliza salarios para hacer comparaciones justas
3. **Genera informes**: Crea documentos profesionales con análisis estadístico y gráficos
4. **Analiza brechas**: Identifica diferencias salariales entre grupos

---

## 💿 Instalación

### Requisitos del Sistema

- **Sistema Operativo:** Windows 10 o superior
- **Espacio en disco:** 500 MB mínimo
- **RAM:** 4 GB mínimo (8 GB recomendado)
- **Resolución:** 1280x720 mínimo

### Proceso de Instalación

1. **Descargar el instalador**
   - Obtenga el archivo `EqualityMomentum_Setup_vX.X.X.exe`

2. **Ejecutar el instalador**
   - Haga doble clic en el archivo descargado
   - Si aparece un aviso de seguridad, haga clic en "Más información" → "Ejecutar de todas formas"

3. **Asistente de instalación**
   - Acepte los términos de licencia
   - Seleccione la carpeta de instalación (por defecto: `C:\Program Files\EqualityMomentum`)
   - Elija si desea crear un acceso directo en el escritorio
   - Haga clic en "Instalar"

4. **Finalizar**
   - Marque "Ejecutar EqualityMomentum" si desea abrirla inmediatamente
   - Haga clic en "Finalizar"

### ¿Qué se instala?

- **Programa principal** en `C:\Program Files\EqualityMomentum`
- **Carpetas de trabajo** en `Documentos\EqualityMomentum`:
  - `Datos`: Para archivos de entrada
  - `Resultados`: Para archivos procesados
  - `Informes`: Para informes generados
  - `Logs`: Para archivos de registro
- **Accesos directos** en el Menú Inicio y (opcionalmente) en el Escritorio

---

## 🚀 Primer Uso

### Abrir la aplicación

1. **Desde el Escritorio**: Haga doble clic en el icono de EqualityMomentum
2. **Desde el Menú Inicio**: Busque "EqualityMomentum" y haga clic

### Pantalla principal

Al abrir la aplicación verá:

```
┌─────────────────────────────────────┐
│      [LOGO EQUALITYMOMENTUM]        │
├─────────────────────────────────────┤
│   Sistema de Gestión de Registros   │
│         Retributivos                │
├─────────────────────────────────────┤
│                                     │
│     ┌─────────────────────┐        │
│     │  PROCESAR DATOS     │        │
│     └─────────────────────┘        │
│                                     │
│     ┌─────────────────────┐        │
│     │  GENERAR INFORME    │        │
│     └─────────────────────┘        │
│                                     │
├─────────────────────────────────────┤
│ [Buscar actualizaciones] [Ayuda]   │
│          Versión 1.0.0              │
└─────────────────────────────────────┘
```

---

## 📊 Procesamiento de Datos

### ¿Cuándo usar esta función?

Use "PROCESAR DATOS" cuando tenga un archivo Excel con información salarial sin procesar y necesite:
- Normalizar los datos
- Calcular equiparaciones salariales
- Preparar los datos para generar informes

### Paso a paso

#### 1. Preparar el archivo Excel

**Formato Estándar:**
- Debe tener la estructura de maestro definida
- Incluir hojas de configuración de complementos

**Formato Triodos:**
- Archivo protegido con contraseña
- Estructura específica de Triodos Bank

#### 2. Iniciar el procesamiento

1. En la pantalla principal, haga clic en **"PROCESAR DATOS"**

2. Se abrirá una nueva ventana

#### 3. Seleccionar archivo

1. Haga clic en **"Examinar..."** junto a "Archivo Excel"
2. Navegue hasta la ubicación de su archivo
3. Seleccione el archivo y haga clic en "Abrir"

#### 4. Configurar opciones

**Tipo de procesamiento:**
- ☐ **Datos estándar** (por defecto)
- ☑ **Datos de Triodos Bank** (si es un archivo de Triodos)

Si seleccionó "Triodos Bank":
- Ingrese la contraseña del archivo (por defecto: `Triodos2025`)

#### 5. Seleccionar carpeta de destino

1. Haga clic en **"Examinar..."** junto a "Carpeta de destino"
2. Seleccione donde desea guardar el archivo procesado
3. Recomendación: Use `Documentos\EqualityMomentum\Resultados`

#### 6. Procesar

1. Haga clic en **"Procesar"**
2. Observe la barra de progreso y el log de procesamiento
3. El proceso puede tardar desde segundos hasta varios minutos, dependiendo del tamaño del archivo

#### 7. Resultado

Al finalizar verá un mensaje:
```
✓ Procesamiento completado exitosamente

Archivo generado:
REPORTE_[TIPO]_YYYYMMDD_HHMMSS.xlsx

Ubicación: [ruta seleccionada]
```

### ⚠️ Problemas comunes

**Error: "El archivo no tiene el formato correcto"**
- Verifique que el archivo Excel tenga la estructura esperada
- Asegúrese de seleccionar el tipo correcto (Estándar o Triodos)

**Error: "Contraseña incorrecta"**
- Verifique la contraseña del archivo
- La contraseña por defecto de Triodos es `Triodos2025`

**El procesamiento es muy lento**
- Esto es normal con archivos grandes (más de 1000 filas)
- No cierre la aplicación, espere a que termine

---

## 📄 Generación de Informes

### ¿Cuándo usar esta función?

Use "GENERAR INFORME" cuando tenga un archivo Excel **YA PROCESADO** y necesite:
- Crear un informe profesional en Word
- Generar gráficos de análisis
- Obtener un documento PDF final

### Paso a paso

#### 1. Tener un archivo procesado

Debe tener un archivo generado por "PROCESAR DATOS" (empieza con `REPORTE_...`)

#### 2. Iniciar la generación

1. En la pantalla principal, haga clic en **"GENERAR INFORME"**
2. Se abrirá una nueva ventana

#### 3. Seleccionar archivo procesado

1. Haga clic en **"Examinar..."** junto a "Archivo Excel"
2. Navegue hasta la ubicación de su archivo procesado
3. Recomendación: Busque en `Documentos\EqualityMomentum\Resultados`
4. Seleccione el archivo y haga clic en "Abrir"

#### 4. Elegir tipo de informe

Seleccione el tipo de informe que desea generar:

- ⚪ **CONSOLIDADO** (Recomendado)
  - Informe completo con todos los análisis
  - Incluye promedios, medianas y análisis de complementos
  - Tablas detalladas y gráficos completos

- ⚪ **PROMEDIO**
  - Solo análisis con promedios
  - Más simple y directo

- ⚪ **MEDIANA**
  - Solo análisis con medianas
  - Útil cuando hay valores atípicos

- ⚪ **COMPLEMENTOS**
  - Solo análisis de complementos salariales
  - Enfocado en componentes específicos del salario

#### 5. Seleccionar carpeta de destino

1. Haga clic en **"Examinar..."** junto a "Carpeta de destino"
2. Seleccione donde desea guardar el informe
3. Recomendación: Use `Documentos\EqualityMomentum\Informes`

#### 6. Generar informe

1. Haga clic en **"Generar Informe"**
2. Observe el progreso:
   - Cargando datos...
   - Analizando información...
   - Generando gráficos...
   - Creando documento Word...
   - Exportando a PDF...

El proceso tarda entre 1 y 5 minutos.

#### 7. Resultado

Al finalizar verá un mensaje:
```
✓ Informe generado exitosamente

Archivos generados:
- registro_retributivo_YYYYMMDD_HHMMSS_CONSOLIDADO.docx
- registro_retributivo_YYYYMMDD_HHMMSS_CONSOLIDADO.pdf

Ubicación: [ruta seleccionada]
```

### 📋 Contenido del informe

El informe incluye:

1. **Portada** con logo corporativo
2. **Resumen ejecutivo**
3. **Análisis por sexo** con brechas salariales
4. **Tablas detalladas** por puesto y categoría
5. **Gráficos profesionales** (donuts, barras)
6. **Análisis de complementos**
7. **Conclusiones y recomendaciones**

### ⚠️ Problemas comunes

**Error: "El archivo no contiene datos procesados"**
- Asegúrese de usar un archivo generado por "PROCESAR DATOS"
- No use archivos Excel originales sin procesar

**El PDF no se genera**
- Verifique que el archivo .docx se haya creado correctamente
- Revise los logs para más detalles

**Valores ocultos por privacidad**
- Esto es normal: la aplicación oculta datos cuando hay pocos empleados (n=1)
- Es un requisito de LOPD/RGPD

---

## 🔄 Actualizaciones

### Verificación automática

La aplicación verifica actualizaciones automáticamente al iniciarse.

Si hay una actualización disponible, verá:
```
┌─────────────────────────────────────┐
│    Actualización disponible         │
├─────────────────────────────────────┤
│  Hay una nueva versión: 1.1.0       │
│                                     │
│  Novedades:                         │
│  • Mejoras en el procesamiento      │
│  • Corrección de errores            │
│  • Nuevas funcionalidades           │
│                                     │
│  ¿Descargar e instalar ahora?       │
│                                     │
│      [Sí]          [No]             │
└─────────────────────────────────────┘
```

### Actualización manual

1. Haga clic en **"Buscar actualizaciones"**
2. La aplicación verificará si hay versiones nuevas
3. Si hay actualización disponible, siga las instrucciones

### Proceso de actualización

1. **Descargar**: La aplicación descarga el instalador
2. **Instalar**: Se ejecuta el nuevo instalador automáticamente
3. **Cerrar**: La aplicación actual se cierra
4. **Reiniciar**: Abra la aplicación actualizada

**Nota:** Sus datos y configuraciones se conservan durante la actualización.

---

## 🔧 Solución de Problemas

### La aplicación no inicia

**Síntomas:**
- Al hacer doble clic no pasa nada
- Aparece un error y se cierra inmediatamente

**Soluciones:**

1. **Verificar requisitos del sistema**
   - Windows 10 o superior
   - 4 GB de RAM mínimo

2. **Reinstalar la aplicación**
   - Desinstale desde "Configuración" → "Aplicaciones"
   - Descargue el instalador más reciente
   - Instale nuevamente

3. **Revisar los logs**
   - Vaya a `Documentos\EqualityMomentum\Logs`
   - Abra el archivo más reciente `app_YYYYMMDD.log`
   - Busque líneas que digan `ERROR` o `CRITICAL`

### Errores durante el procesamiento

**Síntomas:**
- El procesamiento falla con un error
- El archivo resultado no se genera

**Soluciones:**

1. **Verificar formato del archivo**
   - Asegúrese de que el Excel tiene la estructura correcta
   - Si es Triodos, verifique que marcó la opción correspondiente

2. **Verificar contraseña (Triodos)**
   - La contraseña correcta es: `Triodos2025`

3. **Revisar logs de procesamiento**
   - Vaya a `Documentos\EqualityMomentum\Logs`
   - Abra `procesamiento_YYYYMMDD.log` o `procesamiento_triodos_YYYYMMDD.log`

4. **Intentar con otro archivo**
   - Use un archivo de prueba más pequeño
   - Si funciona, el problema está en el archivo original

### Errores al generar informes

**Síntomas:**
- La generación falla antes de terminar
- Se genera el .docx pero no el .pdf

**Soluciones:**

1. **Verificar que el archivo está procesado**
   - Use solo archivos que comiencen con `REPORTE_`
   - Generados por "PROCESAR DATOS"

2. **Si solo falla el PDF**
   - Use el archivo .docx
   - El .docx contiene la misma información

3. **Revisar logs de informes**
   - Vaya a `Documentos\EqualityMomentum\Logs`
   - Abra `informe_YYYYMMDD.log`

### Problemas de actualizaciones

**Síntomas:**
- No detecta actualizaciones
- Falla la descarga

**Soluciones:**

1. **Verificar conexión a internet**
   - La aplicación necesita internet para verificar actualizaciones

2. **Actualización manual**
   - Visite la página de releases
   - Descargue manualmente el instalador
   - Ejecute el instalador

### Reportar problemas

Si el problema persiste:

1. **Recopilar información**
   - Vaya a "Abrir carpeta de logs"
   - Busque archivos `ERROR_REPORT_*.txt` y `ERROR_REPORT_*.json`

2. **Contactar soporte**
   - Envíe los archivos de error
   - Describa el problema y los pasos para reproducirlo

---

## ❓ Preguntas Frecuentes

### ¿Puedo procesar varios archivos a la vez?

No, actualmente la aplicación procesa un archivo a la vez. Debe esperar a que termine uno antes de procesar el siguiente.

### ¿Cuánto tiempo tarda el procesamiento?

Depende del tamaño del archivo:
- Archivos pequeños (<500 filas): 5-30 segundos
- Archivos medianos (500-2000 filas): 30 segundos - 2 minutos
- Archivos grandes (>2000 filas): 2-10 minutos

### ¿Puedo editar el informe generado?

Sí, el archivo `.docx` es totalmente editable. Puede abrirlo con Microsoft Word y hacer cambios.

El archivo `.pdf` no es editable, pero puede regenerarlo después de editar el `.docx`.

### ¿Los datos están seguros?

Sí:
- Todos los datos se procesan localmente en su computadora
- No se envía información a servidores externos
- La aplicación aplica filtros de privacidad LOPD/RGPD automáticamente

### ¿Necesito internet para usar la aplicación?

No para el uso normal. Solo necesita internet para:
- Verificar actualizaciones
- Descargar actualizaciones

El procesamiento y generación de informes funciona sin internet.

### ¿Puedo instalar en varios ordenadores?

Sí, puede instalar la aplicación en tantos ordenadores como necesite.

### ¿Qué pasa con mis datos al actualizar?

Nada. Sus datos en `Documentos\EqualityMomentum` se conservan al actualizar la aplicación.

### ¿Cómo desinstalo la aplicación?

1. Vaya a "Configuración" → "Aplicaciones"
2. Busque "EqualityMomentum"
3. Haga clic en "Desinstalar"
4. Elija si desea conservar o eliminar los datos de usuario

---

## 📞 Soporte Técnico

### Antes de contactar

1. Consulte este manual
2. Revise la sección "Solución de Problemas"
3. Revise los logs en `Documentos\EqualityMomentum\Logs`

### Información para el soporte

Cuando contacte soporte, proporcione:

1. **Versión de la aplicación**
   - Visible en la pantalla principal

2. **Sistema operativo**
   - Windows 10, Windows 11, etc.

3. **Descripción del problema**
   - ¿Qué estaba haciendo?
   - ¿Qué pasó?
   - ¿Mensaje de error (si hay)?

4. **Archivos de log**
   - Los archivos `ERROR_REPORT_*.txt` y `ERROR_REPORT_*.json`
   - O los logs relevantes de `03_LOGS/`

### Contacto

**Email:** soporte@equalitymomentum.com (configurar)

**GitHub Issues:** https://github.com/TU_USUARIO/EqualityMomentum/issues

---

## 📚 Recursos Adicionales

- **README.md**: Información técnica del proyecto
- **GitHub Releases**: Descargar versiones anteriores
- **Logs**: Información detallada de operaciones y errores

---

**© 2025 EqualityMomentum | Desarrollado con ❤️ para promover la igualdad salarial**
