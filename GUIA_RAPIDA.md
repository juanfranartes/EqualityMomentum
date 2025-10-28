# 🚀 Guía Rápida - EqualityMomentum

## Para Comenzar AHORA (Desarrollo)

### 1. Instalar dependencias

```bash
# Activar entorno virtual (si no está activo)
EM\Scripts\activate

# Instalar/actualizar dependencias
pip install -r 04_SCRIPTS\requirements.txt
```

### 2. Ejecutar la aplicación

```bash
# Opción 1: Script .bat (más sencillo)
EJECUTAR_APP.bat

# Opción 2: Directamente
cd 04_SCRIPTS
python app_principal.py
```

---

## Para Crear el Instalador (Distribución)

### Proceso Automatizado (Recomendado)

```bash
cd 04_SCRIPTS
python build_release.py
```

Esto hará TODO automáticamente:
- ✅ Incrementa la versión
- ✅ Actualiza archivos de configuración
- ✅ Compila con PyInstaller
- ✅ Crea el instalador con Inno Setup
- ✅ Genera changelog
- ✅ Crea tag de Git

### Proceso Manual (Si prefieres control paso a paso)

```bash
# 1. Compilar el ejecutable
cd 04_SCRIPTS
build.bat

# 2. Probar el ejecutable
cd dist\EqualityMomentum
EqualityMomentum.exe

# 3. Crear instalador con Inno Setup
# Abrir Inno Setup y compilar: 04_SCRIPTS\installer.iss
```

---

## Flujo de Actualización Semanal

Si modificas los scripts y quieres distribuir a la empresa:

```bash
# 1. Hacer cambios en los scripts
# Editar: procesar_datos.py, generar_informe_optimizado.py, etc.

# 2. Probar los cambios
cd 04_SCRIPTS
python app_principal.py

# 3. Crear nueva versión
python build_release.py
# Seleccionar: PATCH (para cambios menores)
# Ingresar: Descripción de los cambios

# 4. Resultado en: Instaladores\EqualityMomentum_Setup_vX.X.X.exe

# 5. Enviar el instalador a la empresa
# Ellos simplemente ejecutan el instalador y actualiza automáticamente
```

---

## Estructura de Archivos Clave

```
EqualityMomentum/
│
├── 04_SCRIPTS/
│   ├── app_principal.py              ← Aplicación principal (NUEVA)
│   ├── logger_manager.py             ← Sistema de logs (NUEVO)
│   ├── updater.py                    ← Actualizaciones (NUEVO)
│   ├── config.json                   ← Configuración (NUEVO)
│   │
│   ├── procesar_datos.py             ← Tu código existente
│   ├── procesar_datos_triodos.py     ← Tu código existente
│   ├── generar_informe_optimizado.py ← Tu código existente
│   │
│   ├── build.bat                     ← Compilar ejecutable (NUEVO)
│   ├── build_release.py              ← Build automatizado (NUEVO)
│   ├── EqualityMomentum.spec         ← Config PyInstaller (NUEVO)
│   └── installer.iss                 ← Config Inno Setup (NUEVO)
│
├── version.json                      ← Info de versión (NUEVO)
├── README.md                         ← Documentación (NUEVO)
├── MANUAL_USUARIO.md                 ← Manual completo (NUEVO)
├── INSTRUCCIONES_INSTALACION.md      ← Guía instalación (NUEVO)
└── EJECUTAR_APP.bat                  ← Lanzador desarrollo (NUEVO)
```

---

## Características de la Aplicación

### ✨ Lo Nuevo

1. **Interfaz Unificada**
   - Un solo programa con dos botones: PROCESAR DATOS y GENERAR INFORME
   - Identidad corporativa integrada (#1f3c89, #ff5c39)
   - Logo de la empresa visible

2. **Selección de Archivos**
   - El usuario siempre elige el archivo de entrada
   - El usuario siempre elige dónde guardar el resultado
   - Se recuerdan las últimas ubicaciones

3. **Sistema de Logging**
   - TODO error queda registrado automáticamente
   - Genera reportes estructurados para enviar al desarrollador
   - Formato JSON + TXT legible

4. **Actualizaciones Automáticas**
   - Verifica al inicio si hay nueva versión
   - Descarga e instala automáticamente
   - El usuario solo hace clic en "Actualizar"

5. **Instalador Profesional**
   - Instalación con un clic
   - Crea accesos directos
   - Incluye desinstalador
   - Conserva datos al actualizar

### 🔄 Flujo de Actualización para la Empresa

```
Tú (Desarrollador)          Empresa (Usuario)
─────────────────           ──────────────────

1. Modificas código

2. Ejecutas:
   build_release.py

3. Obtienes:
   EqualityMomentum_
   Setup_v1.0.1.exe

4. Envías instalador   →    5. Ejecuta instalador

                            6. Instalación automática

                            7. Aplicación actualizada

                            ✓ Datos conservados
                            ✓ Nuevas funciones
```

---

## Requisitos Previos

### Para Desarrollo

- ✅ Python 3.8+ (ya lo tienes)
- ✅ Entorno virtual EM (ya lo tienes)
- ✅ Dependencias en requirements.txt

### Para Crear Instalador

- ❌ **Inno Setup 6** (necesitas instalarlo)
  - Descargar: https://jrsoftware.org/isdl.php
  - Instalar versión 6.x
  - Gratuito y open source

### Para Distribuir

- ✅ Solo el archivo .exe del instalador
- ✅ Manual de usuario (opcional, pero recomendado)

---

## Comandos Útiles

```bash
# Ver logs
cd 03_LOGS
dir *.log

# Limpiar builds
cd 04_SCRIPTS
rmdir /s /q build dist

# Probar ejecutable sin instalar
cd 04_SCRIPTS\dist\EqualityMomentum
EqualityMomentum.exe

# Ver versión actual
type 04_SCRIPTS\config.json | findstr version
```

---

## Checklist Antes de Distribuir

Antes de enviar el instalador a la empresa:

- [ ] Probaste los cambios en desarrollo
- [ ] Ejecutaste build_release.py
- [ ] Probaste el ejecutable compilado
- [ ] Probaste el instalador
- [ ] Verificaste que procesa datos correctamente
- [ ] Verificaste que genera informes correctamente
- [ ] Revisaste que no hay errores en los logs
- [ ] Documentaste los cambios en el changelog

---

## Solución Rápida de Problemas

### No encuentra módulos al ejecutar

```bash
# Reinstalar dependencias
pip install -r 04_SCRIPTS\requirements.txt --force-reinstall
```

### PyInstaller no funciona

```bash
# Actualizar PyInstaller
pip install --upgrade pyinstaller
```

### Inno Setup no se encuentra

```bash
# Instalar desde: https://jrsoftware.org/isdl.php
# O compilar instalador manualmente abriendo installer.iss
```

### El ejecutable no abre

```bash
# Verificar que se compiló correctamente
cd 04_SCRIPTS\dist\EqualityMomentum
dir EqualityMomentum.exe

# Probar directamente
EqualityMomentum.exe

# Ver errores
type ..\..\03_LOGS\app_*.log
```

---

## Próximos Pasos Sugeridos

1. **Ahora mismo:**
   - Instala PyInstaller: `pip install pyinstaller`
   - Descarga e instala Inno Setup
   - Prueba ejecutar: `python 04_SCRIPTS\app_principal.py`

2. **Después:**
   - Crea tu primer build: `python 04_SCRIPTS\build_release.py`
   - Prueba el instalador generado
   - Instala en una máquina de prueba

3. **Para distribución:**
   - Crea un release en GitHub (opcional)
   - Sube el instalador a un sitio compartido
   - Envía el instalador + manual a la empresa

---

## URLs Importantes (Actualizar)

Edita estos archivos y actualiza los URLs con tu información real:

- `04_SCRIPTS\config.json` → Cambiar URLs de GitHub
- `version.json` → Cambiar URL de descarga
- `04_SCRIPTS\installer.iss` → Cambiar URL de soporte

Busca y reemplaza: `TU_USUARIO` por tu usuario de GitHub

---

## Contacto y Soporte

Si tienes dudas sobre el sistema de distribución:

- 📖 Lee: `INSTRUCCIONES_INSTALACION.md` (detallado)
- 📖 Lee: `README.md` (técnico)
- 📖 Lee: `MANUAL_USUARIO.md` (para usuarios finales)

---

**¡Listo! Ahora tienes un sistema profesional de distribución con actualizaciones automáticas.**

Última actualización: 2025-10-28
