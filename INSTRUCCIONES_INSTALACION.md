# Instrucciones de Instalación y Distribución

## Para Usuarios Finales

### Opción 1: Instalador Ejecutable (Recomendado)

1. Descargue `EqualityMomentum_Setup_vX.X.X.exe`
2. Ejecute el instalador
3. Siga el asistente de instalación
4. Use la aplicación desde el acceso directo

**Ventajas:**
- Instalación profesional con un clic
- No requiere Python instalado
- Incluye desinstalador
- Actualizaciones automáticas

---

## Para Desarrolladores

### Configuración del Entorno de Desarrollo

#### 1. Requisitos previos

- Python 3.8 o superior
- Git
- Inno Setup 6 (para crear instaladores)

#### 2. Clonar el repositorio

```bash
git clone https://github.com/TU_USUARIO/EqualityMomentum.git
cd EqualityMomentum
```

#### 3. Crear entorno virtual

```bash
python -m venv EM
```

#### 4. Activar entorno virtual

**Windows:**
```bash
EM\Scripts\activate
```

**Linux/Mac:**
```bash
source EM/bin/activate
```

#### 5. Instalar dependencias

```bash
pip install -r 04_SCRIPTS\requirements.txt
```

#### 6. Ejecutar aplicación

```bash
# Opción 1: Usando el script bat
EJECUTAR_APP.bat

# Opción 2: Directamente
cd 04_SCRIPTS
python app_principal.py
```

---

## Crear Distribución para Clientes

### Proceso Completo (Recomendado)

Este proceso automatiza todo: incrementa versión, compila y crea instalador.

```bash
cd 04_SCRIPTS
python build_release.py
```

El script preguntará:
1. **Tipo de versión**: PATCH (1.0.0 → 1.0.1), MINOR (1.0.0 → 1.1.0), o MAJOR (1.0.0 → 2.0.0)
2. **Changelog**: Lista de cambios en esta versión

**Resultado:**
- Ejecutable compilado en `04_SCRIPTS\dist\EqualityMomentum\`
- Instalador en `Instaladores\EqualityMomentum_Setup_vX.X.X.exe`
- Versión actualizada en todos los archivos
- Tag de Git creado

### Proceso Manual (Paso a Paso)

#### 1. Actualizar versión

Edite `04_SCRIPTS\config.json`:
```json
{
  "version": "1.0.1",
  ...
}
```

Edite `version.json`:
```json
{
  "version": "1.0.1",
  "release_date": "2025-10-28",
  ...
}
```

#### 2. Compilar con PyInstaller

```bash
cd 04_SCRIPTS
build.bat
```

Esto genera el ejecutable en `dist\EqualityMomentum\EqualityMomentum.exe`

#### 3. Probar el ejecutable

```bash
cd dist\EqualityMomentum
EqualityMomentum.exe
```

Verifique que funciona correctamente:
- ✓ Se abre la interfaz
- ✓ Carga el isotipo
- ✓ Puede procesar datos
- ✓ Puede generar informes

#### 4. Crear instalador con Inno Setup

**Opción A: Interfaz gráfica**
1. Abra Inno Setup Compiler
2. Abra el archivo `04_SCRIPTS\installer.iss`
3. Menú: Build → Compile
4. El instalador se guarda en `Instaladores\`

**Opción B: Línea de comandos**
```bash
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" 04_SCRIPTS\installer.iss
```

#### 5. Probar el instalador

1. Ejecute `EqualityMomentum_Setup_vX.X.X.exe`
2. Instale en una carpeta de prueba
3. Verifique que funciona correctamente
4. Desinstale

#### 6. Crear release en GitHub

```bash
# Commit de cambios
git add .
git commit -m "Release v1.0.1"

# Crear tag
git tag -a v1.0.1 -m "Release v1.0.1"

# Push
git push origin main
git push origin v1.0.1
```

En GitHub:
1. Vaya a "Releases" → "Create a new release"
2. Seleccione el tag `v1.0.1`
3. Título: "EqualityMomentum v1.0.1"
4. Descripción: Copie el changelog
5. Adjunte: `EqualityMomentum_Setup_v1.0.1.exe`
6. Publique

#### 7. Actualizar URL de descarga

Edite `version.json` con la URL real del release:
```json
{
  "version": "1.0.1",
  "download_url": "https://github.com/USUARIO/EqualityMomentum/releases/download/v1.0.1/EqualityMomentum_Setup_v1.0.1.exe",
  ...
}
```

Commit y push:
```bash
git add version.json
git commit -m "Update download URL for v1.0.1"
git push origin main
```

---

## Distribución a Clientes

### Método 1: Instalador Ejecutable

**Entregar:**
- `EqualityMomentum_Setup_vX.X.X.exe`
- `MANUAL_USUARIO.pdf` (convertir desde .md)

**Instrucciones para el cliente:**
1. Ejecutar el instalador
2. Seguir el asistente
3. La aplicación se instalará y creará accesos directos
4. Leer el manual de usuario

### Método 2: Ejecutable Portable (Sin instalador)

Si el cliente prefiere no instalar:

**Preparar:**
1. Copie toda la carpeta `dist\EqualityMomentum`
2. Renombre a `EqualityMomentum_Portable_vX.X.X`
3. Añada un archivo `LEEME.txt`:

```
EqualityMomentum - Versión Portable

INSTRUCCIONES:
1. Extraiga esta carpeta donde desee
2. Ejecute EqualityMomentum.exe
3. La aplicación creará carpetas de trabajo en Documentos

NOTA: Esta versión NO incluye:
- Instalación en el sistema
- Accesos directos automáticos
- Desinstalador

Para instalación completa, use el instalador.
```

4. Comprima en ZIP

**Entregar:**
- `EqualityMomentum_Portable_vX.X.X.zip`
- `MANUAL_USUARIO.pdf`

---

## Actualización de Versiones Instaladas

### Para usuarios con versión previa

El sistema de actualizaciones automáticas:

1. **Detecta** versiones nuevas al iniciar la aplicación
2. **Notifica** al usuario que hay actualización
3. **Descarga** el instalador automáticamente
4. **Instala** sobre la versión anterior

**Los datos del usuario se conservan** en `Documentos\EqualityMomentum`

### Forzar actualización manual

Si un cliente tiene problemas con la actualización automática:

1. Envíele el instalador más reciente
2. Instrucciones:
   - Cierre EqualityMomentum si está abierto
   - Ejecute el nuevo instalador
   - Seleccione "Actualizar" cuando pregunte
   - Sus datos se conservarán

---

## Checklist Pre-Release

Antes de distribuir una nueva versión, verifique:

- [ ] Versión actualizada en `config.json`
- [ ] Versión actualizada en `version.json`
- [ ] Changelog completo y descriptivo
- [ ] Compilación exitosa con PyInstaller
- [ ] Ejecutable probado en Windows 10
- [ ] Ejecutable probado en Windows 11
- [ ] Procesamiento de datos funciona
- [ ] Generación de informes funciona
- [ ] Sistema de logs funciona
- [ ] Instalador creado con Inno Setup
- [ ] Instalación probada
- [ ] Desinstalación probada
- [ ] Actualización desde versión anterior probada
- [ ] Manual de usuario actualizado
- [ ] Tag de Git creado
- [ ] Release en GitHub publicado
- [ ] URL de descarga actualizada en `version.json`

---

## Solución de Problemas en la Distribución

### El ejecutable no funciona en el cliente

**Posibles causas:**
- Windows 7 o anterior (no soportado)
- Antivirus bloquea el ejecutable
- Falta permisos de administrador

**Soluciones:**
- Verificar versión de Windows (mínimo Windows 10)
- Añadir excepción en el antivirus
- Ejecutar como administrador

### El instalador falla

**Posibles causas:**
- Ya hay una versión instalada y bloqueada
- Permisos insuficientes
- Espacio en disco insuficiente

**Soluciones:**
- Cerrar la aplicación antes de instalar
- Ejecutar instalador como administrador
- Liberar espacio en disco (mínimo 500 MB)

### Las actualizaciones no funcionan

**Posibles causas:**
- Sin conexión a internet
- Firewall bloquea la conexión
- URL de descarga incorrecta

**Soluciones:**
- Verificar conexión a internet
- Configurar excepción en firewall
- Actualizar manualmente con el instalador

---

## Contacto para Desarrolladores

Para dudas sobre la compilación o distribución:
- GitHub Issues: https://github.com/TU_USUARIO/EqualityMomentum/issues
- Email: dev@equalitymomentum.com

---

**Última actualización: 2025-10-28**
