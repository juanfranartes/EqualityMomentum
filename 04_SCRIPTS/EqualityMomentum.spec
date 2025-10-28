# -*- mode: python ; coding: utf-8 -*-
"""
Archivo de configuración de PyInstaller para EqualityMomentum
"""

from PyInstaller.utils.hooks import collect_data_files, collect_submodules
import os

block_cipher = None

# Datos adicionales a incluir
datas = []

# Añadir config.json
datas.append(('config.json', '.'))

# Añadir isotipo
isotipo_path = os.path.join('..', '00_DOCUMENTACION', 'isotipo.jpg')
if os.path.exists(isotipo_path):
    datas.append((isotipo_path, '00_DOCUMENTACION'))

# Añadir datos de matplotlib, seaborn, etc.
datas += collect_data_files('matplotlib')
datas += collect_data_files('seaborn')

# Módulos ocultos que PyInstaller podría no detectar
hiddenimports = [
    'PIL._tkinter_finder',
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'tkinter.scrolledtext',
    'pandas',
    'numpy',
    'openpyxl',
    'matplotlib',
    'matplotlib.backends.backend_tkagg',
    'seaborn',
    'docx',
    'yaml',
    'msoffcrypto',
    'dateutil',
    'procesar_datos',
    'procesar_datos_triodos',
    'generar_informe_optimizado',
    'interfaz_procesador',
    'interfaz_generador',
    'logger_manager',
    'updater',
]

# Añadir submódulos de matplotlib y numpy
hiddenimports += collect_submodules('matplotlib')
hiddenimports += collect_submodules('numpy')
hiddenimports += collect_submodules('PIL')

a = Analysis(
    ['app_principal.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='EqualityMomentum',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Sin consola (solo ventana GUI)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join('..', '00_DOCUMENTACION', 'isotipo.jpg') if os.path.exists(os.path.join('..', '00_DOCUMENTACION', 'isotipo.jpg')) else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='EqualityMomentum',
)
