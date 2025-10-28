"""
Script para revertir el último release creado
"""

import json
import subprocess
import sys
import os
import shutil
from pathlib import Path


def rollback_release():
    """Revierte el último release"""
    
    root_dir = Path(__file__).parent.parent
    scripts_dir = Path(__file__).parent
    config_path = scripts_dir / "config.json"
    version_path = root_dir / "version.json"
    
    print("=" * 60)
    print("  ROLLBACK DE RELEASE")
    print("=" * 60)
    
    # 1. Leer versión actual
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        current_version = config.get("version", "1.0.0")
        print(f"\n📌 Versión actual: {current_version}")
    except Exception as e:
        print(f"❌ Error al leer configuración: {e}")
        return
    
    # 2. Pedir nueva versión
    print("\nIngrese la versión a la que desea volver")
    old_version = input(f"Versión anterior [ej: 1.0.0]: ").strip()
    
    if not old_version:
        print("❌ Debe ingresar una versión")
        return
    
    confirm = input(f"\n⚠️  ¿Revertir de v{current_version} a v{old_version}? (s/n): ").strip().lower()
    if confirm != 's':
        print("Cancelado")
        return
    
    # 3. Actualizar config.json
    print("\n🔄 Actualizando configuración...")
    config['version'] = old_version
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)
    print(f"  ✓ config.json actualizado a v{old_version}")
    
    # 4. Limpiar directorios de build
    print("\n🧹 Limpiando directorios de build...")
    dirs_to_clean = ['build', 'dist', '__pycache__']
    for dir_name in dirs_to_clean:
        dir_path = scripts_dir / dir_name
        if dir_path.exists():
            shutil.rmtree(dir_path)
            print(f"  ✓ {dir_name} eliminado")
    
    # 5. Eliminar tag de Git
    print(f"\n🏷️  Eliminando tag de Git v{current_version}...")
    try:
        result = subprocess.run(['git', 'tag', '-d', f'v{current_version}'], 
                              capture_output=True, text=True, cwd=root_dir)
        if result.returncode == 0:
            print(f"  ✓ Tag local v{current_version} eliminado")
            print(f"  ℹ️  Para eliminar del remoto: git push origin :refs/tags/v{current_version}")
        else:
            print(f"  ℹ️  Tag v{current_version} no existe localmente")
    except Exception as e:
        print(f"  ⚠️  No se pudo eliminar tag: {e}")
    
    # 6. Eliminar changelog
    changelog_file = root_dir / f"CHANGELOG_v{current_version}.txt"
    if changelog_file.exists():
        changelog_file.unlink()
        print(f"\n📝 Changelog eliminado: {changelog_file.name}")
    
    # 7. Eliminar instalador si existe
    installer_dir = root_dir.parent / "Instaladores"
    installer_file = installer_dir / f"EqualityMomentum_Setup_v{current_version}.exe"
    if installer_file.exists():
        installer_file.unlink()
        print(f"📦 Instalador eliminado: {installer_file.name}")
    
    print("\n" + "=" * 60)
    print("  ✅ ROLLBACK COMPLETADO")
    print("=" * 60)
    print(f"  Versión revertida: {old_version}")
    print("\nPróximos pasos:")
    print("  1. Revisar cambios en Git: git status")
    print("  2. Hacer commit: git add . && git commit -m 'Revert to v{}'".format(old_version))
    print(f"  3. Eliminar tag remoto: git push origin :refs/tags/v{current_version}")
    print("=" * 60)


if __name__ == "__main__":
    rollback_release()