"""
Script automatizado para crear releases de EqualityMomentum
Automatiza todo el proceso de compilación y creación de instalador
"""

import json
import subprocess
import sys
import os
from pathlib import Path
from datetime import datetime
import shutil


class ReleaseBuilder:
    """Constructor automatizado de releases"""

    def __init__(self):
        self.root_dir = Path(__file__).parent.parent
        self.scripts_dir = Path(__file__).parent
        self.config_path = self.scripts_dir / "config.json"
        self.version_path = self.root_dir / "version.json"
        self.config = self._load_config()
        self.current_version = self.config.get("version", "1.0.0")

    def _load_config(self):
        """Carga la configuración"""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"❌ Error al cargar configuración: {e}")
            sys.exit(1)

    def _save_config(self):
        """Guarda la configuración actualizada"""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            print(f"✓ Configuración actualizada")
        except Exception as e:
            print(f"❌ Error al guardar configuración: {e}")
            sys.exit(1)

    def _update_version_file(self, new_version, changelog):
        """Actualiza el archivo version.json"""
        try:
            version_data = {
                "version": new_version,
                "release_date": datetime.now().strftime("%Y-%m-%d"),
                "download_url": f"https://github.com/juanfranartes/EqualityMomentum/releases/download/v{new_version}/EqualityMomentum_Setup_v{new_version}.exe",
                "changelog": changelog,
                "min_windows_version": "10",
                "required_space_mb": 500,
                "notes": f"Versión {new_version} del sistema EqualityMomentum"
            }

            with open(self.version_path, 'w', encoding='utf-8') as f:
                json.dump(version_data, f, indent=2, ensure_ascii=False)

            print(f"✓ Archivo version.json actualizado")
        except Exception as e:
            print(f"❌ Error al actualizar version.json: {e}")
            sys.exit(1)

    def increment_version(self, part='patch'):
        """
        Incrementa la versión automáticamente

        Args:
            part (str): 'major', 'minor', o 'patch'
        """
        parts = [int(x) for x in self.current_version.split('.')]

        if part == 'major':
            parts[0] += 1
            parts[1] = 0
            parts[2] = 0
        elif part == 'minor':
            parts[1] += 1
            parts[2] = 0
        elif part == 'patch':
            parts[2] += 1

        new_version = '.'.join(map(str, parts))
        self.config['version'] = new_version
        self.current_version = new_version

        print(f"✓ Versión incrementada a: {new_version}")
        return new_version

    def clean_build_dirs(self):
        """Limpia directorios de build anteriores"""
        print("\n📁 Limpiando directorios de build...")

        dirs_to_clean = ['build', 'dist', '__pycache__']

        for dir_name in dirs_to_clean:
            dir_path = self.scripts_dir / dir_name
            if dir_path.exists():
                shutil.rmtree(dir_path)
                print(f"  ✓ Eliminado: {dir_name}")

        print("✓ Directorios limpiados")

    def run_pyinstaller(self):
        """Ejecuta PyInstaller para compilar la aplicación"""
        print("\n🔨 Compilando aplicación con PyInstaller...")
        print("   Esto puede tomar varios minutos...\n")

        try:
            result = subprocess.run(
                [sys.executable, '-m', 'PyInstaller', 'EqualityMomentum.spec', '--clean'],
                cwd=self.scripts_dir,
                capture_output=True,
                text=True
            )

            if result.returncode != 0:
                print(f"❌ Error en PyInstaller:")
                print(result.stderr)
                sys.exit(1)

            print("✓ Compilación exitosa")

            # Crear estructura de carpetas
            self._create_folder_structure()

        except Exception as e:
            print(f"❌ Error al ejecutar PyInstaller: {e}")
            sys.exit(1)

    def _create_folder_structure(self):
        """Crea la estructura de carpetas necesaria"""
        print("\n📂 Creando estructura de carpetas...")

        dist_dir = self.scripts_dir / "dist" / "EqualityMomentum"
        folders = [
            "01_DATOS_SIN_PROCESAR",
            "02_RESULTADOS",
            "03_LOGS",
            "05_INFORMES"
        ]

        for folder in folders:
            folder_path = dist_dir / folder
            folder_path.mkdir(exist_ok=True)
            print(f"  ✓ {folder}")

        # Copiar archivos adicionales
        shutil.copy(self.config_path, dist_dir / "config.json")
        shutil.copy(self.version_path, dist_dir / "version.json")

        print("✓ Estructura creada")

    def run_inno_setup(self):
        """Ejecuta Inno Setup para crear el instalador"""
        print("\n📦 Creando instalador con Inno Setup...")

        # Buscar Inno Setup
        inno_paths = [
            r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
            r"C:\Program Files\Inno Setup 6\ISCC.exe",
            r"C:\Program Files (x86)\Inno Setup 5\ISCC.exe",
            r"C:\Program Files\Inno Setup 5\ISCC.exe",
        ]

        iscc_path = None
        for path in inno_paths:
            if os.path.exists(path):
                iscc_path = path
                break

        if not iscc_path:
            print("⚠️  Inno Setup no encontrado")
            print("   Por favor, ejecute manualmente:")
            print(f"   1. Abra Inno Setup")
            print(f"   2. Compile: {self.scripts_dir / 'installer.iss'}")
            return False

        try:
            # Actualizar versión en el script .iss
            self._update_iss_version()

            result = subprocess.run(
                [iscc_path, str(self.scripts_dir / "installer.iss")],
                capture_output=True,
                text=True
            )

            if result.returncode != 0:
                print(f"❌ Error en Inno Setup:")
                print(result.stderr)
                return False

            print("✓ Instalador creado exitosamente")
            return True

        except Exception as e:
            print(f"❌ Error al ejecutar Inno Setup: {e}")
            return False

    def _update_iss_version(self):
        """Actualiza la versión en el archivo installer.iss"""
        iss_path = self.scripts_dir / "installer.iss"

        try:
            with open(iss_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # Reemplazar la línea de versión
            import re
            content = re.sub(
                r'#define MyAppVersion ".*?"',
                f'#define MyAppVersion "{self.current_version}"',
                content
            )

            with open(iss_path, 'w', encoding='utf-8') as f:
                f.write(content)

            print(f"  ✓ Versión actualizada en installer.iss")

        except Exception as e:
            print(f"⚠️  No se pudo actualizar installer.iss: {e}")

    def create_changelog_file(self, changelog):
        """Crea archivo de notas de versión"""
        print("\n📝 Creando notas de versión...")

        changelog_path = self.root_dir / f"CHANGELOG_v{self.current_version}.txt"

        try:
            with open(changelog_path, 'w', encoding='utf-8') as f:
                f.write(f"EqualityMomentum v{self.current_version}\n")
                f.write(f"{'=' * 50}\n\n")
                f.write(f"Fecha: {datetime.now().strftime('%Y-%m-%d')}\n\n")
                f.write("Cambios:\n")
                for i, change in enumerate(changelog, 1):
                    f.write(f"  {i}. {change}\n")

            print(f"✓ Archivo creado: {changelog_path.name}")

        except Exception as e:
            print(f"⚠️  Error al crear changelog: {e}")

    def create_git_tag(self):
        """Crea un tag en Git para la versión"""
        print(f"\n🏷️  Creando tag de Git v{self.current_version}...")

        try:
            # Verificar si Git está disponible
            result = subprocess.run(['git', '--version'], capture_output=True)
            if result.returncode != 0:
                print("⚠️  Git no encontrado, saltando creación de tag")
                return

            # Crear tag
            subprocess.run(
                ['git', 'tag', '-a', f'v{self.current_version}', '-m', f'Release v{self.current_version}'],
                cwd=self.root_dir,
                capture_output=True
            )

            print(f"✓ Tag v{self.current_version} creado")
            print("  Para subir el tag: git push origin v{self.current_version}")

        except Exception as e:
            print(f"⚠️  No se pudo crear tag de Git: {e}")

    def build_release(self, increment_type='patch', changelog=None):
        """
        Ejecuta todo el proceso de build y release

        Args:
            increment_type (str): 'major', 'minor', o 'patch'
            changelog (list): Lista de cambios en esta versión
        """
        print("=" * 60)
        print("  BUILD AUTOMATIZADO DE EQUALITYMOMENTUM")
        print("=" * 60)

        # 1. Incrementar versión
        new_version = self.increment_version(increment_type)

        # 2. Actualizar archivos de configuración
        self._save_config()
        self._update_version_file(new_version, changelog or [])

        # 3. Limpiar builds anteriores
        self.clean_build_dirs()

        # 4. Compilar con PyInstaller
        self.run_pyinstaller()

        # 5. Crear instalador con Inno Setup
        inno_success = self.run_inno_setup()

        # 6. Crear archivo de changelog
        if changelog:
            self.create_changelog_file(changelog)

        # 7. Crear tag de Git
        self.create_git_tag()

        # Resumen final
        print("\n" + "=" * 60)
        print("  ✅ BUILD COMPLETADO")
        print("=" * 60)
        print(f"  Versión: {new_version}")
        print(f"  Ejecutable: dist/EqualityMomentum/EqualityMomentum.exe")

        if inno_success:
            print(f"  Instalador: ../Instaladores/EqualityMomentum_Setup_v{new_version}.exe")
        else:
            print("  Instalador: Crear manualmente con Inno Setup")

        print("\nPróximos pasos:")
        print("  1. Probar el ejecutable")
        print("  2. Probar el instalador")
        print("  3. Commit y push de cambios")
        print(f"  4. git push origin v{new_version}")
        print("  5. Crear release en GitHub")
        print("=" * 60)


def main():
    """Función principal"""
    builder = ReleaseBuilder()

    print("¿Qué tipo de incremento de versión desea?")
    print("  1. PATCH (bug fixes, cambios menores) - Recomendado")
    print("  2. MINOR (nuevas funcionalidades)")
    print("  3. MAJOR (cambios importantes)")
    print()

    choice = input("Seleccione (1/2/3) [1]: ").strip() or "1"

    increment_map = {
        "1": "patch",
        "2": "minor",
        "3": "major"
    }

    increment_type = increment_map.get(choice, "patch")

    print("\nIngrese los cambios de esta versión (uno por línea, línea vacía para terminar):")
    changelog = []
    while True:
        line = input("  - ").strip()
        if not line:
            break
        changelog.append(line)

    if not changelog:
        changelog = ["Correcciones y mejoras generales"]

    print("\n¿Desea continuar con el build?")
    confirm = input("(s/n) [s]: ").strip().lower() or "s"

    if confirm != 's':
        print("Build cancelado")
        sys.exit(0)

    # Ejecutar build
    builder.build_release(increment_type, changelog)


if __name__ == "__main__":
    main()
