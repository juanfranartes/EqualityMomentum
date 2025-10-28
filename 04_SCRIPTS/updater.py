"""
Sistema de Actualización Automática para EqualityMomentum
Verifica y descarga actualizaciones desde GitHub
"""

import json
import os
import sys
import subprocess
import tempfile
from pathlib import Path
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError
import tkinter as tk
from tkinter import messagebox
import threading


class Updater:
    """Gestor de actualizaciones automáticas"""

    def __init__(self, config_file="config.json"):
        self.config = self._load_config(config_file)
        self.current_version = self.config.get("version", "1.0.0")
        self.update_url = self.config.get("update_url", "")
        self.latest_version_info = None

    def _load_config(self, config_file):
        """Carga la configuración desde el archivo JSON"""
        try:
            config_path = Path(__file__).parent / config_file
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error al cargar configuración: {e}")
            return {"version": "1.0.0", "update_url": ""}

    def check_for_updates(self, silent=False):
        """
        Verifica si hay actualizaciones disponibles

        Args:
            silent (bool): Si es True, no muestra mensajes si no hay actualizaciones

        Returns:
            dict: Información de la versión más reciente, o None si no hay actualizaciones
        """
        try:
            # Descargar información de la última versión
            request = Request(self.update_url)
            request.add_header('User-Agent', 'EqualityMomentum-Updater/1.0')

            with urlopen(request, timeout=10) as response:
                data = response.read()
                version_info = json.loads(data.decode('utf-8'))

            latest_version = version_info.get("version", "0.0.0")

            # Comparar versiones
            if self._compare_versions(latest_version, self.current_version) > 0:
                self.latest_version_info = version_info
                return version_info
            else:
                if not silent:
                    messagebox.showinfo(
                        "Sin actualizaciones",
                        f"Ya tienes la versión más reciente ({self.current_version})"
                    )
                return None

        except (URLError, HTTPError) as e:
            if not silent:
                messagebox.showerror(
                    "Error de conexión",
                    f"No se pudo verificar actualizaciones:\n{str(e)}\n\n"
                    "Verifica tu conexión a internet."
                )
            return None
        except Exception as e:
            if not silent:
                messagebox.showerror(
                    "Error",
                    f"Error al verificar actualizaciones:\n{str(e)}"
                )
            return None

    def _compare_versions(self, version1, version2):
        """
        Compara dos versiones en formato X.Y.Z

        Returns:
            int: 1 si version1 > version2, -1 si version1 < version2, 0 si son iguales
        """
        v1_parts = [int(x) for x in version1.split('.')]
        v2_parts = [int(x) for x in version2.split('.')]

        # Asegurar que ambas tengan 3 partes
        while len(v1_parts) < 3:
            v1_parts.append(0)
        while len(v2_parts) < 3:
            v2_parts.append(0)

        for i in range(3):
            if v1_parts[i] > v2_parts[i]:
                return 1
            elif v1_parts[i] < v2_parts[i]:
                return -1

        return 0

    def prompt_update(self, parent_window=None):
        """
        Muestra un diálogo preguntando si el usuario quiere actualizar

        Args:
            parent_window: Ventana padre de Tkinter (opcional)
        """
        if not self.latest_version_info:
            return

        version = self.latest_version_info.get("version", "desconocida")
        changelog = self.latest_version_info.get("changelog", [])

        # Crear mensaje
        message = f"Hay una nueva versión disponible: {version}\n\n"
        message += "Novedades:\n"
        for i, change in enumerate(changelog[:5], 1):  # Máximo 5 cambios
            message += f"  • {change}\n"

        message += "\n¿Deseas descargar e instalar la actualización ahora?"

        # Mostrar diálogo
        response = messagebox.askyesno(
            "Actualización disponible",
            message,
            parent=parent_window
        )

        if response:
            self.download_and_install()

    def download_and_install(self):
        """Descarga e instala la actualización"""
        if not self.latest_version_info:
            messagebox.showerror("Error", "No hay información de actualización disponible")
            return

        download_url = self.latest_version_info.get("download_url", "")
        if not download_url:
            messagebox.showerror("Error", "No se encontró URL de descarga")
            return

        # Crear ventana de progreso
        progress_window = tk.Toplevel()
        progress_window.title("Descargando actualización")
        progress_window.geometry("400x150")
        progress_window.resizable(False, False)
        progress_window.transient()
        progress_window.grab_set()

        tk.Label(
            progress_window,
            text="Descargando la actualización...\nEsto puede tomar unos minutos.",
            font=('Work Sans', 12),
            pady=20
        ).pack()

        progress_label = tk.Label(progress_window, text="Iniciando descarga...", font=('Work Sans', 10))
        progress_label.pack(pady=10)

        def download_thread():
            try:
                # Descargar instalador
                progress_label.config(text="Descargando instalador...")

                request = Request(download_url)
                request.add_header('User-Agent', 'EqualityMomentum-Updater/1.0')

                with urlopen(request, timeout=300) as response:
                    # Crear archivo temporal
                    temp_dir = tempfile.gettempdir()
                    installer_path = os.path.join(temp_dir, "EqualityMomentum_Update.exe")

                    with open(installer_path, 'wb') as f:
                        total_size = int(response.headers.get('content-length', 0))
                        downloaded = 0
                        chunk_size = 8192

                        while True:
                            chunk = response.read(chunk_size)
                            if not chunk:
                                break

                            f.write(chunk)
                            downloaded += len(chunk)

                            if total_size > 0:
                                percent = int((downloaded / total_size) * 100)
                                progress_label.config(text=f"Descargando: {percent}%")

                progress_label.config(text="Descarga completada. Iniciando instalación...")
                progress_window.update()

                # Cerrar ventana de progreso
                progress_window.destroy()

                # Mostrar mensaje
                messagebox.showinfo(
                    "Actualización lista",
                    "La actualización se ha descargado correctamente.\n\n"
                    "Se abrirá el instalador. La aplicación se cerrará automáticamente."
                )

                # Ejecutar instalador
                subprocess.Popen([installer_path], shell=True)

                # Cerrar aplicación
                sys.exit(0)

            except Exception as e:
                progress_window.destroy()
                messagebox.showerror(
                    "Error de descarga",
                    f"No se pudo descargar la actualización:\n{str(e)}\n\n"
                    "Por favor, descarga manualmente desde:\n"
                    f"{self.config.get('github_releases_url', '')}"
                )

        # Iniciar descarga en thread separado
        thread = threading.Thread(target=download_thread, daemon=True)
        thread.start()

    def check_on_startup(self, parent_window=None, auto_prompt=True):
        """
        Verifica actualizaciones al inicio de la aplicación

        Args:
            parent_window: Ventana padre de Tkinter
            auto_prompt: Si es True, pregunta automáticamente al usuario si quiere actualizar
        """
        def check_thread():
            version_info = self.check_for_updates(silent=True)
            if version_info and auto_prompt:
                # Ejecutar en el hilo principal de Tkinter
                if parent_window:
                    parent_window.after(100, lambda: self.prompt_update(parent_window))

        thread = threading.Thread(target=check_thread, daemon=True)
        thread.start()


if __name__ == "__main__":
    # Test del updater
    root = tk.Tk()
    root.withdraw()  # Ocultar ventana principal

    updater = Updater()

    print(f"Versión actual: {updater.current_version}")
    print("Verificando actualizaciones...")

    version_info = updater.check_for_updates(silent=False)
    if version_info:
        print(f"Nueva versión disponible: {version_info['version']}")
        updater.prompt_update()
    else:
        print("No hay actualizaciones disponibles")

    root.mainloop()
