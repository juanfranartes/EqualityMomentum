"""
EqualityMomentum - Aplicación Principal
Sistema de Gestión de Registros Retributivos

Identidad Corporativa:
- Colores: #1f3c89 (azul), #ff5c39 (naranja), #ffffff (blanco)
- Tipografías: Lusitana (títulos), Work Sans (texto)
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sys
import os
import json
from pathlib import Path
from PIL import Image, ImageTk
import subprocess
import threading
import ctypes

# Habilitar DPI awareness para Windows (mejora la nitidez en pantallas de alta resolución)
try:
    if sys.platform == 'win32':
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass  # Si falla, continuar sin DPI awareness

# Importar módulos personalizados
from logger_manager import get_logger_manager
from updater import Updater


class EqualityMomentumApp:
    """Aplicación principal con interfaz corporativa"""

    def __init__(self, root):
        self.root = root
        self.config = self._load_config()
        self.version = self.config.get("version", "1.0.0")

        # Configurar logger
        self.logger_manager = get_logger_manager("EqualityMomentum", self.version)
        self.logger_manager.setup_exception_handler()
        self.logger = self.logger_manager.get_logger()
        self.logger.info("Iniciando aplicación principal")

        # Configurar updater
        self.updater = Updater()

        # Configurar ventana principal
        self._setup_window()

        # Crear interfaz
        self._create_widgets()

        # Verificar actualizaciones al inicio (después de 2 segundos)
        self.root.after(2000, lambda: self.updater.check_on_startup(self.root, auto_prompt=True))

    def _load_config(self):
        """Carga la configuración desde config.json"""
        try:
            config_path = Path(__file__).parent / "config.json"
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error al cargar configuración: {e}")
            return {
                "version": "1.0.0",
                "colors": {
                    "primary": "#1f3c89",
                    "accent": "#ff5c39",
                    "base": "#ffffff"
                },
                "fonts": {
                    "title": "Lusitana",
                    "body": "Work Sans"
                }
            }

    def _setup_window(self):
        """Configura la ventana principal"""
        self.root.title(f"EqualityMomentum v{self.version}")

        # Poner en pantalla completa (maximizada)
        self.root.state('zoomed')  # Para Windows

        # Alternativa para otros sistemas operativos
        # self.root.attributes('-zoomed', True)  # Para Linux
        # self.root.attributes('-fullscreen', True)  # Para MacOS

        self.root.resizable(True, True)

        # Intentar establecer el icono
        try:
            icon_path = Path(__file__).parent.parent / "00_DOCUMENTACION" / "isotipo.jpg"
            if icon_path.exists():
                # Convertir JPG a ICO temporalmente
                img = Image.open(icon_path)
                img = img.resize((64, 64), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                self.root.iconphoto(True, photo)
        except Exception as e:
            self.logger.warning(f"No se pudo cargar el icono: {e}")

        # Colores corporativos
        colors = self.config.get("colors", {})
        self.primary_color = colors.get("primary", "#1f3c89")
        self.accent_color = colors.get("accent", "#ff5c39")
        self.base_color = colors.get("base", "#ffffff")

        # Configurar estilo
        style = ttk.Style()
        style.theme_use('clam')

        # Estilos personalizados
        style.configure('Title.TLabel',
                       background=self.primary_color,
                       foreground='white',
                       font=('Lusitana', 28, 'bold'),
                       padding=20)

        style.configure('Subtitle.TLabel',
                       background=self.base_color,
                       foreground=self.primary_color,
                       font=('Work Sans', 14),
                       padding=10)

        style.configure('Action.TButton',
                       font=('Work Sans', 16, 'bold'),
                       padding=20,
                       background=self.primary_color,
                       foreground='white')

        style.map('Action.TButton',
                 background=[('active', self.accent_color)])

        style.configure('Secondary.TButton',
                       font=('Work Sans', 11),
                       padding=8)

        style.configure('Version.TLabel',
                       font=('Work Sans', 9),
                       foreground='#666666')

    def _create_widgets(self):
        """Crea los widgets de la interfaz"""
        # Frame principal con color de fondo
        main_frame = tk.Frame(self.root, bg=self.base_color)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Cabecera con logo
        header_frame = tk.Frame(main_frame, bg=self.primary_color, height=220)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)

        # Intentar cargar el isotipo
        try:
            isotipo_path = Path(__file__).parent.parent / "00_DOCUMENTACION" / "isotipo.jpg"
            if isotipo_path.exists():
                img = Image.open(isotipo_path)
                img = img.resize((150, 150), Image.Resampling.LANCZOS)
                self.isotipo_photo = ImageTk.PhotoImage(img)

                logo_label = tk.Label(header_frame, image=self.isotipo_photo, bg=self.primary_color)
                logo_label.pack(pady=30)
        except Exception as e:
            self.logger.warning(f"No se pudo cargar el isotipo: {e}")

            # Título alternativo si no hay logo
            title_label = tk.Label(
                header_frame,
                text="EqualityMomentum",
                font=('Lusitana', 32, 'bold'),
                bg=self.primary_color,
                fg='white'
            )
            title_label.pack(pady=30)

        # Subtítulo
        subtitle_frame = tk.Frame(main_frame, bg=self.base_color)
        subtitle_frame.pack(fill=tk.X, pady=20)

        subtitle = tk.Label(
            subtitle_frame,
            text="Sistema de Gestión de Registros Retributivos",
            font=('Work Sans', 16),
            bg=self.base_color,
            fg=self.primary_color
        )
        subtitle.pack()

        # Frame de botones principales
        buttons_frame = tk.Frame(main_frame, bg=self.base_color)
        buttons_frame.pack(expand=True, pady=50)

        # Botones con tamaño fijo en píxeles para garantizar que sean iguales
        button_width = 700  # Ancho fijo en píxeles
        button_height = 120  # Alto fijo en píxeles

        # Botón PROCESAR DATOS
        btn_procesar_frame = tk.Frame(buttons_frame, width=button_width, height=button_height, bg=self.primary_color)
        btn_procesar_frame.pack(pady=25)
        btn_procesar_frame.pack_propagate(False)  # Evitar que el frame se ajuste al contenido

        btn_procesar = tk.Button(
            btn_procesar_frame,
            text="PROCESAR DATOS",
            font=('Work Sans', 20, 'bold'),
            bg=self.primary_color,
            fg='white',
            activebackground=self.accent_color,
            activeforeground='white',
            cursor='hand2',
            borderwidth=0,
            relief='flat',
            command=self.abrir_procesador
        )
        btn_procesar.place(relx=0.5, rely=0.5, anchor='center', relwidth=1, relheight=1)

        # Botón GENERAR INFORME
        btn_informe_frame = tk.Frame(buttons_frame, width=button_width, height=button_height, bg=self.accent_color)
        btn_informe_frame.pack(pady=25)
        btn_informe_frame.pack_propagate(False)  # Evitar que el frame se ajuste al contenido

        btn_informe = tk.Button(
            btn_informe_frame,
            text="GENERAR INFORME",
            font=('Work Sans', 20, 'bold'),
            bg=self.accent_color,
            fg='white',
            activebackground=self.primary_color,
            activeforeground='white',
            cursor='hand2',
            borderwidth=0,
            relief='flat',
            command=self.abrir_generador
        )
        btn_informe.place(relx=0.5, rely=0.5, anchor='center', relwidth=1, relheight=1)

        # Frame inferior con opciones adicionales
        footer_frame = tk.Frame(main_frame, bg=self.base_color)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=20)

        # Botones secundarios
        secondary_buttons_frame = tk.Frame(footer_frame, bg=self.base_color)
        secondary_buttons_frame.pack()

        btn_actualizar = ttk.Button(
            secondary_buttons_frame,
            text="Buscar actualizaciones",
            style='Secondary.TButton',
            command=self.buscar_actualizaciones
        )
        btn_actualizar.pack(side=tk.LEFT, padx=10)

        btn_abrir_logs = ttk.Button(
            secondary_buttons_frame,
            text="Abrir carpeta de logs",
            style='Secondary.TButton',
            command=self.abrir_carpeta_logs
        )
        btn_abrir_logs.pack(side=tk.LEFT, padx=10)

        btn_ayuda = ttk.Button(
            secondary_buttons_frame,
            text="Ayuda",
            style='Secondary.TButton',
            command=self.mostrar_ayuda
        )
        btn_ayuda.pack(side=tk.LEFT, padx=10)

        # Información de versión
        version_label = ttk.Label(
            footer_frame,
            text=f"Versión {self.version}",
            style='Version.TLabel'
        )
        version_label.pack(pady=10)

    def abrir_procesador(self):
        """Abre la interfaz del procesador de datos"""
        self.logger.info("Abriendo procesador de datos")
        try:
            # Importar y abrir el procesador en una nueva ventana
            # Intentar primero la versión mejorada (v2)
            try:
                from interfaz_procesador_v2 import ProcesadorDatosGUI
            except ImportError:
                from interfaz_procesador import ProcesadorDatosGUI

            procesador_window = tk.Toplevel(self.root)
            # Asegurar que la ventana tenga el tamaño correcto
            procesador_window.update_idletasks()
            ProcesadorDatosGUI(procesador_window)
            # Forzar actualización de geometría
            procesador_window.update()

        except Exception as e:
            self.logger.exception("Error al abrir procesador de datos")
            messagebox.showerror(
                "Error",
                f"No se pudo abrir el procesador de datos:\n{str(e)}\n\n"
                "Revise los logs para más información."
            )

    def abrir_generador(self):
        """Abre la interfaz del generador de informes"""
        self.logger.info("Abriendo generador de informes")
        try:
            # Importar y abrir el generador en una nueva ventana
            # Intentar primero la versión mejorada (v2)
            try:
                from interfaz_generador_v2 import GeneradorInformeGUI
            except ImportError:
                from interfaz_generador import GeneradorInformeGUI

            generador_window = tk.Toplevel(self.root)
            # Asegurar que la ventana tenga el tamaño correcto
            generador_window.update_idletasks()
            GeneradorInformeGUI(generador_window)
            # Forzar actualización de geometría
            generador_window.update()

        except Exception as e:
            self.logger.exception("Error al abrir generador de informes")
            messagebox.showerror(
                "Error",
                f"No se pudo abrir el generador de informes:\n{str(e)}\n\n"
                "Revise los logs para más información."
            )

    def buscar_actualizaciones(self):
        """Busca actualizaciones manualmente"""
        self.logger.info("Buscando actualizaciones manualmente")

        def check_thread():
            version_info = self.updater.check_for_updates(silent=False)
            if version_info:
                self.root.after(0, lambda: self.updater.prompt_update(self.root))

        thread = threading.Thread(target=check_thread, daemon=True)
        thread.start()

    def abrir_carpeta_logs(self):
        """Abre la carpeta de logs en el explorador"""
        self.logger.info("Abriendo carpeta de logs")
        try:
            logs_dir = self.logger_manager.logs_dir
            if sys.platform == 'win32':
                os.startfile(logs_dir)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', logs_dir])
            else:
                subprocess.Popen(['xdg-open', logs_dir])
        except Exception as e:
            self.logger.exception("Error al abrir carpeta de logs")
            messagebox.showerror(
                "Error",
                f"No se pudo abrir la carpeta de logs:\n{str(e)}"
            )

    def mostrar_ayuda(self):
        """Muestra la ventana de ayuda"""
        self.logger.info("Mostrando ayuda")

        ayuda_window = tk.Toplevel(self.root)
        ayuda_window.title("Ayuda - EqualityMomentum")
        ayuda_window.geometry("600x500")
        ayuda_window.resizable(False, False)

        # Título
        titulo = tk.Label(
            ayuda_window,
            text="Ayuda - EqualityMomentum",
            font=('Work Sans', 16, 'bold'),
            fg=self.primary_color
        )
        titulo.pack(pady=20)

        # Texto de ayuda
        texto_ayuda = """
        CÓMO USAR LA APLICACIÓN

        1. PROCESAR DATOS:
           • Haga clic en "PROCESAR DATOS"
           • Seleccione el archivo Excel con los datos sin procesar
           • Elija el tipo de procesamiento (Estándar o Triodos)
           • Seleccione la carpeta donde guardar los resultados
           • Haga clic en "Procesar"

        2. GENERAR INFORME:
           • Haga clic en "GENERAR INFORME"
           • Seleccione el archivo Excel ya procesado
           • Elija el tipo de informe que desea generar
           • Seleccione la carpeta donde guardar el informe
           • Haga clic en "Generar Informe"

        ACTUALIZACIONES:
           • La aplicación verifica actualizaciones al inicio
           • Puede buscar actualizaciones manualmente desde el botón
             "Buscar actualizaciones"

        SOLUCIÓN DE PROBLEMAS:
           • Si encuentra errores, revise la carpeta de logs
           • Los logs contienen información detallada sobre errores
           • Puede enviar los reportes de error al equipo de desarrollo

        CONTACTO:
           • Para soporte técnico, envíe los archivos de log
           • Ubicación: [Ver carpeta de logs]
        """

        texto = tk.Text(
            ayuda_window,
            wrap=tk.WORD,
            font=('Work Sans', 10),
            padx=20,
            pady=20
        )
        texto.insert('1.0', texto_ayuda)
        texto.config(state=tk.DISABLED)
        texto.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # Botón cerrar
        btn_cerrar = ttk.Button(
            ayuda_window,
            text="Cerrar",
            command=ayuda_window.destroy
        )
        btn_cerrar.pack(pady=10)


def main():
    """Función principal"""
    root = tk.Tk()
    app = EqualityMomentumApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
