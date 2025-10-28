"""
Interfaz moderna para el Procesador de Datos
Versión mejorada con mejor diseño visual y UX
"""

import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext, messagebox
import subprocess
import threading
import os
import sys
from datetime import datetime
from pathlib import Path
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


class ProcesadorDatosGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("EqualityMomentum - Procesador de Datos")

        # Abrir en pantalla completa (maximizado)
        self.root.state('zoomed')
        self.root.resizable(True, True)  # Permitir redimensionar
        self.root.minsize(850, 800)  # Tamaño mínimo

        # Colores corporativos
        self.primary_color = "#1f3c89"
        self.accent_color = "#ff5c39"
        self.bg_color = "#f8f9fa"
        self.card_color = "#ffffff"

        # Configurar estilo moderno
        self._configurar_estilo()

        # Variables
        self.archivo_excel = tk.StringVar()
        self.carpeta_destino = tk.StringVar()
        self.es_triodos = tk.BooleanVar(value=False)
        self.password = tk.StringVar(value="Triodos2025")
        self.proceso_activo = False

        # Establecer carpeta destino por defecto
        default_resultados = Path.home() / "Documents" / "EqualityMomentum" / "Resultados"
        if not default_resultados.exists():
            default_resultados = Path(__file__).parent.parent / "02_RESULTADOS"
        self.carpeta_destino.set(str(default_resultados))

        # Configurar interfaz
        self.crear_widgets()

    def _configurar_estilo(self):
        """Configura un estilo moderno y profesional"""
        style = ttk.Style()
        style.theme_use('clam')

        # Colores base
        style.configure('.',
                       font=('Segoe UI', 10),
                       borderwidth=0)

        # Frames
        style.configure('Card.TFrame',
                       background=self.card_color,
                       relief='flat')

        style.configure('Modern.TFrame',
                       background=self.bg_color)

        style.configure('Header.TFrame',
                       background=self.primary_color)

        # Labels
        style.configure('Modern.TLabel',
                       background=self.bg_color,
                       foreground='#333333',
                       font=('Segoe UI', 10))

        style.configure('Card.TLabel',
                       background=self.card_color,
                       foreground='#333333',
                       font=('Segoe UI', 10))

        style.configure('Title.TLabel',
                       background=self.primary_color,
                       foreground='white',
                       font=('Segoe UI', 18, 'bold'),
                       padding=20)

        style.configure('SectionTitle.TLabel',
                       background=self.card_color,
                       foreground=self.primary_color,
                       font=('Segoe UI', 11, 'bold'))

        style.configure('Subtitle.TLabel',
                       background=self.card_color,
                       foreground='#666666',
                       font=('Segoe UI', 9))

        # Buttons
        style.configure('Modern.TButton',
                       font=('Segoe UI', 10),
                       borderwidth=1,
                       relief='flat',
                       padding=10)

        style.configure('Primary.TButton',
                       font=('Segoe UI', 12, 'bold'),
                       borderwidth=0,
                       padding=15)

        style.map('Primary.TButton',
                 background=[('active', self.accent_color), ('!active', self.primary_color)],
                 foreground=[('active', 'white'), ('!active', 'white')])

        # Entry
        style.configure('Modern.TEntry',
                       fieldbackground='white',
                       borderwidth=1,
                       relief='solid')

        # Checkbutton
        style.configure('Card.TCheckbutton',
                       background=self.card_color,
                       font=('Segoe UI', 10))

        # Progressbar
        style.configure('Modern.Horizontal.TProgressbar',
                       troughcolor='#e9ecef',
                       background=self.accent_color,
                       borderwidth=0,
                       thickness=20)

        # Configurar fondo de la ventana
        self.root.configure(bg=self.bg_color)

    def crear_widgets(self):
        # Frame de cabecera
        header_frame = ttk.Frame(self.root, style='Header.TFrame')
        header_frame.pack(fill=tk.X)

        titulo = ttk.Label(header_frame,
                          text="📊 Procesador de Datos",
                          style='Title.TLabel')
        titulo.pack()

        # Frame principal
        main_frame = ttk.Frame(self.root, padding="20", style='Modern.TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Card 1: Selección de archivo
        card1_container = self._crear_card(main_frame, "1  Seleccionar archivo de entrada")
        card1_container.pack(fill=tk.X, pady=(0, 15))

        # Entrada de archivo
        file_frame = ttk.Frame(card1_container.card_content, style='Card.TFrame')
        file_frame.pack(fill=tk.X, padx=15, pady=(0, 15))

        self.entry_archivo = tk.Entry(file_frame,
                                      textvariable=self.archivo_excel,
                                      font=('Segoe UI', 10),
                                      relief='solid',
                                      borderwidth=1,
                                      bg='white')
        self.entry_archivo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        btn_examinar = tk.Button(file_frame,
                                text="📁 Examinar...",
                                command=self.seleccionar_archivo,
                                font=('Segoe UI', 10),
                                bg=self.primary_color,
                                fg='white',
                                activebackground=self.accent_color,
                                activeforeground='white',
                                relief='flat',
                                cursor='hand2',
                                padx=20,
                                pady=8,
                                borderwidth=0)
        btn_examinar.pack(side=tk.RIGHT)

        # Asegurar que el botón responde al hover
        btn_examinar.bind('<Enter>', lambda e: e.widget.config(bg=self.accent_color))
        btn_examinar.bind('<Leave>', lambda e: e.widget.config(bg=self.primary_color))

        # Card 2: Tipo de procesamiento
        card2_container = self._crear_card(main_frame, "2  Tipo de procesamiento")
        card2_container.pack(fill=tk.X, pady=(0, 15))

        type_frame = ttk.Frame(card2_container.card_content, style='Card.TFrame')
        type_frame.pack(fill=tk.X, padx=15, pady=(0, 15))

        # Radio buttons para tipo
        tk.Radiobutton(type_frame,
                      text="📄 Datos estándar",
                      variable=self.es_triodos,
                      value=False,
                      command=self.cambiar_tipo,
                      font=('Segoe UI', 10),
                      bg=self.card_color,
                      activebackground=self.card_color,
                      selectcolor=self.card_color).pack(anchor=tk.W, pady=5)

        tk.Radiobutton(type_frame,
                      text="🏦 Datos de Triodos Bank",
                      variable=self.es_triodos,
                      value=True,
                      command=self.cambiar_tipo,
                      font=('Segoe UI', 10),
                      bg=self.card_color,
                      activebackground=self.card_color,
                      selectcolor=self.card_color).pack(anchor=tk.W, pady=5)

        self.lbl_descripcion = ttk.Label(type_frame,
                                        text="Procesamiento estándar de datos de registro retributivo",
                                        style='Subtitle.TLabel')
        self.lbl_descripcion.pack(anchor=tk.W, pady=(10, 0), padx=(25, 0))

        # Frame de contraseña (oculto por defecto)
        self.frame_password = ttk.Frame(card2_container.card_content, style='Card.TFrame')
        self.frame_password.pack(fill=tk.X, padx=15, pady=(0, 15))

        ttk.Label(self.frame_password,
                 text="🔒 Contraseña del archivo:",
                 style='Card.TLabel').pack(anchor=tk.W, pady=(0, 5))

        self.entry_password = tk.Entry(self.frame_password,
                                       textvariable=self.password,
                                       show="●",
                                       font=('Segoe UI', 10),
                                       relief='solid',
                                       borderwidth=1,
                                       bg='white',
                                       width=30)
        self.entry_password.pack(anchor=tk.W)

        self.frame_password.pack_forget()  # Ocultar inicialmente

        # Card 3: Carpeta de destino
        card3_container = self._crear_card(main_frame, "3  Carpeta de destino")
        card3_container.pack(fill=tk.X, pady=(0, 15))

        dest_frame = ttk.Frame(card3_container.card_content, style='Card.TFrame')
        dest_frame.pack(fill=tk.X, padx=15, pady=(0, 15))

        self.entry_destino = tk.Entry(dest_frame,
                                      textvariable=self.carpeta_destino,
                                      font=('Segoe UI', 10),
                                      relief='solid',
                                      borderwidth=1,
                                      bg='white')
        self.entry_destino.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        btn_destino = tk.Button(dest_frame,
                               text="📂 Seleccionar...",
                               command=self.seleccionar_destino,
                               font=('Segoe UI', 10),
                               bg=self.primary_color,
                               fg='white',
                               activebackground=self.accent_color,
                               activeforeground='white',
                               relief='flat',
                               cursor='hand2',
                               padx=20,
                               pady=8,
                               borderwidth=0)
        btn_destino.pack(side=tk.RIGHT)

        # Hover effect
        btn_destino.bind('<Enter>', lambda e: e.widget.config(bg=self.accent_color))
        btn_destino.bind('<Leave>', lambda e: e.widget.config(bg=self.primary_color))

        # Card 4: Progreso
        card4_container = self._crear_card(main_frame, "Progreso")
        card4_container.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        progress_frame = ttk.Frame(card4_container.card_content, style='Card.TFrame')
        progress_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))

        # Estado
        self.lbl_estado = ttk.Label(progress_frame,
                                    text="✓ Listo para procesar",
                                    style='Card.TLabel',
                                    font=('Segoe UI', 10, 'bold'))
        self.lbl_estado.pack(anchor=tk.W, pady=(0, 10))

        # Barra de progreso
        self.progress = ttk.Progressbar(progress_frame,
                                       mode='indeterminate',
                                       style='Modern.Horizontal.TProgressbar',
                                       length=750)
        self.progress.pack(fill=tk.X, pady=(0, 15))

        # Log
        ttk.Label(progress_frame,
                 text="Detalles del proceso:",
                 style='Card.TLabel',
                 font=('Segoe UI', 9, 'bold')).pack(anchor=tk.W, pady=(0, 5))

        self.log_text = scrolledtext.ScrolledText(progress_frame,
                                                   width=85,
                                                   height=10,
                                                   font=('Consolas', 9),
                                                   bg='#f8f9fa',
                                                   relief='solid',
                                                   borderwidth=1,
                                                   state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Botones de acción
        btn_frame = ttk.Frame(main_frame, style='Modern.TFrame')
        btn_frame.pack(fill=tk.X, pady=(10, 0))

        self.btn_procesar = tk.Button(btn_frame,
                                      text="▶️  Procesar Datos",
                                      command=self.procesar_datos,
                                      font=('Segoe UI', 12, 'bold'),
                                      bg=self.primary_color,
                                      fg='white',
                                      activebackground=self.accent_color,
                                      activeforeground='white',
                                      relief='flat',
                                      cursor='hand2',
                                      padx=30,
                                      pady=12,
                                      borderwidth=0)
        self.btn_procesar.pack(side=tk.LEFT, padx=(0, 10))
        self.btn_procesar.bind('<Enter>', lambda e: e.widget.config(bg=self.accent_color) if str(e.widget['state']) != 'disabled' else None)
        self.btn_procesar.bind('<Leave>', lambda e: e.widget.config(bg=self.primary_color) if str(e.widget['state']) != 'disabled' else None)

        btn_limpiar = tk.Button(btn_frame,
                               text="🔄 Limpiar",
                               command=self.limpiar,
                               font=('Segoe UI', 10),
                               bg='#6c757d',
                               fg='white',
                               activebackground='#5a6268',
                               activeforeground='white',
                               relief='flat',
                               cursor='hand2',
                               padx=20,
                               pady=10,
                               borderwidth=0)
        btn_limpiar.pack(side=tk.LEFT, padx=(0, 10))
        btn_limpiar.bind('<Enter>', lambda e: e.widget.config(bg='#5a6268'))
        btn_limpiar.bind('<Leave>', lambda e: e.widget.config(bg='#6c757d'))

        btn_cerrar = tk.Button(btn_frame,
                              text="✖️ Cerrar",
                              command=self.root.destroy,
                              font=('Segoe UI', 10),
                              bg='#dc3545',
                              fg='white',
                              activebackground='#c82333',
                              activeforeground='white',
                              relief='flat',
                              cursor='hand2',
                              padx=20,
                              pady=10,
                              borderwidth=0)
        btn_cerrar.pack(side=tk.RIGHT)
        btn_cerrar.bind('<Enter>', lambda e: e.widget.config(bg='#c82333'))
        btn_cerrar.bind('<Leave>', lambda e: e.widget.config(bg='#dc3545'))

        # Log inicial
        self.agregar_log("Sistema listo para procesar datos")
        self.agregar_log(f"Carpeta de destino: {self.carpeta_destino.get()}")

    def _crear_card(self, parent, titulo):
        """Crea una tarjeta con sombra y título"""
        # Frame contenedor con borde para simular sombra
        container = tk.Frame(parent, bg='#dee2e6', bd=0)

        # Card principal
        card = tk.Frame(container, bg=self.card_color, bd=0)
        card.pack(padx=2, pady=2, fill=tk.BOTH, expand=True)

        # Título de la card
        title_label = tk.Label(card,
                              text=titulo,
                              font=('Segoe UI', 11, 'bold'),
                              bg=self.card_color,
                              fg=self.primary_color,
                              anchor=tk.W)
        title_label.pack(fill=tk.X, padx=15, pady=(15, 10))

        # Separador
        sep = tk.Frame(card, bg='#e9ecef', height=1)
        sep.pack(fill=tk.X, padx=15, pady=(0, 10))

        # Guardar referencia a la card interna para poder añadir contenido
        container.card_content = card

        return container

    def cambiar_tipo(self):
        """Actualiza la interfaz según el tipo de procesamiento"""
        if self.es_triodos.get():
            self.lbl_descripcion.config(
                text="📑 Archivo protegido de Triodos Bank (requiere contraseña)"
            )
            self.frame_password.pack(fill=tk.X, padx=15, pady=(0, 15))
            self.agregar_log("Tipo cambiado a: Triodos Bank")
        else:
            self.lbl_descripcion.config(
                text="📄 Procesamiento estándar de datos de registro retributivo"
            )
            self.frame_password.pack_forget()
            self.agregar_log("Tipo cambiado a: Estándar")

    def seleccionar_archivo(self):
        """Abre el diálogo para seleccionar archivo Excel"""
        # Determinar directorio inicial
        try:
            initial_dir = Path(__file__).parent.parent / "01_DATOS_SIN_PROCESAR"
            if not initial_dir.exists():
                initial_dir = Path.home() / "Documents" / "EqualityMomentum" / "Datos"
            if not initial_dir.exists():
                initial_dir = Path.home() / "Documents"
        except:
            initial_dir = Path.home() / "Documents"

        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[
                ("Archivos Excel", "*.xlsx *.xls"),
                ("Todos los archivos", "*.*")
            ],
            initialdir=str(initial_dir)
        )

        if archivo:
            self.archivo_excel.set(archivo)
            self.agregar_log(f"✓ Archivo seleccionado: {Path(archivo).name}")
            self.actualizar_estado("✓ Archivo seleccionado, listo para procesar", 'success')

    def seleccionar_destino(self):
        """Abre el diálogo para seleccionar carpeta de destino"""
        # Usar la carpeta actual o una por defecto
        current_dir = self.carpeta_destino.get()
        if not current_dir or not Path(current_dir).exists():
            current_dir = str(Path.home() / "Documents")

        carpeta = filedialog.askdirectory(
            title="Seleccionar carpeta de destino",
            initialdir=current_dir
        )

        if carpeta:
            self.carpeta_destino.set(carpeta)
            self.agregar_log(f"✓ Carpeta de destino: {carpeta}")
            self.actualizar_estado("✓ Carpeta de destino actualizada", 'success')

    def agregar_log(self, mensaje):
        """Agrega un mensaje al área de log"""
        self.log_text.config(state='normal')
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {mensaje}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update()

    def actualizar_estado(self, mensaje, tipo='info'):
        """Actualiza el label de estado"""
        colores = {
            'info': '#0d6efd',
            'success': '#198754',
            'warning': '#ffc107',
            'error': '#dc3545'
        }

        # Actualizar color del texto del label
        self.lbl_estado.config(
            text=mensaje,
            foreground=colores.get(tipo, '#333333')
        )
        self.root.update()

    def procesar_datos(self):
        """Ejecuta el script de procesamiento de datos"""
        # Validaciones
        if not self.archivo_excel.get():
            self.agregar_log("❌ ERROR: No se ha seleccionado ningún archivo")
            messagebox.showerror("Error", "Debe seleccionar un archivo Excel")
            return

        if not os.path.exists(self.archivo_excel.get()):
            self.agregar_log("❌ ERROR: El archivo seleccionado no existe")
            messagebox.showerror("Error", "El archivo seleccionado no existe")
            return

        if not self.carpeta_destino.get():
            self.agregar_log("❌ ERROR: No se ha seleccionado carpeta de destino")
            messagebox.showerror("Error", "Debe seleccionar una carpeta de destino")
            return

        if self.es_triodos.get() and not self.password.get():
            self.agregar_log("❌ ERROR: Debe ingresar la contraseña del archivo")
            messagebox.showerror("Error", "Debe ingresar la contraseña del archivo de Triodos")
            return

        # Deshabilitar botón
        self.btn_procesar.config(state='disabled', bg='#6c757d')
        self.proceso_activo = True

        # Iniciar progreso
        self.progress.start(10)
        self.actualizar_estado("⏳ Procesando datos...", 'warning')

        # Ejecutar en thread
        thread = threading.Thread(target=self.ejecutar_script)
        thread.daemon = True
        thread.start()

    def ejecutar_script(self):
        """Ejecuta el script en un thread separado"""
        try:
            self.agregar_log("=" * 60)
            self.agregar_log("🚀 Iniciando procesamiento...")

            # Determinar qué script ejecutar
            if self.es_triodos.get():
                script = "procesar_datos_triodos.py"
                self.agregar_log("📊 Procesando datos de Triodos Bank...")
            else:
                script = "procesar_datos.py"
                self.agregar_log("📊 Procesando datos estándar...")

            # Preparar comando
            python_exe = sys.executable
            script_path = Path(__file__).parent / script
            archivo_input = self.archivo_excel.get()

            # Preparar variables de entorno
            env = os.environ.copy()
            if self.es_triodos.get():
                env['TRIODOS_PASSWORD'] = self.password.get()
            env['OUTPUT_DIR'] = self.carpeta_destino.get()

            # Ejecutar script
            result = subprocess.run(
                [python_exe, str(script_path), archivo_input],
                capture_output=True,
                text=True,
                env=env,
                cwd=str(Path(__file__).parent)
            )

            # Procesar resultado
            if result.returncode == 0:
                self.agregar_log("✅ Procesamiento completado exitosamente")
                self.actualizar_estado("✅ ¡Procesamiento completado!", 'success')

                # Buscar archivo de salida
                output_files = list(Path(self.carpeta_destino.get()).glob("REPORTE_*.xlsx"))
                if output_files:
                    latest_file = max(output_files, key=lambda p: p.stat().st_mtime)
                    self.agregar_log(f"📁 Archivo generado: {latest_file.name}")
                    self.agregar_log(f"📂 Ubicación: {latest_file.parent}")

                messagebox.showinfo(
                    "Éxito",
                    "El procesamiento se completó correctamente.\n\n"
                    f"Los resultados se guardaron en:\n{self.carpeta_destino.get()}"
                )
            else:
                self.agregar_log("❌ Error durante el procesamiento")
                self.agregar_log(f"Código de error: {result.returncode}")
                if result.stderr:
                    self.agregar_log(f"Detalles: {result.stderr[:500]}")
                self.actualizar_estado("❌ Error en el procesamiento", 'error')
                messagebox.showerror("Error", f"Error durante el procesamiento:\n{result.stderr[:200]}")

        except Exception as e:
            self.agregar_log(f"❌ Excepción: {str(e)}")
            self.actualizar_estado("❌ Error inesperado", 'error')
            messagebox.showerror("Error", f"Error inesperado:\n{str(e)}")

        finally:
            # Restaurar interfaz
            self.progress.stop()
            self.btn_procesar.config(state='normal', bg=self.primary_color)
            self.proceso_activo = False
            self.agregar_log("=" * 60)

    def limpiar(self):
        """Limpia los campos del formulario"""
        self.archivo_excel.set("")
        self.password.set("Triodos2025")
        self.es_triodos.set(False)
        self.cambiar_tipo()

        self.log_text.config(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log_text.config(state='disabled')

        self.actualizar_estado("✓ Listo para procesar", 'info')
        self.agregar_log("Sistema listo para procesar datos")
        self.agregar_log(f"Carpeta de destino: {self.carpeta_destino.get()}")


def main():
    root = tk.Tk()
    app = ProcesadorDatosGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
