"""
Procesador de Registros Retributivos - Versión Web
Refactorizado para trabajar con archivos en memoria (BytesIO)
Utiliza los scripts de 04_SCRIPTS para el procesamiento
"""

import io
import sys
import tempfile
from pathlib import Path
from datetime import datetime

# Agregar la ruta de 04_SCRIPTS al path
scripts_path = Path(__file__).parent.parent / '04_SCRIPTS'
if str(scripts_path) not in sys.path:
    sys.path.insert(0, str(scripts_path))

# Importar los procesadores originales
from procesar_datos import ProcesadorRegistroRetributivo as ProcesadorGeneral
from procesar_datos_triodos import ProcesadorTriodos


class ProcesadorRegistroRetributivo:
    """Procesador de registros retributivos que trabaja con archivos en memoria"""

    def __init__(self):
        """Inicializa el procesador con configuración"""
        self.procesador_general = ProcesadorGeneral()
        self.procesador_triodos = ProcesadorTriodos()

    def procesar_excel_general(self, archivo_bytes):
        """
        Procesa un archivo Excel en formato general desde BytesIO
        Utiliza el script 04_SCRIPTS/procesar_datos.py

        Args:
            archivo_bytes: BytesIO o bytes del archivo Excel

        Returns:
            BytesIO con el archivo Excel procesado
        """
        # Convertir a BytesIO si es necesario
        if isinstance(archivo_bytes, bytes):
            archivo_bytes = io.BytesIO(archivo_bytes)
        
        # Crear archivo temporal
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_input:
            archivo_bytes.seek(0)
            tmp_input.write(archivo_bytes.read())
            tmp_input_path = Path(tmp_input.name)
        
        try:
            # Procesar con el script original
            df_procesado = self.procesador_general.leer_y_procesar_excel(tmp_input_path)
            
            # Crear archivo temporal de salida
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_output:
                tmp_output_path = Path(tmp_output.name)
            
            # Crear reporte usando el método del procesador original
            resultado_path = self.procesador_general.crear_reporte_excel(tmp_input_path, df_procesado)
            
            # Leer el resultado y convertirlo a BytesIO
            output = io.BytesIO()
            with open(resultado_path, 'rb') as f:
                output.write(f.read())
            output.seek(0)
            
            # Limpiar archivos temporales
            tmp_input_path.unlink(missing_ok=True)
            resultado_path.unlink(missing_ok=True)
            if tmp_output_path.exists():
                tmp_output_path.unlink(missing_ok=True)
            
            return output
            
        except Exception as e:
            # Limpiar archivos temporales en caso de error
            tmp_input_path.unlink(missing_ok=True)
            raise Exception(f"Error al procesar archivo general: {str(e)}")

    def procesar_excel_triodos(self, archivo_bytes, password='Triodos2025'):
        """
        Procesa un archivo Excel de Triodos (protegido con contraseña)
        Utiliza el script 04_SCRIPTS/procesar_datos_triodos.py

        Args:
            archivo_bytes: BytesIO o bytes del archivo Excel
            password: Contraseña del archivo

        Returns:
            BytesIO con el archivo Excel procesado
        """
        # Convertir a BytesIO si es necesario
        if isinstance(archivo_bytes, bytes):
            archivo_bytes = io.BytesIO(archivo_bytes)
        
        # Crear archivo temporal
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_input:
            archivo_bytes.seek(0)
            tmp_input.write(archivo_bytes.read())
            tmp_input_path = Path(tmp_input.name)
        
        try:
            # Procesar con el script original de Triodos
            df_procesado = self.procesador_triodos.leer_y_procesar_triodos(tmp_input_path)
            
            # Crear reporte usando el método del procesador Triodos
            resultado_path = self.procesador_triodos.crear_reporte_excel(tmp_input_path, df_procesado)
            
            # Leer el resultado y convertirlo a BytesIO
            output = io.BytesIO()
            with open(resultado_path, 'rb') as f:
                output.write(f.read())
            output.seek(0)
            
            # Limpiar archivos temporales
            tmp_input_path.unlink(missing_ok=True)
            resultado_path.unlink(missing_ok=True)
            
            return output
            
        except Exception as e:
            # Limpiar archivos temporales en caso de error
            if tmp_input_path.exists():
                tmp_input_path.unlink(missing_ok=True)
            raise Exception(f"Error al procesar archivo Triodos: {str(e)}")
