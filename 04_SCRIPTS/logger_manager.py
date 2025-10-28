"""
Sistema de Logging Centralizado para EqualityMomentum
Captura todos los errores y genera reportes estructurados para el desarrollador
"""

import logging
import sys
import os
import traceback
import platform
import json
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path


class LoggerManager:
    """Gestor centralizado de logs con captura de errores estructurada"""

    def __init__(self, app_name="EqualityMomentum", version="1.0.0"):
        self.app_name = app_name
        self.version = version
        self.logs_dir = self._get_logs_directory()
        self.logger = None
        self._setup_logger()

    def _get_logs_directory(self):
        """Determina el directorio de logs según el entorno"""
        # Intentar usar la carpeta de logs del proyecto
        project_logs = Path(__file__).parent.parent / "03_LOGS"
        if project_logs.exists():
            return project_logs

        # Si no existe, usar carpeta de usuario
        user_docs = Path.home() / "Documents" / "EqualityMomentum" / "Logs"
        user_docs.mkdir(parents=True, exist_ok=True)
        return user_docs

    def _setup_logger(self):
        """Configura el logger con formato estructurado"""
        self.logger = logging.getLogger(self.app_name)
        self.logger.setLevel(logging.DEBUG)

        # Evitar duplicados
        if self.logger.handlers:
            return

        # Handler para archivo con rotación
        log_file = self.logs_dir / f"app_{datetime.now().strftime('%Y%m%d')}.log"
        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=10*1024*1024,  # 10 MB
            backupCount=5,
            encoding='utf-8'
        )
        file_handler.setLevel(logging.DEBUG)

        # Handler para consola (solo errores)
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.WARNING)

        # Formato detallado
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - [%(levelname)s] - %(filename)s:%(lineno)d - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )

        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)

        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)

        # Log de inicio
        self.logger.info(f"=" * 80)
        self.logger.info(f"{self.app_name} v{self.version} - Inicio de sesión")
        self.logger.info(f"Sistema: {platform.system()} {platform.version()}")
        self.logger.info(f"Python: {sys.version}")
        self.logger.info(f"Directorio de logs: {self.logs_dir}")
        self.logger.info(f"=" * 80)

    def log_exception(self, exc_type, exc_value, exc_traceback):
        """Captura y registra excepciones no manejadas"""
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return

        self.logger.critical("Excepción no manejada:", exc_info=(exc_type, exc_value, exc_traceback))

        # Generar reporte de error para el desarrollador
        self._generate_error_report(exc_type, exc_value, exc_traceback)

    def _generate_error_report(self, exc_type, exc_value, exc_traceback):
        """Genera un reporte estructurado de error para enviar al desarrollador"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        report_file = self.logs_dir / f"ERROR_REPORT_{timestamp}.json"

        # Información del sistema
        system_info = {
            "os": platform.system(),
            "os_version": platform.version(),
            "os_release": platform.release(),
            "machine": platform.machine(),
            "processor": platform.processor(),
            "python_version": sys.version,
            "python_implementation": platform.python_implementation(),
        }

        # Información de la aplicación
        app_info = {
            "name": self.app_name,
            "version": self.version,
            "timestamp": datetime.now().isoformat(),
        }

        # Información del error
        error_info = {
            "type": exc_type.__name__,
            "message": str(exc_value),
            "traceback": traceback.format_exception(exc_type, exc_value, exc_traceback)
        }

        # Reporte completo
        report = {
            "system_info": system_info,
            "app_info": app_info,
            "error_info": error_info
        }

        # Guardar reporte
        try:
            with open(report_file, 'w', encoding='utf-8') as f:
                json.dump(report, f, indent=2, ensure_ascii=False)

            self.logger.info(f"Reporte de error generado: {report_file}")

            # También generar versión legible
            txt_report = report_file.with_suffix('.txt')
            with open(txt_report, 'w', encoding='utf-8') as f:
                f.write(f"REPORTE DE ERROR - {self.app_name} v{self.version}\n")
                f.write(f"{'=' * 80}\n\n")
                f.write(f"FECHA Y HORA: {app_info['timestamp']}\n\n")
                f.write(f"INFORMACIÓN DEL SISTEMA:\n")
                f.write(f"{'-' * 80}\n")
                for key, value in system_info.items():
                    f.write(f"{key}: {value}\n")
                f.write(f"\n")
                f.write(f"INFORMACIÓN DEL ERROR:\n")
                f.write(f"{'-' * 80}\n")
                f.write(f"Tipo: {error_info['type']}\n")
                f.write(f"Mensaje: {error_info['message']}\n\n")
                f.write(f"Stack Trace:\n")
                f.write(''.join(error_info['traceback']))
                f.write(f"\n{'=' * 80}\n")
                f.write(f"Por favor, envíe este archivo al equipo de desarrollo.\n")

            self.logger.info(f"Reporte legible generado: {txt_report}")

        except Exception as e:
            self.logger.error(f"Error al generar reporte: {e}")

    def get_logger(self):
        """Retorna el logger configurado"""
        return self.logger

    def setup_exception_handler(self):
        """Configura el manejador global de excepciones"""
        sys.excepthook = self.log_exception


# Singleton global
_logger_manager_instance = None

def get_logger_manager(app_name="EqualityMomentum", version="1.0.0"):
    """Obtiene la instancia del gestor de logs (singleton)"""
    global _logger_manager_instance
    if _logger_manager_instance is None:
        _logger_manager_instance = LoggerManager(app_name, version)
    return _logger_manager_instance


if __name__ == "__main__":
    # Test del logger
    lm = get_logger_manager()
    lm.setup_exception_handler()

    logger = lm.get_logger()
    logger.info("Test de logging")
    logger.warning("Test de warning")
    logger.error("Test de error")

    # Test de excepción
    try:
        x = 1 / 0
    except Exception as e:
        logger.exception("Test de excepción capturada")

    print(f"Logs guardados en: {lm.logs_dir}")
