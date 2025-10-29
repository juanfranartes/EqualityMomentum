
# -*- coding: utf-8 -*-
"""
Procesador Automático de Registros Retributivos - Equality Momentum
Versión optimizada sin redundancia con funciones reutilizables
"""

import sys
import os
import re
from pathlib import Path
from datetime import datetime
import pandas as pd
import numpy as np
import warnings
import logging
import traceback

# Importar tkinter solo si está disponible (entornos con GUI)
try:
    import tkinter as tk
    from tkinter import messagebox
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False

# ==================== CONFIGURACIÓN GLOBAL ====================

# Configurar codificación UTF-8
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except:
        pass

# Configurar logging para capturar warnings en archivo
LOG_DIR = Path(__file__).parent.parent / '03_LOGS'
LOG_DIR.mkdir(exist_ok=True)
LOG_FILE = LOG_DIR / f'procesamiento_{datetime.now().strftime("%Y%m%d")}.log'

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

# Suprimir warnings de pandas en consola (pero se guardan en log)
warnings.filterwarnings('ignore', category=pd.errors.SettingWithCopyWarning)
warnings.filterwarnings('ignore', category=RuntimeWarning)
logging.captureWarnings(True)

# ==================== FUNCIONES AUXILIARES ====================

def log(mensaje, tipo='INFO'):
    """Log estandarizado con prefijos visuales"""
    prefijos = {
        'INFO': '[INFO]',
        'OK': '[✓]',
        'ERROR': '[✗]',
        'WARN': '[!]',
        'PROCESO': '[→]'
    }
    logger = logging.getLogger(__name__)
    mensaje_formateado = f"{prefijos.get(tipo, '[INFO]')} {mensaje}"

    if tipo == 'ERROR':
        logger.error(mensaje)
    elif tipo == 'WARN':
        logger.warning(mensaje)
    else:
        logger.info(mensaje)

    print(mensaje_formateado)

# ==================== CLASE PRINCIPAL ====================

class ProcesadorRegistroRetributivo:
    def __init__(self, validador=None):
        """Inicializa el procesador con las rutas y configuración

        Args:
            validador: Instancia de ValidadorMapeoGeneral con mapeo de hojas y variables (opcional)
        """
        # Obtener ruta base del ejecutable
        if hasattr(sys, '_MEIPASS'):
            self.base_path = Path(sys.executable).parent
        else:
            self.base_path = Path(__file__).parent.parent

        # Definir rutas
        self.carpeta_entrada = self.base_path / "01_DATOS_SIN_PROCESAR"
        self.carpeta_resultados = self.base_path / "02_RESULTADOS"
        self.carpeta_logs = LOG_DIR

        # Crear carpetas si no existen
        self.carpeta_resultados.mkdir(exist_ok=True)

        # Inicializar banner
        log("="*60)
        log("PROCESADOR DE REGISTROS RETRIBUTIVOS - EQUALITY MOMENTUM")
        log("="*60)

        # Guardar validador si se proporcionó
        self.validador = validador

        # Configuración de columnas (nombres exactos del Excel)
        # Si se proporciona validador, usar el mapeo; si no, usar valores por defecto
        self.mapeo_columnas = {
            'meses_trabajados': '¿Cuántos meses ha trabajado?',
            'coef_tp': '% de jornada',
            'salario_base_efectivo': 'Salario base anual efectivo',
            'complementos_salariales_efectivo': 'Complementos Salariales efectivo',
            'complementos_extrasalariales_efectivo': 'Complementos Extrasalariales efectivo'
        }

        # Aplicar mapeo del validador si existe
        if self.validador:
            self.mapeo_columnas = self.validador.obtener_mapeo_completo_variables(self.mapeo_columnas)
        
        # Configuración de complementos
        self.configuracion_complementos = {}

        # Cache para configuraciones de complementos (optimización)
        self._config_cache = {}

        # Cache para columnas de complementos (optimización)
        self._columnas_complementos_cache = None
        
        # Definir columnas permitidas
        self.columnas_permitidas = [
            'Reg.',
            'Orden',
            'Sexo',
            'Inicio de Sit. Contractual',
            'Final de Sit. Contractual',
            '¿Es una persona con discapacidad?',
            'Ascendientes con discapacidad',
            'Grupo profesional',
            'Subgrupo profesional',
            'Nivel Convenio Colectivo',
            'Categoría profesional',
            'Puesto de trabajo',
            'Departamento',
            'Nivel de estudios puesto de origen',
            '% de jornada',
            '¿Cuántos meses ha trabajado?',
            'Coeficiente Horas Trabajadas Efectivo',
            '¿Realiza jornada a turnos?',
            'Salario base anual efectivo',
            'Salario base efectivo Total',
            'Salario base anual + complementos',
            'Salario base anual + complementos Total',
            'Salario base anual + complementos + Extrasalariales',
            'Salario base anual + complementos + Extrasalariales Total',
            'Complementos Salariales efectivo',
            'Complementos Salariales efectivo Total',
            'Complementos Extrasalariales efectivo',
            'Complementos Extrasalariales efectivo Total',
            'Compltos Salariales efectivo Total',
            'Compltos Extrasalariales efectivo Total',
            'Nivel SVPT',
            'Puntos',
            'Convenio',
            'Centro de trabajo',
            'Empresa (si forma parte de grupo de empresas)',
            '¿La persona ha sido cesada en el año de referencia?'
        ]
        
        # Añadir columnas PS1-PS100 y PE1-PE27
        for i in range(1, 101):
            self.columnas_permitidas.append(f'PS{i}')
        for i in range(1, 28):
            self.columnas_permitidas.append(f'PE{i}')
        
    def mostrar_mensaje(self, titulo, mensaje, tipo="info"):
        """Muestra mensajes al usuario con GUI (solo si tkinter está disponible)"""
        log(f"Mensaje usuario: {titulo}", 'INFO' if tipo == 'info' else tipo.upper())

        # Solo mostrar GUI si tkinter está disponible
        if not TKINTER_AVAILABLE:
            return

        root = tk.Tk()
        root.withdraw()

        if tipo == "error":
            messagebox.showerror(titulo, mensaje)
        elif tipo == "warning":
            messagebox.showwarning(titulo, mensaje)
        else:
            messagebox.showinfo(titulo, mensaje)
        root.destroy()
        
    def buscar_archivos_excel(self):
        """Busca todos los archivos Excel en la carpeta de entrada"""
        if not self.carpeta_entrada.exists():
            raise Exception(f"No se encontró la carpeta: {self.carpeta_entrada}")

        # Buscar archivos Excel
        archivos_excel = []
        for patron in ['*.xlsx', '*.xls']:
            archivos_excel.extend(list(self.carpeta_entrada.glob(patron)))

        # Filtrar archivos temporales de Excel (comienzan con ~$)
        archivos_excel = [f for f in archivos_excel if not f.name.startswith('~$')]

        if not archivos_excel:
            raise Exception(f"No se encontraron archivos Excel en: {self.carpeta_entrada}")

        log(f"Archivos Excel encontrados: {len(archivos_excel)}", 'OK')
        for archivo in archivos_excel:
            log(f"  • {archivo.name}")

        return archivos_excel

    def is_positive_response(self, value):
        """Verifica si un valor representa una respuesta positiva (Sí/Si/YES)"""
        if pd.isna(value):
            return False
        normalized = str(value).strip().lower()
        return normalized in ['sí', 'si', 'yes', 'y', '1', 'true']

    def filtrar_columnas_permitidas(self, df):
        """Filtra el DataFrame para mantener solo las columnas permitidas"""
        columnas_actuales = df.columns.tolist()
        
        # Crear lista de columnas a mantener
        columnas_a_mantener = []
        columnas_eliminadas = []
        
        for col in columnas_actuales:
            # Verificar si la columna está en la lista permitida
            if col in self.columnas_permitidas:
                columnas_a_mantener.append(col)
            # También verificar si es una columna PS o PE con formato extendido
            # Por ejemplo: "PS1 Antigüedad" debe mantenerse si "PS1" está permitido
            else:
                # Extraer el código base de columnas tipo "PS 1 Antigüedad" -> "PS1"
                match = re.match(r'^(P[SE])\s*(\d+)', col)
                if match:
                    codigo_base = f"{match.group(1)}{match.group(2)}"
                    if codigo_base in self.columnas_permitidas:
                        columnas_a_mantener.append(col)
                    else:
                        columnas_eliminadas.append(col)
                else:
                    columnas_eliminadas.append(col)
        
        # Filtrar DataFrame
        df_filtrado = df[columnas_a_mantener].copy()
        
        if columnas_eliminadas:
            log(f"Columnas eliminadas (no autorizadas): {len(columnas_eliminadas)}", 'INFO')
            if len(columnas_eliminadas) <= 10:
                for col in columnas_eliminadas:
                    log(f"  • {col}", 'INFO')
            else:
                log(f"  • Mostrando primeras 10: {', '.join(columnas_eliminadas[:10])}...", 'INFO')
        
        log(f"Columnas mantenidas: {len(columnas_a_mantener)}", 'OK')
        
        return df_filtrado

    def _cargar_tipo_complemento(self, archivo_path, nombre_hoja, tipo, nombres_columnas_config):
        """Carga un tipo específico de complementos (salarial o extrasalarial)"""
        configuracion = {}
        try:
            df_comp = pd.read_excel(archivo_path, sheet_name=nombre_hoja)
            # Limpiar nombres de columnas
            df_comp.columns = df_comp.columns.str.strip()
            log(f"Procesando {nombre_hoja}...", 'PROCESO')

            for _, row in df_comp.iterrows():
                codigo_val = row.get(nombres_columnas_config['codigo'])
                if pd.notna(codigo_val):
                    codigo = str(codigo_val).strip()
                    nombre_val = row.get(nombres_columnas_config['nombre'])
                    nombre = str(nombre_val).strip() if pd.notna(nombre_val) else ''

                    configuracion[codigo] = {
                        'tipo': tipo,
                        'nombre': nombre,
                        'es_normalizable': self.is_positive_response(row.get(nombres_columnas_config['normalizable'])),
                        'es_anualizable': self.is_positive_response(row.get(nombres_columnas_config['anualizable']))
                    }

            log(f"Complementos {tipo}s configurados: {len(configuracion)}", 'OK')
        except Exception as e:
            log(f"Error cargando complementos {tipo}s: {e}", 'WARN')

        return configuracion

    def cargar_configuracion_complementos(self, excel_file, archivo_path):
        """Carga la configuración de complementos desde las hojas Excel"""
        log("Cargando configuración de complementos...", 'PROCESO')

        # Usar mapeo del validador si existe
        nombres_columnas_config = {
            'codigo': 'Cod',
            'nombre': 'Nombre',
            'normalizable': '¿Es Normalizable?',
            'anualizable': '¿Es Anualizable?'
        }

        if self.validador:
            nombres_columnas_config = self.validador.columnas_config_complementos

        configuracion = {}

        # Cargar complementos salariales y extrasalariales
        hojas_config = [
            ('COMPLEMENTOS SALARIALES', 'salarial'),
            ('COMPLEMENTOS EXTRASALARIALES', 'extrasalarial')
        ]

        for nombre_hoja_esperado, tipo in hojas_config:
            # Usar el nombre mapeado si existe validador
            nombre_hoja = self.validador.obtener_nombre_hoja(nombre_hoja_esperado) if self.validador else nombre_hoja_esperado

            if nombre_hoja in excel_file.sheet_names:
                config_tipo = self._cargar_tipo_complemento(archivo_path, nombre_hoja, tipo, nombres_columnas_config)
                configuracion.update(config_tipo)

        self.configuracion_complementos = configuracion
        log(f"Total complementos configurados: {len(configuracion)}", 'OK')

        return configuracion

    def _normalizar_valor(self, valor, valor_default):
        """Normaliza un valor, retornando el default si es inválido"""
        return valor_default if pd.isna(valor) or valor == 0 else valor

    def calcular_coef_tp(self, valor_coef_tp):
        """Convierte el coeficiente de tiempo parcial a decimal"""
        if pd.isna(valor_coef_tp):
            return 1.0
        return valor_coef_tp / 100 if valor_coef_tp > 1 else valor_coef_tp

    def equiparar_salario_base(self, salario_base_efectivo, coef_tp, meses_trabajados):
        """Equipara el salario base aplicando normalización y anualización"""
        if pd.isna(salario_base_efectivo) or salario_base_efectivo == 0:
            return 0

        coef_tp = self._normalizar_valor(coef_tp, 1.0)
        meses_trabajados = self._normalizar_valor(meses_trabajados, 12)

        return salario_base_efectivo * (1/coef_tp) * (12/meses_trabajados)

    def equiparar_complemento(self, complemento_efectivo, coef_tp, meses_trabajados, es_normalizable, es_anualizable):
        """Equipara un complemento según su configuración"""
        if pd.isna(complemento_efectivo) or complemento_efectivo == 0 or (not es_normalizable and not es_anualizable):
            return complemento_efectivo if pd.notna(complemento_efectivo) else 0

        resultado = complemento_efectivo

        if es_normalizable:
            resultado *= (1 / self._normalizar_valor(coef_tp, 1.0))

        if es_anualizable:
            resultado *= (12 / self._normalizar_valor(meses_trabajados, 12))

        return resultado

    def obtener_config_complemento(self, codigo_complemento):
        """Obtiene la configuración de un complemento específico (con caché)"""
        # Verificar si ya está en caché
        if codigo_complemento in self._config_cache:
            return self._config_cache[codigo_complemento]

        # Estrategias de búsqueda en orden de prioridad
        codigos_a_buscar = [codigo_complemento]

        # Extraer código base de formatos como "PS 1 Antigüedad" -> "PS1"
        match = re.match(r'^(P[SE])\s*(\d+)', codigo_complemento)
        if match:
            codigo_base = f"{match.group(1)}{match.group(2)}"
            codigos_a_buscar.append(codigo_base)

        # Búsqueda alternativa si es solo un número
        if codigo_complemento.isdigit():
            codigos_a_buscar.append(f"PS{codigo_complemento}")

        # Buscar en orden de prioridad
        for codigo in codigos_a_buscar:
            if codigo in self.configuracion_complementos:
                config = self.configuracion_complementos[codigo]
                resultado = (
                    config['es_normalizable'],
                    config['es_anualizable'],
                    config['tipo'],
                    config.get('nombre', '')
                )
                # Guardar en caché
                self._config_cache[codigo_complemento] = resultado
                return resultado

        # Valores por defecto (también cachear)
        log(f"Configuración no encontrada para {codigo_complemento}", 'WARN')
        resultado = (False, False, 'desconocido', '')
        self._config_cache[codigo_complemento] = resultado
        return resultado

    def equiparar_complemento_individual(self, valor_efectivo, coef_tp, meses_trabajados, codigo_ps):
        """Equipara un complemento individual usando su configuración específica"""
        if pd.isna(valor_efectivo) or valor_efectivo == 0:
            return 0

        es_normalizable, es_anualizable, tipo, _ = self.obtener_config_complemento(codigo_ps)
        return self.equiparar_complemento(valor_efectivo, coef_tp, meses_trabajados, es_normalizable, es_anualizable)

    def leer_y_procesar_excel(self, ruta_archivo):
        """Lee y procesa un archivo Excel completo"""
        log(f"Procesando archivo: {ruta_archivo.name}", 'PROCESO')

        # Limpiar cachés para el nuevo archivo
        self._config_cache.clear()
        self._columnas_complementos_cache = None

        try:
            # Cargar información de hojas disponibles
            excel_file = pd.ExcelFile(ruta_archivo)
            log(f"Hojas disponibles: {excel_file.sheet_names}")

            # Cargar hoja principal (BASE GENERAL)
            # Usar el nombre mapeado si existe validador
            nombre_hoja_principal = self.validador.obtener_nombre_hoja("BASE GENERAL") if self.validador else "BASE GENERAL"

            if nombre_hoja_principal not in excel_file.sheet_names:
                raise Exception(f"No se encontró la hoja '{nombre_hoja_principal}' requerida")

            df = pd.read_excel(ruta_archivo, sheet_name=nombre_hoja_principal)
            log(f"Datos cargados: {df.shape[0]} filas x {df.shape[1]} columnas", 'OK')
            
            # IMPORTANTE: Limpiar nombres de columnas (eliminar espacios al inicio/final)
            df.columns = df.columns.str.strip()
            log("Nombres de columnas limpiados (espacios eliminados)", 'OK')
            
            # Buscar columna "Reg." y eliminar todas las columnas anteriores
            if 'Reg.' in df.columns:
                indice_reg = df.columns.get_loc('Reg.')
                if indice_reg > 0:
                    columnas_a_eliminar = df.columns[:indice_reg].tolist()
                    df = df.drop(columns=columnas_a_eliminar)
                    log(f"Eliminadas {len(columnas_a_eliminar)} columnas anteriores a 'Reg.'", 'OK')
                else:
                    log("La columna 'Reg.' ya es la primera columna")
            else:
                log("No se encontró la columna 'Reg.', no se eliminaron columnas", 'WARN')
            
            # Filtrar columnas para mantener solo las permitidas
            df = self.filtrar_columnas_permitidas(df)
            log(f"Datos después del filtrado: {df.shape[0]} filas x {df.shape[1]} columnas", 'OK')
            
            # Cargar configuración de complementos
            self.cargar_configuracion_complementos(excel_file, ruta_archivo)
            
            # Filtrar datos hasta el último registro válido basado en la columna "Orden"
            if 'Orden' in df.columns:
                # Encontrar el último índice con valor no nulo en la columna "Orden"
                indices_con_orden = df[df['Orden'].notna()].index
                if len(indices_con_orden) > 0:
                    ultimo_indice_valido = indices_con_orden.max()
                    registros_originales = len(df)
                    df = df.iloc[:ultimo_indice_valido + 1].copy()
                    registros_filtrados = len(df)
                    log(f"Filtrado: {registros_originales} → {registros_filtrados} registros", 'OK')
                else:
                    log("No se encontraron datos válidos en 'Orden'", 'WARN')
            else:
                log("No se encontró la columna 'Orden'", 'WARN')
            
            # Verificar columnas críticas
            columnas_encontradas = {}
            for clave, nombre_col in self.mapeo_columnas.items():
                if nombre_col in df.columns:
                    columnas_encontradas[clave] = nombre_col
                    log(f"✓ {clave}: {nombre_col}", 'OK')
                else:
                    log(f"✗ {clave}: '{nombre_col}' NO ENCONTRADA", 'WARN')
            
            if len(columnas_encontradas) < 3:  # Mínimo necesario
                raise Exception(f"Faltan columnas críticas. Encontradas: {list(columnas_encontradas.keys())}")
            
            # Procesar datos de equiparación
            df_equiparado = self.procesar_equiparacion(df, columnas_encontradas)

            log(f"Procesamiento completado: {df_equiparado.shape}", 'OK')
            return df_equiparado

        except Exception as e:
            log(f"Error procesando {ruta_archivo.name}: {str(e)}", 'ERROR')
            raise

    def calcular_complementos_efectivos(self, df):
        """Calcula las columnas de complementos efectivos como suma de PS y PE"""
        log("Calculando complementos efectivos totales...", 'PROCESO')

        # Obtener columnas PS y PE (sin las equiparadas)
        columnas_ps = [col for col in df.columns if re.match(r'^PS\d+', col) and not col.endswith('_equiparado')]
        columnas_pe = [col for col in df.columns if re.match(r'^PE\d+', col) and not col.endswith('_equiparado')]

        # Calcular Complementos Salariales efectivo (suma de todas las PS)
        if 'Complementos Salariales efectivo' not in df.columns and columnas_ps:
            df['Complementos Salariales efectivo'] = df[columnas_ps].fillna(0).sum(axis=1)
            promedio_cs = df['Complementos Salariales efectivo'].mean()
            log(f"Complementos Salariales efectivo calculado: {promedio_cs:.2f} € (suma de {len(columnas_ps)} columnas PS)")
        elif 'Complementos Salariales efectivo' in df.columns:
            log("Complementos Salariales efectivo ya existe en los datos")

        # Calcular Complementos Extrasalariales efectivo (suma de todas las PE)
        if 'Complementos Extrasalariales efectivo' not in df.columns and columnas_pe:
            df['Complementos Extrasalariales efectivo'] = df[columnas_pe].fillna(0).sum(axis=1)
            promedio_ce = df['Complementos Extrasalariales efectivo'].mean()
            log(f"Complementos Extrasalariales efectivo calculado: {promedio_ce:.2f} € (suma de {len(columnas_pe)} columnas PE)")
        elif 'Complementos Extrasalariales efectivo' in df.columns:
            log("Complementos Extrasalariales efectivo ya existe en los datos")

        # Renombrar columnas abreviadas "Compltos" a "Complementos" para consistencia
        if 'Compltos Salariales efectivo Total' in df.columns:
            df.rename(columns={'Compltos Salariales efectivo Total': 'Complementos Salariales efectivo Total'}, inplace=True)
            log("Renombrada: 'Compltos Salariales efectivo Total' → 'Complementos Salariales efectivo Total'")

        if 'Compltos Extrasalariales efectivo Total' in df.columns:
            df.rename(columns={'Compltos Extrasalariales efectivo Total': 'Complementos Extrasalariales efectivo Total'}, inplace=True)
            log("Renombrada: 'Compltos Extrasalariales efectivo Total' → 'Complementos Extrasalariales efectivo Total'")

    def procesar_equiparacion(self, df, columnas_encontradas):
        """Procesa la equiparación de todos los valores"""
        log("Iniciando cálculos de equiparación...", 'PROCESO')
        
        df_equiparado = df.copy()
        
        # Obtener columnas necesarias
        col_meses = columnas_encontradas.get('meses_trabajados')
        col_coef_tp = columnas_encontradas.get('coef_tp')
        col_sb_efectivo = columnas_encontradas.get('salario_base_efectivo')
        
        if not all([col_meses, col_coef_tp, col_sb_efectivo]):
            missing = [k for k, v in [('meses', col_meses), ('coef_tp', col_coef_tp), ('sb_efectivo', col_sb_efectivo)] if not v]
            raise Exception(f"Columnas críticas faltantes: {missing}")
        
        # Calcular coeficiente TP corregido (vectorizado)
        coef_tp_values = df_equiparado[col_coef_tp].fillna(1.0)
        df_equiparado['coef_tp_calculado'] = np.where(
            coef_tp_values > 1,
            coef_tp_values / 100,
            coef_tp_values
        )

        # Equiparar salario base (VECTORIZADO - optimización crítica)
        sb_efectivo = df_equiparado[col_sb_efectivo].fillna(0)
        coef_tp_norm = df_equiparado['coef_tp_calculado'].replace(0, 1.0).fillna(1.0)
        meses_norm = df_equiparado[col_meses].replace(0, 12).fillna(12)

        df_equiparado['salario_base_equiparado'] = np.where(
            sb_efectivo == 0,
            0,
            sb_efectivo * (1 / coef_tp_norm) * (12 / meses_norm)
        )
        
        sb_efectivo_promedio = df_equiparado[col_sb_efectivo].mean()
        sb_equiparado_promedio = df_equiparado['salario_base_equiparado'].mean()
        log(f"SB efectivo promedio: {sb_efectivo_promedio:.2f} €")
        log(f"SB equiparado promedio: {sb_equiparado_promedio:.2f} €")

        # Calcular columnas de complementos efectivos si no existen
        self.calcular_complementos_efectivos(df_equiparado)

        # Procesar complementos individuales
        complementos_procesados = self.procesar_complementos_individuales(df_equiparado, col_meses)
        
        # Calcular totales de complementos equiparados
        self.calcular_totales_complementos(df_equiparado)

        log(f"Complementos individuales procesados: {complementos_procesados}", 'OK')
        return df_equiparado

    def _obtener_columnas_complementos(self, df, prefijos=['PS', 'PE']):
        """Obtiene columnas de complementos por prefijos (con caché)"""
        # Si ya está en caché, retornar directamente
        if self._columnas_complementos_cache is not None:
            return self._columnas_complementos_cache

        columnas_por_tipo = {}
        for prefijo in prefijos:
            columnas_por_tipo[prefijo] = [
                col for col in df.columns
                if (col.startswith(prefijo) and
                    bool(re.match(rf'^{prefijo}\s*\d+', col)) and
                    not col.endswith('_equiparado'))
            ]

        # Guardar en caché
        self._columnas_complementos_cache = columnas_por_tipo
        return columnas_por_tipo

    def procesar_complementos_individuales(self, df_equiparado, col_meses):
        """Procesa todos los complementos PS y PE individuales"""
        if not self.configuracion_complementos:
            log("No hay configuración de complementos disponible", 'WARN')
            return 0

        log("Procesando complementos individuales...", 'PROCESO')
        complementos_procesados = 0

        # Obtener columnas de complementos
        columnas_por_tipo = self._obtener_columnas_complementos(df_equiparado)

        for tipo, columnas in columnas_por_tipo.items():
            log(f"Columnas {tipo} encontradas: {len(columnas)}")

            for col_comp in columnas:
                es_normalizable, es_anualizable, _, nombre_comp = self.obtener_config_complemento(col_comp)

                # Extraer código base (ej: "PS2" de "PS2" o "PS 2")
                match = re.match(r'^(P[SE])\s*(\d+)', col_comp)
                if match:
                    codigo_base = f"{match.group(1)}{match.group(2)}"
                else:
                    codigo_base = col_comp

                # Solo procesar (equiparar) si es normalizable O anualizable
                if es_normalizable or es_anualizable:
                    datos_no_nulos = df_equiparado[col_comp].dropna()
                    if len(datos_no_nulos) > 0:
                        # Usar código base para la columna equiparada (ej: "PS2_equiparado")
                        col_equiparado = f"{codigo_base}_equiparado"

                        nombre_display = f"{codigo_base} {nombre_comp}" if nombre_comp else codigo_base
                        log(f"  {nombre_display}: {len(datos_no_nulos)} registros (N:{es_normalizable}, A:{es_anualizable})")

                        # VECTORIZACIÓN: Calcular complemento equiparado sin apply
                        comp_efectivo = df_equiparado[col_comp].fillna(0)
                        resultado = comp_efectivo.copy()

                        # Aplicar normalización si corresponde
                        if es_normalizable:
                            coef_tp_norm = df_equiparado['coef_tp_calculado'].replace(0, 1.0).fillna(1.0)
                            resultado = resultado * (1 / coef_tp_norm)

                        # Aplicar anualización si corresponde
                        if es_anualizable:
                            meses_norm = df_equiparado[col_meses].replace(0, 12).fillna(12)
                            resultado = resultado * (12 / meses_norm)

                        # Mantener 0 donde el valor efectivo era 0 o NaN
                        df_equiparado[col_equiparado] = np.where(
                            comp_efectivo == 0,
                            0,
                            resultado
                        )

                        complementos_procesados += 1

                        prom_efectivo = df_equiparado[col_comp].mean()
                        prom_equiparado = df_equiparado[col_equiparado].mean()
                        log(f"    Efectivo: {prom_efectivo:.2f} € | Equiparado: {prom_equiparado:.2f} €")

        return complementos_procesados

    def _calcular_total_correcto(self, row, columnas_base, df_equiparado):
        """Calcula total: equiparado si es procesable, efectivo si no"""
        total = 0
        for col_base in columnas_base:
            # Verificar que la columna existe
            if col_base not in row.index:
                continue

            valor_original = row[col_base]
            if pd.notna(valor_original) and valor_original != 0:
                # Extraer código base para buscar la columna equiparada
                match = re.match(r'^(P[SE])\s*(\d+)', col_base)
                if match:
                    codigo_base = f"{match.group(1)}{match.group(2)}"
                else:
                    codigo_base = col_base

                col_equiparada = f"{codigo_base}_equiparado"
                if col_equiparada in df_equiparado.columns:
                    valor_equiparado = row[col_equiparada]
                    total += valor_equiparado if pd.notna(valor_equiparado) else valor_original
                else:
                    total += valor_original
        return total

    def calcular_totales_complementos(self, df_equiparado):
        """Calcula los totales correctos de complementos equiparados"""
        log("Calculando totales de complementos...", 'PROCESO')

        # Obtener columnas una sola vez (ya usa caché internamente)
        columnas_por_tipo = self._obtener_columnas_complementos(df_equiparado)

        # Calcular totales
        df_equiparado['complementos_salariales_equiparados'] = df_equiparado.apply(
            lambda row: self._calcular_total_correcto(row, columnas_por_tipo['PS'], df_equiparado), axis=1
        )

        df_equiparado['complementos_extrasalariales_equiparados'] = df_equiparado.apply(
            lambda row: self._calcular_total_correcto(row, columnas_por_tipo['PE'], df_equiparado), axis=1
        )

        cs_promedio = df_equiparado['complementos_salariales_equiparados'].mean()
        ce_promedio = df_equiparado['complementos_extrasalariales_equiparados'].mean()
        log(f"CS equiparados promedio: {cs_promedio:.2f} €", 'OK')
        log(f"CE equiparados promedio: {ce_promedio:.2f} €", 'OK')

        # Calcular columnas combinadas adicionales
        self.calcular_columnas_combinadas(df_equiparado)

    def calcular_columnas_combinadas(self, df_equiparado):
        """Calcula las columnas combinadas necesarias para las visualizaciones"""
        log("Calculando columnas combinadas...", 'PROCESO')
        
        # sb_mas_comp_equiparado = salario_base_equiparado + complementos_salariales_equiparados
        if 'salario_base_equiparado' in df_equiparado.columns and 'complementos_salariales_equiparados' in df_equiparado.columns:
            df_equiparado['sb_mas_comp_salariales_equiparado'] = (
                df_equiparado['salario_base_equiparado'].fillna(0) + 
                df_equiparado['complementos_salariales_equiparados'].fillna(0)
            )
            promedio_sb_comp = df_equiparado['sb_mas_comp_salariales_equiparado'].mean()
            log(f"SB + Comp. Salariales promedio: {promedio_sb_comp:.2f} €")
        
        # sb_mas_comp_extrasal_equiparado = sb_mas_comp_equiparado + complementos_extrasalariales_equiparados
        if ('sb_mas_comp_salariales_equiparado' in df_equiparado.columns and 
            'complementos_extrasalariales_equiparados' in df_equiparado.columns):
            df_equiparado['sb_mas_comp_total_equiparado'] = (
                df_equiparado['sb_mas_comp_salariales_equiparado'].fillna(0) + 
                df_equiparado['complementos_extrasalariales_equiparados'].fillna(0)
            )
            promedio_total = df_equiparado['sb_mas_comp_total_equiparado'].mean()
            log(f"SB + Comp. Total promedio: {promedio_total:.2f} €")

    def crear_reporte_excel(self, archivo_original, df_procesado):
        """Crea el archivo Excel de resultados simplificado"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_base = archivo_original.stem
        nombre_resultado = f"REPORTE_{nombre_base}_{timestamp}.xlsx"
        ruta_resultado = self.carpeta_resultados / nombre_resultado

        log(f"Creando reporte: {nombre_resultado}", 'PROCESO')

        try:
            # Asegurar que las columnas estén limpias antes de guardar
            df_procesado.columns = df_procesado.columns.str.strip()
            
            # Escribir archivo Excel con solo los datos procesados
            with pd.ExcelWriter(ruta_resultado, engine='openpyxl') as writer:
                # Solo datos procesados - sin hojas adicionales
                df_procesado.to_excel(writer, sheet_name='DATOS_PROCESADOS', index=False)

            # Verificar archivo creado
            if ruta_resultado.exists():
                tamano_mb = ruta_resultado.stat().st_size / (1024 * 1024)
                log(f"Reporte creado: {tamano_mb:.2f} MB", 'OK')
                return ruta_resultado
            else:
                raise Exception("El archivo no se pudo crear correctamente")

        except Exception as e:
            log(f"Error creando reporte: {str(e)}", 'ERROR')
            raise

    def procesar_archivo(self, ruta_archivo):
        """Procesa un archivo Excel individual"""
        inicio_tiempo = datetime.now()

        try:
            log(f"INICIANDO: {ruta_archivo.name}", 'PROCESO')
            
            # Procesar datos
            df_procesado = self.leer_y_procesar_excel(ruta_archivo)
            
            # Crear reporte
            ruta_resultado = self.crear_reporte_excel(ruta_archivo, df_procesado)
            
            # Calcular tiempo transcurrido
            tiempo_total = (datetime.now() - inicio_tiempo).total_seconds()

            log(f"COMPLETADO: {ruta_archivo.name} en {tiempo_total:.1f}s", 'OK')
            
            return {
                'archivo_original': ruta_archivo.name,
                'archivo_resultado': ruta_resultado.name,
                'registros_procesados': len(df_procesado),
                'tiempo_procesamiento': tiempo_total,
                'estado': 'ÉXITO'
            }
            
        except Exception as e:
            tiempo_total = (datetime.now() - inicio_tiempo).total_seconds()
            error_msg = str(e)

            log(f"ERROR en {ruta_archivo.name}: {error_msg}", 'ERROR')
            logging.error(f"Detalles técnicos: {traceback.format_exc()}")
            
            return {
                'archivo_original': ruta_archivo.name,
                'archivo_resultado': None,
                'registros_procesados': 0,
                'tiempo_procesamiento': tiempo_total,
                'estado': 'ERROR',
                'error': error_msg
            }

    def ejecutar_procesamiento(self, archivo_especifico=None):
        """Función principal que ejecuta todo el procesamiento

        Args:
            archivo_especifico: Ruta completa del archivo específico a procesar (opcional)
                               Si no se proporciona, procesa todos los archivos en la carpeta
        """
        inicio_total = datetime.now()

        try:
            # Si se proporciona un archivo específico, solo procesar ese
            if archivo_especifico:
                archivo_path = Path(archivo_especifico)
                if not archivo_path.exists():
                    raise Exception(f"El archivo especificado no existe: {archivo_especifico}")

                log(f"Procesando archivo específico: {archivo_path.name}", 'PROCESO')
                archivos_excel = [archivo_path]
            else:
                log("Buscando archivos Excel...", 'PROCESO')
                archivos_excel = self.buscar_archivos_excel()

            # Procesar cada archivo
            resultados = []
            for i, archivo in enumerate(archivos_excel, 1):
                log(f"\nArchivo {i}/{len(archivos_excel)}: {archivo.name}", 'PROCESO')
                resultado = self.procesar_archivo(archivo)
                resultados.append(resultado)
            
            # Crear resumen final
            exitosos = [r for r in resultados if r['estado'] == 'ÉXITO']
            errores = [r for r in resultados if r['estado'] == 'ERROR']
            
            tiempo_total_proceso = (datetime.now() - inicio_total).total_seconds()

            # Log resumen final
            log("="*60)
            log("RESUMEN FINAL DEL PROCESAMIENTO")
            log("="*60)
            log(f"Archivos exitosos: {len(exitosos)}", 'OK')
            log(f"Archivos con errores: {len(errores)}", 'ERROR' if errores else 'INFO')
            log(f"Tiempo total: {tiempo_total_proceso:.1f}s")

            if exitosos:
                log("\nArchivos exitosos:")
                total_registros = 0
                for r in exitosos:
                    total_registros += r['registros_procesados']
                    log(f"  ✓ {r['archivo_original']} → {r['archivo_resultado']} ({r['registros_procesados']} registros)", 'OK')
                log(f"Total registros procesados: {total_registros}", 'OK')

            if errores:
                log("\nArchivos con errores:")
                for r in errores:
                    log(f"  ✗ {r['archivo_original']}: {r['error']}", 'ERROR')
            
            # Mostrar mensaje final al usuario
            self.mostrar_mensaje_final(exitosos, errores, tiempo_total_proceso)
                
            return len(errores) == 0  # True si no hay errores
            
        except Exception as e:
            tiempo_total_proceso = (datetime.now() - inicio_total).total_seconds()
            error_msg = f"Error crítico en el procesamiento:\n\n{str(e)}\n\nTiempo transcurrido: {tiempo_total_proceso:.1f}s"

            log(f"ERROR CRÍTICO: {error_msg}", 'ERROR')
            logging.error(f"Detalles técnicos: {traceback.format_exc()}")

            self.mostrar_mensaje("Error Crítico", error_msg, "error")
            return False

    def mostrar_mensaje_final(self, exitosos, errores, tiempo_total):
        """Muestra el mensaje final al usuario"""
        if errores:
            mensaje_final = f"""PROCESAMIENTO COMPLETADO CON ADVERTENCIAS

⏱️  Tiempo total: {tiempo_total:.1f} segundos
✅ Archivos exitosos: {len(exitosos)}
❌ Archivos con errores: {len(errores)}

📁 Los reportes están guardados en:
{self.carpeta_resultados}

📋 Los logs detallados están en:
{self.carpeta_logs}

⚠️  Revisar archivos con errores en los logs."""
            
            self.mostrar_mensaje("Procesamiento Completado", mensaje_final, "warning")
        else:
            total_registros = sum([r['registros_procesados'] for r in exitosos])
            mensaje_final = f"""¡PROCESAMIENTO COMPLETADO EXITOSAMENTE!

⏱️  Tiempo total: {tiempo_total:.1f} segundos
📊 Archivos procesados: {len(exitosos)}
👥 Total de registros: {total_registros}

📁 Todos los reportes están guardados en:
{self.carpeta_resultados}

📋 Los logs están en:
{self.carpeta_logs}

✨ Puede revisar los resultados abriendo los archivos Excel generados."""
            
            self.mostrar_mensaje("¡Procesamiento Exitoso!", mensaje_final)


def main():
    """Función principal del programa"""
    # Detectar si se ejecuta desde workflow
    ejecutado_desde_workflow = "--workflow" in sys.argv

    # Detectar si se pasa un archivo específico como argumento
    archivo_especifico = None
    if len(sys.argv) > 1 and not sys.argv[1].startswith('--'):
        archivo_especifico = sys.argv[1]

    try:
        # Crear instancia del procesador
        procesador = ProcesadorRegistroRetributivo()

        # Ejecutar procesamiento
        exito = procesador.ejecutar_procesamiento(archivo_especifico)

        # Solo pausar si se ejecuta directamente (no desde workflow)
        if not ejecutado_desde_workflow:
            input("\nPresiona Enter para cerrar...")

        # Código de salida
        sys.exit(0 if exito else 1)

    except KeyboardInterrupt:
        print("\n\nProcesamiento interrumpido por el usuario.")
        if not ejecutado_desde_workflow:
            input("Presiona Enter para cerrar...")
        sys.exit(1)

    except Exception as e:
        print(f"\nError crítico no manejado: {str(e)}")
        print("\nContacte con soporte técnico.")
        if not ejecutado_desde_workflow:
            input("Presiona Enter para cerrar...")
        sys.exit(1)


if __name__ == "__main__":
    main()