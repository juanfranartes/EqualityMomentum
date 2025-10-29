# -*- coding: utf-8 -*-
"""
Procesador Automático de Registros Retributivos - TRIODOS BANK
Adaptado del procesador general para el formato específico de Triodos
"""

import sys
import os
import re
import io
from pathlib import Path
from datetime import datetime
from dateutil.relativedelta import relativedelta
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

import msoffcrypto

# ==================== CONFIGURACIÓN GLOBAL ====================

# Configurar codificación UTF-8
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except:
        pass

# Configurar logging
LOG_DIR = Path(__file__).parent.parent / '03_LOGS'
LOG_DIR.mkdir(exist_ok=True)
LOG_FILE = LOG_DIR / f'procesamiento_triodos_{datetime.now().strftime("%Y%m%d")}.log'

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

warnings.filterwarnings('ignore', category=pd.errors.SettingWithCopyWarning)
warnings.filterwarnings('ignore', category=RuntimeWarning)
logging.captureWarnings(True)

# Contraseña del archivo Triodos
TRIODOS_PASSWORD = 'Triodos2025'

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

class ProcesadorTriodos:
    def __init__(self):
        """Inicializa el procesador para Triodos"""
        # Obtener ruta base
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
        log("PROCESADOR TRIODOS BANK - EQUALITY MOMENTUM")
        log("="*60)

        # Mapeo de columnas Triodos → Formato Maestro
        self.mapeo_columnas = {
            'num_personal': 'Nº personal',
            'sexo': 'Sexo',
            'fecha_inicio_sit': 'Fecha inicio sit.',
            'fecha_fin_sit': 'Fecha fin sit.',
            'grupo_prof': 'Grupo prof.',
            'clasif_interna': 'Clasif. interna',
            'valoracion_puesto': 'Valoración puesto',
            'puesto': 'Puesto',
            'departamento': 'Departamento',
            'jornada_pct': '% Jornada',
            'reduccion_pct': '% Reducción',
            'salario_base_efectivo': 'A154-Salario base de nivel*CT',
            'bruto_pagado': 'Bruto pagado'
        }

        # Configuración de complementos
        self.configuracion_complementos = {}
        self._config_cache = {}
        self._columnas_complementos_cache = None

        # Contraseña para archivos protegidos (puede ser configurada externamente)
        self.password = TRIODOS_PASSWORD

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

    def buscar_archivo_triodos(self):
        """Busca el archivo Triodos.xlsx en la carpeta de entrada"""
        if not self.carpeta_entrada.exists():
            raise Exception(f"No se encontró la carpeta: {self.carpeta_entrada}")

        archivo_triodos = self.carpeta_entrada / "Triodos.xlsx"

        if not archivo_triodos.exists():
            raise Exception(f"No se encontró el archivo Triodos.xlsx en: {self.carpeta_entrada}")

        log(f"Archivo Triodos encontrado: {archivo_triodos.name}", 'OK')
        return archivo_triodos

    def abrir_archivo_protegido(self, ruta_archivo):
        """Abre un archivo Excel protegido con contraseña"""
        try:
            log("Desencriptando archivo protegido...", 'PROCESO')
            with open(ruta_archivo, 'rb') as f:
                office_file = msoffcrypto.OfficeFile(f)
                office_file.load_key(password=self.password)

                decrypted = io.BytesIO()
                office_file.decrypt(decrypted)
                decrypted.seek(0)

            log("Archivo desencriptado correctamente", 'OK')
            return decrypted
        except Exception as e:
            log(f"Error al desencriptar archivo: {str(e)}", 'ERROR')
            raise

    def calcular_meses_trabajados(self, fecha_inicio, fecha_fin):
        """
        Calcula los meses trabajados con precisión decimal.

        Fórmula del usuario:
        meses = (días_inicio × 12/365) + meses_completos + (días_fin × 12/365)

        Ejemplo 24/01/2024 al 15/04/2024:
        - Días enero: 8 días (del 24 al 31)
        - Meses completos: febrero y marzo = 2 meses
        - Días abril: 15 días (del 1 al 15)
        - Total = (8×12/365) + 2 + (15×12/365) = 2.756 meses
        """
        if pd.isna(fecha_inicio) or pd.isna(fecha_fin):
            return 12.0  # Por defecto

        from calendar import monthrange

        # Convertir a datetime si es necesario
        if isinstance(fecha_inicio, pd.Timestamp):
            fecha_inicio = fecha_inicio.to_pydatetime()
        if isinstance(fecha_fin, pd.Timestamp):
            fecha_fin = fecha_fin.to_pydatetime()

        # CASO ESPECIAL: Período completo de meses enteros (del 1 al último día)
        if fecha_inicio.day == 1 and fecha_fin.day == monthrange(fecha_fin.year, fecha_fin.month)[1]:
            delta = relativedelta(fecha_fin, fecha_inicio)
            return float(delta.years * 12 + delta.months + 1)

        # CASO GENERAL: Calcular días parciales
        ultimo_dia_mes_inicio = monthrange(fecha_inicio.year, fecha_inicio.month)[1]
        ultimo_dia_mes_fin = monthrange(fecha_fin.year, fecha_fin.month)[1]

        # 1. Días del mes de inicio (si no es día 1)
        if fecha_inicio.day == 1:
            dias_inicio = 0
            mes_inicio_es_completo = True
        else:
            dias_inicio = ultimo_dia_mes_inicio - fecha_inicio.day + 1
            mes_inicio_es_completo = False

        # 2. Días del mes de fin (si no es último día del mes)
        if fecha_fin.day == ultimo_dia_mes_fin:
            dias_fin = 0
            mes_fin_es_completo = True
        else:
            dias_fin = fecha_fin.day
            mes_fin_es_completo = False

        # 3. Contar meses completos entre el mes de inicio y el mes de fin
        delta = relativedelta(fecha_fin, fecha_inicio)
        total_meses_diff = delta.years * 12 + delta.months

        # Ajustar según si los meses parciales son completos o no
        if mes_inicio_es_completo and mes_fin_es_completo:
            # Ambos meses son completos, contar todos los meses
            meses_completos = total_meses_diff + 1
        elif mes_inicio_es_completo:
            # Solo el inicio es completo, contar desde inicio hasta el mes anterior al fin
            meses_completos = total_meses_diff
        elif mes_fin_es_completo:
            # Solo el fin es completo, contar desde el mes siguiente al inicio hasta el fin
            meses_completos = total_meses_diff
        else:
            # Ninguno es completo, contar solo los meses intermedios
            meses_completos = total_meses_diff - 1 if total_meses_diff > 0 else 0

        # Aplicar fórmula: (días_inicio × 12/365) + meses_completos + (días_fin × 12/365)
        meses = (dias_inicio * 12.0 / 365.0) + meses_completos + (dias_fin * 12.0 / 365.0)

        # Limitar entre 0.01 y 12
        return max(0.01, min(12.0, meses))

    def asignar_reg_por_empleado(self, df):
        """
        Asigna valores a la columna 'Reg.' según la lógica:
        - Si un empleado tiene múltiples situaciones contractuales:
          * 'Ex' para las situaciones antiguas (todas excepto la última)
          * '' (vacío) para la situación más reciente
        - Si un empleado tiene una sola situación:
          * '' (vacío)
        """
        log("Asignando valores a columna 'Reg.'...", 'PROCESO')

        # Inicializar con cadena vacía y convertir columna a tipo string
        df['Reg.'] = pd.Series([''] * len(df), index=df.index, dtype='object')

        # Agrupar por número de personal
        for num_personal in df[self.mapeo_columnas['num_personal']].unique():
            if pd.isna(num_personal):
                continue

            # Obtener todas las filas de este empleado
            mask_empleado = df[self.mapeo_columnas['num_personal']] == num_personal
            indices_empleado = df[mask_empleado].index.tolist()

            # Si tiene más de una situación contractual
            if len(indices_empleado) > 1:
                # Ordenar por fecha de fin (la más reciente al final)
                filas_empleado = df.loc[indices_empleado].copy()
                filas_empleado_sorted = filas_empleado.sort_values(
                    by=self.mapeo_columnas['fecha_fin_sit'],
                    na_position='last'
                )

                # Marcar todas excepto la última como 'Ex'
                indices_antiguos = filas_empleado_sorted.index[:-1]
                df.loc[indices_antiguos, 'Reg.'] = 'Ex'

                # Asegurar explícitamente que la última tiene cadena vacía (no NaN)
                indice_ultimo = filas_empleado_sorted.index[-1]
                df.at[indice_ultimo, 'Reg.'] = ''

        num_ex = (df['Reg.'] == 'Ex').sum()
        num_vacios = (df['Reg.'] == '').sum()
        log(f"Registros marcados como 'Ex': {num_ex}, Vacíos: {num_vacios}", 'OK')

    def is_positive_response(self, value):
        """Verifica si un valor representa una respuesta positiva"""
        if pd.isna(value):
            return False
        normalized = str(value).strip().lower()
        return normalized in ['sí', 'si', 'yes', 'y', '1', 'true']

    def cargar_configuracion_complementos(self, archivo_decrypted):
        """Carga la configuración de complementos desde las hojas Excel"""
        log("Cargando configuración de complementos de Triodos...", 'PROCESO')

        nombres_columnas_config = {
            'codigo': 'Cod',
            'nombre': 'Nombre',
            'normalizable': '¿Es Normalizable?',
            'anualizable': '¿Es Anualizable?'
        }

        configuracion = {}

        # Cargar complementos salariales y extrasalariales
        hojas_config = [
            ('COMPLEMENTOS SALARIALES', 'salarial'),
            ('COMPLEMENTOS EXTRASALARIALES', 'extrasalarial')
        ]

        for nombre_hoja, tipo in hojas_config:
            try:
                df_comp = pd.read_excel(archivo_decrypted, sheet_name=nombre_hoja, engine='openpyxl')
                archivo_decrypted.seek(0)
                # Limpiar nombres de columnas
                df_comp.columns = df_comp.columns.str.strip()
                log(f"Procesando {nombre_hoja}...", 'PROCESO')

                for _, row in df_comp.iterrows():
                    # En Triodos, la columna 'Nombre' contiene el código real (ej: A001-Trienios)
                    # y 'Cod' contiene PS1, PS2, etc.
                    nombre_val = row.get(nombres_columnas_config['nombre'])

                    if pd.notna(nombre_val):
                        nombre_completo = str(nombre_val).strip()

                        # Extraer el código A### de la columna Nombre
                        # Formato: "A001-Trienios" -> usamos "A001" como clave
                        codigo_a = nombre_completo.split('-')[0].strip() if '-' in nombre_completo else nombre_completo

                        configuracion[codigo_a] = {
                            'tipo': tipo,
                            'nombre_completo': nombre_completo,
                            'es_normalizable': self.is_positive_response(row.get(nombres_columnas_config['normalizable'])),
                            'es_anualizable': self.is_positive_response(row.get(nombres_columnas_config['anualizable']))
                        }

                log(f"Complementos {tipo}s configurados: {len([c for c in configuracion.values() if c['tipo'] == tipo])}", 'OK')
            except Exception as e:
                log(f"Error cargando complementos {tipo}s: {e}", 'WARN')

        self.configuracion_complementos = configuracion
        log(f"Total complementos configurados: {len(configuracion)}", 'OK')

        return configuracion

    def obtener_config_complemento(self, codigo_complemento):
        """Obtiene la configuración de un complemento específico (con caché)"""
        if codigo_complemento in self._config_cache:
            return self._config_cache[codigo_complemento]

        # Buscar directamente por el código
        if codigo_complemento in self.configuracion_complementos:
            config = self.configuracion_complementos[codigo_complemento]
            resultado = (
                config['es_normalizable'],
                config['es_anualizable'],
                config['tipo'],
                config.get('nombre_completo', '')
            )
            self._config_cache[codigo_complemento] = resultado
            return resultado

        # Valores por defecto
        log(f"Configuración no encontrada para {codigo_complemento}", 'WARN')
        resultado = (False, False, 'desconocido', '')
        self._config_cache[codigo_complemento] = resultado
        return resultado

    def calcular_coef_tp(self, valor_jornada):
        """
        Calcula el Coeficiente Horas Trabajadas Efectivo.
        Fórmula CORRECTA: coef_tp = % jornada en decimal

        Si % jornada = 80%, entonces coef_tp = 0.8
        Si % jornada = 100%, entonces coef_tp = 1.0

        Args:
            valor_jornada: Porcentaje de jornada (ej: 80 para 80%)
        """
        if pd.isna(valor_jornada):
            jornada = 100.0
        else:
            jornada = valor_jornada

        # Convertir jornada a decimal si está en porcentaje
        jornada_dec = jornada / 100.0 if jornada > 1 else jornada

        return jornada_dec

    def _normalizar_valor(self, valor, valor_default):
        """Normaliza un valor, retornando el default si es inválido"""
        return valor_default if pd.isna(valor) or valor == 0 else valor

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

    def leer_y_procesar_triodos(self, ruta_archivo):
        """Lee y procesa el archivo Triodos"""
        log(f"Procesando archivo: {ruta_archivo.name}", 'PROCESO')

        try:
            # Abrir archivo protegido
            archivo_decrypted = self.abrir_archivo_protegido(ruta_archivo)

            # Cargar hoja principal (BASE GENERAL)
            df = pd.read_excel(archivo_decrypted, sheet_name="BASE GENERAL", engine='openpyxl')
            archivo_decrypted.seek(0)
            log(f"Datos cargados: {df.shape[0]} filas x {df.shape[1]} columnas", 'OK')
            
            # IMPORTANTE: Limpiar nombres de columnas (eliminar espacios al inicio/final)
            df.columns = df.columns.str.strip()
            log("Nombres de columnas limpiados (espacios eliminados)", 'OK')

            # Cargar configuración de complementos
            self.cargar_configuracion_complementos(archivo_decrypted)

            # CRÍTICO: Filtrar filas de totales (tienen Nº personal pero NO tienen fechas)
            # Estas son las filas que Triodos agrega al final de cada empleado
            log("Filtrando filas de totales...", 'PROCESO')
            filas_originales = len(df)

            # Filtrar: Mantener solo filas CON fechas O la fila de total general (última, sin Nº personal)
            df = df[
                (df[self.mapeo_columnas['fecha_inicio_sit']].notna()) |
                (df[self.mapeo_columnas['num_personal']].isna())
            ].copy()

            filas_filtradas = len(df)
            log(f"Filas filtradas: {filas_originales} → {filas_filtradas} (eliminadas {filas_originales - filas_filtradas} filas de totales)", 'OK')

            # Calcular meses trabajados
            log("Calculando meses trabajados...", 'PROCESO')
            df['¿Cuántos meses ha trabajado?'] = df.apply(
                lambda row: self.calcular_meses_trabajados(
                    row[self.mapeo_columnas['fecha_inicio_sit']],
                    row[self.mapeo_columnas['fecha_fin_sit']]
                ),
                axis=1
            )

            # Mapear columnas básicas al formato maestro
            log("Mapeando columnas al formato maestro...", 'PROCESO')
            df['Orden'] = df[self.mapeo_columnas['num_personal']]

            # Convertir Sexo: Masculino/Femenino → Hombres/Mujeres
            df['Sexo'] = df[self.mapeo_columnas['sexo']].map({
                'Masculino': 'Hombres',
                'Femenino': 'Mujeres'
            })

            df['Inicio de Sit. Contractual'] = df[self.mapeo_columnas['fecha_inicio_sit']]
            df['Final de Sit. Contractual'] = df[self.mapeo_columnas['fecha_fin_sit']]

            # Capitalizar valores de texto (excepto Grupo profesional)
            df['Grupo profesional'] = df[self.mapeo_columnas['grupo_prof']].astype(str)
            df['Categoría profesional'] = df[self.mapeo_columnas['clasif_interna']].astype(str).str.title()
            df['Puesto de trabajo'] = df[self.mapeo_columnas['puesto']].astype(str).str.title()
            df['Nivel Convenio Colectivo'] = df[self.mapeo_columnas['clasif_interna']].astype(str).str.title()
            df['Departamento'] = df[self.mapeo_columnas['departamento']].astype(str)
            df['Nivel SVPT'] = df[self.mapeo_columnas['valoracion_puesto']].astype(str)

            df['% de jornada'] = df[self.mapeo_columnas['jornada_pct']]

            # Calcular Coeficiente Horas Trabajadas Efectivo
            # FÓRMULA CORREGIDA: coef_tp = % jornada en decimal
            log("Calculando coeficiente de tiempo parcial...", 'PROCESO')
            df['Coeficiente Horas Trabajadas Efectivo'] = df.apply(
                lambda row: self.calcular_coef_tp(row[self.mapeo_columnas['jornada_pct']]),
                axis=1
            )

            # Asignar salario base
            df['Salario base anual efectivo'] = df[self.mapeo_columnas['salario_base_efectivo']]

            # Procesar equiparación
            df_equiparado = self.procesar_equiparacion(df)

            # IMPORTANTE: Calcular totales acumulativos por empleado (como en formato maestro)
            df_equiparado = self.calcular_totales_acumulativos(df_equiparado)

            # Asignar columna Reg. DESPUÉS de todos los cálculos para evitar pérdidas
            self.asignar_reg_por_empleado(df_equiparado)

            log(f"Procesamiento completado: {df_equiparado.shape}", 'OK')
            return df_equiparado

        except Exception as e:
            log(f"Error procesando {ruta_archivo.name}: {str(e)}", 'ERROR')
            raise

    def calcular_totales_acumulativos(self, df):
        """
        Calcula totales acumulativos por empleado (formato maestro).
        Para empleados con múltiples situaciones contractuales:
        - Cada fila tiene su 'Salario base anual efectivo' individual
        - 'Salario base efectivo Total' es la suma acumulativa hasta esa fila
        """
        log("Calculando totales acumulativos por empleado...", 'PROCESO')

        # Ordenar por empleado y fecha de inicio
        df_sorted = df.sort_values(['Orden', 'Inicio de Sit. Contractual']).copy()

        # Columnas para acumular (todas las que terminen en 'efectivo' o contengan complementos)
        columnas_base_efectivo = ['Salario base anual efectivo']
        columnas_complementos_efectivos = [col for col in df.columns if 'Compltos' in col and 'efectivo' in col]

        # Calcular acumulativos para Salario base efectivo Total
        df_sorted['Salario base efectivo Total'] = df_sorted.groupby('Orden')['Salario base anual efectivo'].cumsum()

        # Calcular acumulativos para complementos
        # IMPORTANTE: Los nombres deben coincidir EXACTAMENTE con el formato maestro (con espacio al final)
        for col in columnas_complementos_efectivos:
            if 'Salariales' in col:
                # Para complementos salariales - con espacio al final como en formato maestro
                df_sorted['Compltos Salariales efectivo Total '] = df_sorted.groupby('Orden')[col].cumsum()
            elif 'Extrasalariales' in col:
                # Para complementos extrasalariales - con espacio al final como en formato maestro
                df_sorted['Compltos Extrasalariales efectivo Total '] = df_sorted.groupby('Orden')[col].cumsum()

        # Calcular totales combinados acumulativos
        if 'Salario base anual + complementos' in df_sorted.columns:
            df_sorted['Salario base anual + complementos Total'] = df_sorted.groupby('Orden')['Salario base anual + complementos'].cumsum()

        if 'Salario base anual + complementos + Extrasalariales' in df_sorted.columns:
            df_sorted['Salario base anual + complementos + Extrasalariales Total'] = df_sorted.groupby('Orden')['Salario base anual + complementos + Extrasalariales'].cumsum()

        # Restaurar orden original
        df_sorted = df_sorted.sort_index()

        log("Totales acumulativos calculados correctamente", 'OK')
        return df_sorted

    def _obtener_columnas_complementos_triodos(self, df):
        """Obtiene columnas de complementos en formato Triodos (A###, PA###, PC###)"""
        if self._columnas_complementos_cache is not None:
            return self._columnas_complementos_cache

        # Buscar TODAS las columnas que empiezan con A, PA o PC
        # IMPORTANTE: No filtrar por posición, porque los complementos están dispersos
        columnas_comp = []
        for col in df.columns:
            # Columnas A### (ej: A001-Trienios, A052-Kilom. Exento)
            # EXCEPTO A154 que es el salario base
            if col.startswith('A') and col != 'A154-Salario base de nivel*CT':
                columnas_comp.append(col)
            # Columnas PA### (ej: PA10-Prestaciones IT, PA40-Prestac.oblig.Empresa-E)
            elif col.startswith('PA'):
                columnas_comp.append(col)
            # Columnas PC### (ej: PC10-Complementos IT, PC20-Compl.AT y EP)
            elif col.startswith('PC'):
                columnas_comp.append(col)

        log(f"Total complementos encontrados: {len(columnas_comp)}", 'INFO')

        # Clasificar en salariales y extrasalariales según configuración
        columnas_por_tipo = {'PS': [], 'PE': []}

        for col in columnas_comp:
            # Extraer el código (ej: "A001" de "A001-Trienios", "PA10" de "PA10-Prestaciones IT")
            if '-' in col:
                codigo = col.split('-')[0].strip()
            else:
                # Para columnas sin guión, tomar los primeros caracteres
                codigo = col[:4].strip() if col.startswith('A') else col.strip()

            if codigo in self.configuracion_complementos:
                tipo = self.configuracion_complementos[codigo]['tipo']
                if tipo == 'salarial':
                    columnas_por_tipo['PS'].append(col)
                elif tipo == 'extrasalarial':
                    columnas_por_tipo['PE'].append(col)
            else:
                # Para PA, PC y complementos A sin configuración:
                # - Por defecto salariales EXCEPTO si tienen marcadores de extrasalarial
                if 'CE' in col or 'Exento' in col or col.startswith('PA') or col.startswith('PC'):
                    # Probablemente extrasalarial
                    log(f"Complemento {col} sin configuración, asumiendo como EXTRASALARIAL", 'WARN')
                    columnas_por_tipo['PE'].append(col)
                else:
                    log(f"Complemento {col} sin configuración, asumiendo como SALARIAL", 'WARN')
                    columnas_por_tipo['PS'].append(col)

        self._columnas_complementos_cache = columnas_por_tipo
        log(f"Complementos salariales: {len(columnas_por_tipo['PS'])}, Extrasalariales: {len(columnas_por_tipo['PE'])}", 'INFO')
        return columnas_por_tipo

    def procesar_equiparacion(self, df):
        """Procesa la equiparación de todos los valores"""
        log("Iniciando cálculos de equiparación...", 'PROCESO')

        df_equiparado = df.copy()

        col_meses = '¿Cuántos meses ha trabajado?'
        col_sb_efectivo = 'Salario base anual efectivo'

        # Equiparar salario base (VECTORIZADO)
        sb_efectivo = df_equiparado[col_sb_efectivo].fillna(0)
        coef_tp_norm = df_equiparado['Coeficiente Horas Trabajadas Efectivo'].replace(0, 1.0).fillna(1.0)
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

        # Procesar complementos individuales
        complementos_procesados = self.procesar_complementos_triodos(df_equiparado, col_meses)

        # Calcular totales de complementos equiparados
        self.calcular_totales_complementos(df_equiparado)

        log(f"Complementos procesados: {complementos_procesados}", 'OK')
        return df_equiparado

    def procesar_complementos_triodos(self, df_equiparado, col_meses):
        """Procesa todos los complementos de Triodos"""
        if not self.configuracion_complementos:
            log("No hay configuración de complementos disponible", 'WARN')
            return 0

        log("Procesando complementos de Triodos...", 'PROCESO')
        complementos_procesados = 0

        # Obtener columnas de complementos
        columnas_por_tipo = self._obtener_columnas_complementos_triodos(df_equiparado)

        for tipo, columnas in columnas_por_tipo.items():
            log(f"Columnas {tipo} encontradas: {len(columnas)}")

            for col_comp in columnas:
                # Extraer el código
                if '-' in col_comp:
                    codigo = col_comp.split('-')[0].strip()
                else:
                    codigo = col_comp[:4].strip() if col_comp.startswith('A') else col_comp.strip()

                es_normalizable, es_anualizable, _, nombre_comp = self.obtener_config_complemento(codigo)

                # PROCESAR TODOS LOS COMPLEMENTOS (aunque no se equiparen)
                datos_no_nulos = df_equiparado[col_comp].dropna()
                if len(datos_no_nulos) > 0:
                    col_equiparado = f"{col_comp}_equiparado"

                    # Si es normalizable O anualizable, equiparar
                    if es_normalizable or es_anualizable:
                        log(f"  {col_comp}: {len(datos_no_nulos)} registros (N:{es_normalizable}, A:{es_anualizable})")

                        # VECTORIZACIÓN
                        comp_efectivo = df_equiparado[col_comp].fillna(0)
                        resultado = comp_efectivo.copy()

                        # Aplicar normalización si corresponde
                        if es_normalizable:
                            coef_tp_norm = df_equiparado['Coeficiente Horas Trabajadas Efectivo'].replace(0, 1.0).fillna(1.0)
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

                        prom_efectivo = df_equiparado[col_comp].mean()
                        prom_equiparado = df_equiparado[col_equiparado].mean()
                        log(f"    Efectivo: {prom_efectivo:.2f} € | Equiparado: {prom_equiparado:.2f} €")
                    else:
                        # Si NO es normalizable ni anualizable, copiar el valor tal cual
                        log(f"  {col_comp}: {len(datos_no_nulos)} registros (SIN equiparación)")
                        df_equiparado[col_equiparado] = df_equiparado[col_comp].copy()

                    complementos_procesados += 1

        return complementos_procesados

    def _calcular_total_correcto(self, row, columnas_base, df_equiparado):
        """Calcula total: equiparado si es procesable, efectivo si no"""
        total = 0
        for col_base in columnas_base:
            if col_base not in row.index:
                continue

            valor_original = row[col_base]
            if pd.notna(valor_original) and valor_original != 0:
                col_equiparada = f"{col_base}_equiparado"
                if col_equiparada in df_equiparado.columns:
                    valor_equiparado = row[col_equiparada]
                    total += valor_equiparado if pd.notna(valor_equiparado) else valor_original
                else:
                    total += valor_original
        return total

    def calcular_totales_complementos(self, df_equiparado):
        """Calcula los totales correctos de complementos equiparados"""
        log("Calculando totales de complementos...", 'PROCESO')

        columnas_por_tipo = self._obtener_columnas_complementos_triodos(df_equiparado)

        # Calcular totales
        df_equiparado['Compltos Salariales efectivo'] = df_equiparado.apply(
            lambda row: sum([row[col] for col in columnas_por_tipo['PS'] if col in row.index and pd.notna(row[col])]),
            axis=1
        )

        df_equiparado['Compltos Extrasalariales efectivo'] = df_equiparado.apply(
            lambda row: sum([row[col] for col in columnas_por_tipo['PE'] if col in row.index and pd.notna(row[col])]),
            axis=1
        )

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
        """Calcula las columnas combinadas necesarias"""
        log("Calculando columnas combinadas...", 'PROCESO')

        if 'salario_base_equiparado' in df_equiparado.columns and 'complementos_salariales_equiparados' in df_equiparado.columns:
            df_equiparado['sb_mas_comp_salariales_equiparado'] = (
                df_equiparado['salario_base_equiparado'].fillna(0) +
                df_equiparado['complementos_salariales_equiparados'].fillna(0)
            )
            promedio_sb_comp = df_equiparado['sb_mas_comp_salariales_equiparado'].mean()
            log(f"SB + Comp. Salariales promedio: {promedio_sb_comp:.2f} €")

        if ('sb_mas_comp_salariales_equiparado' in df_equiparado.columns and
            'complementos_extrasalariales_equiparados' in df_equiparado.columns):
            df_equiparado['sb_mas_comp_total_equiparado'] = (
                df_equiparado['sb_mas_comp_salariales_equiparado'].fillna(0) +
                df_equiparado['complementos_extrasalariales_equiparados'].fillna(0)
            )
            promedio_total = df_equiparado['sb_mas_comp_total_equiparado'].mean()
            log(f"SB + Comp. Total promedio: {promedio_total:.2f} €")

        # Calcular totales adicionales
        if 'Salario base anual efectivo' in df_equiparado.columns and 'Compltos Salariales efectivo' in df_equiparado.columns:
            df_equiparado['Salario base anual + complementos'] = (
                df_equiparado['Salario base anual efectivo'].fillna(0) +
                df_equiparado['Compltos Salariales efectivo'].fillna(0)
            )

        if 'Salario base anual + complementos' in df_equiparado.columns and 'Compltos Extrasalariales efectivo' in df_equiparado.columns:
            df_equiparado['Salario base anual + complementos + Extrasalariales'] = (
                df_equiparado['Salario base anual + complementos'].fillna(0) +
                df_equiparado['Compltos Extrasalariales efectivo'].fillna(0)
            )

    def crear_reporte_excel(self, archivo_original, df_procesado):
        """Crea el archivo Excel de resultados"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_resultado = f"REPORTE_TRIODOS_{timestamp}.xlsx"
        ruta_resultado = self.carpeta_resultados / nombre_resultado

        log(f"Creando reporte: {nombre_resultado}", 'PROCESO')

        try:
            # Asegurar que las columnas estén limpias antes de procesar
            df_procesado.columns = df_procesado.columns.str.strip()
            
            # Eliminar columnas originales de Triodos que no deben estar en el reporte final
            columnas_eliminar = [
                'Nº personal', 'Fecha inicio contr.', 'Fecha de salida', 'Fecha inicio sit.',
                'Fecha fin sit.', 'Motivo cambio', 'Fecha de nacimiento', 'Tipo contrato',
                '% Jornada', '% Reducción', 'Motivo reducción', 'Puesto', 'Unidad Organizativa',
                'División', 'Nivel 4', 'Nivel 5', 'Nivel 6', 'Subdivisión Personal',
                'División Personal', 'Valoración puesto', 'Clasif. interna', 'Convenio colectivo',
                'Área convenio', 'Grupo prof.', 'Nivel prof.', 'Categ. prof.',
                'Dias baja enf.', 'Dias baja mat.', 'Dias baja pat.', 'Dias baja otro',
                'Bruto pagado', 'Salario pactado', 'A154-Salario base de nivel*CT',
                'Comprobación', 'Dif'
            ]

            # Mantener columnas relevantes:
            # - Columnas maestras (Orden, Sexo, etc.)
            # - TODOS los complementos efectivos (A###, PA###, PC###)
            # - Complementos equiparados (A###_equiparado, PA###_equiparado, PC###_equiparado)
            # - Columnas calculadas
            # ELIMINAR: solo columnas administrativas de Triodos
            columnas_a_mantener = []
            for col in df_procesado.columns:
                # Eliminar si está en la lista de columnas administrativas a eliminar
                if col in columnas_eliminar:
                    continue
                # MANTENER TODO LO DEMÁS (incluidos complementos efectivos y equiparados)
                columnas_a_mantener.append(col)

            df_final = df_procesado[columnas_a_mantener].copy()

            # Ordenar columnas según el formato maestro CORREGIDO
            orden_columnas_maestro = [
                'Reg.', 'Orden', 'Sexo',
                'Inicio de Sit. Contractual', 'Final de Sit. Contractual',
                'Grupo profesional', 'Nivel Convenio Colectivo', 'Categoría profesional',
                'Puesto de trabajo', 'Departamento', 'Nivel SVPT', '% de jornada',
                '¿Cuántos meses ha trabajado?', 'Coeficiente Horas Trabajadas Efectivo',
                'Salario base anual efectivo', 'Salario base efectivo Total',
                'Salario base anual + complementos', 'Salario base anual + complementos Total',
                'Salario base anual + complementos + Extrasalariales', 'Salario base anual + complementos + Extrasalariales Total',
                'Compltos Salariales efectivo', 'Compltos Salariales efectivo Total ',
                'Compltos Extrasalariales efectivo', 'Compltos Extrasalariales efectivo Total ',
            ]

            # Extraer complementos efectivos (A###, PA###, PC### sin _equiparado)
            complementos_efectivos = [
                col for col in df_final.columns
                if '_equiparado' not in col and (
                    (col.startswith('A') and col != 'A154-Salario base de nivel*CT') or
                    col.startswith('PA') or
                    col.startswith('PC')
                )
            ]
            # Ordenar alfabéticamente
            complementos_efectivos_ordenados = sorted(complementos_efectivos)
            orden_columnas_maestro.extend(complementos_efectivos_ordenados)

            # DESPUÉS de los efectivos, añadir los complementos equiparados
            complementos_equiparados = [
                col for col in df_final.columns
                if '_equiparado' in col and (
                    col.startswith('A') or
                    col.startswith('PA') or
                    col.startswith('PC')
                )
            ]
            # Ordenar alfabéticamente
            complementos_equiparados_ordenados = sorted(complementos_equiparados)
            orden_columnas_maestro.extend(complementos_equiparados_ordenados)

            # Añadir columnas calculadas al final
            columnas_calculadas = [
                'salario_base_equiparado',
                'complementos_salariales_equiparados',
                'complementos_extrasalariales_equiparados',
                'sb_mas_comp_salariales_equiparado',
                'sb_mas_comp_total_equiparado'
            ]
            orden_columnas_maestro.extend(columnas_calculadas)

            # Añadir cualquier otra columna que no esté en el orden definido
            for col in df_final.columns:
                if col not in orden_columnas_maestro:
                    orden_columnas_maestro.append(col)

            # Reordenar solo las columnas que existan en el DataFrame
            columnas_finales = [col for col in orden_columnas_maestro if col in df_final.columns]
            df_final = df_final[columnas_finales]

            log(f"Columnas en reporte final: {len(df_final.columns)} (eliminadas {len(df_procesado.columns) - len(df_final.columns)} columnas de Triodos)", 'INFO')

            with pd.ExcelWriter(ruta_resultado, engine='openpyxl') as writer:
                df_final.to_excel(writer, sheet_name='DATOS_PROCESADOS', index=False)

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
        """Procesa el archivo Triodos"""
        inicio_tiempo = datetime.now()

        try:
            log(f"INICIANDO: {ruta_archivo.name}", 'PROCESO')

            # Procesar datos
            df_procesado = self.leer_y_procesar_triodos(ruta_archivo)

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

    def ejecutar_procesamiento(self, archivo_especifico=None, password=None):
        """Función principal que ejecuta todo el procesamiento

        Args:
            archivo_especifico: Ruta completa del archivo específico a procesar (opcional)
            password: Contraseña para desbloquear el archivo Excel (opcional)
        """
        inicio_total = datetime.now()

        try:
            # Si se proporciona un archivo específico, usarlo
            if archivo_especifico:
                archivo_path = Path(archivo_especifico)
                if not archivo_path.exists():
                    raise Exception(f"El archivo especificado no existe: {archivo_especifico}")
                log(f"Procesando archivo específico: {archivo_path.name}", 'PROCESO')
                archivo_triodos = archivo_path
            else:
                log("Buscando archivo Triodos.xlsx...", 'PROCESO')
                archivo_triodos = self.buscar_archivo_triodos()

            # Guardar la contraseña si se proporciona
            if password:
                self.password = password
                log("Usando contraseña proporcionada", 'INFO')

            # Procesar archivo
            resultado = self.procesar_archivo(archivo_triodos)

            tiempo_total_proceso = (datetime.now() - inicio_total).total_seconds()

            # Log resumen final
            log("="*60)
            log("RESUMEN FINAL DEL PROCESAMIENTO")
            log("="*60)

            if resultado['estado'] == 'ÉXITO':
                log(f"Estado: ÉXITO", 'OK')
                log(f"Archivo generado: {resultado['archivo_resultado']}", 'OK')
                log(f"Registros procesados: {resultado['registros_procesados']}", 'OK')
                log(f"Tiempo total: {tiempo_total_proceso:.1f}s", 'OK')

                mensaje_final = f"""¡PROCESAMIENTO COMPLETADO EXITOSAMENTE!

⏱️  Tiempo total: {tiempo_total_proceso:.1f} segundos
📊 Registros procesados: {resultado['registros_procesados']}

📁 El reporte está guardado en:
{self.carpeta_resultados / resultado['archivo_resultado']}

📋 Los logs están en:
{self.carpeta_logs}

✨ Puede revisar los resultados abriendo el archivo Excel generado."""

                self.mostrar_mensaje("¡Procesamiento Exitoso!", mensaje_final)
                return True
            else:
                log(f"Estado: ERROR", 'ERROR')
                log(f"Error: {resultado.get('error', 'Desconocido')}", 'ERROR')

                mensaje_final = f"""ERROR EN EL PROCESAMIENTO

❌ Error: {resultado.get('error', 'Desconocido')}

📋 Revisar logs para más detalles:
{self.carpeta_logs}"""

                self.mostrar_mensaje("Error en Procesamiento", mensaje_final, "error")
                return False

        except Exception as e:
            tiempo_total_proceso = (datetime.now() - inicio_total).total_seconds()
            error_msg = f"Error crítico en el procesamiento:\n\n{str(e)}\n\nTiempo transcurrido: {tiempo_total_proceso:.1f}s"

            log(f"ERROR CRÍTICO: {error_msg}", 'ERROR')
            logging.error(f"Detalles técnicos: {traceback.format_exc()}")

            self.mostrar_mensaje("Error Crítico", error_msg, "error")
            return False


def main():
    """Función principal del programa"""
    ejecutado_desde_workflow = "--workflow" in sys.argv

    # Detectar si se pasa un archivo específico y contraseña como argumentos
    archivo_especifico = None
    password = None

    if len(sys.argv) > 1 and not sys.argv[1].startswith('--'):
        archivo_especifico = sys.argv[1]

    if len(sys.argv) > 2 and not sys.argv[2].startswith('--'):
        password = sys.argv[2]

    try:
        # Crear instancia del procesador
        procesador = ProcesadorTriodos()

        # Ejecutar procesamiento
        exito = procesador.ejecutar_procesamiento(archivo_especifico, password)

        # Solo pausar si se ejecuta directamente
        if not ejecutado_desde_workflow:
            input("\nPresiona Enter para cerrar...")

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
