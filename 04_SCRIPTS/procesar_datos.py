
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PROCESADOR AUTOMÁTICO DE REGISTROS RETRIBUTIVOS - EQUALITY MOMENTUM
Versión optimizada para ejecutable empresarial
Autor: Equality Momentum
Fecha: 2025
"""

import os
import sys
import pandas as pd
import numpy as np
import warnings
from pathlib import Path
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import traceback
import logging

# Suprimir warnings innecesarios
warnings.filterwarnings('ignore')

class LogManager:
    """Maneja los logs del sistema"""
    def __init__(self, log_dir):
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(exist_ok=True)
        
        # Configurar logging
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = self.log_dir / f"procesamiento_{timestamp}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        self.logger = logging.getLogger(__name__)
        self.logger.info("="*60)
        self.logger.info("INICIANDO PROCESADOR DE REGISTROS RETRIBUTIVOS")
        self.logger.info("="*60)

class ProcesadorRegistroRetributivo:
    def __init__(self):
        """Inicializa el procesador con las rutas y configuración"""
        # Obtener ruta base del ejecutable
        if hasattr(sys, '_MEIPASS'):
            self.base_path = Path(sys.executable).parent
        else:
            self.base_path = Path(__file__).parent.parent
            
        # Definir rutas
        self.carpeta_entrada = self.base_path / "01_DATOS_SIN_PROCESAR"
        self.carpeta_resultados = self.base_path / "02_RESULTADOS" 
        self.carpeta_logs = self.base_path / "03_LOGS"
        
        # Crear carpetas si no existen
        self.carpeta_resultados.mkdir(exist_ok=True)
        self.carpeta_logs.mkdir(exist_ok=True)
        
        # Inicializar logging
        self.log_manager = LogManager(self.carpeta_logs)
        self.logger = self.log_manager.logger
        
        # Configuración de columnas (nombres exactos del Excel)
        self.mapeo_columnas = {
            'meses_trabajados': '¿Cuántos meses ha trabajado?',
            'coef_tp': '% de jornada',
            'salario_base_efectivo': 'Salario base anual efectivo',
            'complementos_salariales_efectivo': 'Compltos Salariales efectivo',
            'complementos_extrasalariales_efectivo': 'Compltos Extrasalariales efectivo'
        }
        
        # Configuración de complementos
        self.configuracion_complementos = {}
        
    def mostrar_mensaje(self, titulo, mensaje, tipo="info"):
        """Muestra mensajes al usuario"""
        self.logger.info(f"MENSAJE USUARIO ({tipo.upper()}): {titulo} - {mensaje}")
        
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
            
        if not archivos_excel:
            raise Exception(f"No se encontraron archivos Excel en: {self.carpeta_entrada}")
            
        self.logger.info(f"Archivos Excel encontrados: {len(archivos_excel)}")
        for archivo in archivos_excel:
            self.logger.info(f"  - {archivo.name}")
            
        return archivos_excel

    def is_positive_response(self, value):
        """Verifica si un valor representa una respuesta positiva (Sí/Si/YES)"""
        if pd.isna(value):
            return False
        normalized = str(value).strip().lower()
        return normalized in ['sí', 'si', 'yes', 'y', '1', 'true']

    def cargar_configuracion_complementos(self, excel_file, archivo_path):
        """Carga la configuración de complementos desde las hojas Excel"""
        self.logger.info("Cargando configuración de complementos...")
        
        nombres_columnas_config = {
            'codigo': 'Cod',
            'normalizable': '¿Es Normalizable?',
            'anualizable': '¿Es Anualizable?'
        }
        
        configuracion = {}
        
        # Cargar complementos salariales
        if 'COMPLEMENTOS SALARIALES' in excel_file.sheet_names:
            try:
                df_comp_sal = pd.read_excel(archivo_path, sheet_name='COMPLEMENTOS SALARIALES')
                self.logger.info("Procesando COMPLEMENTOS SALARIALES...")
                
                for _, row in df_comp_sal.iterrows():
                    codigo_val = row.get(nombres_columnas_config['codigo'])
                    if pd.notna(codigo_val):
                        codigo = str(codigo_val).strip()
                        
                        normalizable_val = row.get(nombres_columnas_config['normalizable'])
                        anualizable_val = row.get(nombres_columnas_config['anualizable'])
                        
                        configuracion[codigo] = {
                            'tipo': 'salarial',
                            'es_normalizable': self.is_positive_response(normalizable_val),
                            'es_anualizable': self.is_positive_response(anualizable_val)
                        }
                
                salariales = len([k for k, v in configuracion.items() if v['tipo'] == 'salarial'])
                self.logger.info(f"Complementos salariales configurados: {salariales}")
                
            except Exception as e:
                self.logger.warning(f"Error cargando complementos salariales: {e}")
        
        # Cargar complementos extrasalariales
        if 'COMPLEMENTOS EXTRASALARIALES' in excel_file.sheet_names:
            try:
                df_comp_extra = pd.read_excel(archivo_path, sheet_name='COMPLEMENTOS EXTRASALARIALES')
                self.logger.info("Procesando COMPLEMENTOS EXTRASALARIALES...")
                
                for _, row in df_comp_extra.iterrows():
                    codigo_val = row.get(nombres_columnas_config['codigo'])
                    if pd.notna(codigo_val):
                        codigo = str(codigo_val).strip()
                        
                        normalizable_val = row.get(nombres_columnas_config['normalizable'])
                        anualizable_val = row.get(nombres_columnas_config['anualizable'])
                        
                        configuracion[codigo] = {
                            'tipo': 'extrasalarial',
                            'es_normalizable': self.is_positive_response(normalizable_val),
                            'es_anualizable': self.is_positive_response(anualizable_val)
                        }
                
                extrasalariales = len([k for k, v in configuracion.items() if v['tipo'] == 'extrasalarial'])
                self.logger.info(f"Complementos extrasalariales configurados: {extrasalariales}")
                
            except Exception as e:
                self.logger.warning(f"Error cargando complementos extrasalariales: {e}")
        
        self.configuracion_complementos = configuracion
        self.logger.info(f"Total complementos configurados: {len(configuracion)}")
        
        return configuracion

    def calcular_coef_tp(self, valor_coef_tp):
        """Convierte el coeficiente de tiempo parcial a decimal"""
        if pd.isna(valor_coef_tp):
            return 1.0
        # Si el valor es mayor que 1, asumimos que está en porcentaje
        if valor_coef_tp > 1:
            return valor_coef_tp / 100
        return valor_coef_tp

    def equiparar_salario_base(self, salario_base_efectivo, coef_tp, meses_trabajados):
        """Equipara el salario base aplicando normalización y anualización"""
        if pd.isna(salario_base_efectivo) or salario_base_efectivo == 0:
            return 0
        
        # Normalización (jornada completa)
        if pd.isna(coef_tp) or coef_tp == 0:
            coef_tp = 1.0
        salario_normalizado = salario_base_efectivo * (1/coef_tp)
        
        # Anualización (12 meses)
        if pd.isna(meses_trabajados) or meses_trabajados == 0:
            meses_trabajados = 12
        salario_equiparado = salario_normalizado * (12/meses_trabajados)
        
        return salario_equiparado

    def equiparar_complemento(self, complemento_efectivo, coef_tp, meses_trabajados, es_normalizable, es_anualizable):
        """Equipara un complemento según su configuración"""
        if pd.isna(complemento_efectivo) or complemento_efectivo == 0:
            return 0
        
        # Si no es ni normalizable ni anualizable, retornar valor original
        if not es_normalizable and not es_anualizable:
            return complemento_efectivo
        
        resultado = complemento_efectivo
        
        # Aplicar normalización si corresponde
        if es_normalizable:
            if pd.isna(coef_tp) or coef_tp == 0:
                coef_tp = 1.0
            resultado = resultado * (1/coef_tp)
        
        # Aplicar anualización si corresponde
        if es_anualizable:
            if pd.isna(meses_trabajados) or meses_trabajados == 0:
                meses_trabajados = 12
            resultado = resultado * (12/meses_trabajados)
        
        return resultado

    def obtener_config_complemento(self, codigo_complemento):
        """Obtiene la configuración de un complemento específico"""
        # Búsqueda directa
        if codigo_complemento in self.configuracion_complementos:
            config = self.configuracion_complementos[codigo_complemento]
            return config['es_normalizable'], config['es_anualizable'], config['tipo']
        
        # Búsquedas alternativas para diferentes formatos de códigos
        if codigo_complemento.isdigit():
            codigo_ps = f"PS{codigo_complemento}"
            if codigo_ps in self.configuracion_complementos:
                config = self.configuracion_complementos[codigo_ps]
                return config['es_normalizable'], config['es_anualizable'], config['tipo']
        
        if codigo_complemento.startswith('PS') and codigo_complemento[2:].isdigit():
            codigo_num = codigo_complemento[2:]
            if codigo_num in self.configuracion_complementos:
                config = self.configuracion_complementos[codigo_num]
                return config['es_normalizable'], config['es_anualizable'], config['tipo']
        
        # Valores por defecto
        self.logger.warning(f"Configuración no encontrada para {codigo_complemento}")
        return False, False, 'desconocido'

    def equiparar_complemento_individual(self, valor_efectivo, coef_tp, meses_trabajados, codigo_ps):
        """Equipara un complemento individual usando su configuración específica"""
        if pd.isna(valor_efectivo) or valor_efectivo == 0:
            return 0
        
        es_normalizable, es_anualizable, tipo = self.obtener_config_complemento(codigo_ps)
        return self.equiparar_complemento(valor_efectivo, coef_tp, meses_trabajados, es_normalizable, es_anualizable)

    def leer_y_procesar_excel(self, ruta_archivo):
        """Lee y procesa un archivo Excel completo"""
        self.logger.info(f"Procesando archivo: {ruta_archivo.name}")
        
        try:
            # Cargar información de hojas disponibles
            excel_file = pd.ExcelFile(ruta_archivo)
            self.logger.info(f"Hojas disponibles: {excel_file.sheet_names}")
            
            # Cargar hoja principal (BASE GENERAL)
            if "BASE GENERAL" not in excel_file.sheet_names:
                raise Exception("No se encontró la hoja 'BASE GENERAL' requerida")
            
            df = pd.read_excel(ruta_archivo, sheet_name="BASE GENERAL")
            self.logger.info(f"Datos cargados: {df.shape[0]} filas x {df.shape[1]} columnas")
            
            # Buscar columna "Reg." y eliminar todas las columnas anteriores
            if 'Reg.' in df.columns:
                indice_reg = df.columns.get_loc('Reg.')
                if indice_reg > 0:
                    columnas_a_eliminar = df.columns[:indice_reg].tolist()
                    df = df.drop(columns=columnas_a_eliminar)
                    self.logger.info(f"Eliminadas {len(columnas_a_eliminar)} columnas anteriores a 'Reg.': {columnas_a_eliminar}")
                else:
                    self.logger.info("La columna 'Reg.' ya es la primera columna")
            else:
                self.logger.warning("No se encontró la columna 'Reg.', no se eliminaron columnas")
            
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
                    self.logger.info(f"Filtrado por último dato con 'Orden': {registros_originales} → {registros_filtrados} registros")
                else:
                    self.logger.warning("No se encontraron datos válidos en la columna 'Orden'")
            else:
                self.logger.warning("No se encontró la columna 'Orden', no se aplicó filtro de registros")
            
            # Verificar columnas críticas
            columnas_encontradas = {}
            for clave, nombre_col in self.mapeo_columnas.items():
                if nombre_col in df.columns:
                    columnas_encontradas[clave] = nombre_col
                    self.logger.info(f"✓ {clave}: {nombre_col}")
                else:
                    self.logger.warning(f"✗ {clave}: '{nombre_col}' NO ENCONTRADA")
            
            if len(columnas_encontradas) < 3:  # Mínimo necesario
                raise Exception(f"Faltan columnas críticas. Encontradas: {list(columnas_encontradas.keys())}")
            
            # Procesar datos de equiparación
            df_equiparado = self.procesar_equiparacion(df, columnas_encontradas)
            
            self.logger.info(f"Procesamiento completado: {df_equiparado.shape}")
            return df_equiparado
            
        except Exception as e:
            self.logger.error(f"Error procesando {ruta_archivo.name}: {str(e)}")
            raise

    def procesar_equiparacion(self, df, columnas_encontradas):
        """Procesa la equiparación de todos los valores"""
        self.logger.info("Iniciando cálculos de equiparación...")
        
        df_equiparado = df.copy()
        
        # Obtener columnas necesarias
        col_meses = columnas_encontradas.get('meses_trabajados')
        col_coef_tp = columnas_encontradas.get('coef_tp')
        col_sb_efectivo = columnas_encontradas.get('salario_base_efectivo')
        
        if not all([col_meses, col_coef_tp, col_sb_efectivo]):
            missing = [k for k, v in [('meses', col_meses), ('coef_tp', col_coef_tp), ('sb_efectivo', col_sb_efectivo)] if not v]
            raise Exception(f"Columnas críticas faltantes: {missing}")
        
        # Calcular coeficiente TP corregido
        df_equiparado['coef_tp_calculado'] = df_equiparado[col_coef_tp].apply(self.calcular_coef_tp)
        
        # Equiparar salario base
        df_equiparado['salario_base_equiparado'] = df_equiparado.apply(
            lambda row: self.equiparar_salario_base(
                row[col_sb_efectivo], 
                row['coef_tp_calculado'], 
                row[col_meses]
            ), axis=1
        )
        
        sb_efectivo_promedio = df_equiparado[col_sb_efectivo].mean()
        sb_equiparado_promedio = df_equiparado['salario_base_equiparado'].mean()
        self.logger.info(f"SB efectivo promedio: {sb_efectivo_promedio:.2f}")
        self.logger.info(f"SB equiparado promedio: {sb_equiparado_promedio:.2f}")
        
        # Procesar complementos individuales
        complementos_procesados = self.procesar_complementos_individuales(df_equiparado, col_meses)
        
        # Calcular totales de complementos equiparados
        self.calcular_totales_complementos(df_equiparado)
        
        self.logger.info(f"Complementos individuales procesados: {complementos_procesados}")
        return df_equiparado

    def procesar_complementos_individuales(self, df_equiparado, col_meses):
        """Procesa todos los complementos PS y PE individuales"""
        if not self.configuracion_complementos:
            self.logger.warning("No hay configuración de complementos disponible")
            return 0
        
        self.logger.info("Procesando complementos individuales...")
        complementos_procesados = 0
        
        # Buscar columnas PS y PE
        columnas_ps = [col for col in df_equiparado.columns 
                      if col.startswith('PS') and col[2:].isdigit() and not col.endswith('_equiparado')]
        columnas_pe = [col for col in df_equiparado.columns 
                      if col.startswith('PE') and col[2:].isdigit() and not col.endswith('_equiparado')]
        
        self.logger.info(f"Columnas PS encontradas: {len(columnas_ps)}")
        self.logger.info(f"Columnas PE encontradas: {len(columnas_pe)}")
        
        # Procesar complementos PS y PE
        for columnas, tipo in [(columnas_ps, 'PS'), (columnas_pe, 'PE')]:
            for col_comp in columnas:
                codigo_comp = col_comp
                
                # Obtener configuración
                es_normalizable, es_anualizable, tipo_config = self.obtener_config_complemento(codigo_comp)
                
                # Solo procesar si es normalizable O anualizable
                if es_normalizable or es_anualizable:
                    col_equiparado = f"{codigo_comp}_equiparado"
                    
                    # Verificar si hay datos
                    datos_no_nulos = df_equiparado[col_comp].dropna()
                    if len(datos_no_nulos) > 0:
                        self.logger.info(f"Procesando {codigo_comp}: {len(datos_no_nulos)} registros (N:{es_normalizable}, A:{es_anualizable})")
                        
                        # Equiparar
                        df_equiparado[col_equiparado] = df_equiparado.apply(
                            lambda row: self.equiparar_complemento_individual(
                                row[col_comp], 
                                row['coef_tp_calculado'], 
                                row[col_meses],
                                codigo_comp
                            ), axis=1
                        )
                        complementos_procesados += 1
                        
                        # Log promedios
                        prom_efectivo = df_equiparado[col_comp].mean()
                        prom_equiparado = df_equiparado[col_equiparado].mean()
                        self.logger.info(f"  {codigo_comp} - Efectivo: {prom_efectivo:.2f}, Equiparado: {prom_equiparado:.2f}")
                else:
                    self.logger.info(f"Saltando {codigo_comp}: No procesable")
        
        return complementos_procesados

    def calcular_totales_complementos(self, df_equiparado):
        """Calcula los totales correctos de complementos equiparados"""
        self.logger.info("Calculando totales de complementos...")
        
        def calcular_total_correcto(row, tipo_complemento='PS'):
            """Calcula total: equiparado si es procesable, efectivo si no"""
            total = 0
            
            if tipo_complemento == 'PS':
                columnas_base = [col for col in df_equiparado.columns 
                               if col.startswith('PS') and col[2:].isdigit() and not col.endswith('_equiparado')]
            else:  # PE
                columnas_base = [col for col in df_equiparado.columns 
                               if col.startswith('PE') and col[2:].isdigit() and not col.endswith('_equiparado')]
            
            for col_base in columnas_base:
                valor_original = row[col_base]
                
                if pd.notna(valor_original) and valor_original != 0:
                    col_equiparada = f"{col_base}_equiparado"
                    
                    if col_equiparada in df_equiparado.columns:
                        # Usar valor equiparado si existe
                        valor_equiparado = row[col_equiparada]
                        total += valor_equiparado if pd.notna(valor_equiparado) else valor_original
                    else:
                        # Usar valor original si no hay equiparación
                        total += valor_original
            
            return total
        
        # Calcular totales
        df_equiparado['complementos_salariales_equiparados'] = df_equiparado.apply(
            lambda row: calcular_total_correcto(row, 'PS'), axis=1
        )
        
        df_equiparado['complementos_extrasalariales_equiparados'] = df_equiparado.apply(
            lambda row: calcular_total_correcto(row, 'PE'), axis=1
        )
        
        cs_promedio = df_equiparado['complementos_salariales_equiparados'].mean()
        ce_promedio = df_equiparado['complementos_extrasalariales_equiparados'].mean()
        self.logger.info(f"CS equiparados promedio: {cs_promedio:.2f}")
        self.logger.info(f"CE equiparados promedio: {ce_promedio:.2f}")
        
        # Calcular columnas combinadas adicionales
        self.calcular_columnas_combinadas(df_equiparado)

    def calcular_columnas_combinadas(self, df_equiparado):
        """Calcula las columnas combinadas necesarias para las visualizaciones"""
        self.logger.info("Calculando columnas combinadas...")
        
        # sb_mas_comp_equiparado = salario_base_equiparado + complementos_salariales_equiparados
        if 'salario_base_equiparado' in df_equiparado.columns and 'complementos_salariales_equiparados' in df_equiparado.columns:
            df_equiparado['sb_mas_comp_salariales_equiparado'] = (
                df_equiparado['salario_base_equiparado'].fillna(0) + 
                df_equiparado['complementos_salariales_equiparados'].fillna(0)
            )
            promedio_sb_comp = df_equiparado['sb_mas_comp_salariales_equiparado'].mean()
            self.logger.info(f"SB + Complementos Salariales equiparado promedio: {promedio_sb_comp:.2f}")
        
        # sb_mas_comp_extrasal_equiparado = sb_mas_comp_equiparado + complementos_extrasalariales_equiparados
        if ('sb_mas_comp_salariales_equiparado' in df_equiparado.columns and 
            'complementos_extrasalariales_equiparados' in df_equiparado.columns):
            df_equiparado['sb_mas_comp_total_equiparado'] = (
                df_equiparado['sb_mas_comp_salariales_equiparado'].fillna(0) + 
                df_equiparado['complementos_extrasalariales_equiparados'].fillna(0)
            )
            promedio_total = df_equiparado['sb_mas_comp_total_equiparado'].mean()
            self.logger.info(f"SB + Complementos Total equiparado promedio: {promedio_total:.2f}")

    def crear_reporte_excel(self, archivo_original, df_procesado):
        """Crea el archivo Excel de resultados simplificado"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_base = archivo_original.stem
        nombre_resultado = f"REPORTE_{nombre_base}_{timestamp}.xlsx"
        ruta_resultado = self.carpeta_resultados / nombre_resultado
        
        self.logger.info(f"Creando reporte: {nombre_resultado}")
        
        try:
            # Escribir archivo Excel con solo los datos procesados
            with pd.ExcelWriter(ruta_resultado, engine='openpyxl') as writer:
                # Solo datos procesados - sin hojas adicionales
                df_procesado.to_excel(writer, sheet_name='DATOS_PROCESADOS', index=False)
            
            # Verificar archivo creado
            if ruta_resultado.exists():
                tamano_mb = ruta_resultado.stat().st_size / (1024 * 1024)
                self.logger.info(f"Reporte creado exitosamente: {tamano_mb:.2f} MB")
                return ruta_resultado
            else:
                raise Exception("El archivo no se pudo crear correctamente")
                
        except Exception as e:
            self.logger.error(f"Error creando reporte: {str(e)}")
            raise

    def procesar_archivo(self, ruta_archivo):
        """Procesa un archivo Excel individual"""
        inicio_tiempo = datetime.now()
        
        try:
            self.logger.info(f"INICIANDO: {ruta_archivo.name}")
            
            # Procesar datos
            df_procesado = self.leer_y_procesar_excel(ruta_archivo)
            
            # Crear reporte
            ruta_resultado = self.crear_reporte_excel(ruta_archivo, df_procesado)
            
            # Calcular tiempo transcurrido
            tiempo_total = (datetime.now() - inicio_tiempo).total_seconds()
            
            self.logger.info(f"COMPLETADO: {ruta_archivo.name} en {tiempo_total:.1f}s")
            
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
            
            self.logger.error(f"ERROR en {ruta_archivo.name}: {error_msg}")
            self.logger.error(f"Detalles técnicos: {traceback.format_exc()}")
            
            return {
                'archivo_original': ruta_archivo.name,
                'archivo_resultado': None,
                'registros_procesados': 0,
                'tiempo_procesamiento': tiempo_total,
                'estado': 'ERROR',
                'error': error_msg
            }

    def ejecutar_procesamiento(self):
        """Función principal que ejecuta todo el procesamiento"""
        inicio_total = datetime.now()
        
        try:
            self.logger.info("Buscando archivos Excel...")
            archivos_excel = self.buscar_archivos_excel()
            
            # Procesar cada archivo
            resultados = []
            for i, archivo in enumerate(archivos_excel, 1):
                self.logger.info(f"Procesando archivo {i}/{len(archivos_excel)}: {archivo.name}")
                resultado = self.procesar_archivo(archivo)
                resultados.append(resultado)
            
            # Crear resumen final
            exitosos = [r for r in resultados if r['estado'] == 'ÉXITO']
            errores = [r for r in resultados if r['estado'] == 'ERROR']
            
            tiempo_total_proceso = (datetime.now() - inicio_total).total_seconds()
            
            # Log resumen final
            self.logger.info("="*60)
            self.logger.info("RESUMEN FINAL DEL PROCESAMIENTO")
            self.logger.info("="*60)
            self.logger.info(f"Archivos procesados exitosamente: {len(exitosos)}")
            self.logger.info(f"Archivos con errores: {len(errores)}")
            self.logger.info(f"Tiempo total de procesamiento: {tiempo_total_proceso:.1f} segundos")
            
            if exitosos:
                self.logger.info("\nArchivos exitosos:")
                total_registros = 0
                for r in exitosos:
                    total_registros += r['registros_procesados']
                    self.logger.info(f"  ✓ {r['archivo_original']} → {r['archivo_resultado']} ({r['registros_procesados']} registros)")
                self.logger.info(f"Total de registros procesados: {total_registros}")
                    
            if errores:
                self.logger.info("\nArchivos con errores:")
                for r in errores:
                    self.logger.info(f"  ✗ {r['archivo_original']}: {r['error']}")
            
            # Mostrar mensaje final al usuario
            self.mostrar_mensaje_final(exitosos, errores, tiempo_total_proceso)
                
            return len(errores) == 0  # True si no hay errores
            
        except Exception as e:
            tiempo_total_proceso = (datetime.now() - inicio_total).total_seconds()
            error_msg = f"Error crítico en el procesamiento:\n\n{str(e)}\n\nTiempo transcurrido: {tiempo_total_proceso:.1f}s"
            
            self.logger.error(f"ERROR CRÍTICO: {error_msg}")
            self.logger.error(f"Detalles técnicos: {traceback.format_exc()}")
            
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
    try:
        # Crear instancia del procesador
        procesador = ProcesadorRegistroRetributivo()
        
        # Ejecutar procesamiento
        exito = procesador.ejecutar_procesamiento()
        
        # Pausa para que el usuario pueda leer el mensaje
        input("\nPresiona Enter para cerrar...")
        
        # Código de salida
        sys.exit(0 if exito else 1)
        
    except KeyboardInterrupt:
        print("\n\nProcesamiento interrumpido por el usuario.")
        input("Presiona Enter para cerrar...")
        sys.exit(1)
        
    except Exception as e:
        print(f"\nError crítico no manejado: {str(e)}")
        print("\nContacte con soporte técnico.")
        input("Presiona Enter para cerrar...")
        sys.exit(1)


if __name__ == "__main__":
    main()