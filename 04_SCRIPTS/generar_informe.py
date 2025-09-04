# generar_informe.py
"""
Sistema completo para automatizar la creación de informes con visualizaciones
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from docx import Document
from docx.shared import Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import yaml
import json
from datetime import datetime
import os
from pathlib import Path

class AutomatedReportSystem:
    def __init__(self, config_file="report_config.yaml"):
        """
        Sistema automatizado de generación de reportes
        """
        self.config = self.load_config(config_file)
        self.data = None
        self.charts_created = {}
        
        # Configuración de visualización
        plt.rcParams['figure.figsize'] = (12, 8)
        plt.rcParams['font.size'] = 10
        sns.set_style("whitegrid")
        
        # Paleta de colores para género
        self.colores_genero = {
            'H': '#DC2626',  # Rojo para hombres
            'M': '#2E86AB',  # Azul para mujeres
        }
    
    def load_config(self, config_file):
        """Carga la configuración desde un archivo YAML"""
        default_config = {
            'excel_file': '',  # Se determina automáticamente
            'template_word': 'plantilla_informe.docx',  # Opcional
            'output_file': f'informe_registro_retributivo_{datetime.now().strftime("%Y%m%d_%H%M")}.docx',
            'charts': {
                'salario_base_efectivo': {
                    'type': 'donut',
                    'data_columns': ['Salario base efectivo Total'],
                    'metodo': 'efectivos_sb',
                    'title': 'Comparación Salario Base Efectivo Total por Género',
                    'subtitulo': 'Análisis de igualdad retributiva - Salario base efectivamente percibido (solo SB > 0)',
                    'marker': '{grafico_sb_efectivo}'
                },
                'sb_complementos_efectivo': {
                    'type': 'donut',
                    'data_columns': ['Salario base anual + complementos Total'],
                    'metodo': 'efectivos_sb_complementos',
                    'title': 'Salario Base + Complementos Salariales Efectivos por Género',
                    'subtitulo': 'Incluye salario base y complementos salariales efectivamente percibidos (todas las personas)',
                    'marker': '{grafico_sb_comp_efectivo}'
                },
                'sb_total_efectivo': {
                    'type': 'donut',
                    'data_columns': ['Salario base anual + complementos + Extrasalariales Total'],
                    'metodo': 'efectivos_sb_complementos',
                    'title': 'SB + Complementos + Extrasalariales Efectivos por Género',
                    'subtitulo': 'Retribución total efectiva incluyendo todos los conceptos (todas las personas)',
                    'marker': '{grafico_sb_total_efectivo}'
                },
                'salario_base_equiparado': {
                    'type': 'donut',
                    'data_columns': ['salario_base_equiparado'],
                    'metodo': 'equiparados_sb',
                    'title': 'Comparación Salario Base Equiparado por Género',
                    'subtitulo': 'Salario base normalizado a jornada completa y 12 meses (solo SB > 0)',
                    'marker': '{grafico_sb_equiparado}'
                },
                'sb_complementos_equiparado': {
                    'type': 'donut',
                    'data_columns': ['sb_mas_comp_salariales_equiparado'],
                    'metodo': 'equiparados_sb_complementos',
                    'title': 'Salario Base + Complementos Salariales Equiparados por Género',
                    'subtitulo': 'SB + complementos salariales normalizados a jornada completa y 12 meses (todas las personas)',
                    'marker': '{grafico_sb_comp_equiparado}'
                },
                'sb_total_equiparado': {
                    'type': 'donut',
                    'data_columns': ['sb_mas_comp_total_equiparado'],
                    'metodo': 'equiparados_sb_complementos',
                    'title': 'SB + Complementos + Extrasalariales Equiparados por Género',
                    'subtitulo': 'Retribución total equiparada: SB + complementos salariales y extrasalariales (todas las personas)',
                    'marker': '{grafico_sb_total_equiparado}'
                }
            }
        }
        
        try:
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    user_config = yaml.safe_load(f)
                # Combinar configuraciones
                default_config.update(user_config)
        except Exception as e:
            print(f"Error cargando configuración: {e}")
            print("Usando configuración por defecto")
        
        return default_config
    
    def calcular_brecha_salarial(self):
        """Calcula la brecha salarial por grupo profesional"""
        if 'Sexo' not in self.data.columns or 'Salario base anual efectivo' not in self.data.columns:
            print("No se pueden calcular brechas: faltan columnas necesarias")
            return None
        
        # Calcular brecha por grupo profesional si existe
        if 'Grupo profesional' in self.data.columns:
            grupos = self.data['Grupo profesional'].unique()
            brechas = []
            
            for grupo in grupos:
                data_grupo = self.data[self.data['Grupo profesional'] == grupo]
                
                # Calcular salarios promedio por género
                salarios_genero = data_grupo.groupby('Sexo')['Salario base anual efectivo'].mean()
                
                if 'H' in salarios_genero.index and 'M' in salarios_genero.index:
                    salario_h = salarios_genero['H']
                    salario_m = salarios_genero['M']
                    
                    # Calcular brecha porcentual (diferencia respecto al salario mayor)
                    brecha = ((salario_h - salario_m) / max(salario_h, salario_m)) * 100
                    brechas.append({'Grupo profesional': grupo, 'brecha_porcentual': brecha})
                    print(f"Grupo {grupo}: Brecha salarial = {brecha:.2f}%")
            
            if brechas:
                df_brechas = pd.DataFrame(brechas)
                # Añadir las brechas al dataframe principal
                self.data = self.data.merge(df_brechas, on='Grupo profesional', how='left')
                print("Brechas salariales calculadas y añadidas al dataset")
                
                # Añadir gráfico de brecha si no existe en la configuración
                if 'brecha_por_grupo' not in self.config['charts']:
                    self.config['charts']['brecha_por_grupo'] = {
                        'type': 'bar',
                        'data_columns': ['Grupo profesional', 'brecha_porcentual'],
                        'title': 'Brecha Salarial por Grupo Profesional (%)',
                        'marker': '{grafico_brecha_grupo}'
                    }
                    print("Gráfico de brecha salarial añadido a la configuración")
                
                return df_brechas
        
        return None
    
    def formato_numero_es(self, numero, decimales=2):
        """Formatea un número con estilo español (punto como separador de miles)"""
        if pd.isna(numero):
            return "0,00"
        
        # Formatear el número con decimales
        numero_formateado = f"{numero:,.{decimales}f}"
        
        # Cambiar punto por coma para decimales y coma por punto para miles (estilo español)
        numero_formateado = numero_formateado.replace(',', 'X').replace('.', ',').replace('X', '.')
        
        return numero_formateado
    
    def calcular_promedios_efectivos_sb(self, df, columna_salario):
        """
        Calcula promedios para SALARIO BASE EFECTIVOS siguiendo las reglas:
        - Solo registros actuales (excluir 'Ex')
        - Solo donde SB > 0
        - Usar columnas 'Total efectivo'
        """
        # Filtrar registros actuales (sin "Ex" en primera columna)
        df_actual = df[df.iloc[:, 0] != 'Ex'].copy()
        
        # Filtrar datos válidos (SB > 0) y calcular promedio solo de estos registros
        df_sb_mayor_0 = df_actual[(df_actual[columna_salario].notna()) & (df_actual[columna_salario] > 0)]
        
        # Calcular promedios por género solo de personas con SB > 0
        promedios = df_sb_mayor_0.groupby('SEXO')[columna_salario].mean()
        
        return {
            'H': promedios.get('H', 0),
            'M': promedios.get('M', 0)
        }
    
    def calcular_promedios_efectivos_sb_complementos(self, df, columna_salario):
        """
        Calcula promedios para SB + COMPLEMENTOS EFECTIVOS siguiendo las reglas:
        - Solo registros actuales (excluir 'Ex')
        - TODAS las personas (incluir SB = 0)
        - Usar columnas 'Total efectivo'
        """
        # Filtrar registros actuales (sin "Ex" en primera columna)
        df_actual = df[df.iloc[:, 0] != 'Ex'].copy()
        
        # Filtrar datos válidos (solo quitar nulos, incluir SB = 0)
        df_valido = df_actual[df_actual[columna_salario].notna()]
        
        # Calcular promedios por género (incluir todas las personas)
        promedios = df_valido.groupby('SEXO')[columna_salario].mean()
        
        return {
            'H': promedios.get('H', 0),
            'M': promedios.get('M', 0)
        }
    
    def calcular_promedios_equiparados_sb(self, df, columna_salario):
        """
        Calcula promedios para SALARIO BASE EQUIPARADOS siguiendo las reglas:
        - Solo registros actuales (excluir 'Ex')
        - Solo donde SB base > 0
        - Usar columnas equiparadas basadas en situación contractual actual
        """
        # Filtrar registros actuales (sin "Ex" en primera columna)
        df_actual = df[df.iloc[:, 0] != 'Ex'].copy()
        
        # Filtrar datos válidos (SB base > 0) - usar la columna de salario base efectivo como referencia
        col_sb_efectivo = "Salario base anual efectivo"
        if col_sb_efectivo in df.columns:
            df_sb_mayor_0 = df_actual[(df_actual[col_sb_efectivo].notna()) & (df_actual[col_sb_efectivo] > 0)]
        else:
            df_sb_mayor_0 = df_actual[df_actual[columna_salario].notna() & (df_actual[columna_salario] > 0)]
        
        # Calcular promedios por género
        promedios = df_sb_mayor_0.groupby('SEXO')[columna_salario].mean()
        
        return {
            'H': promedios.get('H', 0),
            'M': promedios.get('M', 0)
        }
    
    def calcular_promedios_equiparados_sb_complementos(self, df, columna_salario):
        """
        Calcula promedios para SB + COMPLEMENTOS EQUIPARADOS siguiendo las reglas:
        - Solo registros actuales (excluir 'Ex')
        - TODAS las personas (incluir SB = 0)
        - Usar columnas equiparadas basadas en situación contractual actual
        """
        # Filtrar registros actuales (sin "Ex" en primera columna)
        df_actual = df[df.iloc[:, 0] != 'Ex'].copy()
        
        # Filtrar datos válidos (solo quitar nulos, incluir SB = 0)
        df_valido = df_actual[df_actual[columna_salario].notna()]
        
        # Calcular promedios por género
        promedios = df_valido.groupby('SEXO')[columna_salario].mean()
        
        return {
            'H': promedios.get('H', 0),
            'M': promedios.get('M', 0)
        }
    
    def crear_grafico_donut(self, datos_genero, titulo, subtitulo="", formato_moneda=True):
        """
        Crea un gráfico de donut con la brecha salarial en el centro
        """
        # Preparar datos
        valores = [datos_genero['H'], datos_genero['M']]
        etiquetas = ['Hombres', 'Mujeres']
        colores = [self.colores_genero['H'], self.colores_genero['M']]
        
        # Calcular la brecha salarial
        if datos_genero['M'] > 0:
            brecha = ((datos_genero['H'] - datos_genero['M']) / datos_genero['M']) * 100
        else:
            brecha = 0
            
        # Configurar la figura
        fig, ax = plt.subplots(figsize=(10, 8))
        
        # Crear el gráfico de donut
        wedges, texts, autotexts = ax.pie(valores, labels=etiquetas, autopct='%1.1f%%',
                                          colors=colores, startangle=90, 
                                          wedgeprops=dict(width=0.4, edgecolor='white', linewidth=2))
        
        # Mejorar la apariencia de los porcentajes
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
            autotext.set_fontsize(14)
        
        # Añadir texto de brecha en el centro
        color_brecha = '#e74c3c' if brecha > 0 else '#27ae60' if brecha < 0 else '#95a5a6'
        signo = '+' if brecha > 0 else ''
        
        ax.text(0, 0.05, 'Brecha Salarial',
                horizontalalignment='center', verticalalignment='center',
                fontsize=14, fontweight='bold', color='#2c3e50')
        
        porcentaje_texto = f'{signo}{self.formato_numero_es(brecha, 2)}%'
        ax.text(0, -0.15, porcentaje_texto, 
                horizontalalignment='center', verticalalignment='center',
                fontsize=16, fontweight='bold', color=color_brecha)
        
        # Título y subtítulo
        ax.set_title(titulo, fontsize=16, fontweight='bold', pad=20, color='#2c3e50')
        if subtitulo:
            ax.text(0, -1.3, subtitulo, horizontalalignment='center', 
                    fontsize=12, style='italic', color='#7f8c8d')
        
        # Leyenda personalizada con formato español
        if formato_moneda:
            leyenda_labels = [f'{etiqueta}: {self.formato_numero_es(valor, 0)}€' for etiqueta, valor in zip(etiquetas, valores)]
        else:
            leyenda_labels = [f'{etiqueta}: {self.formato_numero_es(valor, 2)}' for etiqueta, valor in zip(etiquetas, valores)]
        
        ax.legend(wedges, leyenda_labels, 
                  loc="center left", bbox_to_anchor=(1, 0, 0.5, 1),
                  fontsize=11)
        
        # Ajustar el aspecto
        ax.axis('equal')
        plt.tight_layout()
        
        return fig
    
    def load_data(self):
        """Carga los datos desde Excel usando la lógica de registro retributivo"""
        try:
            # Obtener la ruta base del proyecto
            ruta_base = Path.cwd().parent if Path.cwd().name == '04_SCRIPTS' else Path.cwd()
            
            # Buscar el archivo más reciente en la carpeta de resultados
            carpeta_resultados = ruta_base / '02_RESULTADOS'
            print(f"Buscando archivos en: {carpeta_resultados}")
            
            # Buscar archivos que empiecen con REPORTE_
            archivos_reporte = list(carpeta_resultados.glob('REPORTE_*.xlsx'))
            if archivos_reporte:
                # Tomar el archivo más reciente
                archivo_mas_reciente = max(archivos_reporte, key=os.path.getctime)
                ruta_datos = archivo_mas_reciente
                print(f"Archivo más reciente encontrado: {ruta_datos.name}")
            else:
                # Fallback al archivo específico
                ruta_datos = carpeta_resultados / 'REPORTE_DATOS registro retributivo_20250902_235323.xlsx'
                print(f"Usando archivo específico: {ruta_datos}")

            print(f"Cargando datos desde: {ruta_datos}")

            # Leer todas las hojas del archivo procesado
            datos_procesados = {}
            archivo_excel = pd.ExcelFile(ruta_datos)
            for hoja in archivo_excel.sheet_names:
                datos_procesados[hoja] = pd.read_excel(archivo_excel, sheet_name=hoja)
                print(f"Cargada hoja '{hoja}': {len(datos_procesados[hoja])} registros")

            print("\nHojas disponibles en el archivo procesado:")
            for hoja in datos_procesados.keys():
                print(f"  - {hoja}")

            # Verificar si existe la hoja con datos procesados (actualizada)
            if 'DATOS_PROCESADOS' in datos_procesados:
                self.data = datos_procesados['DATOS_PROCESADOS']
                print(f"\nDataset principal cargado: {len(self.data)} registros")
                print(f"Columnas principales: {list(self.data.columns[:5])}...")
                
                # Verificar columnas clave
                columnas_clave = ['Sexo', 'Salario base anual efectivo', 'salario_base_equiparado']
                for col in columnas_clave:
                    if col in self.data.columns:
                        print(f"Columna '{col}' encontrada")
                    else:
                        print(f"Columna '{col}' NO encontrada")
                
                # Mostrar todas las columnas disponibles para verificar
                print(f"\nTodas las columnas disponibles ({len(self.data.columns)}):")
                for i, col in enumerate(self.data.columns, 1):
                    print(f"   {i:2d}. {col}")
                
                # Almacenar todas las hojas para uso posterior
                self.datos_procesados = datos_procesados
                
                # Configurar columna SEXO para compatibilidad
                if 'Sexo' in self.data.columns:
                    self.data['SEXO'] = self.data['Sexo'].map({'Hombres': 'H', 'Mujeres': 'M'})
                    print("✅ Columna SEXO configurada correctamente")
                    distribucion = self.data['SEXO'].value_counts()
                    print(f"   Distribución: {distribucion.to_dict()}")
                
                return True
            else:
                print("No se encontró la hoja 'DATOS_PROCESADOS'")
                if datos_procesados:
                    print("Hojas disponibles:")
                    for hoja in datos_procesados.keys():
                        print(f"  - {hoja}")
                return False
                
        except Exception as e:
            print(f"Error cargando datos: {e}")
            return False
    
    def prepare_chart_data(self, chart_config):
        """Prepara los datos para un gráfico específico"""
        columns = chart_config['data_columns']
        chart_type = chart_config['type']
        
        if len(columns) == 1:
            # Datos con una columna (para conteos o distribuciones)
            column = columns[0]
            if chart_type == 'pie':
                data = self.data[column].value_counts()
            else:
                data = self.data[column].dropna()
            
            if 'limit' in chart_config:
                data = data.head(chart_config['limit'])
            
            return data
            
        elif len(columns) == 2:
            # Datos con dos columnas (x, y)
            x_col, y_col = columns
            
            if chart_type == 'box':
                # Para boxplot, devolver los datos sin procesar
                return self.data[[x_col, y_col]].dropna()
            
            elif chart_type == 'scatter':
                # Para scatter, devolver los datos sin procesar
                return self.data[[x_col, y_col]].dropna()
            
            elif chart_type in ['bar', 'line']:
                # Para bar/line, agrupar y calcular promedio o suma según el contexto
                if 'brecha' in y_col.lower() or 'porcentual' in y_col.lower():
                    # Para brechas, calcular promedio
                    data = self.data.groupby(x_col)[y_col].mean()
                elif 'salario' in y_col.lower():
                    # Para salarios, calcular promedio
                    data = self.data.groupby(x_col)[y_col].mean()
                else:
                    # Por defecto, sumar
                    data = self.data.groupby(x_col)[y_col].sum()
                
                # Aplicar límite si existe
                if 'limit' in chart_config:
                    if chart_type == 'bar':
                        data = data.nlargest(chart_config['limit'])
                    else:
                        data = data.head(chart_config['limit'])
                
                return data
        
        # Fallback: devolver datos sin procesar
        return self.data[columns].dropna()
    
    def create_chart(self, chart_id, chart_config):
        """Crea un gráfico individual - especializado para gráficos de donut del registro retributivo"""
        chart_type = chart_config['type']
        title = chart_config['title']
        columns = chart_config['data_columns']
        
        if chart_type == 'donut':
            # Gráfico de donut específico del registro retributivo
            columna_datos = columns[0]  # La columna de salario a analizar
            metodo_calculo = chart_config.get('metodo', 'efectivos_sb')
            subtitulo = chart_config.get('subtitulo', '')
            
            # Seleccionar el método de cálculo apropiado
            if metodo_calculo == 'efectivos_sb':
                datos_genero = self.calcular_promedios_efectivos_sb(self.data, columna_datos)
            elif metodo_calculo == 'efectivos_sb_complementos':
                datos_genero = self.calcular_promedios_efectivos_sb_complementos(self.data, columna_datos)
            elif metodo_calculo == 'equiparados_sb':
                datos_genero = self.calcular_promedios_equiparados_sb(self.data, columna_datos)
            elif metodo_calculo == 'equiparados_sb_complementos':
                datos_genero = self.calcular_promedios_equiparados_sb_complementos(self.data, columna_datos)
            else:
                print(f"Método de cálculo no reconocido: {metodo_calculo}")
                return None
            
            # Crear el gráfico de donut
            fig = self.crear_grafico_donut(datos_genero, title, subtitulo)
            
            # Mostrar estadísticas
            diferencia = datos_genero['H'] - datos_genero['M']
            porcentaje_diferencia = (diferencia / datos_genero['M']) * 100 if datos_genero['M'] > 0 else 0
            
            print(f"📊 {title}")
            print(f"   Hombres: {self.formato_numero_es(datos_genero['H'], 2)}€")
            print(f"   Mujeres: {self.formato_numero_es(datos_genero['M'], 2)}€")
            print(f"   Brecha: {self.formato_numero_es(porcentaje_diferencia, 2)}%")
            print("-" * 50)
            
            # Guardar el gráfico
            chart_filename = f"temp_chart_{chart_id}.png"
            fig.savefig(chart_filename, dpi=300, bbox_inches='tight', 
                       facecolor='white', edgecolor='none')
            plt.close(fig)
            
            self.charts_created[chart_id] = {
                'filename': chart_filename,
                'marker': chart_config['marker'],
                'title': title
            }
            
            return chart_filename
        
        else:
            # Mantener funcionalidad para otros tipos de gráficos
            # Configurar tamaño de figura
            plt.figure(figsize=(12, 8))
            
            if chart_type == 'bar':
                data = self.prepare_chart_data(chart_config)
                bars = plt.bar(range(len(data)), data.values, color=sns.color_palette("husl", len(data)))
                plt.xticks(range(len(data)), data.index, rotation=45, ha='right')
                
                # Añadir valores encima de las barras
                for bar, value in zip(bars, data.values):
                    plt.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(data.values)*0.01,
                            f'{value:,.1f}%' if 'porcentual' in columns[1] else f'{value:,.0f}', 
                            ha='center', va='bottom', fontweight='bold')
            
            elif chart_type == 'pie':
                if len(columns) == 1:
                    # Para una sola columna, usar value_counts
                    data = self.data[columns[0]].value_counts()
                    colors = [self.colores_genero.get(idx, '#7C7C7C') for idx in data.index] if columns[0] == 'Sexo' else sns.color_palette("husl", len(data))
                else:
                    data = self.prepare_chart_data(chart_config)
                    colors = sns.color_palette("husl", len(data))
                
                wedges, texts, autotexts = plt.pie(data.values, labels=data.index, autopct='%1.1f%%', 
                                                  colors=colors, startangle=90)
                
                # Mejorar la apariencia del texto
                for autotext in autotexts:
                    autotext.set_color('white')
                    autotext.set_fontweight('bold')
            
            # Configurar el gráfico
            plt.title(title, fontsize=14, fontweight='bold', pad=20)
            plt.tight_layout()
            
            # Guardar el gráfico
            chart_filename = f"temp_chart_{chart_id}.png"
            plt.savefig(chart_filename, dpi=300, bbox_inches='tight', 
                       facecolor='white', edgecolor='none')
            plt.close()
            
            self.charts_created[chart_id] = {
                'filename': chart_filename,
                'marker': chart_config['marker'],
                'title': title
            }
            
            return chart_filename
    
    def create_all_charts(self):
        """Crea todos los gráficos definidos en la configuración"""
        if self.data is None:
            print("Error: No hay datos cargados")
            return False
        
        for chart_id, chart_config in self.config['charts'].items():
            try:
                print(f"Creando gráfico: {chart_id}")
                self.create_chart(chart_id, chart_config)
            except Exception as e:
                print(f"Error creando gráfico {chart_id}: {e}")
        
        return True
    
    def create_word_document(self):
        """Crea o actualiza el documento Word"""
        template_path = self.config.get('template_word')
        
        # Cargar plantilla o crear documento nuevo
        if template_path and os.path.exists(template_path):
            doc = Document(template_path)
            print(f"Usando plantilla: {template_path}")
        else:
            doc = Document()
            print("Creando documento nuevo")
            
            # Añadir contenido básico si no hay plantilla
            doc.add_heading('Informe Automatizado de Análisis', 0)
            doc.add_paragraph(f'Generado automáticamente el {datetime.now().strftime("%d/%m/%Y %H:%M")}')
            
            for chart_id, chart_info in self.charts_created.items():
                doc.add_heading(chart_info['title'], level=2)
                doc.add_paragraph('A continuación se presenta la visualización correspondiente:')
                doc.add_paragraph(chart_info['marker'])
                doc.add_page_break()
        
        # Reemplazar marcadores con gráficos
        self.replace_markers_with_charts(doc)
        
        # Guardar documento
        output_path = self.config['output_file']
        doc.save(output_path)
        print(f"Documento guardado en: {output_path}")
        
        return output_path
    
    def replace_markers_with_charts(self, doc):
        """Reemplaza los marcadores en el documento con los gráficos"""
        for paragraph in doc.paragraphs:
            for chart_id, chart_info in self.charts_created.items():
                marker = chart_info['marker']
                
                if marker in paragraph.text:
                    print(f"Reemplazando marcador {marker}")
                    
                    # Limpiar el párrafo
                    paragraph.clear()
                    
                    # Añadir la imagen
                    run = paragraph.add_run()
                    run.add_picture(chart_info['filename'], width=Inches(6.5))
                    
                    # Centrar la imagen
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    def cleanup_temp_files(self):
        """Limpia los archivos temporales"""
        for chart_info in self.charts_created.values():
            filename = chart_info['filename']
            if os.path.exists(filename):
                os.remove(filename)
                print(f"Archivo temporal eliminado: {filename}")
    
    def generate_report(self):
        """Ejecuta todo el flujo de generación del reporte"""
        print("🚀 Iniciando generación automatizada del reporte de registro retributivo...")
        
        # 1. Cargar datos
        if not self.load_data():
            return False
        
        # 2. Calcular brechas salariales
        print("📊 Calculando brechas salariales...")
        self.calcular_brecha_salarial()
        
        # 3. Crear todos los gráficos
        if not self.create_all_charts():
            return False
        
        # 4. Crear documento Word
        output_path = self.create_word_document()
        
        # 5. Limpiar archivos temporales
        self.cleanup_temp_files()
        
        print(f"✅ Reporte de registro retributivo generado exitosamente: {output_path}")
        return True

# Archivo de configuración de ejemplo (report_config.yaml)
def create_sample_config():
    """Crea un archivo de configuración de ejemplo para registro retributivo"""
    config = {
        'excel_file': '',  # Se determina automáticamente
        'template_word': 'plantilla_informe.docx',  # Opcional
        'output_file': f'informe_registro_retributivo_{datetime.now().strftime("%Y%m%d_%H%M")}.docx',
        'charts': {
            'salario_base_efectivo': {
                'type': 'donut',
                'data_columns': ['Salario base efectivo Total'],
                'metodo': 'efectivos_sb',
                'title': 'Comparación Salario Base Efectivo Total por Género',
                'subtitulo': 'Análisis de igualdad retributiva - Salario base efectivamente percibido (solo SB > 0)',
                'marker': '{grafico_sb_efectivo}'
            },
            'sb_complementos_efectivo': {
                'type': 'donut',
                'data_columns': ['Salario base anual + complementos Total'],
                'metodo': 'efectivos_sb_complementos',
                'title': 'Salario Base + Complementos Salariales Efectivos por Género',
                'subtitulo': 'Incluye salario base y complementos salariales efectivamente percibidos (todas las personas)',
                'marker': '{grafico_sb_comp_efectivo}'
            },
            'sb_total_efectivo': {
                'type': 'donut',
                'data_columns': ['Salario base anual + complementos + Extrasalariales Total'],
                'metodo': 'efectivos_sb_complementos',
                'title': 'SB + Complementos + Extrasalariales Efectivos por Género',
                'subtitulo': 'Retribución total efectiva incluyendo todos los conceptos (todas las personas)',
                'marker': '{grafico_sb_total_efectivo}'
            },
            'salario_base_equiparado': {
                'type': 'donut',
                'data_columns': ['salario_base_equiparado'],
                'metodo': 'equiparados_sb',
                'title': 'Comparación Salario Base Equiparado por Género',
                'subtitulo': 'Salario base normalizado a jornada completa y 12 meses (solo SB > 0)',
                'marker': '{grafico_sb_equiparado}'
            },
            'sb_complementos_equiparado': {
                'type': 'donut',
                'data_columns': ['sb_mas_comp_salariales_equiparado'],
                'metodo': 'equiparados_sb_complementos',
                'title': 'Salario Base + Complementos Salariales Equiparados por Género',
                'subtitulo': 'SB + complementos salariales normalizados a jornada completa y 12 meses (todas las personas)',
                'marker': '{grafico_sb_comp_equiparado}'
            },
            'sb_total_equiparado': {
                'type': 'donut',
                'data_columns': ['sb_mas_comp_total_equiparado'],
                'metodo': 'equiparados_sb_complementos',
                'title': 'SB + Complementos + Extrasalariales Equiparados por Género',
                'subtitulo': 'Retribución total equiparada: SB + complementos salariales y extrasalariales (todas las personas)',
                'marker': '{grafico_sb_total_equiparado}'
            }
        }
    }
    
    with open('report_config.yaml', 'w', encoding='utf-8') as f:
        yaml.dump(config, f, default_flow_style=False, allow_unicode=True)
    
    print("Archivo de configuración creado: report_config.yaml")

# Script principal
if __name__ == "__main__":
    # Crear configuración de ejemplo si no existe
    if not os.path.exists('report_config.yaml'):
        create_sample_config()
    
    # Generar el reporte
    system = AutomatedReportSystem()
    system.generate_report()