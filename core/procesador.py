"""
Procesador de Registros Retributivos - Versión Web
Refactorizado para trabajar con archivos en memoria (BytesIO)
"""

import io
import re
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
import numpy as np
import msoffcrypto

from .utils import is_positive_response, normalizar_valor, calcular_coef_tp


class ProcesadorRegistroRetributivo:
    """Procesador de registros retributivos que trabaja con archivos en memoria"""

    def __init__(self):
        """Inicializa el procesador con configuración"""
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
        self._config_cache = {}
        self._columnas_complementos_cache = None

    def procesar_excel_general(self, archivo_bytes):
        """
        Procesa un archivo Excel en formato general desde BytesIO

        Args:
            archivo_bytes: BytesIO o bytes del archivo Excel

        Returns:
            BytesIO con el archivo Excel procesado
        """
        # Limpiar cachés
        self._config_cache.clear()
        self._columnas_complementos_cache = None

        # Leer Excel desde bytes
        if isinstance(archivo_bytes, bytes):
            archivo_bytes = io.BytesIO(archivo_bytes)

        # Cargar información de hojas disponibles
        excel_file = pd.ExcelFile(archivo_bytes)

        # Cargar hoja principal
        if "BASE GENERAL" not in excel_file.sheet_names:
            raise Exception("No se encontró la hoja 'BASE GENERAL' requerida")

        archivo_bytes.seek(0)
        df = pd.read_excel(archivo_bytes, sheet_name="BASE GENERAL")

        # LIMPIAR ESPACIOS EN NOMBRES DE COLUMNAS
        df.columns = df.columns.str.strip()

        # Buscar columna "Reg." y eliminar columnas anteriores
        if 'Reg.' in df.columns:
            indice_reg = df.columns.get_loc('Reg.')
            if indice_reg > 0:
                columnas_a_eliminar = df.columns[:indice_reg].tolist()
                df = df.drop(columns=columnas_a_eliminar)

        # Cargar configuración de complementos
        archivo_bytes.seek(0)
        self._cargar_configuracion_complementos(archivo_bytes)

        # Filtrar datos hasta el último registro válido
        if 'Orden' in df.columns:
            indices_con_orden = df[df['Orden'].notna()].index
            if len(indices_con_orden) > 0:
                ultimo_indice_valido = indices_con_orden.max()
                df = df.iloc[:ultimo_indice_valido + 1].copy()

        # Verificar columnas críticas
        columnas_encontradas = {}
        for clave, nombre_col in self.mapeo_columnas.items():
            if nombre_col in df.columns:
                columnas_encontradas[clave] = nombre_col

        if len(columnas_encontradas) < 3:
            raise Exception(f"Faltan columnas críticas. Encontradas: {list(columnas_encontradas.keys())}")

        # Procesar equiparación
        df_equiparado = self._procesar_equiparacion(df, columnas_encontradas)

        # Generar Excel en memoria
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_equiparado.to_excel(writer, sheet_name='DATOS_PROCESADOS', index=False)
        output.seek(0)

        return output

    def procesar_excel_triodos(self, archivo_bytes, password='Triodos2025'):
        """
        Procesa un archivo Excel de Triodos (protegido con contraseña)

        Args:
            archivo_bytes: BytesIO o bytes del archivo Excel
            password: Contraseña del archivo

        Returns:
            BytesIO con el archivo Excel procesado
        """
        # Limpiar cachés
        self._config_cache.clear()
        self._columnas_complementos_cache = None

        # Desencriptar archivo
        if isinstance(archivo_bytes, bytes):
            archivo_bytes = io.BytesIO(archivo_bytes)

        archivo_bytes.seek(0)
        office_file = msoffcrypto.OfficeFile(archivo_bytes)
        office_file.load_key(password=password)

        decrypted = io.BytesIO()
        office_file.decrypt(decrypted)
        decrypted.seek(0)

        # Cargar hoja principal
        df = pd.read_excel(decrypted, sheet_name="BASE GENERAL", engine='openpyxl')
        decrypted.seek(0)

        # LIMPIAR ESPACIOS EN NOMBRES DE COLUMNAS
        df.columns = df.columns.str.strip()

        # Cargar configuración de complementos
        self._cargar_configuracion_complementos_triodos(decrypted)
        decrypted.seek(0)

        # Mapeo de columnas Triodos
        mapeo_triodos = {
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
            'salario_base_efectivo': 'A154-Salario base de nivel*CT',
        }

        # Filtrar filas de totales (sin fechas)
        df = df[
            (df[mapeo_triodos['fecha_inicio_sit']].notna()) |
            (df[mapeo_triodos['num_personal']].isna())
        ].copy()

        # Calcular meses trabajados
        df['¿Cuántos meses ha trabajado?'] = df.apply(
            lambda row: self._calcular_meses_trabajados(
                row[mapeo_triodos['fecha_inicio_sit']],
                row[mapeo_triodos['fecha_fin_sit']]
            ),
            axis=1
        )

        # Mapear columnas al formato maestro
        df['Orden'] = df[mapeo_triodos['num_personal']]
        df['Sexo'] = df[mapeo_triodos['sexo']].map({
            'Masculino': 'Hombres',
            'Femenino': 'Mujeres'
        })
        df['Inicio de Sit. Contractual'] = df[mapeo_triodos['fecha_inicio_sit']]
        df['Final de Sit. Contractual'] = df[mapeo_triodos['fecha_fin_sit']]
        df['Grupo profesional'] = df[mapeo_triodos['grupo_prof']].astype(str)
        df['Categoría profesional'] = df[mapeo_triodos['clasif_interna']].astype(str).str.title()
        df['Puesto de trabajo'] = df[mapeo_triodos['puesto']].astype(str).str.title()
        df['Nivel Convenio Colectivo'] = df[mapeo_triodos['clasif_interna']].astype(str).str.title()
        df['Departamento'] = df[mapeo_triodos['departamento']].astype(str)
        df['Nivel SVPT'] = df[mapeo_triodos['valoracion_puesto']].astype(str)
        df['% de jornada'] = df[mapeo_triodos['jornada_pct']]

        # Calcular coeficiente
        df['Coeficiente Horas Trabajadas Efectivo'] = df.apply(
            lambda row: calcular_coef_tp(row[mapeo_triodos['jornada_pct']]),
            axis=1
        )

        df['Salario base anual efectivo'] = df[mapeo_triodos['salario_base_efectivo']]

        # Procesar equiparación
        columnas_encontradas = {
            'meses_trabajados': '¿Cuántos meses ha trabajado?',
            'coef_tp': '% de jornada',
            'salario_base_efectivo': 'Salario base anual efectivo'
        }
        df_equiparado = self._procesar_equiparacion_triodos(df, columnas_encontradas)

        # Asignar Reg. por empleado
        self._asignar_reg_por_empleado(df_equiparado, mapeo_triodos['num_personal'])

        # Generar Excel en memoria
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_equiparado.to_excel(writer, sheet_name='DATOS_PROCESADOS', index=False)
        output.seek(0)

        return output

    def _cargar_configuracion_complementos(self, archivo_bytes):
        """Carga configuración de complementos desde hojas Excel"""
        nombres_columnas_config = {
            'codigo': 'Cod',
            'nombre': 'Nombre',
            'normalizable': '¿Es Normalizable?',
            'anualizable': '¿Es Anualizable?'
        }

        configuracion = {}
        hojas_config = [
            ('COMPLEMENTOS SALARIALES', 'salarial'),
            ('COMPLEMENTOS EXTRASALARIALES', 'extrasalarial')
        ]

        excel_file = pd.ExcelFile(archivo_bytes)

        for nombre_hoja, tipo in hojas_config:
            if nombre_hoja in excel_file.sheet_names:
                archivo_bytes.seek(0)
                df_comp = pd.read_excel(archivo_bytes, sheet_name=nombre_hoja)

                for _, row in df_comp.iterrows():
                    codigo_val = row.get(nombres_columnas_config['codigo'])
                    if pd.notna(codigo_val):
                        codigo = str(codigo_val).strip()
                        nombre_val = row.get(nombres_columnas_config['nombre'])
                        nombre = str(nombre_val).strip() if pd.notna(nombre_val) else ''

                        configuracion[codigo] = {
                            'tipo': tipo,
                            'nombre': nombre,
                            'es_normalizable': is_positive_response(row.get(nombres_columnas_config['normalizable'])),
                            'es_anualizable': is_positive_response(row.get(nombres_columnas_config['anualizable']))
                        }

        self.configuracion_complementos = configuracion

    def _cargar_configuracion_complementos_triodos(self, archivo_bytes):
        """Carga configuración de complementos de Triodos"""
        nombres_columnas_config = {
            'codigo': 'Cod',
            'nombre': 'Nombre',
            'normalizable': '¿Es Normalizable?',
            'anualizable': '¿Es Anualizable?'
        }

        configuracion = {}
        hojas_config = [
            ('COMPLEMENTOS SALARIALES', 'salarial'),
            ('COMPLEMENTOS EXTRASALARIALES', 'extrasalarial')
        ]

        for nombre_hoja, tipo in hojas_config:
            try:
                archivo_bytes.seek(0)
                df_comp = pd.read_excel(archivo_bytes, sheet_name=nombre_hoja, engine='openpyxl')

                for _, row in df_comp.iterrows():
                    nombre_val = row.get(nombres_columnas_config['nombre'])

                    if pd.notna(nombre_val):
                        nombre_completo = str(nombre_val).strip()
                        codigo_a = nombre_completo.split('-')[0].strip() if '-' in nombre_completo else nombre_completo

                        configuracion[codigo_a] = {
                            'tipo': tipo,
                            'nombre_completo': nombre_completo,
                            'es_normalizable': is_positive_response(row.get(nombres_columnas_config['normalizable'])),
                            'es_anualizable': is_positive_response(row.get(nombres_columnas_config['anualizable']))
                        }
            except Exception:
                pass

        self.configuracion_complementos = configuracion

    def _procesar_equiparacion(self, df, columnas_encontradas):
        """Procesa la equiparación de todos los valores"""
        df_equiparado = df.copy()

        col_meses = columnas_encontradas.get('meses_trabajados')
        col_coef_tp = columnas_encontradas.get('coef_tp')
        col_sb_efectivo = columnas_encontradas.get('salario_base_efectivo')

        # Calcular coeficiente TP
        coef_tp_values = df_equiparado[col_coef_tp].fillna(1.0)
        df_equiparado['coef_tp_calculado'] = np.where(
            coef_tp_values > 1,
            coef_tp_values / 100,
            coef_tp_values
        )

        # Equiparar salario base (vectorizado)
        sb_efectivo = df_equiparado[col_sb_efectivo].fillna(0)
        coef_tp_norm = df_equiparado['coef_tp_calculado'].replace(0, 1.0).fillna(1.0)
        meses_norm = df_equiparado[col_meses].replace(0, 12).fillna(12)

        df_equiparado['salario_base_equiparado'] = np.where(
            sb_efectivo == 0,
            0,
            sb_efectivo * (1 / coef_tp_norm) * (12 / meses_norm)
        )

        # Procesar complementos
        self._procesar_complementos_individuales(df_equiparado, col_meses)
        self._calcular_totales_complementos(df_equiparado)

        return df_equiparado

    def _procesar_equiparacion_triodos(self, df, columnas_encontradas):
        """Procesa equiparación para formato Triodos"""
        df_equiparado = df.copy()

        col_meses = '¿Cuántos meses ha trabajado?'
        col_sb_efectivo = 'Salario base anual efectivo'

        # Equiparar salario base
        sb_efectivo = df_equiparado[col_sb_efectivo].fillna(0)
        coef_tp_norm = df_equiparado['Coeficiente Horas Trabajadas Efectivo'].replace(0, 1.0).fillna(1.0)
        meses_norm = df_equiparado[col_meses].replace(0, 12).fillna(12)

        df_equiparado['salario_base_equiparado'] = np.where(
            sb_efectivo == 0,
            0,
            sb_efectivo * (1 / coef_tp_norm) * (12 / meses_norm)
        )

        # Procesar complementos Triodos
        self._procesar_complementos_triodos(df_equiparado, col_meses)
        self._calcular_totales_complementos_triodos(df_equiparado)

        return df_equiparado

    def _procesar_complementos_individuales(self, df_equiparado, col_meses):
        """Procesa complementos PS y PE individuales"""
        if not self.configuracion_complementos:
            return

        columnas_por_tipo = self._obtener_columnas_complementos(df_equiparado)
        renombrar_dict = {}

        for tipo, columnas in columnas_por_tipo.items():
            for col_comp in columnas:
                es_normalizable, es_anualizable, _, nombre_comp = self._obtener_config_complemento(col_comp)

                if nombre_comp and nombre_comp not in col_comp:
                    nombre_completo = f"{col_comp} {nombre_comp}".strip()
                    renombrar_dict[col_comp] = nombre_completo
                else:
                    nombre_completo = col_comp

                if es_normalizable or es_anualizable:
                    datos_no_nulos = df_equiparado[col_comp].dropna()
                    if len(datos_no_nulos) > 0:
                        col_equiparado = f"{nombre_completo}_equiparado"

                        comp_efectivo = df_equiparado[col_comp].fillna(0)
                        resultado = comp_efectivo.copy()

                        if es_normalizable:
                            coef_tp_norm = df_equiparado['coef_tp_calculado'].replace(0, 1.0).fillna(1.0)
                            resultado = resultado * (1 / coef_tp_norm)

                        if es_anualizable:
                            meses_norm = df_equiparado[col_meses].replace(0, 12).fillna(12)
                            resultado = resultado * (12 / meses_norm)

                        df_equiparado[col_equiparado] = np.where(
                            comp_efectivo == 0,
                            0,
                            resultado
                        )

        if renombrar_dict:
            df_equiparado.rename(columns=renombrar_dict, inplace=True)

    def _procesar_complementos_triodos(self, df_equiparado, col_meses):
        """Procesa complementos de Triodos (A###, PA###, PC###)"""
        if not self.configuracion_complementos:
            return

        columnas_por_tipo = self._obtener_columnas_complementos_triodos(df_equiparado)

        for tipo, columnas in columnas_por_tipo.items():
            for col_comp in columnas:
                codigo = col_comp.split('-')[0].strip() if '-' in col_comp else col_comp[:4].strip()
                es_normalizable, es_anualizable, _, _ = self._obtener_config_complemento(codigo)

                datos_no_nulos = df_equiparado[col_comp].dropna()
                if len(datos_no_nulos) > 0:
                    col_equiparado = f"{col_comp}_equiparado"

                    if es_normalizable or es_anualizable:
                        comp_efectivo = df_equiparado[col_comp].fillna(0)
                        resultado = comp_efectivo.copy()

                        if es_normalizable:
                            coef_tp_norm = df_equiparado['Coeficiente Horas Trabajadas Efectivo'].replace(0, 1.0).fillna(1.0)
                            resultado = resultado * (1 / coef_tp_norm)

                        if es_anualizable:
                            meses_norm = df_equiparado[col_meses].replace(0, 12).fillna(12)
                            resultado = resultado * (12 / meses_norm)

                        df_equiparado[col_equiparado] = np.where(
                            comp_efectivo == 0,
                            0,
                            resultado
                        )
                    else:
                        df_equiparado[col_equiparado] = df_equiparado[col_comp].copy()

    def _calcular_totales_complementos(self, df_equiparado):
        """Calcula totales de complementos equiparados"""
        columnas_por_tipo = self._obtener_columnas_complementos(df_equiparado)

        df_equiparado['complementos_salariales_equiparados'] = df_equiparado.apply(
            lambda row: self._calcular_total_correcto(row, columnas_por_tipo['PS'], df_equiparado), axis=1
        )

        df_equiparado['complementos_extrasalariales_equiparados'] = df_equiparado.apply(
            lambda row: self._calcular_total_correcto(row, columnas_por_tipo['PE'], df_equiparado), axis=1
        )

        self._calcular_columnas_combinadas(df_equiparado)

    def _calcular_totales_complementos_triodos(self, df_equiparado):
        """Calcula totales de complementos para Triodos"""
        columnas_por_tipo = self._obtener_columnas_complementos_triodos(df_equiparado)

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

        self._calcular_columnas_combinadas(df_equiparado)

    def _calcular_columnas_combinadas(self, df_equiparado):
        """Calcula columnas combinadas (SB + Comp)"""
        if 'salario_base_equiparado' in df_equiparado.columns and 'complementos_salariales_equiparados' in df_equiparado.columns:
            df_equiparado['sb_mas_comp_salariales_equiparado'] = (
                df_equiparado['salario_base_equiparado'].fillna(0) +
                df_equiparado['complementos_salariales_equiparados'].fillna(0)
            )

        if ('sb_mas_comp_salariales_equiparado' in df_equiparado.columns and
            'complementos_extrasalariales_equiparados' in df_equiparado.columns):
            df_equiparado['sb_mas_comp_total_equiparado'] = (
                df_equiparado['sb_mas_comp_salariales_equiparado'].fillna(0) +
                df_equiparado['complementos_extrasalariales_equiparados'].fillna(0)
            )

    def _calcular_total_correcto(self, row, columnas_base, df_equiparado):
        """Calcula total: equiparado si existe, efectivo si no"""
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

    def _obtener_config_complemento(self, codigo_complemento):
        """Obtiene configuración de un complemento (con caché)"""
        if codigo_complemento in self._config_cache:
            return self._config_cache[codigo_complemento]

        codigos_a_buscar = [codigo_complemento]
        match = re.match(r'^(P[SE])\s*(\d+)', codigo_complemento)
        if match:
            codigo_base = f"{match.group(1)}{match.group(2)}"
            codigos_a_buscar.append(codigo_base)

        if codigo_complemento.isdigit():
            codigos_a_buscar.append(f"PS{codigo_complemento}")

        for codigo in codigos_a_buscar:
            if codigo in self.configuracion_complementos:
                config = self.configuracion_complementos[codigo]
                resultado = (
                    config['es_normalizable'],
                    config['es_anualizable'],
                    config['tipo'],
                    config.get('nombre', config.get('nombre_completo', ''))
                )
                self._config_cache[codigo_complemento] = resultado
                return resultado

        resultado = (False, False, 'desconocido', '')
        self._config_cache[codigo_complemento] = resultado
        return resultado

    def _obtener_columnas_complementos(self, df, prefijos=['PS', 'PE']):
        """Obtiene columnas de complementos (con caché)"""
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

        self._columnas_complementos_cache = columnas_por_tipo
        return columnas_por_tipo

    def _obtener_columnas_complementos_triodos(self, df):
        """Obtiene columnas de complementos Triodos (A###, PA###, PC###)"""
        if self._columnas_complementos_cache is not None:
            return self._columnas_complementos_cache

        columnas_comp = []
        for col in df.columns:
            if col.startswith('A') and col != 'A154-Salario base de nivel*CT':
                columnas_comp.append(col)
            elif col.startswith('PA') or col.startswith('PC'):
                columnas_comp.append(col)

        columnas_por_tipo = {'PS': [], 'PE': []}
        for col in columnas_comp:
            codigo = col.split('-')[0].strip() if '-' in col else col[:4].strip()

            if codigo in self.configuracion_complementos:
                tipo = self.configuracion_complementos[codigo]['tipo']
                if tipo == 'salarial':
                    columnas_por_tipo['PS'].append(col)
                elif tipo == 'extrasalarial':
                    columnas_por_tipo['PE'].append(col)
            else:
                if 'CE' in col or 'Exento' in col or col.startswith('PA') or col.startswith('PC'):
                    columnas_por_tipo['PE'].append(col)
                else:
                    columnas_por_tipo['PS'].append(col)

        self._columnas_complementos_cache = columnas_por_tipo
        return columnas_por_tipo

    def _calcular_meses_trabajados(self, fecha_inicio, fecha_fin):
        """Calcula meses trabajados con precisión decimal"""
        if pd.isna(fecha_inicio) or pd.isna(fecha_fin):
            return 12.0

        from calendar import monthrange

        if isinstance(fecha_inicio, pd.Timestamp):
            fecha_inicio = fecha_inicio.to_pydatetime()
        if isinstance(fecha_fin, pd.Timestamp):
            fecha_fin = fecha_fin.to_pydatetime()

        # Caso especial: período completo
        if fecha_inicio.day == 1 and fecha_fin.day == monthrange(fecha_fin.year, fecha_fin.month)[1]:
            delta = relativedelta(fecha_fin, fecha_inicio)
            return float(delta.years * 12 + delta.months + 1)

        # Caso general
        ultimo_dia_mes_inicio = monthrange(fecha_inicio.year, fecha_inicio.month)[1]
        ultimo_dia_mes_fin = monthrange(fecha_fin.year, fecha_fin.month)[1]

        if fecha_inicio.day == 1:
            dias_inicio = 0
            mes_inicio_es_completo = True
        else:
            dias_inicio = ultimo_dia_mes_inicio - fecha_inicio.day + 1
            mes_inicio_es_completo = False

        if fecha_fin.day == ultimo_dia_mes_fin:
            dias_fin = 0
            mes_fin_es_completo = True
        else:
            dias_fin = fecha_fin.day
            mes_fin_es_completo = False

        delta = relativedelta(fecha_fin, fecha_inicio)
        total_meses_diff = delta.years * 12 + delta.months

        if mes_inicio_es_completo and mes_fin_es_completo:
            meses_completos = total_meses_diff + 1
        elif mes_inicio_es_completo:
            meses_completos = total_meses_diff
        elif mes_fin_es_completo:
            meses_completos = total_meses_diff
        else:
            meses_completos = total_meses_diff - 1 if total_meses_diff > 0 else 0

        meses = (dias_inicio * 12.0 / 365.0) + meses_completos + (dias_fin * 12.0 / 365.0)
        return max(0.01, min(12.0, meses))

    def _asignar_reg_por_empleado(self, df, col_num_personal):
        """Asigna valores 'Ex' o vacío a columna Reg. según situaciones contractuales"""
        df['Reg.'] = pd.Series([''] * len(df), index=df.index, dtype='object')

        for num_personal in df[col_num_personal].unique():
            if pd.isna(num_personal):
                continue

            mask_empleado = df[col_num_personal] == num_personal
            indices_empleado = df[mask_empleado].index.tolist()

            if len(indices_empleado) > 1:
                filas_empleado = df.loc[indices_empleado].copy()
                filas_empleado_sorted = filas_empleado.sort_values(
                    by='Final de Sit. Contractual',
                    na_position='last'
                )

                indices_antiguos = filas_empleado_sorted.index[:-1]
                df.loc[indices_antiguos, 'Reg.'] = 'Ex'

                indice_ultimo = filas_empleado_sorted.index[-1]
                df.at[indice_ultimo, 'Reg.'] = ''
