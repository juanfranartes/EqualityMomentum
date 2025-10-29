# -*- coding: utf-8 -*-
"""
Validador y Mapeador de Hojas y Variables
Permite mapeo manual cuando los nombres esperados no coinciden
"""

import pandas as pd
from typing import Dict, List, Tuple, Optional


class ValidadorMapeo:
    """
    Clase para validar y mapear hojas y variables de archivos Excel
    """

    def __init__(self):
        """Inicializa el validador"""
        self.mapeo_hojas = {}
        self.mapeo_variables = {}

    def validar_hojas(self, excel_file: pd.ExcelFile, hojas_requeridas: List[str]) -> Dict:
        """
        Valida que existan las hojas requeridas en el archivo Excel.

        Args:
            excel_file: Archivo Excel abierto con pd.ExcelFile
            hojas_requeridas: Lista de nombres de hojas requeridas

        Returns:
            Dict con:
                - 'encontradas': Dict {nombre_esperado: nombre_real}
                - 'faltantes': List de nombres de hojas no encontradas
                - 'disponibles': List de todas las hojas disponibles
        """
        hojas_disponibles = excel_file.sheet_names
        encontradas = {}
        faltantes = []

        for hoja_requerida in hojas_requeridas:
            if hoja_requerida in hojas_disponibles:
                # La hoja existe con el nombre esperado
                encontradas[hoja_requerida] = hoja_requerida
            else:
                # La hoja no existe con el nombre esperado
                faltantes.append(hoja_requerida)

        return {
            'encontradas': encontradas,
            'faltantes': faltantes,
            'disponibles': hojas_disponibles
        }

    def validar_variables(self, df: pd.DataFrame, variables_requeridas: Dict[str, str]) -> Dict:
        """
        Valida que existan las variables (columnas) requeridas en el DataFrame.

        Args:
            df: DataFrame con los datos
            variables_requeridas: Dict {clave_interna: nombre_esperado_columna}

        Returns:
            Dict con:
                - 'encontradas': Dict {clave_interna: nombre_real_columna}
                - 'faltantes': Dict {clave_interna: nombre_esperado_columna}
                - 'disponibles': List de todas las columnas disponibles
        """
        # Limpiar nombres de columnas
        columnas_disponibles = [str(col).strip() for col in df.columns]
        encontradas = {}
        faltantes = {}

        for clave_interna, nombre_esperado in variables_requeridas.items():
            nombre_esperado_limpio = str(nombre_esperado).strip()

            if nombre_esperado_limpio in columnas_disponibles:
                # La columna existe con el nombre esperado
                encontradas[clave_interna] = nombre_esperado_limpio
            else:
                # La columna no existe con el nombre esperado
                faltantes[clave_interna] = nombre_esperado_limpio

        return {
            'encontradas': encontradas,
            'faltantes': faltantes,
            'disponibles': columnas_disponibles
        }

    def aplicar_mapeo_hojas(self, mapeo_usuario: Dict[str, str]):
        """
        Aplica el mapeo de hojas proporcionado por el usuario.

        Args:
            mapeo_usuario: Dict {nombre_esperado: nombre_real_seleccionado}
        """
        self.mapeo_hojas.update(mapeo_usuario)

    def aplicar_mapeo_variables(self, mapeo_usuario: Dict[str, str]):
        """
        Aplica el mapeo de variables proporcionado por el usuario.

        Args:
            mapeo_usuario: Dict {clave_interna: nombre_real_columna_seleccionado}
        """
        self.mapeo_variables.update(mapeo_usuario)

    def obtener_nombre_hoja(self, nombre_esperado: str) -> str:
        """
        Obtiene el nombre real de una hoja, usando el mapeo si existe.

        Args:
            nombre_esperado: Nombre de hoja esperado por el script

        Returns:
            Nombre real de la hoja (mapeado o esperado)
        """
        return self.mapeo_hojas.get(nombre_esperado, nombre_esperado)

    def obtener_nombre_variable(self, clave_interna: str, nombre_esperado: str) -> str:
        """
        Obtiene el nombre real de una variable, usando el mapeo si existe.

        Args:
            clave_interna: Clave interna usada por el script
            nombre_esperado: Nombre de columna esperado por defecto

        Returns:
            Nombre real de la columna (mapeado o esperado)
        """
        return self.mapeo_variables.get(clave_interna, nombre_esperado)

    def obtener_mapeo_completo_variables(self, variables_requeridas: Dict[str, str]) -> Dict[str, str]:
        """
        Obtiene el mapeo completo de variables, combinando el mapeo por defecto con el del usuario.

        Args:
            variables_requeridas: Dict {clave_interna: nombre_esperado_columna}

        Returns:
            Dict {clave_interna: nombre_real_columna}
        """
        mapeo_completo = {}
        for clave_interna, nombre_esperado in variables_requeridas.items():
            mapeo_completo[clave_interna] = self.obtener_nombre_variable(clave_interna, nombre_esperado)
        return mapeo_completo


class ValidadorMapeoGeneral(ValidadorMapeo):
    """
    Validador específico para archivos del tipo General (procesar_datos.py)
    """

    def __init__(self):
        super().__init__()

        # Hojas requeridas para el procesador general
        self.hojas_requeridas = [
            'BASE GENERAL',
            'COMPLEMENTOS SALARIALES',
            'COMPLEMENTOS EXTRASALARIALES'
        ]

        # Variables críticas requeridas (mapeo del archivo procesar_datos.py)
        self.variables_criticas = {
            'meses_trabajados': '¿Cuántos meses ha trabajado?',
            'coef_tp': '% de jornada',
            'salario_base_efectivo': 'Salario base anual efectivo',
            'complementos_salariales_efectivo': 'Complementos Salariales efectivo',
            'complementos_extrasalariales_efectivo': 'Complementos Extrasalariales efectivo'
        }

        # Columnas de configuración de complementos
        self.columnas_config_complementos = {
            'codigo': 'Cod',
            'nombre': 'Nombre',
            'normalizable': '¿Es Normalizable?',
            'anualizable': '¿Es Anualizable?'
        }


class ValidadorMapeoTriodos(ValidadorMapeo):
    """
    Validador específico para archivos de Triodos Bank (procesar_datos_triodos.py)
    """

    def __init__(self):
        super().__init__()

        # Hojas requeridas para el procesador Triodos
        self.hojas_requeridas = [
            'BASE GENERAL',
            'COMPLEMENTOS SALARIALES',
            'COMPLEMENTOS EXTRASALARIALES'
        ]

        # Variables críticas requeridas (mapeo del archivo procesar_datos_triodos.py)
        self.variables_criticas = {
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

        # Columnas de configuración de complementos
        self.columnas_config_complementos = {
            'codigo': 'Cod',
            'nombre': 'Nombre',
            'normalizable': '¿Es Normalizable?',
            'anualizable': '¿Es Anualizable?'
        }
