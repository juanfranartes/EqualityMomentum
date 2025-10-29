"""
Utilidades comunes para procesamiento de datos
"""

import pandas as pd
import numpy as np
import re


def formato_numero_es(numero, decimales=2):
    """Formatea número con estilo español"""
    if pd.isna(numero) or numero is None:
        return "0,00" if decimales == 2 else "0"
    formato = f"{{:,.{decimales}f}}"
    num_str = formato.format(numero)
    return num_str.replace(',', 'X').replace('.', ',').replace('X', '.')


def calcular_brecha(valor_h, valor_m):
    """Calcula brecha salarial: ((H - M) / H) * 100"""
    if valor_h is None or valor_h <= 0 or valor_m is None or valor_m < 0:
        return None
    return ((valor_h - valor_m) / valor_h) * 100


def is_positive_response(value):
    """Verifica si un valor representa una respuesta positiva (Sí/Si/YES)"""
    if pd.isna(value):
        return False
    normalized = str(value).strip().lower()
    return normalized in ['sí', 'si', 'yes', 'y', '1', 'true']


def normalizar_valor(valor, valor_default):
    """Normaliza un valor, retornando el default si es inválido"""
    return valor_default if pd.isna(valor) or valor == 0 else valor


def calcular_coef_tp(valor_coef_tp):
    """Convierte el coeficiente de tiempo parcial a decimal"""
    if pd.isna(valor_coef_tp):
        return 1.0
    return valor_coef_tp / 100 if valor_coef_tp > 1 else valor_coef_tp


def reformatear_etiqueta_escala(etiqueta):
    """
    Reformatea etiquetas de 'Escala X + Nombre' a 'Nombre - EX'
    Ejemplo: 'Escala 2 + Offside Leader' -> 'Offside Leader - E2'
    """
    if not etiqueta or not isinstance(etiqueta, str):
        return etiqueta

    patron = r'Escala\s+(\d+)\s*\+\s*(.+)'
    match = re.match(patron, etiqueta.strip(), re.IGNORECASE)

    if match:
        numero_escala = match.group(1)
        nombre = match.group(2).strip()
        return f'{nombre} - E{numero_escala}'

    return etiqueta
