"""
lector.py — Lee libros de remuneraciones del cliente y extrae sus columnas/conceptos.

Estructura esperada de los libros:
  Fila 1 (Excel): Nombre de la empresa
  Fila 2 (Excel): Encabezados reales de columna
  Fila 3+:        Datos
"""

import pandas as pd
import unicodedata
import re
from io import BytesIO


# Patrones de columnas metadata (no son conceptos de remuneración)
PATRONES_META = [
    'año', 'mes', 'rut', 'nombre', 'apellido', 'trabajador', 'empleado',
    'empresa', 'area', 'zona', 'sucursal', 'centro costo', 'centro de costo',
    'cargo', 'contrato', 'jornada', 'n°', 'nro', 'numero', 'numero', 'folio',
    'dias trabajados', 'dias licencia', 'dias ausencia', 'hhs.', 'horas a pago',
    'remuneracion imponible', 'remuneracion no imponible', 'remuneracion total',
    'renta tributable', 'descuentos legales', 'otros descuentos',
    'sueldo liquido', 'anticipos',
]


def normalizar(texto: str) -> str:
    """Minúsculas, sin tildes, sin espacios extra."""
    if not isinstance(texto, str):
        texto = str(texto)
    texto = texto.strip().lower()
    texto = unicodedata.normalize('NFD', texto)
    texto = ''.join(c for c in texto if unicodedata.category(c) != 'Mn')
    return texto


def es_columna_meta(nombre_col: str) -> bool:
    """Detecta si una columna es metadata del trabajador (no concepto de remuneración)."""
    col_norm = normalizar(nombre_col)
    return any(patron in col_norm for patron in PATRONES_META)


def detectar_fila_header(df_raw: pd.DataFrame) -> int:
    """
    Detecta en qué fila están los encabezados reales del libro.
    Los libros pueden tener una fila de título de empresa antes de los headers.

    Returns:
        Índice de fila (0-based) donde están los encabezados reales.
    """
    # Si las columnas son todas Unnamed, los headers reales están en una fila de datos
    unnamed_count = sum(1 for c in df_raw.columns if str(c).startswith('Unnamed'))
    if unnamed_count > len(df_raw.columns) * 0.7:
        # Buscar la primera fila que tenga varios valores de texto (los headers reales)
        for i in range(min(5, len(df_raw))):
            row = df_raw.iloc[i]
            text_vals = sum(1 for v in row if isinstance(v, str) and len(str(v).strip()) > 1)
            if text_vals >= 5:
                return i + 1  # header está en esta fila → leer con header=i+1
    return 0  # los headers ya están bien (fila 0 del Excel)


def leer_libro(archivo_bytes: bytes) -> tuple:
    """
    Lee un libro de remuneraciones desde bytes y clasifica sus columnas.

    Returns:
        (df, conceptos, metadata, fila_header_usada)
    """
    # Primer pase: detectar dónde están los headers
    df_raw = pd.read_excel(BytesIO(archivo_bytes), header=0, nrows=10)
    fila_header = detectar_fila_header(df_raw)

    # Segundo pase: leer con el header correcto
    df = pd.read_excel(BytesIO(archivo_bytes), header=fila_header)

    # Filtrar columnas vacías o sin nombre útil
    columnas_validas = []
    for col in df.columns:
        col_str = str(col).strip()
        if col_str and not col_str.startswith('Unnamed') and col_str.lower() != 'nan':
            columnas_validas.append(col)

    conceptos = []
    metadata = []
    for col in columnas_validas:
        if es_columna_meta(str(col)):
            metadata.append(col)
        else:
            conceptos.append(col)

    return df[columnas_validas], conceptos, metadata


def leer_multiples_libros(archivos_streamlit: list) -> tuple:
    """
    Lee múltiples libros y retorna el conjunto unificado de conceptos.

    Returns:
        (todos_conceptos: dict{col -> [archivos]}, errores: list[str])
    """
    todos_conceptos: dict[str, list] = {}
    errores: list[str] = []

    for archivo in archivos_streamlit:
        nombre = archivo.name
        try:
            archivo_bytes = archivo.read()
            archivo.seek(0)
            _, conceptos, _ = leer_libro(archivo_bytes)

            for concepto in conceptos:
                col_str = str(concepto)
                if col_str not in todos_conceptos:
                    todos_conceptos[col_str] = []
                todos_conceptos[col_str].append(nombre)

        except Exception as e:
            errores.append(f"{nombre}: {str(e)}")

    return todos_conceptos, errores
