"""
comparador.py — Lógica de cruce entre conceptos del cliente y el listado de Rex+.
"""

import re
import pandas as pd
import unicodedata
from io import BytesIO
from rapidfuzz import process, fuzz


def normalizar(texto: str) -> str:
    """Minúsculas, sin tildes, sin espacios extra."""
    if not isinstance(texto, str):
        texto = str(texto)
    texto = texto.strip().lower()
    texto = unicodedata.normalize('NFD', texto)
    texto = ''.join(c for c in texto if unicodedata.category(c) != 'Mn')
    return texto


def limpiar_concepto_cliente(col: str) -> str:
    """
    Limpia el nombre de columna del libro del cliente para mejorar el matching.
    Elimina:
      - Prefijo numérico estilo "0001 " o "1522 "
      - Sufijos indicadores como "(M)", "(A)", "(GROSS UP)"
    Ejemplo: "0001 SUELDO BASE (M)" → "SUELDO BASE"
    """
    # Quitar prefijo numérico (2-4 dígitos seguido de espacio)
    col = re.sub(r'^\d{2,4}\s+', '', col.strip())
    # Quitar sufijos entre paréntesis al final
    col = re.sub(r'\s*\([^)]*\)\s*$', '', col)
    return col.strip()


def cargar_conceptos_rex(archivo_bytes: bytes) -> pd.DataFrame:
    """
    Carga el listado de conceptos de Rex+.
    El archivo tiene una fila de título ('Lista de conceptos') y luego los encabezados.

    Returns:
        DataFrame con columnas: Concepto, Nombre, Tipo
    """
    df = pd.read_excel(BytesIO(archivo_bytes), header=1)
    df = df[['Concepto', 'Nombre', 'Tipo']].dropna(subset=['Concepto'])
    df = df[df['Concepto'].astype(str).str.strip() != '']
    return df.reset_index(drop=True)


def cruzar_conceptos(
    conceptos_cliente: list[str],
    df_rex: pd.DataFrame,
    umbral_alto: int = 85,
    umbral_bajo: int = 60
) -> dict:
    """
    Cruza los conceptos del cliente con los de Rex+.

    Estrategia de matching (en orden):
    1. Exacto por código Rex (Concepto)
    2. Exacto por nombre Rex (Nombre)
    3. Exacto tras limpiar prefijo numérico y sufijos
    4. Fuzzy match sobre nombre limpio vs. Concepto y Nombre de Rex
    5. Sin match

    Returns:
        dict con claves 'match', 'dudoso', 'sin_match'
    """
    conceptos_rex = df_rex['Concepto'].tolist()
    nombres_rex   = df_rex['Nombre'].tolist()

    # Índices normalizados para búsqueda exacta
    norm_concepto = {normalizar(c): i for i, c in enumerate(conceptos_rex)}
    norm_nombre   = {normalizar(n): i for i, n in enumerate(nombres_rex)}

    # Versiones normalizadas para fuzzy
    conceptos_rex_norm = [normalizar(c) for c in conceptos_rex]
    nombres_rex_norm   = [normalizar(n) for n in nombres_rex]

    resultados: dict[str, list] = {
        'match':     [],
        'dudoso':    [],
        'sin_match': [],
    }

    for col in conceptos_cliente:
        col_norm    = normalizar(col)
        col_limpio  = limpiar_concepto_cliente(col)
        col_limpio_norm = normalizar(col_limpio)

        # ── 1. Exacto por código Rex ──────────────────────────────────
        if col_norm in norm_concepto:
            idx = norm_concepto[col_norm]
            row = df_rex.iloc[idx]
            resultados['match'].append(_mk_match(col, row, 100, 'Exacto (código)'))
            continue

        # ── 2. Exacto por nombre Rex ──────────────────────────────────
        if col_norm in norm_nombre:
            idx = norm_nombre[col_norm]
            row = df_rex.iloc[idx]
            resultados['match'].append(_mk_match(col, row, 100, 'Exacto (nombre)'))
            continue

        # ── 3. Exacto tras limpiar prefijo/sufijo ─────────────────────
        if col_limpio_norm and col_limpio_norm != col_norm:
            if col_limpio_norm in norm_concepto:
                idx = norm_concepto[col_limpio_norm]
                row = df_rex.iloc[idx]
                resultados['match'].append(_mk_match(col, row, 98, f'Exacto limpio (código) ← "{col_limpio}"'))
                continue
            if col_limpio_norm in norm_nombre:
                idx = norm_nombre[col_limpio_norm]
                row = df_rex.iloc[idx]
                resultados['match'].append(_mk_match(col, row, 98, f'Exacto limpio (nombre) ← "{col_limpio}"'))
                continue

        # ── 4. Fuzzy match ────────────────────────────────────────────
        # Usar el nombre limpio para mejorar la calidad del fuzzy
        query = col_limpio_norm if col_limpio_norm else col_norm

        m_cod = process.extract(query, conceptos_rex_norm, scorer=fuzz.token_sort_ratio, limit=4)
        m_nom = process.extract(query, nombres_rex_norm,   scorer=fuzz.token_sort_ratio, limit=4)

        candidatos = []
        for _, score, idx in m_cod:
            row = df_rex.iloc[idx]
            candidatos.append(_mk_sug(row, score))
        for _, score, idx in m_nom:
            row = df_rex.iloc[idx]
            candidatos.append(_mk_sug(row, score))

        # Deduplicar por concepto Rex, conservar score más alto
        seen: dict[str, dict] = {}
        for c in candidatos:
            key = c['concepto']
            if key not in seen or c['score'] > seen[key]['score']:
                seen[key] = c
        candidatos_uniq = sorted(seen.values(), key=lambda x: -x['score'])

        mejor = candidatos_uniq[0] if candidatos_uniq else None

        if mejor and mejor['score'] >= umbral_alto:
            resultados['match'].append({
                'col_cliente':  col,
                'concepto_rex': mejor['concepto'],
                'nombre_rex':   mejor['nombre'],
                'tipo_rex':     mejor['tipo'],
                'score':        mejor['score'],
                'metodo':       f'Fuzzy ({mejor["score"]:.0f}%) ← "{col_limpio}"'
            })
        elif mejor and mejor['score'] >= umbral_bajo:
            resultados['dudoso'].append({
                'col_cliente': col,
                'col_limpia':  col_limpio,
                'sugerencias': candidatos_uniq[:4]
            })
        else:
            resultados['sin_match'].append({'col_cliente': col})

    return resultados


# ── Helpers internos ───────────────────────────────────────────────────────────

def _mk_match(col: str, row: pd.Series, score: int | float, metodo: str) -> dict:
    return {
        'col_cliente':  col,
        'concepto_rex': row['Concepto'],
        'nombre_rex':   row['Nombre'],
        'tipo_rex':     row['Tipo'],
        'score':        score,
        'metodo':       metodo,
    }

def _mk_sug(row: pd.Series, score: float) -> dict:
    return {
        'concepto': row['Concepto'],
        'nombre':   row['Nombre'],
        'tipo':     row['Tipo'],
        'score':    score,
    }
