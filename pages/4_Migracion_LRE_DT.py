"""
Rex+ Tools · 4 — Migración desde LRE DT
Etapa 1: Armado del maestro (migracion_lre_dt.py) y Etapa 2: archivo de carga (migracion_lre_dt_etapa2.py);
esta página solo lo ejecuta dentro de Rex+ Tools.
"""
import os
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

from migracion_lre_dt import main  # noqa: E402

main()
