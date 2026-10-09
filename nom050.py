"""Reglas NOM-050-SCFI-2004: leyenda del CONTENIDO según su unidad."""
from __future__ import annotations

import os
import re

from catalogos import cargar_lista

NUMERO_NORMA = 50

# Catálogo editable de unidades de masa o volumen: si el CONTENIDO trae una de
# ellas lleva "CONTENIDO NETO"; si no (piezas, latas...), solo "CONTENIDO".
CATALOGO_UNIDADES_PATH = os.path.join("data", "unidades_nom050.json")

UNIDADES_CONTENIDO_NETO = [
    "MILILITRO", "MILILITROS", "ML", "MLS", "LITRO", "LITROS", "LT", "LTS", "L",
    "GRAMO", "GRAMOS", "GR", "GRS", "G", "KILOGRAMO", "KILOGRAMOS", "KILO", "KILOS", "KG", "KGS", "K",
]

_LEYENDA_CONTENIDO = re.compile(r"^CONTENIDO(\s+NETO)?\s*:?\s*", re.IGNORECASE)

# Patrón compilado + fecha de modificación del catálogo con que se armó: se
# reconstruye solo cuando el catálogo cambia (ej. desde Configuración).
_cache: tuple[float, re.Pattern[str]] | None = None


def cargar_unidades(ruta: str = CATALOGO_UNIDADES_PATH) -> list[str]:
    """Unidades del catálogo editable (se crea si no existe)."""
    return cargar_lista(ruta, "contenido_neto", UNIDADES_CONTENIDO_NETO)


def _patron_unidades() -> re.Pattern[str]:
    global _cache
    if not os.path.exists(CATALOGO_UNIDADES_PATH):
        cargar_unidades()
    modificado = os.path.getmtime(CATALOGO_UNIDADES_PATH)
    if _cache is None or _cache[0] != modificado:
        unidades = sorted(cargar_unidades(), key=len, reverse=True)  # "MLS" antes que "ML"
        # La unidad puede venir pegada al número ("500ml"); [^\W\d_] = cualquier
        # letra, así la "L" de "LATA" no cuenta como litro. Sin unidades, nada
        # lleva "CONTENIDO NETO".
        opciones = "|".join(re.escape(u) for u in unidades) or r"(?!)"
        _cache = (modificado, re.compile(rf"(?<![^\W\d_])({opciones})(?![^\W\d_])", re.IGNORECASE))
    return _cache[1]


def formatear_contenido(texto: str) -> str:
    """Leyenda del CONTENIDO. Si la celda ya trae la leyenda se quita y se pone
    la correcta, sin repetirla."""
    valor = _LEYENDA_CONTENIDO.sub("", texto.strip()).strip()
    prefijo = "CONTENIDO NETO " if _patron_unidades().search(valor) else "CONTENIDO "
    return prefijo + valor
