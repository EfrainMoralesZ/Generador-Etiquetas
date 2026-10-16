"""Reglas NOM-141-SSA1/SCFI-2012: leyendas precautorias según el tipo de
producto (detectado en la DESCRIPCION) y declaración del CONTENIDO."""
from __future__ import annotations

import os
import re

from catalogos import cargar_lista
from nom004 import normalizar

NUMERO_NORMA = 141

# Leyendas editables por tipo de producto (una lista por tipo en el mismo JSON).
CATALOGO_LEYENDAS_PATH = os.path.join("data", "leyendas_nom141.json")

PRIMEROS_AUXILIOS = (
    "Primeros auxilios: en caso de contacto con los ojos, enjuáguese inmediatamente con abundante agua; "
    "en caso de ingestión, no provoque el vómito y consulte a un médico."
)

# Tipos de producto, en el orden en que se imprimen sus leyendas si una
# descripción cae en varios. "palabras" se buscan en la DESCRIPCION sin
# importar mayúsculas ni acentos, como palabra completa (así "colorante" no
# se confunde con "decolorante").
TIPOS = {
    "solares": {
        "titulo": "Protectores solares",
        "palabras": ["PROTECTOR SOLAR", "PROTECTORES SOLARES", "BLOQUEADOR SOLAR", "BLOQUEADORES SOLARES",
                     "FILTRO SOLAR", "BRONCEADOR", "BRONCEADORES", "AUTOBRONCEADOR", "AUTOBRONCEADORES"],
        "leyendas": [
            "Aplíquese antes de la exposición al sol.",
            "Para mantener la protección, repita con frecuencia la aplicación del producto, especialmente "
            "después de transpirar, bañarse o secarse.",
            "Aplique a la piel la cantidad suficiente.",
        ],
    },
    "desodorantes": {
        "titulo": "Desodorantes",
        "palabras": ["DESODORANTE", "DESODORANTES", "ANTITRANSPIRANTE", "ANTITRANSPIRANTES"],
        "leyendas": [
            "No se aplique sobre piel irritada o lastimada.",
            "Descontinúe su uso en caso de presentarse irritación, enrojecimiento o alguna molestia.",
            "No se deje al alcance de los niños.",
        ],
    },
    "tintes": {
        "titulo": "Tintes",
        "palabras": ["TINTE", "TINTES", "COLORANTE", "COLORANTES", "COLORACION", "COLORACIONES"],
        "leyendas": [
            "Los colorantes del cabello pueden causar reacciones alérgicas graves.",
            "Lea y siga las instrucciones.",
            "Este producto no está destinado a utilizarse en personas menores de dieciséis años.",
            "Los tatuajes temporales de «henna negra» pueden aumentar el riesgo de alergia.",
            "No utilice el tinte capilar:",
            "Si tiene una erupción cutánea en el rostro o tiene el cuero cabelludo sensible, irritado o dañado.",
            "Si alguna vez ha experimentado cualquier tipo de reacción después de la coloración del cabello.",
            "Si alguna vez ha experimentado una reacción a los tatuajes temporales de «henna negra».",
            "Realice una prueba preliminar de acuerdo a las instrucciones.",
            "Puede causar alergia en algunas personas.",
            "Suspenda su empleo en caso de irritación.",
            "No se aplique en cejas o pestañas.",
            "Evite el contacto con los ojos.",
            "No se deje al alcance de los niños.",
            PRIMEROS_AUXILIOS,
            "Use guantes apropiados.",
            "En caso de que el producto entre en contacto con los ojos, enjuáguelos inmediatamente con agua.",
        ],
    },
    "permanentes": {
        "titulo": "Permanentes",
        "palabras": ["PERMANENTE", "PERMANENTES", "ALACIADOR PERMANENTE", "ALACIADORES PERMANENTES",
                     "ONDULADO PERMANENTE"],
        "leyendas": [
            "No se aplique en cejas o pestañas.",
            "Evite el contacto con los ojos. Puede causar ceguera.",
            "Úsese exclusivamente conforme al instructivo.",
            "No se deje al alcance de los niños.",
            PRIMEROS_AUXILIOS,
        ],
    },
    "decolorantes": {
        "titulo": "Decolorantes",
        "palabras": ["DECOLORANTE", "DECOLORANTES"],
        "leyendas": [
            "Evite el contacto con los ojos.",
            "En caso de que el producto entre en contacto con los ojos, enjuáguelos inmediatamente con agua.",
            "Use guantes apropiados.",
            "No se aplique si la piel cabelluda está irritada.",
            "Suspenda su empleo en caso de irritación.",
            "No se aplique en cejas o pestañas.",
            "No se deje al alcance de los niños.",
            PRIMEROS_AUXILIOS,
        ],
    },
    "depilatorios": {
        "titulo": "Depilatorios",
        "palabras": ["DEPILATORIO", "DEPILATORIOS", "EPILATORIO", "EPILATORIOS"],
        "leyendas": [
            "No se aplique sobre piel irritada o lastimada.",
            "No se deje al alcance de los niños.",
            "Evite el contacto con los ojos.",
        ],
    },
}

# Sustancias que, si vienen en los INSUMOS/INGREDIENTES de un tipo, se
# destacan con "Contiene ..." (palabra a buscar, texto a imprimir).
SUSTANCIAS = {
    "desodorantes": [("FENOLSULFONATO DE ZINC", "Contiene fenolsulfonato de zinc. Evite el contacto con los ojos.")],
    "tintes": [
        ("DIAMINOBENCENO", "Contiene diaminobenceno."),
        ("PHENYLENEDIAMINE", "Contiene diaminobenceno."),
        ("FENILENDIAMINA", "Contiene diaminobenceno."),
        ("DIAMINOTOLUENO", "Contiene diaminotolueno."),
        ("TOLUENE-2,5-DIAMINE", "Contiene diaminotolueno."),
        ("DIAMINOFENOL", "Contiene diaminofenol."),
        ("DIAMINOPHENOL", "Contiene diaminofenol."),
        ("HIDROQUINONA", "Contiene hidroquinona."),
        ("HYDROQUINONE", "Contiene hidroquinona."),
        ("RESORCINOL", "Contiene resorcinol."),
    ],
    "permanentes": [
        ("TIOGLICOLATO", "Contiene tioglicolato, puede causar daño al cabello y a la piel cabelluda."),
        ("THIOGLYCOLATE", "Contiene tioglicolato, puede causar daño al cabello y a la piel cabelluda."),
        ("HIDROXIDO DE SODIO", "Contiene un agente alcalino."),
        ("SODIUM HYDROXIDE", "Contiene un agente alcalino."),
        ("HIDROXIDO DE POTASIO", "Contiene un agente alcalino."),
        ("POTASSIUM HYDROXIDE", "Contiene un agente alcalino."),
        ("HIDROXIDO DE LITIO", "Contiene un agente alcalino."),
        ("LITHIUM HYDROXIDE", "Contiene un agente alcalino."),
        ("HIDROXIDO DE CALCIO", "Contiene un agente alcalino."),
        ("CALCIUM HYDROXIDE", "Contiene un agente alcalino."),
    ],
    "decolorantes": [
        ("PEROXIDO DE HIDROGENO", "Contiene peróxido de hidrógeno, puede causar daño."),
        ("HYDROGEN PEROXIDE", "Contiene peróxido de hidrógeno, puede causar daño."),
        ("PERSULFATO", "Contiene persulfatos, puede causar daño."),
        ("PERSULFATE", "Contiene persulfatos, puede causar daño."),
    ],
}

# En tintes, permanentes y decolorantes de uso profesional.
TIPOS_PROFESIONALES = ("tintes", "permanentes", "decolorantes")
PALABRAS_PROFESIONAL = ("PROFESIONAL", "PROFESIONALES", "PROFESSIONAL")
LEYENDA_PROFESIONAL = "Reservado a profesionales."

# "CONTENIDO", "CONTENIDO NETO", "CONT", "CONT. NET." no son necesarias: solo
# va el dato cuantitativo y su unidad.
_LEYENDA_CONTENIDO = re.compile(r"^\s*CONT(ENIDO)?\.?(\s*NET(O)?\.?)?\s*:?\s*", re.IGNORECASE)


def _patron(palabras) -> re.Pattern[str]:
    opciones = "|".join(re.escape(normalizar(p)) for p in sorted(palabras, key=len, reverse=True))
    return re.compile(rf"(?<![a-z0-9])(?:{opciones})(?![a-z0-9])")


_PATRONES_TIPO = {tipo: _patron(datos["palabras"]) for tipo, datos in TIPOS.items()}
_PATRON_PROFESIONAL = _patron(PALABRAS_PROFESIONAL)
# "Tinte permanente" es un tinte, no un permanente; solo cuenta como
# permanente si además dice alaciador u ondulado.
_PATRON_PERMANENTE_REAL = _patron(["ALACIADOR", "ALACIADORES", "ALACIADO", "ONDULADO"])


def cargar_leyendas(tipo: str) -> list[str]:
    """Leyendas del catálogo editable de un tipo (se crea si no existe)."""
    return cargar_lista(CATALOGO_LEYENDAS_PATH, tipo, TIPOS[tipo]["leyendas"])


def _es_el_producto(texto: str, inicio: int) -> bool:
    """La palabra nombra al producto solo si abre la descripción ("DESODORANTE",
    "PROTECTOR SOLAR 50 FPS") o va tras "de" ("DÚO DE DESODORANTES"); como
    adjetivo no cuenta ("TOALLITAS DESODORANTES", "MAQUILLAJE CON EFECTO BRONCEADOR")."""
    antes = texto[:inicio].strip()
    return not antes or bool(re.search(r"(?:^|\s)del?$", antes))


def detectar_tipos(descripcion: str) -> list[str]:
    """Tipos de producto que menciona la descripción."""
    texto = normalizar(descripcion or "")
    tipos = [
        tipo for tipo, patron in _PATRONES_TIPO.items()
        if any(_es_el_producto(texto, m.start()) for m in patron.finditer(texto))
    ]
    if "tintes" in tipos and "permanentes" in tipos and not _PATRON_PERMANENTE_REAL.search(texto):
        tipos.remove("permanentes")
    return tipos


def generar_leyendas(descripcion: str, ingredientes: str = "") -> list[str]:
    """Leyendas que exige la norma según la descripción (y las sustancias que
    trae el producto), sin repetir las que comparten varios tipos."""
    tipos = detectar_tipos(descripcion)
    ingredientes_norm = normalizar(ingredientes or "")
    leyendas: list[str] = []
    for tipo in tipos:
        leyendas += cargar_leyendas(tipo)
        leyendas += [texto for palabra, texto in SUSTANCIAS.get(tipo, []) if _patron([palabra]).search(ingredientes_norm)]
    if any(t in TIPOS_PROFESIONALES for t in tipos) and _PATRON_PROFESIONAL.search(normalizar(descripcion or "")):
        leyendas.append(LEYENDA_PROFESIONAL)
    return _sin_repetidas(leyendas)


def combinar_leyendas(generadas: list[str], del_excel: str) -> str:
    """Leyendas de la norma seguidas de las que ya traía el Excel que no estén
    repetidas. Cada leyenda va en su propio renglón."""
    extra = [l.strip() for l in re.split(r"[\r\n]+", del_excel or "") if l.strip()]
    return "\n".join(_sin_repetidas(generadas + extra))


def _sin_repetidas(leyendas: list[str]) -> list[str]:
    vistas, resultado = set(), []
    for leyenda in leyendas:
        llave = normalizar(leyenda).strip(" .:-")
        if llave and llave not in vistas:
            vistas.add(llave)
            resultado.append(leyenda)
    return resultado


def formatear_contenido(texto: str) -> str:
    """Solo el dato cuantitativo y su unidad (ej. "CONT. NET. 250 ml" -> "250 ml")."""
    return _LEYENDA_CONTENIDO.sub("", texto.strip()).strip()


def _comparable(texto: str) -> str:
    """Sin mayúsculas, acentos, puntuación ni saltos de línea, para buscar una
    leyenda dentro de la celda aunque venga en otro renglón o sin punto final."""
    return " ".join(re.sub(r"[^a-z0-9]+", " ", normalizar(texto or "")).split())


def revisar(descripcion: str, leyendas: str, ingredientes: str = "") -> list[dict]:
    """Si la DESCRIPCION es de un tipo que exige leyendas precautorias, la
    celda LEYENDAS del Excel debe traer todas las de ese tipo (catálogo, más
    "Contiene ..." y "Reservado a profesionales" cuando apliquen):
    [{columna, regla, ok, mensaje}]."""
    tipos = detectar_tipos(descripcion)
    regla = {"columna": "LEYENDAS", "regla": "Leyendas precautorias"}
    if not tipos:
        return [{**regla, "ok": True, "mensaje": "El producto no es de un tipo que requiera leyendas precautorias."}]
    titulos = ", ".join(TIPOS[t]["titulo"].lower() for t in tipos)

    celda = (leyendas or "").strip()
    if not celda or celda.upper() in ("NAN", "N/A", "NONE"):
        return [{**regla, "ok": False,
                 "mensaje": f"El producto \"{descripcion}\" ({titulos}) requiere leyendas precautorias y la "
                            "celda LEYENDAS está vacía."}]

    requeridas = generar_leyendas(descripcion, ingredientes)
    celda_comparable = f" {_comparable(celda)} "
    faltan = [l for l in requeridas if f" {_comparable(l)} " not in celda_comparable]
    if not faltan:
        return [{**regla, "ok": True,
                 "mensaje": f"Trae las {len(requeridas)} leyendas precautorias de {titulos}."}]
    return [{**regla, "ok": False,
             "mensaje": f"El producto \"{descripcion}\" ({titulos}) requiere leyendas precautorias y en la celda "
                        f"LEYENDAS faltan {len(faltan)} de {len(requeridas)}:\n"
                        + "\n".join(f"   – {l}" for l in faltan)}]
