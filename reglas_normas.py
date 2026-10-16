# -- REGLAS Y CATÁLOGOS DE CADA NORMA (para la pantalla Configuración) -- #
"""Describe, por número de norma (igual que el generador las identifica), qué
reglas aplica el sistema y qué catálogos editables usan esas reglas.

Las reglas viven en código (nom004.py, nom050.py, nom141.py, armadoEtiqueta.formatear_valor);
aquí solo se describen para mostrarlas. Los catálogos sí se editan desde la app."""
import re

import nom004
import nom050
import nom141

_PATRON_NUMERO_NORMA = re.compile(r"NOM-(\d+)", re.IGNORECASE)

def _es_plural(palabra, palabras):
    """Si `palabra` es el plural de otra de la lista (palabra por palabra)."""
    partes = palabra.split()
    return any(
        len(otra.split()) == len(partes) and otra != palabra
        and all(p in (o, o + "S", o + "ES") for p, o in zip(partes, otra.split()))
        for otra in palabras
    )


CATALOGOS = {
    nom004.NUMERO_NORMA: [
        {
            "id": "fibras_naturales",
            "titulo": "Fibras naturales",
            "descripcion": "Fibras válidas en el insumo y el forro. Se comparan sin importar mayúsculas, "
                           "pero los acentos deben coincidir.",
            "ruta": nom004.CATALOGO_FIBRAS_PATH, "clave": "naturales", "inicial": nom004.FIBRAS_NATURALES,
        },
        {
            "id": "fibras_quimicas",
            "titulo": "Fibras químicas",
            "descripcion": "Fibras válidas en el insumo y el forro. Se comparan sin importar mayúsculas, "
                           "pero los acentos deben coincidir.",
            "ruta": nom004.CATALOGO_FIBRAS_PATH, "clave": "quimicas", "inicial": nom004.FIBRAS_QUIMICAS,
        },
        {
            "id": "cuidado_adicionales",
            "titulo": "Frases adicionales de cuidado",
            "descripcion": "Frases que pueden ir en cualquier parte de las instrucciones de cuidado sin "
                           "afectar el orden (ej. \"Lavar por separado\").",
            "ruta": nom004.CATALOGO_CUIDADO_PATH, "clave": "adicionales", "inicial": nom004.ADICIONALES,
        },
    ],
    nom050.NUMERO_NORMA: [
        {
            "id": "unidades_contenido_neto",
            "titulo": "Unidades de contenido neto",
            "descripcion": "Si el CONTENIDO trae una de estas unidades (pegada o no al número, sin importar "
                           "mayúsculas) se imprime \"CONTENIDO NETO\"; si no, solo \"CONTENIDO\".",
            "ruta": nom050.CATALOGO_UNIDADES_PATH, "clave": "contenido_neto",
            "inicial": nom050.UNIDADES_CONTENIDO_NETO,
        },
    ],
    nom141.NUMERO_NORMA: [
        {
            "id": f"leyendas_{tipo}",
            "titulo": datos["titulo"],
            "descripcion": "Leyendas que se imprimen cuando la DESCRIPCION empieza con (o dice \"de ...\"): "
                           + ", ".join(p.lower() for p in datos["palabras"] if not _es_plural(p, datos["palabras"]))
                           + " (y sus plurales).",
            "ruta": nom141.CATALOGO_LEYENDAS_PATH, "clave": tipo, "inicial": datos["leyendas"],
        }
        for tipo, datos in nom141.TIPOS.items()
    ],
}

_NO_REVISADAS = ", ".join(v for v in sorted(nom004.NO_APLICA) if v not in ("", "NAN", "NONE"))

VALIDACIONES = {
    nom004.NUMERO_NORMA: [
        ("Insumo y forro: formato \"NN% fibra\"",
         "Cada parte va como \"NN% fibra\" (o \"fibra NN%\"), separadas por coma, punto y coma o salto "
         "de línea. El porcentaje debe ser mayor a 0."),
        ("Fibras del catálogo",
         "Cada fibra del insumo y del forro debe estar en el catálogo de fibras naturales o químicas y "
         "escribirse igual (con acentos)."),
        ("Suma 100%", "Los porcentajes del insumo y los del forro deben sumar 100% cada uno."),
        ("Orden de predominancia", "Las fibras van de mayor a menor porcentaje."),
        ("Instrucciones de cuidado completas",
         "Si la celda de cuidado trae texto, debe incluir tipo de lavado, temperatura (cuando el lavado "
         "es a mano o a máquina), blanqueo, secado y planchado. Si viene vacía, la etiqueta se imprime "
         "sin cuidado."),
        ("Orden de las instrucciones",
         "Las partes del cuidado deben ir en este orden: "
         + ", ".join(nombre for nombre, _ in nom004.PARTES_CUIDADO.values())
         + ". Las frases adicionales del catálogo pueden ir en cualquier parte."),
        ("Celdas que no se revisan", f"Insumo, forro o cuidado vacíos o con {_NO_REVISADAS}."),
    ],
    nom141.NUMERO_NORMA: [
        ("Leyendas precautorias",
         "Si la DESCRIPCION es de un tipo con leyendas (tintes, decolorantes, permanentes, protectores "
         "solares, desodorantes, depilatorios...), la celda LEYENDAS del Excel debe traer todas las "
         "leyendas del catálogo de ese tipo (y \"Contiene ...\" o \"Reservado a profesionales\" cuando "
         "apliquen). Se comparan sin importar mayúsculas, acentos, puntuación ni renglones. "
         "El tipo cuenta solo si es el producto: al inicio de la descripción (\"DESODORANTE\") o tras "
         "\"de\" (\"DÚO DE DESODORANTES\"); como adjetivo no (\"TOALLITAS DESODORANTES\", "
         "\"MAQUILLAJE CON EFECTO BRONCEADOR\")."),
    ],
}

# Formato que armadoEtiqueta.formatear_valor aplica a ciertos campos en todas
# las normas (por nombre de campo).
_FORMATO_GENERAL = [
    (("PAIS ORIGEN", "PAIS DE ORIGEN", "PAIS"), "Se antepone \"HECHO EN\" y va en mayúsculas."),
    (("FORRO",), "Se antepone \"FORRO\"."),
    (("TALLA",), "Se antepone \"TALLA\"."),
    (("INSUMOS/INGREDIENTES", "INGREDIENTES"), "Se antepone \"Ingredientes:\" si la celda no lo trae ya."),
    (("CONTENIDO",), "Se imprime solo el valor, en negrita y más grande."),
    (("MEDIDAS",),
     "Con dos columnas MEDIDAS en el Excel, la primera va en la etiqueta y la segunda es la altura del "
     "contenido (fuera del recuadro); con una sola, es solo la altura."),
    (("TIPO", "TIPO DE ETIQUETA"), "No va dentro del recuadro: se imprime bajo el título (Costura / Adherible)."),
]

# Formatos propios de una norma; reemplazan al general del mismo campo.
_FORMATO_NORMA = {
    nom141.NUMERO_NORMA: [
        (("LEYENDAS",),
         "Según la DESCRIPCION (tintes, colorantes, coloración, decolorantes, permanentes, alaciadores "
         "permanentes, protectores solares, bronceadores, autobronceadores, desodorantes, antitranspirantes, "
         "depilatorios, epilatorios) se agregan las leyendas precautorias del catálogo de ese tipo (la celda "
         "LEYENDAS del Excel debe traerlas; si viene vacía la fila queda con observación). Si los ingredientes traen sustancias a destacar (ej. resorcinol, "
         "hidroquinona, hidróxido de sodio) se agrega \"Contiene ...\"; si el producto es profesional, "
         "\"Reservado a profesionales\". Las leyendas que ya traía el Excel se conservan al final sin repetir."),
        (("CONTENIDO", "CONTENIDO NETO"),
         "Solo el dato y su unidad (ej. 250 ml): se quita \"CONTENIDO\", \"CONTENIDO NETO\", \"CONT\" o "
         "\"CONT. NET.\" si la celda los trae. Va en negrita y más grande."),
    ],
    nom050.NUMERO_NORMA: [
        (("CONTENIDO",),
         "Si trae una unidad del catálogo de unidades de contenido neto (ej. 500 ml, 1 kg) se antepone "
         "\"CONTENIDO NETO\"; si no (piezas, latas...) se antepone \"CONTENIDO\". Si la celda ya trae "
         "la leyenda, se reemplaza por la correcta sin repetirla. Va en negrita y más grande."),
    ],
}


def numero_norma(nombre):
    match = _PATRON_NUMERO_NORMA.search(nombre or "")
    return int(match.group(1)) if match else None


def catalogos_de(nombre):
    return CATALOGOS.get(numero_norma(nombre), [])


def validaciones_de(nombre):
    return VALIDACIONES.get(numero_norma(nombre), [])


def formatos_de(nombre, campos):
    """[(campo, descripción)] del formato que se aplica a los campos de la norma."""
    propios = _FORMATO_NORMA.get(numero_norma(nombre), [])
    resultado = []
    for campo in campos:
        campo_norm = campo.strip().upper()
        for nombres, descripcion in propios + _FORMATO_GENERAL:
            if campo_norm in nombres:
                resultado.append((campo, descripcion))
                break
    return resultado


def cuenta_reglas(nombre):
    """Reglas propias de la norma (validaciones + formatos exclusivos)."""
    numero = numero_norma(nombre)
    return len(VALIDACIONES.get(numero, [])) + len(_FORMATO_NORMA.get(numero, []))
