"""NOM-189-SSA1/SCFI-2018 (productos de aseo de uso doméstico): catálogos de
leyendas precautorias por riesgo (inciso 6.1.6), solo de referencia.

No hay un catálogo de qué productos llevan cada leyenda (depende de la
fórmula), así que el sistema no las valida ni las agrega: el ejecutivo técnico
evalúa cada fila en la ventana de análisis y marca "Libre de leyendas
precautorias" cuando aplica; esas filas se imprimen sin LEYENDAS."""
from __future__ import annotations

import os

from catalogos import cargar_lista

NUMERO_NORMA = 189

# Leyendas editables por tipo de riesgo (una lista por tipo en el mismo JSON).
CATALOGO_LEYENDAS_PATH = os.path.join("data", "leyendas_nom189.json")

NO_ALCANCE_NINOS = "No se deje al alcance de los niños."
GUANTES = "Utilice guantes al emplear este producto."

# "criterio" resume cuándo aplica cada tipo según la norma (se muestra en
# Configuración junto al catálogo).
TIPOS = {
    "toxicos": {
        "titulo": "Productos tóxicos",
        "criterio": "Pueden causar daños a la salud o la muerte según su concentración, vía y tiempo de "
                    "exposición (3.29, 6.1.6.1). Llevan además un pictograma de toxicidad.",
        "leyendas": [
            "Contiene sustancias tóxicas cuya ingestión, inhalación o contacto directo inmediato o prolongado "
            "pueden originar graves daños a la salud.",
            NO_ALCANCE_NINOS,
        ],
    },
    "corrosivos": {
        "titulo": "Productos corrosivos",
        "criterio": "Destruyen o alteran de forma irreversible el tejido vivo en el sitio de contacto (3.18, "
                    "6.1.6.2). \"Lea la etiqueta antes de usar este producto\" va en la superficie principal "
                    "de exhibición. Llevan además un pictograma de corrosivo.",
        "leyendas": [
            "Lea la etiqueta antes de usar este producto.",
            "CORROSIVO",
            "No reutilice el envase; una vez terminado el contenido, deséchelo.",
            GUANTES,
            NO_ALCANCE_NINOS,
            "En caso de ingestión accidental, no provoque el vómito y obtenga ayuda médica de inmediato.",
            "En caso de contacto con la piel o con los ojos, enjuague con abundante agua y obtenga ayuda médica.",
        ],
    },
    "inflamables": {
        "titulo": "Productos inflamables",
        "criterio": "Punto de inflamación menor o igual a 37.8 °C (3.22, 6.1.6.3). Llevan también las "
                    "leyendas de tóxicos o corrosivos que les apliquen por su fórmula.",
        "leyendas": [
            "INFLAMABLE",
            "Mantener lejos del fuego y de instalaciones eléctricas.",
        ],
    },
    "presurizados": {
        "titulo": "Envases presurizados",
        "criterio": "Envase a presión con válvula y gas propulsor, como los aerosoles (3.6, 6.1.6.4).",
        "leyendas": [
            "No se use cerca de los ojos o flama.",
            "No se queme ni perfore el envase.",
            "No se exponga al calor.",
            NO_ALCANCE_NINOS,
        ],
    },
    "no_mezclar": {
        "titulo": "No mezclar con otras sustancias",
        "criterio": "La fórmula representa un riesgo a la salud si se mezcla con otras sustancias (6.1.6.5). "
                    "Conviene precisar con qué no se mezcla (ej. \"No se mezcle con cloro\").",
        "leyendas": [
            "No se mezcle con otros productos.",
        ],
    },
    "irritantes": {
        "titulo": "Irritantes o sensibilizantes",
        "criterio": "Contienen enzimas, oxidantes u otras sustancias que por su concentración irritan o "
                    "sensibilizan la piel o las mucosas en uso normal (6.1.6.6).",
        "leyendas": [
            "Puede provocar irritación en la piel y en mucosas.",
            GUANTES,
        ],
    },
}


def cargar_leyendas(tipo: str) -> list[str]:
    """Leyendas del catálogo editable de un tipo (se crea si no existe)."""
    return cargar_lista(CATALOGO_LEYENDAS_PATH, tipo, TIPOS[tipo]["leyendas"])

