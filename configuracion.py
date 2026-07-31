# -- CONFIGURACIÓN DE NORMAS PARA ETIQUETAS -- #
import json
import os
import re

CONFIG_PATH = os.path.join("data", "config_etiquetas.json")

# Misma expresión que usa armadoEtiqueta.construir_mapa_numero_a_norma para
# identificar la norma de cada fila del Excel: exige "NOM-" seguido de un
# número. Si una norma nueva no cumple este patrón, el generador nunca la
# va a poder emparejar con la columna "CODIGO FORMATO".
_PATRON_NUMERO_NORMA = re.compile(r"NOM-(\d+)", re.IGNORECASE)

ORIENTACION_DEFECTO = "vertical"
ORIENTACIONES_VALIDAS = ("vertical", "horizontal")


def cargar_config(config_path=CONFIG_PATH):
    if not os.path.exists(config_path):
        return {}
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def guardar_config(config, config_path=CONFIG_PATH):
    os.makedirs(os.path.dirname(config_path), exist_ok=True)
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)


def listar_normas(config):
    return sorted(config.keys())


def obtener_campos(config, norma):
    return list(config.get(norma, {}).get("campos", []))


def obtener_orientacion(config, norma):
    return config.get(norma, {}).get("orientacion", ORIENTACION_DEFECTO)


def validar_nombre_norma(nombre, config=None, excluir=None):
    """Valida que el nombre de una norma sea utilizable.

    `excluir` es el nombre original cuando se está renombrando (para no
    chocar consigo misma al validar)."""
    nombre = (nombre or "").strip()
    if not nombre:
        return "El nombre de la norma no puede estar vacío."

    match = _PATRON_NUMERO_NORMA.search(nombre)
    if not match:
        return "El nombre debe incluir 'NOM-' seguido de un número, ej. NOM-004-SE-2021."

    if config is not None:
        if nombre in config and nombre != excluir:
            return f"Ya existe una norma llamada '{nombre}'."

        numero = int(match.group(1))
        for otra in config:
            if otra == excluir:
                continue
            otro_match = _PATRON_NUMERO_NORMA.search(otra)
            if otro_match and int(otro_match.group(1)) == numero:
                return (
                    f"El número {numero} ya lo usa la norma '{otra}'. El generador "
                    "identifica la norma por ese número, así que no puede repetirse."
                )

    return None


def agregar_norma(config, nombre, campos=None, orientacion=ORIENTACION_DEFECTO):
    error = validar_nombre_norma(nombre, config)
    if error:
        raise ValueError(error)
    if orientacion not in ORIENTACIONES_VALIDAS:
        orientacion = ORIENTACION_DEFECTO
    config[nombre] = {"campos": list(campos or []), "orientacion": orientacion}
    return config


def eliminar_norma(config, nombre):
    config.pop(nombre, None)
    return config


def actualizar_campos_norma(config, nombre, campos):
    if nombre not in config:
        raise KeyError(f"La norma '{nombre}' no existe.")
    config[nombre]["campos"] = list(campos)
    return config


def actualizar_orientacion_norma(config, nombre, orientacion):
    if nombre not in config:
        raise KeyError(f"La norma '{nombre}' no existe.")
    if orientacion not in ORIENTACIONES_VALIDAS:
        raise ValueError(f"Orientación inválida: '{orientacion}'.")
    config[nombre]["orientacion"] = orientacion
    return config


def agregar_campo(config, nombre, campo):
    campo = (campo or "").strip().upper()
    if not campo:
        raise ValueError("El nombre del campo no puede estar vacío.")
    campos = config.setdefault(nombre, {"campos": []}).setdefault("campos", [])
    if campo in campos:
        raise ValueError(f"El campo '{campo}' ya existe en esta norma.")
    campos.append(campo)
    return config


def eliminar_campo(config, nombre, campo):
    campos = config.get(nombre, {}).get("campos", [])
    if campo in campos:
        campos.remove(campo)
    return config
