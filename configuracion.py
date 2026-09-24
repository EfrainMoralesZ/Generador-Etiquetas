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
    # Se escribe primero a un temporal y luego se reemplaza el original, para
    # que un fallo a medio guardar (o OneDrive bloqueando el archivo) no deje
    # el JSON truncado ahora que se guarda en cada cambio de campo.
    os.makedirs(os.path.dirname(config_path), exist_ok=True)
    temporal = config_path + ".tmp"
    with open(temporal, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)
    os.replace(temporal, config_path)


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


def mover_campo(config, nombre, campo, desplazamiento):
    """Mueve un campo `desplazamiento` posiciones (-1 sube, +1 baja). El orden
    de la lista es el orden en que se imprimen los campos en la etiqueta."""
    campos = config.get(nombre, {}).get("campos", [])
    if campo not in campos:
        return config
    origen = campos.index(campo)
    destino = max(0, min(len(campos) - 1, origen + desplazamiento))
    campos.insert(destino, campos.pop(origen))
    return config


def mover_campo_y_guardar(nombre, campo, desplazamiento, config_path=CONFIG_PATH):
    """Reordena un campo de una norma existente y lo escribe en el JSON.
    Devuelve el config actualizado."""
    config = cargar_config(config_path)
    if nombre not in config:
        raise KeyError(f"La norma '{nombre}' no existe.")
    mover_campo(config, nombre, campo, desplazamiento)
    guardar_config(config, config_path)
    return config


def agregar_campo_y_guardar(nombre, campo, config_path=CONFIG_PATH):
    """Agrega un campo a una norma existente y lo escribe en el JSON.

    Relee el archivo antes de modificarlo para no pisar cambios hechos por
    fuera de la app. Devuelve el config actualizado."""
    config = cargar_config(config_path)
    if nombre not in config:
        raise KeyError(f"La norma '{nombre}' no existe.")
    agregar_campo(config, nombre, campo)
    guardar_config(config, config_path)
    return config


def eliminar_campo_y_guardar(nombre, campo, config_path=CONFIG_PATH):
    """Quita un campo de una norma existente y lo escribe en el JSON.

    Una norma debe conservar al menos un campo. Devuelve el config
    actualizado."""
    config = cargar_config(config_path)
    if nombre not in config:
        raise KeyError(f"La norma '{nombre}' no existe.")
    campos = obtener_campos(config, nombre)
    if campo in campos and len(campos) <= 1:
        raise ValueError("Una norma debe tener al menos un campo.")
    eliminar_campo(config, nombre, campo)
    guardar_config(config, config_path)
    return config
