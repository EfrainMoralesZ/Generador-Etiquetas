# -- CATÁLOGOS EDITABLES DE LAS NORMAS -- #
"""Listas editables (fibras, unidades, frases...) guardadas en JSON como
{"clave": [valores]}. Un mismo archivo puede tener varias listas (ej. las
fibras naturales y químicas de la NOM-004). Sin dependencias de Tkinter."""
import json
import os


def _leer(ruta):
    if not os.path.exists(ruta):
        return {}
    with open(ruta, "r", encoding="utf-8") as f:
        return json.load(f)


def _escribir(ruta, datos):
    # Temporal + reemplazo, igual que configuracion.guardar_config, para no
    # dejar el JSON truncado si algo falla a medio guardar.
    os.makedirs(os.path.dirname(ruta) or ".", exist_ok=True)
    temporal = ruta + ".tmp"
    with open(temporal, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=4)
    os.replace(temporal, ruta)


def cargar_lista(ruta, clave, inicial):
    """Valores de la lista `clave`. Si el archivo o la lista no existen, se
    crean con los valores de `inicial`."""
    datos = _leer(ruta)
    if clave not in datos:
        datos[clave] = list(inicial)
        _escribir(ruta, datos)
    return [str(v).strip() for v in datos[clave] if str(v).strip()]


def guardar_lista(ruta, clave, valores):
    """Reemplaza la lista `clave` conservando las demás listas del archivo."""
    datos = _leer(ruta)
    datos[clave] = [str(v).strip() for v in valores if str(v).strip()]
    _escribir(ruta, datos)
    return datos[clave]


def _repetido(valores, valor, excluir=None):
    llave = valor.strip().lower()
    return next((v for v in valores if v.strip().lower() == llave and v != excluir), None)


def agregar_valor(ruta, clave, inicial, valor):
    valor = (valor or "").strip()
    if not valor:
        raise ValueError("El valor no puede estar vacío.")
    valores = cargar_lista(ruta, clave, inicial)
    repetido = _repetido(valores, valor)
    if repetido:
        raise ValueError(f"'{repetido}' ya está en el catálogo.")
    valores.append(valor)
    return guardar_lista(ruta, clave, valores)


def renombrar_valor(ruta, clave, inicial, anterior, nuevo):
    nuevo = (nuevo or "").strip()
    if not nuevo:
        raise ValueError("El valor no puede estar vacío.")
    valores = cargar_lista(ruta, clave, inicial)
    if anterior not in valores:
        raise KeyError(f"'{anterior}' ya no existe en el catálogo.")
    repetido = _repetido(valores, nuevo, excluir=anterior)
    if repetido:
        raise ValueError(f"'{repetido}' ya está en el catálogo.")
    valores[valores.index(anterior)] = nuevo
    return guardar_lista(ruta, clave, valores)


def eliminar_valor(ruta, clave, inicial, valor):
    valores = cargar_lista(ruta, clave, inicial)
    if valor in valores:
        valores.remove(valor)
    return guardar_lista(ruta, clave, valores)
