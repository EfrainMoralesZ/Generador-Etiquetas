# ARMADO DE ETIQUETAS SEGUN NORMA (NOM) A PARTIR DE UN EXCEL #
import json
import os
import re
import textwrap
from functools import lru_cache

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader

DEFAULT_CONFIG_PATH = os.path.join("data", "config_etiquetas.json")
DEFAULT_JSON_DIR = os.path.join("data", "etiquetas")

COLUMNA_NORMA = "CODIGO FORMATO"
CAMPOS_ENCABEZADO = ("EAN", "MARCA")
CAMPOS_PIE = ("IMPORTADOR",)

DPI = 300
MARGIN_X = 45
MARGIN_Y_TOP = 40
MARGIN_Y_BOTTOM = 40
HEADER_GAP = 25
FOOTER_GAP = 25
FIELD_SPACER = 18
MAX_CHARS_HEADER = 32
MAX_CHARS_BODY = 38
MIN_WIDTH_PX = 700
MIN_HEIGHT_PX = 380
FONT_SIZE_HEADER = 30
FONT_SIZE_BODY = 24

RUTAS_FUENTE = [
    "arialbd.ttf",
    "Arial Bold.ttf",
    "C:/Windows/Fonts/arialbd.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
]

@lru_cache(maxsize=None)
def _cargar_fuente(tamano):
    for ruta in RUTAS_FUENTE:
        try:
            return ImageFont.truetype(ruta, tamano)
        except Exception:
            continue
    return ImageFont.load_default()

_CARACTERES_INVALIDOS = re.compile(r'[<>:"/\\|?*]')

# Unidades de peso/volumen que indican "contenido neto" (gramos, mililitros,
# litros, kilos...) en vez de una cantidad de piezas.
_UNIDADES_CONTENIDO_NETO = re.compile(r"\b(ML|MLS?|LTS?|L|KGS?|GRS?|G)\b", re.IGNORECASE)

def _nombre_archivo_seguro(texto):
    texto = _CARACTERES_INVALIDOS.sub("_", str(texto).strip())
    return texto or "SIN_DATO"

def _medir_texto(draw, texto, font):
    bbox = draw.textbbox((0, 0), texto, font=font)
    return bbox[2] - bbox[0]

def cargar_config_etiquetas(config_path=DEFAULT_CONFIG_PATH):
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)

def construir_mapa_numero_a_norma(config):
    """Mapea el número de una norma (ej. 4, 15, 50) a su clave completa del JSON."""
    mapa = {}
    for norma in config:
        match = re.search(r"NOM-(\d+)", norma.upper())
        if match:
            mapa[int(match.group(1))] = norma
    return mapa

def extraer_numero_norma(valor):
    """Extrae el primer número encontrado (ej. 'NOM004TEXX' -> 4)."""
    match = re.search(r"\d+", str(valor))
    if not match:
        return None
    return int(match.group())

def buscar_valor_columna(fila, campo):
    """Busca un valor en la fila tolerando variaciones de mayúsculas/espacios."""
    if campo in fila:
        return fila[campo]
    campo_norm = campo.strip().upper()
    for clave, valor in fila.items():
        if str(clave).strip().upper() == campo_norm:
            return valor
    return None

#Reglas para anteponer titulos ej: HECHO EN... CONTENIDO, CONTENIDO NETO, FORRO, TALLA, PAIS ORIGEN.
def formatear_valor(campo, valor):
    if valor is None:
        return None
    texto = str(valor).strip()
    if texto == "" or texto.upper() in ("NAN", "N/A", "NONE"):
        return None

    campo_norm = campo.strip().upper()
    # ANTEPONER HECHO EN ANTES DEL PAIS ORIGEN
    if campo_norm in ("PAIS DE ORIGEN", "PAIS", "PAIS ORIGEN"):
        return f"HECHO EN {texto.upper()}"
    # ANTEPONER FORRO ANTES DEL TEXTO DE FORRO
    if campo_norm == "FORRO":
        return f"FORRO {texto}"
    # ANTEPONER TALLA ANTES DEL TEXTO DE TALLA
    if campo_norm == "TALLA":
        return f"TALLA {texto}"
    # ANTEPONER CONTENIDO (piezas) O CONTENIDO NETO (peso/volumen) SEGUN LA UNIDAD
    if campo_norm == "CONTENIDO":
        texto_mayus = texto.upper()
        if texto_mayus.startswith("CONTENIDO"):
            return texto_mayus
        if _UNIDADES_CONTENIDO_NETO.search(texto_mayus):
            return f"CONTENIDO NETO {texto}"
        return f"CONTENIDO {texto}"
    return texto

def extraer_campos_etiqueta(fila, campos):
    """Devuelve la lista de (campo, texto_formateado) para los campos con valor."""
    resultado = []
    for campo in campos:
        valor = buscar_valor_columna(fila, campo)
        texto = formatear_valor(campo, valor)
        if texto:
            resultado.append((campo, texto))
    return resultado

def excel_a_json(excel_path, carpeta_salida=DEFAULT_JSON_DIR):
    """Lee el Excel y lo guarda de inmediato como .json en `carpeta_salida`
    (se necesita ahí desde la subida, no solo tras generar, para poder
    inspeccionar cómo queda extraído el texto). Si el usuario nunca genera
    las etiquetas con este archivo, la app se encarga de borrar ese .json
    (ver app._cargar_archivo / _quitar_archivo)."""
    df = pd.read_excel(excel_path, dtype=str)
    df = df.fillna("")
    registros = df.to_dict(orient="records")

    os.makedirs(carpeta_salida, exist_ok=True)
    nombre_base = os.path.splitext(os.path.basename(excel_path))[0]
    json_path = os.path.join(carpeta_salida, f"{nombre_base}.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(registros, f, ensure_ascii=False, indent=2)

    return registros, json_path

def _analizar_fila(fila, idx, mapa_numero_a_norma, config):
    """Determina la norma, los campos y el estado de una fila del Excel.

    Se usa tanto en la previsualización como en la generación real para que
    ambas coincidan exactamente en qué fila es válida y con qué norma."""
    ean = buscar_valor_columna(fila, "EAN")
    marca = buscar_valor_columna(fila, "MARCA")
    codigo_formato = buscar_valor_columna(fila, COLUMNA_NORMA)

    item = {
        "fila": idx,
        "ean": str(ean).strip() if ean else "",
        "marca": str(marca).strip() if marca else "",
        "codigo_formato": str(codigo_formato).strip() if codigo_formato else "",
        "norma": None,
        "campos_texto": [],
        "orientacion": "vertical",
        "error": None,
    }

    if not codigo_formato or not str(codigo_formato).strip():
        item["error"] = f"Falta la columna '{COLUMNA_NORMA}'"
        return item

    numero = extraer_numero_norma(codigo_formato)
    if numero is None or numero not in mapa_numero_a_norma:
        item["error"] = f"Código '{codigo_formato}' no coincide con ninguna norma"
        return item

    norma = mapa_numero_a_norma[numero]
    campos = config[norma]["campos"]
    campos_texto = extraer_campos_etiqueta(fila, campos)

    item["norma"] = norma
    item["campos_texto"] = campos_texto
    item["orientacion"] = config[norma].get("orientacion", "vertical")
    if not campos_texto:
        item["error"] = f"Sin datos para los campos de la norma {norma}"

    return item

def _filas_sin_codigo_formato(registros):
    """Números de fila (1-based) donde falta un valor en la columna
    'CODIGO FORMATO' (ya sea porque la columna no existe en el Excel o
    porque esa celda está en blanco). Esa columna es la que le dice al
    sistema qué norma y qué armado le corresponde a cada etiqueta, así que
    si falta en cualquier fila no se debe generar nada hasta corregirla."""
    return [
        idx for idx, fila in enumerate(registros, start=1)
        if not (buscar_valor_columna(fila, COLUMNA_NORMA) or "").strip()
    ]

def previsualizar_etiquetas_desde_excel(excel_path, config_path=DEFAULT_CONFIG_PATH, json_dir=DEFAULT_JSON_DIR):
    """Analiza el Excel y arma un resumen (fila, EAN, marca, norma, campos, error)
    sin generar imágenes ni PDFs, para que el usuario verifique antes de generar."""
    config = cargar_config_etiquetas(config_path)
    mapa_numero_a_norma = construir_mapa_numero_a_norma(config)

    registros, json_path = excel_a_json(excel_path, json_dir)

    detalle = []
    for idx, fila in enumerate(registros, start=1):
        item = _analizar_fila(fila, idx, mapa_numero_a_norma, config)
        detalle.append({
            "fila": item["fila"],
            "ean": item["ean"],
            "marca": item["marca"],
            "codigo_formato": item["codigo_formato"],
            "norma": item["norma"],
            "campos": [c for c, _ in item["campos_texto"]],
            "error": item["error"],
        })

    return {
        "total_filas": len(registros),
        "listas": sum(1 for d in detalle if not d["error"]),
        "detalle": detalle,
        "json_path": json_path,
        "filas_sin_codigo_formato": _filas_sin_codigo_formato(registros),
    }

def _envolver_campos(campos_texto, max_chars):
    return [textwrap.wrap(texto, width=max_chars) or [texto] for _, texto in campos_texto]

def crear_imagen_etiqueta(campos_texto, orientacion="vertical"):
    """Genera la imagen de la etiqueta ajustando ancho y alto al contenido.

    El contenido siempre se arma apilado (igual que en vertical); si la norma
    pide "horizontal" la imagen final se rota 90°, ya que ese formato es para
    material que se alimenta apaisado en la impresora de etiquetas."""
    if not campos_texto:
        raise ValueError("No hay campos con datos para generar la etiqueta")

    header_items = [(c, t) for c, t in campos_texto if c.strip().upper() == "EAN"]
    header_items += [(c, t) for c, t in campos_texto if c.strip().upper() == "MARCA"]
    footer_items = [(c, t) for c, t in campos_texto if c.strip().upper() in CAMPOS_PIE]
    footer_items += [(c, t) for c, t in campos_texto if c.strip().upper() == "TALLA"]
    excluidos = {c.strip().upper() for c in CAMPOS_ENCABEZADO} | {c.strip().upper() for c in CAMPOS_PIE} | {"TALLA"}
    center_items = [(c, t) for c, t in campos_texto if c.strip().upper() not in excluidos]

    font_header = _cargar_fuente(FONT_SIZE_HEADER)
    font_body = _cargar_fuente(FONT_SIZE_BODY)

    dummy = ImageDraw.Draw(Image.new("RGB", (10, 10)))

    grupos_header = _envolver_campos(header_items, MAX_CHARS_HEADER)
    grupos_center = _envolver_campos(center_items, MAX_CHARS_BODY)
    grupos_footer = _envolver_campos(footer_items, MAX_CHARS_BODY)

    line_height_header = font_header.size + 14
    line_height_body = font_body.size + 12

    max_width = 0
    for grupo in grupos_header:
        for linea in grupo:
            max_width = max(max_width, _medir_texto(dummy, linea, font_header))
    for grupo in grupos_center + grupos_footer:
        for linea in grupo:
            max_width = max(max_width, _medir_texto(dummy, linea, font_body))

    def _altura_grupos(grupos, line_height):
        altura = 0
        for i, grupo in enumerate(grupos):
            altura += len(grupo) * line_height
            if i < len(grupos) - 1:
                altura += FIELD_SPACER
        return altura

    altura_total = MARGIN_Y_TOP
    altura_total += _altura_grupos(grupos_header, line_height_header)
    if grupos_header:
        altura_total += HEADER_GAP
    altura_total += _altura_grupos(grupos_center, line_height_body)
    if grupos_footer:
        altura_total += FOOTER_GAP
        altura_total += _altura_grupos(grupos_footer, line_height_body)
    altura_total += MARGIN_Y_BOTTOM

    ancho = max(MIN_WIDTH_PX, int(max_width) + 2 * MARGIN_X)
    alto = max(MIN_HEIGHT_PX, int(altura_total))

    img = Image.new("RGB", (ancho, alto), "white")
    draw = ImageDraw.Draw(img)
    draw.rectangle([0, 0, ancho - 1, alto - 1], outline="black", width=2)

    def _dibujar_grupos(y, grupos, font, line_height):
        for i, grupo in enumerate(grupos):
            for linea in grupo:
                w = _medir_texto(draw, linea, font)
                x = (ancho - w) / 2
                draw.text((x, y), linea, font=font, fill="black")
                y += line_height
            if i < len(grupos) - 1:
                y += FIELD_SPACER
        return y

    y = MARGIN_Y_TOP
    y = _dibujar_grupos(y, grupos_header, font_header, line_height_header)
    if grupos_header:
        y += HEADER_GAP
    y = _dibujar_grupos(y, grupos_center, font_body, line_height_body)
    if grupos_footer:
        y += FOOTER_GAP
        _dibujar_grupos(y, grupos_footer, font_body, line_height_body)

    ancho_cm = ancho / DPI * 2.54
    alto_cm = alto / DPI * 2.54

    if orientacion == "horizontal":
        img = img.transpose(Image.ROTATE_90)
        ancho_cm, alto_cm = alto_cm, ancho_cm

    return img, ancho_cm, alto_cm

def guardar_etiqueta_pdf(imagen, ancho_cm, alto_cm, output_dir, nombre_base, nombres_usados):
    """Guarda una etiqueta como PDF individual (tamaño exacto de la etiqueta).

    `nombres_usados` es un dict compartido entre llamadas para evitar que dos
    etiquetas con el mismo EAN + norma se sobreescriban entre sí.
    """
    nombre = _nombre_archivo_seguro(nombre_base)
    contador = nombres_usados.get(nombre, 0)
    nombres_usados[nombre] = contador + 1
    nombre_final = nombre if contador == 0 else f"{nombre}_{contador + 1}"

    os.makedirs(output_dir, exist_ok=True)
    ruta_salida = os.path.join(output_dir, f"{nombre_final}.pdf")

    ancho_pt = ancho_cm * cm
    alto_pt = alto_cm * cm
    c = canvas.Canvas(ruta_salida, pagesize=(ancho_pt, alto_pt))
    c.drawImage(ImageReader(imagen), 0, 0, width=ancho_pt, height=alto_pt)
    c.showPage()
    c.save()

    return ruta_salida

def generar_etiquetas_desde_excel(
    excel_path,
    output_dir,
    config_path=DEFAULT_CONFIG_PATH,
    json_dir=DEFAULT_JSON_DIR,
    log_callback=None,
):
    def log(mensaje):
        if log_callback:
            log_callback(mensaje)

    config = cargar_config_etiquetas(config_path)
    mapa_numero_a_norma = construir_mapa_numero_a_norma(config)

    registros, json_path = excel_a_json(excel_path, json_dir)
    log(f"Excel convertido a JSON: {json_path}")

    filas_sin_codigo_formato = _filas_sin_codigo_formato(registros)
    if filas_sin_codigo_formato:
        filas_txt = ", ".join(str(f) for f in filas_sin_codigo_formato[:15])
        extra = "…" if len(filas_sin_codigo_formato) > 15 else ""
        raise ValueError(
            f"Falta la columna '{COLUMNA_NORMA}' en la(s) fila(s): {filas_txt}{extra}. "
            "Esa columna es indispensable para saber qué norma y qué armado le corresponde a "
            "cada etiqueta, así que hay que completarla en todas las filas antes de generar."
        )

    os.makedirs(output_dir, exist_ok=True)

    archivos_generados = []
    errores = []
    detalle = []
    nombres_usados = {}

    for idx, fila in enumerate(registros, start=1):
        item = _analizar_fila(fila, idx, mapa_numero_a_norma, config)
        registro = {
            "fila": item["fila"],
            "ean": item["ean"],
            "marca": item["marca"],
            "norma": item["norma"],
            "campos": [c for c, _ in item["campos_texto"]],
            "error": item["error"],
            "pdf_path": None,
        }

        if item["error"]:
            errores.append(f"Fila {idx}: {item['error']}")
            detalle.append(registro)
            continue

        try:
            img, ancho_cm, alto_cm = crear_imagen_etiqueta(item["campos_texto"], item["orientacion"])
        except Exception as e:
            mensaje = f"error generando la etiqueta ({e})"
            errores.append(f"Fila {idx}: {mensaje}")
            registro["error"] = mensaje
            detalle.append(registro)
            continue

        ean = item["ean"] or f"FILA{idx}"
        nombre_base = f"{ean}_{item['norma']}"
        carpeta_norma = os.path.join(output_dir, _nombre_archivo_seguro(item["norma"]))

        try:
            ruta_salida = guardar_etiqueta_pdf(img, ancho_cm, alto_cm, carpeta_norma, nombre_base, nombres_usados)
        except Exception as e:
            mensaje = f"error guardando la etiqueta ({e})"
            errores.append(f"Fila {idx}: {mensaje}")
            registro["error"] = mensaje
            detalle.append(registro)
            continue

        registro["pdf_path"] = ruta_salida
        archivos_generados.append(ruta_salida)
        log(f"Fila {idx}: etiqueta guardada -> {os.path.basename(ruta_salida)}")
        detalle.append(registro)

    if not archivos_generados:
        raise ValueError("No se generó ninguna etiqueta. Revisa el archivo Excel y la columna '" + COLUMNA_NORMA + "'.")

    log(f"Carpeta de salida: {output_dir}")

    return {
        "total_filas": len(registros),
        "generadas": len(archivos_generados),
        "errores": errores,
        "json_path": json_path,
        "output_dir": output_dir,
        "archivos": archivos_generados,
        "detalle": detalle,
    }

def main():
    import sys

    if len(sys.argv) < 3:
        print("Uso: python armadoEtiqueta.py <excel> <carpeta_salida>")
        return

    resultado = generar_etiquetas_desde_excel(sys.argv[1], sys.argv[2], log_callback=print)
    print(resultado)

if __name__ == "__main__":
    main()
    