# ARMADO DE ETIQUETAS SEGUN NORMA (NOM) A PARTIR DE UN EXCEL #
import json
import os
import re
import textwrap
from functools import lru_cache

import pandas as pd
from docx import Document
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

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

_CARACTERES_INVALIDOS = re.compile(r'[<>:"/\\|?*]')

# Unidades de peso/volumen que indican "contenido neto" (gramos, mililitros,
# litros, kilos...) en vez de una cantidad de piezas.
_UNIDADES_CONTENIDO_NETO = re.compile(r"\b(ML|MLS?|LTS?|L|KGS?|GRS?|G)\b", re.IGNORECASE)

def _nombre_archivo_seguro(texto):
    texto = _CARACTERES_INVALIDOS.sub("_", str(texto).strip())
    return texto or "SIN_DATO"

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

# El diseño se define en "píxeles a 300 DPI" (constantes de arriba) y se
# convierte a puntos tipográficos, que es la unidad de PDF y Word.
_PX_A_PT = 72 / DPI

def _pt(px):
    return px * _PX_A_PT

NOMBRE_FUENTE_PDF = "EtiquetaBold"

@lru_cache(maxsize=None)
def _fuente_pdf():
    """Registra Arial Bold en reportlab para escribir el texto como texto real
    (seleccionable/copiable). Si no está disponible usa Helvetica-Bold."""
    for ruta in RUTAS_FUENTE:
        try:
            pdfmetrics.registerFont(TTFont(NOMBRE_FUENTE_PDF, ruta))
            return NOMBRE_FUENTE_PDF
        except Exception:
            continue
    return "Helvetica-Bold"

def calcular_layout_etiqueta(campos_texto):
    """Calcula el armado de la etiqueta (bloques, líneas, tamaños y medidas en
    puntos) ajustando ancho y alto al contenido. El mismo layout se usa para
    el PDF y para el Word, así ambos salen idénticos."""
    if not campos_texto:
        raise ValueError("No hay campos con datos para generar la etiqueta")

    header_items = [(c, t) for c, t in campos_texto if c.strip().upper() == "EAN"]
    header_items += [(c, t) for c, t in campos_texto if c.strip().upper() == "MARCA"]
    footer_items = [(c, t) for c, t in campos_texto if c.strip().upper() in CAMPOS_PIE]
    footer_items += [(c, t) for c, t in campos_texto if c.strip().upper() == "TALLA"]
    excluidos = {c.strip().upper() for c in CAMPOS_ENCABEZADO} | {c.strip().upper() for c in CAMPOS_PIE} | {"TALLA"}
    center_items = [(c, t) for c, t in campos_texto if c.strip().upper() not in excluidos]

    fuente = _fuente_pdf()
    bloques = []
    grupos_header = _envolver_campos(header_items, MAX_CHARS_HEADER)
    if grupos_header:
        bloques.append({
            "grupos": grupos_header, "tamano": _pt(FONT_SIZE_HEADER),
            "interlineado": _pt(FONT_SIZE_HEADER + 14), "espacio_antes": 0,
        })
    grupos_center = _envolver_campos(center_items, MAX_CHARS_BODY)
    if grupos_center:
        bloques.append({
            "grupos": grupos_center, "tamano": _pt(FONT_SIZE_BODY),
            "interlineado": _pt(FONT_SIZE_BODY + 12),
            "espacio_antes": _pt(HEADER_GAP) if grupos_header else 0,
        })
    grupos_footer = _envolver_campos(footer_items, MAX_CHARS_BODY)
    if grupos_footer:
        espacio = _pt(FOOTER_GAP)
        if grupos_header and not grupos_center:
            espacio += _pt(HEADER_GAP)
        bloques.append({
            "grupos": grupos_footer, "tamano": _pt(FONT_SIZE_BODY),
            "interlineado": _pt(FONT_SIZE_BODY + 12), "espacio_antes": espacio,
        })

    max_width = 0
    altura = _pt(MARGIN_Y_TOP)
    for bloque in bloques:
        altura += bloque["espacio_antes"]
        for i, grupo in enumerate(bloque["grupos"]):
            for linea in grupo:
                max_width = max(max_width, pdfmetrics.stringWidth(linea, fuente, bloque["tamano"]))
            altura += len(grupo) * bloque["interlineado"]
            if i < len(bloque["grupos"]) - 1:
                altura += _pt(FIELD_SPACER)
    altura += _pt(MARGIN_Y_BOTTOM)

    return {
        "fuente": fuente,
        "bloques": bloques,
        "ancho": max(_pt(MIN_WIDTH_PX), max_width + 2 * _pt(MARGIN_X)),
        "alto": max(_pt(MIN_HEIGHT_PX), altura),
    }

def _ruta_salida_unica(output_dir, nombre_base, nombres_usados):
    """Ruta base (sin extensión) para los archivos de una etiqueta.

    `nombres_usados` es un dict compartido entre llamadas para evitar que dos
    etiquetas con el mismo EAN + norma se sobreescriban entre sí.
    """
    nombre = _nombre_archivo_seguro(nombre_base)
    contador = nombres_usados.get(nombre, 0)
    nombres_usados[nombre] = contador + 1
    nombre_final = nombre if contador == 0 else f"{nombre}_{contador + 1}"

    os.makedirs(output_dir, exist_ok=True)
    return os.path.join(output_dir, nombre_final)

def guardar_etiqueta_pdf(layout, orientacion, ruta_salida):
    """Guarda la etiqueta como PDF (tamaño exacto de la etiqueta) escribiendo
    el texto como texto real, para que se pueda seleccionar, copiar y pegar.

    El contenido siempre se arma apilado (igual que en vertical); si la norma
    pide "horizontal" la página se rota 90°, ya que ese formato es para
    material que se alimenta apaisado en la impresora de etiquetas."""
    ancho, alto = layout["ancho"], layout["alto"]
    fuente = layout["fuente"]
    pagina = (alto, ancho) if orientacion == "horizontal" else (ancho, alto)

    c = canvas.Canvas(ruta_salida, pagesize=pagina)
    c.setTitle(os.path.splitext(os.path.basename(ruta_salida))[0])
    if orientacion == "horizontal":
        c.translate(alto, 0)
        c.rotate(90)

    grosor = _pt(2)
    c.setLineWidth(grosor)
    c.rect(grosor / 2, grosor / 2, ancho - grosor, alto - grosor, stroke=1, fill=0)

    y = _pt(MARGIN_Y_TOP)  # medido desde arriba
    for bloque in layout["bloques"]:
        y += bloque["espacio_antes"]
        tamano = bloque["tamano"]
        ascenso = pdfmetrics.getAscent(fuente, tamano)
        c.setFont(fuente, tamano)
        for i, grupo in enumerate(bloque["grupos"]):
            for linea in grupo:
                c.drawCentredString(ancho / 2, alto - y - ascenso, linea)
                y += bloque["interlineado"]
            if i < len(bloque["grupos"]) - 1:
                y += _pt(FIELD_SPACER)

    c.showPage()
    c.save()
    return ruta_salida

def _agregar_xml(padre, etiqueta, atributos=None, hijos=()):
    elemento = OxmlElement(etiqueta)
    for clave, valor in (atributos or {}).items():
        elemento.set(qn(clave), valor)
    for hijo_tag, hijo_attrs in hijos:
        _agregar_xml(elemento, hijo_tag, hijo_attrs)
    padre.append(elemento)
    return elemento

def guardar_etiqueta_docx(layout, orientacion, ruta_salida):
    """Guarda la etiqueta como documento de Word (.docx) con el mismo tamaño,
    borde, fuente y acomodo que el PDF, para poder editarla o copiar su texto.

    La etiqueta es una tabla de una celda con borde que ocupa toda la página;
    en orientación horizontal se gira el texto de la celda 90° (abajo→arriba)."""
    ancho, alto = layout["ancho"], layout["alto"]
    horizontal = orientacion == "horizontal"
    pag_ancho, pag_alto = (alto, ancho) if horizontal else (ancho, alto)

    doc = Document()
    doc.core_properties.title = os.path.splitext(os.path.basename(ruta_salida))[0]
    seccion = doc.sections[0]
    seccion.page_width = Pt(pag_ancho)
    seccion.page_height = Pt(pag_alto)
    # Un pequeño margen evita que Word recorte el borde contra la orilla de la página.
    margen_borde = 2
    for margen in ("header_distance", "footer_distance", "gutter"):
        setattr(seccion, margen, Pt(0))
    for margen in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(seccion, margen, Pt(margen_borde))

    tabla = doc.add_table(rows=1, cols=1)
    tabla.autofit = False
    _agregar_xml(tabla._tbl.tblPr, "w:tblLayout", {"w:type": "fixed"})

    fila = tabla.rows[0]
    fila.height = Pt(pag_alto - 2 * margen_borde)
    fila.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    tabla.columns[0].width = Pt(pag_ancho - 2 * margen_borde)
    celda = fila.cells[0]
    celda.width = Pt(pag_ancho - 2 * margen_borde)

    tc_pr = celda._tc.get_or_add_tcPr()
    lados = ("top", "left", "bottom", "right")
    borde = {"w:val": "single", "w:sz": "4", "w:space": "0", "w:color": "000000"}
    _agregar_xml(tc_pr, "w:tcBorders", hijos=[(f"w:{lado}", borde) for lado in lados])
    _agregar_xml(tc_pr, "w:tcMar", hijos=[(f"w:{lado}", {"w:w": "0", "w:type": "dxa"}) for lado in lados])
    if horizontal:
        _agregar_xml(tc_pr, "w:textDirection", {"w:val": "btLr"})

    primero = True
    for bloque in layout["bloques"]:
        for i, grupo in enumerate(bloque["grupos"]):
            parrafo = celda.paragraphs[0] if primero else celda.add_paragraph()
            if primero:
                antes = _pt(MARGIN_Y_TOP) - margen_borde + bloque["espacio_antes"]
            else:
                antes = bloque["espacio_antes"] if i == 0 else _pt(FIELD_SPACER)
            primero = False

            formato = parrafo.paragraph_format
            formato.alignment = WD_ALIGN_PARAGRAPH.CENTER
            formato.space_before = Pt(antes)
            formato.space_after = Pt(0)
            formato.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            formato.line_spacing = Pt(bloque["interlineado"])

            for j, linea in enumerate(grupo):
                run = parrafo.add_run(linea)
                run.bold = True
                run.font.name = "Arial"
                run.font.size = Pt(bloque["tamano"])
                if j < len(grupo) - 1:
                    run.add_break()

    # Word exige un párrafo después de la tabla; se deja oculto y de 1 pt para
    # que no genere una segunda página.
    final = doc.add_paragraph()
    formato = final.paragraph_format
    formato.space_before = Pt(0)
    formato.space_after = Pt(0)
    formato.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    formato.line_spacing = Pt(1)
    _agregar_xml(final._p.get_or_add_pPr(), "w:rPr", hijos=[("w:vanish", {}), ("w:sz", {"w:val": "2"})])

    doc.save(ruta_salida)
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
    archivos_word = []
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
            "docx_path": None,
        }

        if item["error"]:
            errores.append(f"Fila {idx}: {item['error']}")
            detalle.append(registro)
            continue

        try:
            layout = calcular_layout_etiqueta(item["campos_texto"])
        except Exception as e:
            mensaje = f"error generando la etiqueta ({e})"
            errores.append(f"Fila {idx}: {mensaje}")
            registro["error"] = mensaje
            detalle.append(registro)
            continue

        ean = item["ean"] or f"FILA{idx}"
        nombre_base = f"{ean}_{item['norma']}"
        carpeta_norma = os.path.join(output_dir, _nombre_archivo_seguro(item["norma"]))

        ruta_base = _ruta_salida_unica(carpeta_norma, nombre_base, nombres_usados)
        try:
            ruta_pdf = guardar_etiqueta_pdf(layout, item["orientacion"], ruta_base + ".pdf")
            ruta_docx = guardar_etiqueta_docx(layout, item["orientacion"], ruta_base + ".docx")
        except Exception as e:
            mensaje = f"error guardando la etiqueta ({e})"
            errores.append(f"Fila {idx}: {mensaje}")
            registro["error"] = mensaje
            detalle.append(registro)
            continue

        registro["pdf_path"] = ruta_pdf
        registro["docx_path"] = ruta_docx
        archivos_generados.append(ruta_pdf)
        archivos_word.append(ruta_docx)
        log(f"Fila {idx}: etiqueta guardada -> {os.path.basename(ruta_base)} (.pdf y .docx)")
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
        "archivos_word": archivos_word,
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
    