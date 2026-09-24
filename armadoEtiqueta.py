# ARMADO DE ETIQUETAS SEGUN NORMA (NOM) A PARTIR DE UN EXCEL #
import json
import os
import re

import pandas as pd

from plantilla_asignaciones import (
    MEMBRETE_PATH, calcular_plantilla, guardar_plantilla_docx, guardar_plantilla_pdf,
)

DEFAULT_CONFIG_PATH = os.path.join("data", "config_etiquetas.json")
DEFAULT_JSON_DIR = os.path.join("data", "etiquetas")

COLUMNA_NORMA = "CODIGO FORMATO"

# Columnas opcionales del Excel para el encabezado de la hoja (plantilla de
# asignaciones). Se aceptan varios nombres porque cada cliente los escribe
# distinto; si no vienen, esa parte del encabezado simplemente se omite.
COLUMNAS_ASIGNACION = ("ASIGN", "ASIGNACION", "ASIGNACIÓN")
COLUMNAS_DESCRIPCION = ("DESCRIPCION", "DESCRIPCIÓN", "DENOMINACION", "DENOMINACIÓN")
# El tipo (Costura / Adherible) va fuera del recuadro, bajo el subtítulo, así
# que estas columnas nunca se imprimen dentro de la etiqueta aunque estén
# entre los campos de la norma.
COLUMNAS_TIPO = ("TIPO DE ETIQUETA", "TIPO")
# La altura del contenido (ej. "2mm") se imprime fuera del recuadro, junto a
# la leyenda "Altura del contenido". Sale de una columna MEDIDAS: si el Excel
# trae dos, la segunda es la altura y la primera son las medidas del producto
# (van dentro de la etiqueta en las normas que las llevan); si trae una sola,
# esa es la altura. Ver _altura_contenido.
COLUMNA_MEDIDAS = "MEDIDAS"

_CARACTERES_INVALIDOS = re.compile(r'[<>:"/\\|?*]')

# pandas renombra los encabezados repetidos como "MEDIDAS.1", "MEDIDAS.2"...
_SUFIJO_COLUMNA_REPETIDA = re.compile(r"\.\d+$")

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

def _primer_valor(fila, columnas):
    """Primer valor no vacío entre varias columnas alternativas."""
    for columna in columnas:
        valor = buscar_valor_columna(fila, columna)
        texto = str(valor).strip() if valor is not None else ""
        if texto and texto.upper() not in ("NAN", "N/A", "NONE"):
            return texto
    return ""

#Reglas para anteponer titulos ej: HECHO EN... FORRO, TALLA, PAIS ORIGEN.
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
    # CONTENIDO, INGREDIENTES E IMPORTADOR SE IMPRIMEN SOLO CON SU VALOR, SIN PREFIJO
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

def _columnas_medidas(fila):
    """Nombres de las columnas MEDIDAS de la fila, en el orden del Excel."""
    return [
        clave for clave in fila
        if _SUFIJO_COLUMNA_REPETIDA.sub("", str(clave)).strip().upper() == COLUMNA_MEDIDAS
    ]

def _altura_contenido(fila):
    """Devuelve (altura, medidas_van_en_etiqueta). Con dos columnas MEDIDAS la
    altura es la segunda y la primera se queda para la etiqueta; con una sola,
    esa es la altura y ya no se imprime dentro del recuadro."""
    columnas = _columnas_medidas(fila)
    if not columnas:
        return "", True
    columna_altura = columnas[1] if len(columnas) > 1 else columnas[0]
    return _primer_valor(fila, (columna_altura,)), len(columnas) > 1

def _titulo_hoja(fila, ean):
    """Título de la hoja, ej. 'TJX028 CREMA FACIAL 50 ml': asignación (o el
    EAN si el Excel no trae asignación), descripción y contenido."""
    partes = [
        _primer_valor(fila, COLUMNAS_ASIGNACION) or ean,
        _primer_valor(fila, COLUMNAS_DESCRIPCION),
        _primer_valor(fila, ("CONTENIDO",)),
    ]
    return " ".join(p for p in partes if p)

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
    altura_contenido, medidas_en_etiqueta = _altura_contenido(fila)
    fuera_de_etiqueta = COLUMNAS_TIPO if medidas_en_etiqueta else COLUMNAS_TIPO + (COLUMNA_MEDIDAS,)
    campos = [c for c in config[norma]["campos"] if c.strip().upper() not in fuera_de_etiqueta]
    campos_texto = extraer_campos_etiqueta(fila, campos)

    item["norma"] = norma
    item["campos_texto"] = campos_texto
    item["orientacion"] = config[norma].get("orientacion", "vertical")
    item["asignacion"] = _primer_valor(fila, COLUMNAS_ASIGNACION)
    item["titulo"] = _titulo_hoja(fila, item["ean"])
    item["tipo"] = _primer_valor(fila, COLUMNAS_TIPO).capitalize()
    item["altura_contenido"] = altura_contenido
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

    # Sin el membrete no se puede armar ninguna hoja; se avisa una sola vez
    # en lugar de marcar error en cada fila.
    if not os.path.exists(MEMBRETE_PATH):
        raise FileNotFoundError(
            f"No se encontró el membrete '{MEMBRETE_PATH}'. Debe estar en la carpeta img "
            "junto a la aplicación."
        )

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
            plantilla = calcular_plantilla(
                item["campos_texto"], item["titulo"], item["tipo"],
                item["altura_contenido"], item["orientacion"],
            )
        except Exception as e:
            mensaje = f"error generando la etiqueta ({e})"
            errores.append(f"Fila {idx}: {mensaje}")
            registro["error"] = mensaje
            detalle.append(registro)
            continue

        # Con asignación el archivo se llama como el título de la hoja (ej.
        # "TJX028 CREMA FACIAL 50 ml"); si no, como antes: EAN + norma.
        if item["asignacion"]:
            nombre_base = item["titulo"]
        else:
            ean = item["ean"] or f"FILA{idx}"
            nombre_base = f"{ean}_{item['norma']}"
        carpeta_norma = os.path.join(output_dir, _nombre_archivo_seguro(item["norma"]))

        ruta_base = _ruta_salida_unica(carpeta_norma, nombre_base, nombres_usados)
        try:
            ruta_pdf = guardar_plantilla_pdf(plantilla, ruta_base + ".pdf")
            ruta_docx = guardar_plantilla_docx(plantilla, ruta_base + ".docx")
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
    