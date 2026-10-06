"""Reglas NOM-004-SE-2021 (insumo, forro e instrucciones de cuidado) para revisar una fila del Excel."""
from __future__ import annotations

import json
import os
import re
import unicodedata
from typing import Any

NUMERO_NORMA = 4

# Catálogo editable de fibras: si no existe se crea con las fibras iniciales.
CATALOGO_FIBRAS_PATH = os.path.join("data", "fibras_nom004.json")

FIBRAS_NATURALES = ["Algodón", "Lana", "Seda", "Lino", "Cáñamo", "Yute", "Ramio", "Cashmere", "Mohair", "Alpaca", "Angora"]
FIBRAS_QUIMICAS = ["Poliéster", "Viscosa", "Nylon", "Poliamida", "Acrílico", "Elastano", "Polipropileno", "Acetato", "Spandex", "Lycra"]

COLUMNAS_INSUMO = ["INSUMOS", "INSUMO PRINCIPAL", "INSUMO PRICIPAL", "INSUMOS/INGREDIENTES"]
COLUMNAS_FORRO = ["FORRO"]
COLUMNAS_CUIDADO = ["CUIDADO", "INSTRUCCIONES DE CUIDADO"]

NO_APLICA = {"", "NO APLICA", "SI APLICA", "N/A", "NA", "NAN", "NONE", "VER ETIQUETA"}

# Orden de generate_care_instructions_text(); patrones sobre texto en minúsculas y sin acentos.
PARTES_CUIDADO: dict[str, tuple[str, str]] = {
    "lavado": ("tipo de lavado", r"lava(do|r)\s+(a\s+mano|a\s+maquina|en\s+seco|profesional)|lavado\s+profesional"),
    "temperatura": ("temperatura de lavado", r"temperatura|agua\s+(fria|tibia|caliente)|\bg?ua\s+(fria|tibia|caliente)|\b(fria|tibia|caliente)\b|\d+\s*°?\s*c\b"),
    "ciclo": ("ciclo", r"\bciclo\b"),
    "uso_de": ("uso de jabón o detergente", r"\b(jabon|detergente)\b"),
    "blanqueo": ("blanqueo", r"blanqueador|\bcloro\b"),
    "secado_maquina": ("secado en máquina", r"secado\s+en\s+maquina|secadora|secado\s+profesional"),
    "secado_exprimido": ("exprimido", r"exprimi"),
    "secado_centrifugado": ("centrifugado", r"centrifug"),
    "secado_colgado": ("colgado (línea u horizontal)", r"secado\s+de\s+linea|secado\s+horizontal|\blinea\b|tender|colgad"),
    "secado_ubicacion": ("ubicación del secado (sol o sombra)", r"al\s+sol\b|a\s+la\s+sombra"),
    "planchar": ("planchado", r"planch"),
}
PARTES_SECADO = ["secado_maquina", "secado_exprimido", "secado_centrifugado", "secado_colgado", "secado_ubicacion"]

ADICIONALES = [
    "Lavar por separado", "Lavar con colores similares", "Lavar antes de utilizar", "Lavar por el revés",
    "No exprimir o no torcer", "Limpiar sólo con paño húmedo", "No agregar suavizante para telas",
    "Retirar inmediatamente", "Planchar solo por el revés", "No planchar los adornos", "Utilizar paño para planchado",
    "No utilizar blanqueadores ópticos", "Utilizar red de lavado", "No planchar con vapor", "Sólo vapor", "No remojar",
    "Se recomienda plancha de vapor", "Secar lejos del calor directo", "Dar forma nuevamente mientras está húmedo",
    "Dar forma nuevamente y secar en posición horizontal",
    "Planchar con un paño para evitar abrillantamiento o amarillamiento",
]


def cargar_catalogo(ruta: str = CATALOGO_FIBRAS_PATH) -> list[str]:
    """Fibras del catálogo editable (JSON). Si no existe, lo crea con las fibras iniciales."""
    if not os.path.exists(ruta):
        os.makedirs(os.path.dirname(ruta) or ".", exist_ok=True)
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump({"naturales": FIBRAS_NATURALES, "quimicas": FIBRAS_QUIMICAS}, f, ensure_ascii=False, indent=4)
    with open(ruta, "r", encoding="utf-8") as f:
        catalogo = json.load(f)
    return [str(fibra).strip() for lista in catalogo.values() for fibra in lista if str(fibra).strip()]


def normalizar(texto: str) -> str:
    """Minúsculas y sin acentos (ñ -> n)."""
    sin_acentos = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    return sin_acentos.lower()


def formato_porcentaje(valor: float) -> str:
    return f"{valor:.2f}".rstrip("0").rstrip(".")


class Nom004Validador:
    def __init__(self, fibras: list[str] | None = None) -> None:
        # En producción pásale el catálogo editable (cargar_catalogo); sin él usa las fibras iniciales.
        self.fibras = fibras if fibras is not None else FIBRAS_NATURALES + FIBRAS_QUIMICAS

    # ---------- fila completa ----------

    def revisar(self, fila: dict[str, Any]) -> list[dict[str, Any]]:
        """Todas las reglas revisadas: [{columna, regla, ok, mensaje}]."""
        revision: list[dict[str, Any]] = []

        for columnas, que in ((COLUMNAS_INSUMO, "insumo"), (COLUMNAS_FORRO, "forro")):
            columna, valor = self._celda(fila, columnas)
            if columna is not None and not self._no_aplica(valor):
                revision += [{"columna": columna, **r} for r in self.revisar_composicion(valor, que)]

        columna, valor = self._celda(fila, COLUMNAS_CUIDADO)
        if columna is not None:
            revision += [{"columna": columna, **r} for r in self.revisar_cuidado(valor)]

        return revision

    @staticmethod
    def observaciones(revision: list[dict[str, Any]]) -> list[dict[str, str]]:
        """Solo lo que no cumple; si una fila tiene alguna, el lote no se genera."""
        return [{"columna": r["columna"], "mensaje": r["mensaje"]} for r in revision if not r["ok"]]

    # ---------- insumo / forro ----------

    def revisar_composicion(self, texto: str, que: str) -> list[dict[str, Any]]:
        texto = re.sub(r"^(insumos?|forro)\s*:?\s*", "", texto.strip(), flags=re.I).strip()
        partes = [p.strip() for p in re.split(r"[,;\n]+", texto) if p.strip()]
        Que = que[:1].upper() + que[1:]

        fibras: list[tuple[float, str]] = []
        mal_formadas: list[str] = []
        for parte in partes:
            m = re.match(r"^(\d+(?:[.,]\d+)?)\s*%\s*(.+)$", parte)
            if m:
                fibras.append((float(m.group(1).replace(",", ".")), m.group(2).strip()))
                continue
            m = re.match(r"^(.+?)\s+(\d+(?:[.,]\d+)?)\s*%$", parte)
            if m:
                fibras.append((float(m.group(2).replace(",", ".")), m.group(1).strip()))
                continue
            mal_formadas.append(f'"{parte}"')
        en_cero = [f for p, f in fibras if p <= 0]

        reglas: list[dict[str, Any]] = []
        if not mal_formadas and not en_cero:
            reglas.append({"regla": 'Formato "NN% fibra"', "ok": True, "mensaje": f"{Que}: cada parte trae porcentaje y fibra."})
        else:
            mensaje = ""
            if mal_formadas:
                mensaje += f"En el {que}, {', '.join(mal_formadas)} no tiene la forma \"NN% fibra\". "
            if en_cero:
                mensaje += f"En el {que}, el porcentaje de {', '.join(en_cero)} debe ser mayor a 0."
            reglas.append({"regla": 'Formato "NN% fibra"', "ok": False, "mensaje": mensaje.strip()})

        if not fibras:
            return reglas

        errores = [e for e in (self._revisar_fibra(f) for _, f in fibras) if e]
        reglas.append(
            {"regla": "Fibras del catálogo", "ok": True, "mensaje": f"{Que}: todas las fibras están en el catálogo y bien escritas."}
            if not errores else
            {"regla": "Fibras del catálogo", "ok": False, "mensaje": f"En el {que}, " + " ".join(errores)}
        )

        porcentajes = [p for p, _ in fibras]
        total = sum(porcentajes)
        suma = " + ".join(formato_porcentaje(p) for p in porcentajes)
        reglas.append(
            {"regla": "Suma 100%", "ok": True, "mensaje": f"{Que}: suma 100% ({suma})."}
            if abs(total - 100.0) <= 0.01 else
            {"regla": "Suma 100%", "ok": False, "mensaje": f"El {que} suma {formato_porcentaje(total)}% ({suma}); debe sumar 100%."}
        )

        ordenadas = sorted(fibras, key=lambda f: f[0], reverse=True)  # estable: los empates conservan su orden
        if [p for p, _ in ordenadas] == porcentajes:
            lista = ", ".join(f"{formato_porcentaje(p)}%" for p in porcentajes)
            reglas.append({"regla": "Orden de predominancia", "ok": True, "mensaje": f"{Que}: va de mayor a menor ({lista})."})
        else:
            debe = ", ".join(f"{formato_porcentaje(p)}% {f}" for p, f in ordenadas)
            reglas.append({"regla": "Orden de predominancia", "ok": False, "mensaje": f"El {que} no va de mayor a menor porcentaje; debe ir: {debe}."})

        return reglas

    def _revisar_fibra(self, fibra: str) -> str | None:
        for correcta in self.fibras:
            if correcta.lower() == fibra.lower():
                return None
        for correcta in self.fibras:
            if normalizar(correcta) == normalizar(fibra):
                return f'"{fibra}" está mal escrita; debe ser "{correcta.lower()}".'
        return f'la fibra "{fibra}" no está en el catálogo de fibras.'

    # ---------- instrucciones de cuidado ----------

    def revisar_cuidado(self, texto: str) -> list[dict[str, Any]]:
        if self._no_aplica(texto):
            return [{"regla": "Instrucciones completas", "ok": False, "mensaje": "Faltan las instrucciones de cuidado."}]

        normalizado = normalizar(texto)
        for adicional in ADICIONALES:
            frase = normalizar(adicional)
            normalizado = normalizado.replace(frase, " " * len(frase))  # conserva posiciones

        posiciones: dict[str, int] = {}
        for clave, (_, patron) in PARTES_CUIDADO.items():
            m = re.search(patron, normalizado)
            if m:
                posiciones[clave] = m.start()

        nombre = lambda clave: PARTES_CUIDADO[clave][0]
        faltan: list[str] = []
        if "lavado" not in posiciones:
            faltan.append(nombre("lavado"))
        elif re.search(r"a\s+(mano|maquina)", normalizado) and "temperatura" not in posiciones:
            faltan.append(nombre("temperatura"))
        if "blanqueo" not in posiciones:
            faltan.append(nombre("blanqueo"))
        if not any(p in posiciones for p in PARTES_SECADO):
            faltan.append("secado")
        if "planchar" not in posiciones:
            faltan.append(nombre("planchar"))

        reglas: list[dict[str, Any]] = [
            {"regla": "Instrucciones completas", "ok": True, "mensaje": "Cuidado: trae lavado, temperatura, blanqueo, secado y planchado."}
            if not faltan else
            {"regla": "Instrucciones completas", "ok": False, "mensaje": "Faltan en las instrucciones de cuidado: " + ", ".join(faltan) + "."}
        ]

        en_texto = sorted(posiciones, key=lambda c: posiciones[c])
        esperado = [c for c in PARTES_CUIDADO if c in posiciones]
        reglas.append(
            {"regla": "Orden de las instrucciones", "ok": True, "mensaje": "Cuidado: en el orden establecido."}
            if en_texto == esperado else
            {"regla": "Orden de las instrucciones", "ok": False,
             "mensaje": "Instrucciones de cuidado fuera de orden. Vienen: " + ", ".join(map(nombre, en_texto))
                        + ". Deben ir: " + ", ".join(map(nombre, esperado)) + "."}
        )
        return reglas

    # ---------- utilidades ----------

    @staticmethod
    def _celda(fila: dict[str, Any], columnas: list[str]) -> tuple[str | None, str]:
        for columna in columnas:
            for clave, valor in fila.items():
                if str(clave).strip().upper() == columna:
                    return str(clave), "" if valor is None else str(valor).strip()
        return None, ""

    @staticmethod
    def _no_aplica(valor: str) -> bool:
        return valor.strip().upper() in NO_APLICA
