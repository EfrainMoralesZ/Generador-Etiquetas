# -- SISTEMA V&C - GENERADOR DE ETIQUETAS -- #
import os
import json
import re
import shutil
import threading
from datetime import datetime

import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image
import fitz  # PyMuPDF

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:
    DND_FILES = None
    TkinterDnD = None

from armadoEtiqueta import generar_etiquetas_desde_excel, previsualizar_etiquetas_desde_excel
import configuracion

APP_VERSION = "1.0.0"
ESTADO_PATH = os.path.join("data", "estado_app.json")
LOTES_DIR = os.path.join("data", "lotes")

# ---------- ESTILO VISUAL V&C ---------- #
STYLE = {
    "primario": "#ECD925",
    "primario_hover": "#d9c520",
    "secundario": "#282828",
    "secundario_hover": "#3d3d3d",
    "exito": "#008D53",
    "exito_suave": "#E7F6EF",
    "advertencia": "#c0392b",
    "advertencia_suave": "#FDEBEA",
    "fondo": "#F8F9FA",
    "sidebar": "#FFFFFF",
    "surface": "#FFFFFF",
    "surface_alt": "#FFF9E3",
    "borde": "#E7E7E7",
    "texto_oscuro": "#282828",
    "texto_secundario": "#6B7280",
    "texto_claro": "#ffffff",
}

FONT_TITLE = ("Segoe UI", 20, "bold")
FONT_SUBTITLE = ("Segoe UI", 15, "bold")
FONT_LABEL = ("Segoe UI", 12)
FONT_SMALL = ("Segoe UI", 11)
FONT_TINY = ("Segoe UI", 10)
FONT_EMOJI = ("Segoe UI Emoji", 16)

COLUMNAS_ETIQUETAS = [("EAN", 2), ("Marca", 2), ("Norma", 3), ("Estado", 2), ("", 1), ("", 2), ("", 1)]
ETIQUETAS_POR_PAGINA = 50

ctk.set_appearance_mode("light")


def _formato_tamano(num_bytes):
    tamano = float(num_bytes)
    for unidad in ("B", "KB", "MB", "GB"):
        if tamano < 1024 or unidad == "GB":
            return f"{tamano:.0f} {unidad}" if unidad == "B" else f"{tamano:.1f} {unidad}"
        tamano /= 1024
    return f"{tamano:.1f} GB"


_CARACTERES_NO_SEGUROS = re.compile(r"[^A-Za-z0-9]+")


def _slug(texto):
    texto = _CARACTERES_NO_SEGUROS.sub("_", str(texto)).strip("_")
    return texto or "lote"


def _ruta_lote_unica(lote_id):
    ruta = os.path.join(LOTES_DIR, f"{lote_id}.jsonl")
    base, contador = lote_id, 1
    while os.path.exists(ruta):
        lote_id = f"{base}_{contador}"
        ruta = os.path.join(LOTES_DIR, f"{lote_id}.jsonl")
        contador += 1
    return lote_id, ruta


def _guardar_detalle_lote(detalle, lote_id):
    """Guarda el detalle de un lote en su propio archivo .jsonl (una etiqueta
    por línea) para no tener que cargar TODO el historial en memoria."""
    os.makedirs(LOTES_DIR, exist_ok=True)
    lote_id, ruta = _ruta_lote_unica(lote_id)
    with open(ruta, "w", encoding="utf-8") as f:
        for item in detalle:
            f.write(json.dumps(item, ensure_ascii=False))
            f.write("\n")
    return lote_id, ruta


def _migrar_lote_a_jsonl(lote):
    """Convierte un lote del formato antiguo (detalle embebido) al nuevo
    formato modular, escribiendo su detalle a un .jsonl aparte."""
    detalle = lote.pop("detalle", [])
    lote_id_base = f"{_slug(lote.get('nombre_excel', 'lote'))}_{_slug(lote.get('fecha', ''))}"
    lote_id, ruta = _guardar_detalle_lote(detalle, lote_id_base)
    lote["id"] = lote_id
    lote["detalle_path"] = ruta
    lote["total_detalle"] = len(detalle)
    lote.setdefault("total_filas", len(detalle))
    lote.setdefault("generadas", sum(1 for d in detalle if not d.get("error")))
    return lote


def _cargar_manifiesto_lotes():
    """Carga solo los metadatos de cada lote (ligero); el detalle de cada uno
    vive en su propio .jsonl y se lee bajo demanda, no aquí.

    Formatos soportados: el nuevo ({"lotes": [...]} con detalle_path) y el
    antiguo (detalle embebido, de versiones previas), migrando este último
    automáticamente para no perder el historial ni seguir arrastrando el
    archivo de estado pesado."""
    if not os.path.exists(ESTADO_PATH):
        return []
    try:
        with open(ESTADO_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return []

    if isinstance(data, dict) and "lotes" in data:
        lotes = data.get("lotes") or []
    elif isinstance(data, dict) and "detalle" in data:
        lotes = [data]
    else:
        lotes = []

    migrado = False
    lotes_normalizados = []
    for lote in lotes:
        if "detalle" in lote:
            lote = _migrar_lote_a_jsonl(lote)
            migrado = True
        lotes_normalizados.append(lote)

    if migrado:
        _guardar_manifiesto_lotes(lotes_normalizados)

    return lotes_normalizados


def _guardar_manifiesto_lotes(lotes):
    os.makedirs(os.path.dirname(ESTADO_PATH), exist_ok=True)
    with open(ESTADO_PATH, "w", encoding="utf-8") as f:
        json.dump({"lotes": lotes}, f, ensure_ascii=False, indent=2)


class GenerdorEtiquetas:

    PASOS = ["Archivo cargado", "Analizando datos", "Generando PDF", "Finalizado"]

    def __init__(self):
        self.excel_path = None
        self.resultado_analisis = None
        self._json_path_actual = None
        self._archivo_generado = False
        self.lotes = _cargar_manifiesto_lotes()
        self.estado_lote = self.lotes[-1] if self.lotes else None
        self.estado_pasos = ["pendiente"] * 4
        self.dnd_activo = False
        self.etiquetas_filtradas = []
        self.pagina_etiquetas_actual = 0
        self.modo_busqueda_etiquetas = False
        self._busqueda_after_id = None
        self._peticion_pagina_id = 0
        self._indices_lote = {}
        self._ventana_preview = None

        self._normas_config = {}
        self._norma_seleccionada = None
        self._creando_norma = False
        self._campos_editor = []
        self._orientacion_editor = configuracion.ORIENTACION_DEFECTO

        self.root = ctk.CTk()
        self.root.title("Generador de Etiquetas")
        self.root.geometry("1180x700")
        self.root.minsize(1020, 640)
        self.root.configure(fg_color=STYLE["fondo"])

        if TkinterDnD is not None:
            try:
                TkinterDnD.require(self.root)
                self.dnd_activo = True
            except Exception:
                self.dnd_activo = False

        self._construir_interfaz()
        self.root.mainloop()

    # ---------------------------------------------------------------- #
    # Estructura general: sidebar + páginas
    # ---------------------------------------------------------------- #
    def _construir_interfaz(self):
        self._construir_sidebar()

        self.contenedor_paginas = ctk.CTkFrame(self.root, fg_color=STYLE["fondo"], corner_radius=0)
        self.contenedor_paginas.pack(side="right", fill="both", expand=True)

        self.pagina_generador = self._crear_pagina_generador(self.contenedor_paginas)
        self.pagina_etiquetas = self._crear_pagina_etiquetas(self.contenedor_paginas)
        self.pagina_configuracion = self._crear_pagina_configuracion(self.contenedor_paginas)

        self._mostrar_pagina("generador")

    def _construir_sidebar(self):
        sidebar = ctk.CTkFrame(self.root, fg_color=STYLE["sidebar"], corner_radius=0, width=230)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        ctk.CTkFrame(sidebar, fg_color=STYLE["borde"], width=1).place(relx=1.0, rely=0, relheight=1, anchor="ne")

        top = ctk.CTkFrame(sidebar, fg_color="transparent")
        top.pack(fill="x", padx=20, pady=(24, 20))
        ctk.CTkLabel(
            top, text="🏷️", font=FONT_EMOJI, fg_color=STYLE["primario"],
            corner_radius=8, width=38, height=38
        ).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(
            top, text="Generador de\nEtiquetas", font=("Segoe UI", 14, "bold"),
            text_color=STYLE["texto_oscuro"], justify="left"
        ).pack(side="left")

        self.btn_nav = {}
        nav_items = [
            ("generador", "🏠", "Generador"),
            ("etiquetas", "🔎", "Etiquetas generadas"),
            ("config", "⚙️", "Configuración"),
        ]
        for clave, icono, texto in nav_items:
            btn = ctk.CTkButton(
                sidebar, text=f"  {icono}   {texto}", anchor="w", font=FONT_LABEL,
                fg_color="transparent", hover_color=STYLE["surface_alt"],
                text_color=STYLE["texto_oscuro"], corner_radius=8, height=38,
                command=lambda c=clave: self._mostrar_pagina(c)
            )
            btn.pack(fill="x", padx=14, pady=3)
            self.btn_nav[clave] = btn

        ctk.CTkFrame(sidebar, fg_color="transparent").pack(fill="both", expand=True)

        self.card_ultimo = ctk.CTkFrame(
            sidebar, fg_color=STYLE["fondo"], corner_radius=10,
            border_width=1, border_color=STYLE["borde"]
        )
        self.card_ultimo.pack(fill="x", padx=14, pady=(0, 10))
        self._refrescar_card_ultimo()

        ctk.CTkLabel(
            sidebar, text=f"Versión {APP_VERSION}", font=FONT_TINY,
            text_color=STYLE["texto_secundario"]
        ).pack(pady=(0, 16))

    def _refrescar_card_ultimo(self):
        for w in self.card_ultimo.winfo_children():
            w.destroy()

        ctk.CTkLabel(
            self.card_ultimo, text="📄  Último archivo generado", font=FONT_TINY,
            text_color=STYLE["texto_secundario"], anchor="w"
        ).pack(fill="x", padx=12, pady=(12, 4))

        if not self.estado_lote:
            ctk.CTkLabel(
                self.card_ultimo, text="Aún no has generado etiquetas.", font=FONT_SMALL,
                text_color=STYLE["texto_secundario"], anchor="w", justify="left", wraplength=170
            ).pack(fill="x", padx=12, pady=(0, 12))
            return

        nombre = self.estado_lote.get("nombre_excel", "—")
        fecha = self.estado_lote.get("fecha", "")
        ctk.CTkLabel(
            self.card_ultimo, text=nombre, font=("Segoe UI", 12, "bold"),
            text_color=STYLE["texto_oscuro"], anchor="w", justify="left", wraplength=180
        ).pack(fill="x", padx=12)
        ctk.CTkLabel(
            self.card_ultimo, text=fecha, font=FONT_TINY,
            text_color=STYLE["texto_secundario"], anchor="w"
        ).pack(fill="x", padx=12, pady=(0, 8))

        contenedor_btns = ctk.CTkFrame(self.card_ultimo, fg_color="transparent")
        contenedor_btns.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkButton(
            contenedor_btns, text="Abrir carpeta", font=FONT_TINY, height=28,
            fg_color=STYLE["surface"], hover_color=STYLE["surface_alt"],
            text_color=STYLE["texto_oscuro"], border_width=1, border_color=STYLE["borde"],
            corner_radius=6, command=self._abrir_carpeta_ultimo
        ).pack(fill="x", pady=(0, 6))
        ctk.CTkButton(
            contenedor_btns, text="Ver etiquetas", font=FONT_TINY, height=28,
            fg_color=STYLE["secundario"], hover_color=STYLE["secundario_hover"],
            text_color=STYLE["texto_claro"], corner_radius=6,
            command=lambda: self._mostrar_pagina("etiquetas")
        ).pack(fill="x")

    def _abrir_carpeta_ultimo(self):
        if not self.estado_lote:
            return
        carpeta = self.estado_lote.get("output_dir")
        if carpeta and os.path.isdir(carpeta):
            os.startfile(carpeta)
        else:
            messagebox.showwarning(
                "Carpeta no encontrada",
                "La carpeta de este lote ya no existe o fue movida."
            )

    def _mostrar_pagina(self, clave):
        paginas = {
            "generador": self.pagina_generador,
            "etiquetas": self.pagina_etiquetas,
            "config": self.pagina_configuracion,
        }
        for pagina in paginas.values():
            pagina.pack_forget()
        paginas[clave].pack(fill="both", expand=True)

        for c, btn in self.btn_nav.items():
            btn.configure(fg_color=STYLE["surface_alt"] if c == clave else "transparent")

        if clave == "etiquetas":
            self._refrescar_pagina_etiquetas()
        elif clave == "config":
            self._refrescar_pagina_configuracion()

    @staticmethod
    def _limpiar_frame(frame):
        for w in frame.winfo_children():
            w.destroy()

    # ---------------------------------------------------------------- #
    # Página: Generador
    # ---------------------------------------------------------------- #
    def _crear_pagina_generador(self, master):
        pagina = ctk.CTkFrame(master, fg_color=STYLE["fondo"], corner_radius=0)

        # header = ctk.CTkFrame(pagina, fg_color="transparent")
        # header.pack(fill="x", padx=30, pady=(26, 10))
        # ctk.CTkLabel(
        #     header, text="🏷️  Generador de Etiquetas", font=FONT_TITLE,
        #     text_color=STYLE["texto_oscuro"]
        # ).pack(anchor="w")
        # ctk.CTkLabel(
        #     header, text="Convierte tu archivo Excel en etiquetas PDF de manera rápida y sencilla.",
        #     font=FONT_LABEL, text_color=STYLE["texto_secundario"]
        # ).pack(anchor="w", pady=(4, 0))

        cuerpo = ctk.CTkFrame(pagina, fg_color="transparent")
        cuerpo.pack(fill="both", expand=True, padx=30, pady=(10, 24))
        cuerpo.grid_columnconfigure(0, weight=3)
        cuerpo.grid_columnconfigure(1, weight=2)
        cuerpo.grid_rowconfigure(0, weight=1)

        centro = ctk.CTkFrame(cuerpo, fg_color="transparent")
        centro.grid(row=0, column=0, sticky="nsew", padx=(0, 20))

        self.dropzone = ctk.CTkFrame(
            centro, fg_color=STYLE["surface_alt"], corner_radius=14,
            border_width=2, border_color=STYLE["primario"]
        )
        self.dropzone.pack(fill="x", pady=(0, 16))
        self._render_dropzone_vacio()
        self._registrar_drop_target()

        stepper_card = ctk.CTkFrame(
            centro, fg_color=STYLE["surface"], corner_radius=14,
            border_width=1, border_color=STYLE["borde"]
        )
        stepper_card.pack(fill="x", pady=(0, 16))
        self._construir_stepper(stepper_card)

        self.btn_generar = ctk.CTkButton(
            centro, text="🚀  Generar Etiquetas", font=FONT_SUBTITLE, height=54,
            fg_color=STYLE["primario"], hover_color=STYLE["primario_hover"],
            text_color=STYLE["texto_oscuro"], corner_radius=12,
            state="disabled", command=self.generar_pdf
        )
        self.btn_generar.pack(fill="x")
        ctk.CTkLabel(
            centro, text="Se generará un archivo PDF por cada etiqueta detectada.",
            font=FONT_TINY, text_color=STYLE["texto_secundario"]
        ).pack(anchor="w", pady=(6, 0))

        actividad_card = ctk.CTkFrame(
            cuerpo, fg_color=STYLE["surface"], corner_radius=14,
            border_width=1, border_color=STYLE["borde"]
        )
        actividad_card.grid(row=0, column=1, sticky="nsew")
        ctk.CTkLabel(
            actividad_card, text="🕐  Actividad reciente", font=FONT_SUBTITLE,
            text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", padx=16, pady=(16, 8))

        self.actividad_scroll = ctk.CTkScrollableFrame(actividad_card, fg_color="transparent")
        self.actividad_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        self.banner_resultado = ctk.CTkFrame(actividad_card, fg_color="transparent")
        self.banner_resultado.pack(fill="x", padx=16, pady=(0, 16))

        return pagina

    def _registrar_drop_target(self):
        if not self.dnd_activo:
            return
        self.dropzone.drop_target_register(DND_FILES)
        self.dropzone.dnd_bind("<<Drop>>", self._on_drop)

    def _render_dropzone_vacio(self):
        self._limpiar_frame(self.dropzone)
        contenido = ctk.CTkFrame(self.dropzone, fg_color="transparent")
        contenido.pack(fill="x", padx=20, pady=28)

        ctk.CTkLabel(contenido, text="📄", font=("Segoe UI Emoji", 34)).pack()

        titulo = "Arrastra tu archivo Excel aquí" if self.dnd_activo else "Selecciona tu archivo Excel"
        ctk.CTkLabel(
            contenido, text=titulo, font=("Segoe UI", 14, "bold"),
            text_color=STYLE["texto_oscuro"]
        ).pack(pady=(8, 2))

        subtitulo = "o haz clic para seleccionarlo" if self.dnd_activo else "Formatos permitidos: .xlsx, .xls"
        ctk.CTkLabel(
            contenido, text=subtitulo, font=FONT_SMALL, text_color=STYLE["texto_secundario"]
        ).pack()

        ctk.CTkButton(
            contenido, text="📁  Seleccionar archivo", font=FONT_LABEL, height=36, width=210,
            fg_color=STYLE["primario"], hover_color=STYLE["primario_hover"],
            text_color=STYLE["texto_oscuro"], corner_radius=8,
            command=self.seleccionar_excel
        ).pack(pady=(14, 0))

    def _render_dropzone_archivo(self, ruta):
        self._limpiar_frame(self.dropzone)
        fila = ctk.CTkFrame(self.dropzone, fg_color="transparent")
        fila.pack(fill="x", padx=18, pady=18)

        ctk.CTkLabel(fila, text="📊", font=("Segoe UI Emoji", 26)).pack(side="left", padx=(0, 12))

        info = ctk.CTkFrame(fila, fg_color="transparent")
        info.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(
            info, text=os.path.basename(ruta), font=("Segoe UI", 13, "bold"),
            text_color=STYLE["texto_oscuro"], anchor="w"
        ).pack(fill="x")

        try:
            tamano = _formato_tamano(os.path.getsize(ruta))
        except OSError:
            tamano = "—"
        self.lbl_info_archivo = ctk.CTkLabel(
            info, text=tamano, font=FONT_TINY, text_color=STYLE["texto_secundario"], anchor="w"
        )
        self.lbl_info_archivo.pack(fill="x")

        ctk.CTkButton(
            fila, text="✕", width=32, height=32, font=FONT_LABEL,
            fg_color="transparent", hover_color=STYLE["advertencia_suave"],
            text_color=STYLE["texto_secundario"], corner_radius=8,
            command=self._quitar_archivo
        ).pack(side="right")

    def _construir_stepper(self, master):
        contenedor = ctk.CTkFrame(master, fg_color="transparent")
        contenedor.pack(fill="x", padx=18, pady=18)

        fila_pasos = ctk.CTkFrame(contenedor, fg_color="transparent")
        fila_pasos.pack(fill="x")

        self.paso_widgets = []
        for i, nombre in enumerate(self.PASOS):
            columna = ctk.CTkFrame(fila_pasos, fg_color="transparent", width=1, height=1)
            columna.grid(row=0, column=i * 2, sticky="n")
            fila_pasos.grid_columnconfigure(i * 2, weight=1)

            circulo = ctk.CTkLabel(
                columna, text=str(i + 1), width=32, height=32, corner_radius=16,
                fg_color=STYLE["borde"], text_color=STYLE["texto_secundario"],
                font=("Segoe UI", 13, "bold")
            )
            circulo.pack()
            titulo = ctk.CTkLabel(columna, text=nombre, font=FONT_SMALL, text_color=STYLE["texto_oscuro"])
            titulo.pack(pady=(6, 0))
            estado = ctk.CTkLabel(columna, text="Pendiente", font=FONT_TINY, text_color=STYLE["texto_secundario"])
            estado.pack()

            widgets = {"circulo": circulo, "titulo": titulo, "estado": estado}

            if i < len(self.PASOS) - 1:
                linea = ctk.CTkFrame(fila_pasos, fg_color=STYLE["borde"], width=1, height=2)
                linea.grid(row=0, column=i * 2 + 1, sticky="ew", pady=(15, 0))
                fila_pasos.grid_columnconfigure(i * 2 + 1, weight=2)
                widgets["linea_siguiente"] = linea

            self.paso_widgets.append(widgets)

        self.progress = ctk.CTkProgressBar(contenedor, fg_color=STYLE["borde"], progress_color=STYLE["primario"])
        self.progress.set(0)
        self.progress.pack(fill="x", pady=(18, 4))

        self.lbl_progreso = ctk.CTkLabel(contenedor, text="", font=FONT_TINY, text_color=STYLE["texto_secundario"])
        self.lbl_progreso.pack(anchor="e")

    def _actualizar_stepper(self):
        for i, widgets in enumerate(self.paso_widgets):
            estado = self.estado_pasos[i]
            if estado == "completado":
                widgets["circulo"].configure(text="✓", fg_color=STYLE["exito"], text_color=STYLE["texto_claro"])
                widgets["estado"].configure(text="Completado", text_color=STYLE["exito"])
            elif estado == "progreso":
                widgets["circulo"].configure(text=str(i + 1), fg_color=STYLE["primario"], text_color=STYLE["texto_oscuro"])
                widgets["estado"].configure(text="En progreso", text_color=STYLE["texto_oscuro"])
            else:
                widgets["circulo"].configure(text=str(i + 1), fg_color=STYLE["borde"], text_color=STYLE["texto_secundario"])
                widgets["estado"].configure(text="Pendiente", text_color=STYLE["texto_secundario"])

            if "linea_siguiente" in widgets:
                color = STYLE["exito"] if estado == "completado" else STYLE["borde"]
                widgets["linea_siguiente"].configure(fg_color=color)

    def _agregar_actividad(self, icono, titulo, subtitulo=""):
        fila = ctk.CTkFrame(self.actividad_scroll, fg_color="transparent")
        fila.pack(fill="x", pady=6)

        hora = datetime.now().strftime("%H:%M")
        ctk.CTkLabel(fila, text=icono, font=FONT_EMOJI, width=26).pack(side="left", anchor="n")

        info = ctk.CTkFrame(fila, fg_color="transparent")
        info.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(
            info, text=f"{hora}   {titulo}", font=("Segoe UI", 11, "bold"),
            text_color=STYLE["texto_oscuro"], anchor="w", justify="left", wraplength=210
        ).pack(fill="x")
        if subtitulo:
            ctk.CTkLabel(
                info, text=subtitulo, font=FONT_TINY, text_color=STYLE["texto_secundario"],
                anchor="w", justify="left", wraplength=210
            ).pack(fill="x")

        self.root.update_idletasks()
        try:
            self.actividad_scroll._parent_canvas.yview_moveto(1.0)
        except Exception:
            pass

    # ---------------------------------------------------------------- #
    # Flujo: carga -> análisis -> generación
    # ---------------------------------------------------------------- #
    def seleccionar_excel(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if ruta:
            self._cargar_archivo(ruta)

    def _on_drop(self, event):
        rutas = self.root.tk.splitlist(event.data)
        for ruta in rutas:
            if ruta.lower().endswith((".xlsx", ".xls")):
                self._cargar_archivo(ruta)
                return
        messagebox.showwarning("Archivo no válido", "Arrastra un archivo Excel (.xlsx o .xls).")

    def _limpiar_json_abandonado(self):
        """Si el archivo cargado hasta ahora nunca llegó a generarse, borra
        el .json que se guardó en data/etiquetas al analizarlo (se guarda
        desde la subida para poder inspeccionarlo, pero no debe quedar
        huérfano si el usuario nunca le da a 'Generar Etiquetas')."""
        if self._json_path_actual and not self._archivo_generado:
            try:
                if os.path.exists(self._json_path_actual):
                    os.remove(self._json_path_actual)
            except OSError:
                pass
        self._json_path_actual = None
        self._archivo_generado = False

    def _cargar_archivo(self, ruta):
        self._limpiar_json_abandonado()

        self.excel_path = ruta
        self.resultado_analisis = None
        self._render_dropzone_archivo(ruta)
        self.btn_generar.configure(state="disabled")

        self.estado_pasos = ["completado", "progreso", "pendiente", "pendiente"]
        self._actualizar_stepper()
        self.progress.set(0)
        self.lbl_progreso.configure(text="")
        self._limpiar_frame(self.banner_resultado)
        self._limpiar_frame(self.actividad_scroll)

        self._agregar_actividad("📥", "Archivo cargado", os.path.basename(ruta))
        self._agregar_actividad("🔍", "Analizando datos", "Validando información y aplicando reglas")

        hilo = threading.Thread(target=self._analizar_en_hilo, args=(ruta,), daemon=True)
        hilo.start()

    def _quitar_archivo(self):
        self._limpiar_json_abandonado()

        self.excel_path = None
        self.resultado_analisis = None
        self.estado_pasos = ["pendiente"] * 4
        self._actualizar_stepper()
        self.progress.set(0)
        self.lbl_progreso.configure(text="")
        self.btn_generar.configure(state="disabled")
        self._render_dropzone_vacio()
        self._registrar_drop_target()

    def _analizar_en_hilo(self, ruta):
        try:
            resultado = previsualizar_etiquetas_desde_excel(ruta)
            self.root.after(0, self._analisis_completado, resultado)
        except Exception as e:
            self.root.after(0, self._analisis_fallido, str(e))

    def _analisis_completado(self, resultado):
        self.resultado_analisis = resultado
        self._json_path_actual = resultado.get("json_path")
        self._archivo_generado = False
        self.estado_pasos[1] = "completado"
        self._actualizar_stepper()

        if hasattr(self, "lbl_info_archivo"):
            self.lbl_info_archivo.configure(
                text=f"{resultado['listas']} de {resultado['total_filas']} filas listas"
            )

        filas_sin_codigo = resultado.get("filas_sin_codigo_formato") or []
        if filas_sin_codigo:
            self.btn_generar.configure(state="disabled")
            filas_txt = ", ".join(str(f) for f in filas_sin_codigo[:15])
            extra = "…" if len(filas_sin_codigo) > 15 else ""
            self._agregar_actividad(
                "❌", "Falta la columna 'CODIGO FORMATO'",
                f"Falta en {len(filas_sin_codigo)} fila(s): {filas_txt}{extra}"
            )
            messagebox.showerror(
                "Falta la columna 'CODIGO FORMATO'",
                f"Falta un valor en la columna 'CODIGO FORMATO' en {len(filas_sin_codigo)} "
                f"fila(s): {filas_txt}{extra}\n\n"
                "Esa columna es indispensable para saber qué norma y qué armado le corresponde "
                "a cada etiqueta, así que hay que completarla en todas las filas antes de poder "
                "generar el lote."
            )
        elif resultado["listas"] > 0:
            self.btn_generar.configure(state="normal")
            self._agregar_actividad(
                "✅", "Datos analizados",
                f"{resultado['listas']} de {resultado['total_filas']} filas listas para generar etiqueta"
            )
        else:
            self.btn_generar.configure(state="disabled")
            self._agregar_actividad(
                "⚠️", "Sin filas válidas",
                "Ninguna fila coincide con una norma configurada"
            )

    def _analisis_fallido(self, mensaje):
        self.estado_pasos[1] = "pendiente"
        self._actualizar_stepper()
        self._agregar_actividad("❌", "Error al analizar", mensaje)
        messagebox.showerror("Error", mensaje)

    def generar_pdf(self):
        if not self.excel_path or not self.resultado_analisis:
            return

        carpeta_padre = filedialog.askdirectory(title="Selecciona dónde crear la carpeta de etiquetas")
        if not carpeta_padre:
            return

        nombre_excel = os.path.splitext(os.path.basename(self.excel_path))[0]
        marca_tiempo = datetime.now().strftime("%Y%m%d_%H%M%S")
        carpeta_salida = os.path.join(carpeta_padre, f"Etiquetas_{nombre_excel}_{marca_tiempo}")

        self.btn_generar.configure(state="disabled")
        self.estado_pasos[2] = "progreso"
        self._actualizar_stepper()
        self.progress.set(0)
        self.lbl_progreso.configure(text="0%")
        self._limpiar_frame(self.banner_resultado)

        total = self.resultado_analisis["total_filas"] or 1
        self._contador_generadas = 0

        self._agregar_actividad("🖨️", "Generando PDF", "Creando etiquetas… esto puede tardar unos segundos")

        def on_log(mensaje):
            if mensaje.startswith("Fila "):
                self._contador_generadas += 1
                pct = min(self._contador_generadas / total, 1.0)
                self.root.after(0, self._actualizar_progreso, pct)

        hilo = threading.Thread(
            target=self._generar_en_hilo, args=(self.excel_path, carpeta_salida, on_log), daemon=True
        )
        hilo.start()

    def _actualizar_progreso(self, pct):
        self.progress.set(pct)
        self.lbl_progreso.configure(text=f"{int(pct * 100)}%")

    def _generar_en_hilo(self, excel_path, carpeta_salida, on_log):
        try:
            resultado = generar_etiquetas_desde_excel(excel_path, carpeta_salida, log_callback=on_log)
            self.root.after(0, self._generacion_completada, resultado)
        except Exception as e:
            self.root.after(0, self._generacion_fallida, str(e))

    def _generacion_completada(self, resultado):
        self._archivo_generado = True
        self.estado_pasos[2] = "completado"
        self.estado_pasos[3] = "completado"
        self._actualizar_stepper()
        self.progress.set(1.0)
        self.lbl_progreso.configure(text="100%")
        self.btn_generar.configure(state="normal")

        self._agregar_actividad(
            "📦", "PDF generado",
            f"{resultado['generadas']} de {resultado['total_filas']} etiquetas generadas"
        )

        if resultado["errores"]:
            extra = "…" if len(resultado["errores"]) > 3 else ""
            self._agregar_actividad(
                "⚠️", f"{len(resultado['errores'])} fila(s) con problemas",
                "; ".join(resultado["errores"][:3]) + extra
            )

        nombre_excel = os.path.basename(self.excel_path)
        lote_id_base = f"{_slug(nombre_excel)}_{datetime.now().strftime('%Y%m%d%H%M%S')}"
        lote_id, detalle_path = _guardar_detalle_lote(resultado["detalle"], lote_id_base)

        self.estado_lote = {
            "id": lote_id,
            "nombre_excel": nombre_excel,
            "fecha": datetime.now().strftime("%d/%m/%Y · %H:%M"),
            "output_dir": resultado["output_dir"],
            "json_path": resultado["json_path"],
            "detalle_path": detalle_path,
            "total_filas": resultado["total_filas"],
            "generadas": resultado["generadas"],
            "total_detalle": len(resultado["detalle"]),
        }
        self.lotes.append(self.estado_lote)
        _guardar_manifiesto_lotes(self.lotes)
        self._refrescar_card_ultimo()
        self._mostrar_banner_resultado(True, resultado)

        try:
            os.startfile(resultado["output_dir"])
        except Exception:
            pass

    def _generacion_fallida(self, mensaje):
        self.estado_pasos[2] = "pendiente"
        self._actualizar_stepper()
        self.btn_generar.configure(state="normal")
        self._agregar_actividad("❌", "Error al generar", mensaje)
        self._mostrar_banner_resultado(False, {"mensaje": mensaje})
        messagebox.showerror("Error", mensaje)

    def _mostrar_banner_resultado(self, exito, resultado):
        self._limpiar_frame(self.banner_resultado)
        color_fondo = STYLE["exito_suave"] if exito else STYLE["advertencia_suave"]
        color_texto = STYLE["exito"] if exito else STYLE["advertencia"]

        card = ctk.CTkFrame(self.banner_resultado, fg_color=color_fondo, corner_radius=10)
        card.pack(fill="x")

        titulo = "✅ Proceso completado" if exito else "❌ Proceso con errores"
        subtitulo = (
            "Tu archivo de etiquetas ha sido generado exitosamente."
            if exito else resultado.get("mensaje", "")
        )
        ctk.CTkLabel(
            card, text=titulo, font=("Segoe UI", 12, "bold"), text_color=color_texto, anchor="w"
        ).pack(fill="x", padx=12, pady=(10, 0))
        ctk.CTkLabel(
            card, text=subtitulo, font=FONT_TINY, text_color=color_texto, anchor="w",
            justify="left", wraplength=230
        ).pack(fill="x", padx=12, pady=(2, 10))

    # ---------------------------------------------------------------- #
    # Página: Etiquetas generadas
    # ---------------------------------------------------------------- #
    def _crear_pagina_etiquetas(self, master):
        pagina = ctk.CTkFrame(master, fg_color=STYLE["fondo"], corner_radius=0)

        header = ctk.CTkFrame(pagina, fg_color="transparent")
        header.pack(fill="x", padx=30, pady=(26, 10))
        ctk.CTkLabel(
            header, text="🔎  Etiquetas generadas", font=FONT_TITLE, text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w")
        self.lbl_subtitulo_etiquetas = ctk.CTkLabel(
            header, text="Busca por EAN o por norma y descarga el PDF de cada etiqueta.",
            font=FONT_LABEL, text_color=STYLE["texto_secundario"]
        )
        self.lbl_subtitulo_etiquetas.pack(anchor="w", pady=(4, 0))

        barra = ctk.CTkFrame(pagina, fg_color="transparent")
        barra.pack(fill="x", padx=30, pady=(0, 12))
        self.entrada_busqueda = ctk.CTkEntry(
            barra, placeholder_text="🔍  Buscar por EAN o norma...", font=FONT_LABEL, height=38
        )
        self.entrada_busqueda.pack(fill="x")
        self.entrada_busqueda.bind("<KeyRelease>", lambda e: self._filtrar_etiquetas())

        encabezados = ctk.CTkFrame(pagina, fg_color="transparent")
        encabezados.pack(fill="x", padx=34)
        self._configurar_columnas(encabezados)
        for i, (texto, _) in enumerate(COLUMNAS_ETIQUETAS):
            ctk.CTkLabel(
                encabezados, text=texto, font=("Segoe UI", 11, "bold"),
                text_color=STYLE["texto_secundario"], anchor="w"
            ).grid(row=0, column=i, sticky="ew", padx=6, pady=(0, 6))

        self.lista_etiquetas_frame = ctk.CTkScrollableFrame(pagina, fg_color="transparent")
        self.lista_etiquetas_frame.pack(fill="both", expand=True, padx=24, pady=(0, 6))

        self.lbl_mensaje_etiquetas = ctk.CTkLabel(
            self.lista_etiquetas_frame, text="", font=FONT_LABEL, text_color=STYLE["texto_secundario"]
        )
        self._filas_pool_etiquetas = []

        paginacion = ctk.CTkFrame(pagina, fg_color="transparent")
        paginacion.pack(fill="x", padx=24, pady=(0, 20))
        self.btn_pagina_anterior = ctk.CTkButton(
            paginacion, text="◀  Anterior", font=FONT_SMALL, height=32, width=110,
            fg_color=STYLE["surface"], hover_color=STYLE["surface_alt"],
            text_color=STYLE["texto_oscuro"], border_width=1, border_color=STYLE["borde"],
            corner_radius=6, command=lambda: self._cambiar_pagina_etiquetas(-1)
        )
        self.btn_pagina_anterior.pack(side="left")
        self.lbl_pagina_etiquetas = ctk.CTkLabel(
            paginacion, text="", font=FONT_SMALL, text_color=STYLE["texto_secundario"]
        )
        self.lbl_pagina_etiquetas.pack(side="left", expand=True)
        self.btn_pagina_siguiente = ctk.CTkButton(
            paginacion, text="Siguiente  ▶", font=FONT_SMALL, height=32, width=110,
            fg_color=STYLE["surface"], hover_color=STYLE["surface_alt"],
            text_color=STYLE["texto_oscuro"], border_width=1, border_color=STYLE["borde"],
            corner_radius=6, command=lambda: self._cambiar_pagina_etiquetas(1)
        )
        self.btn_pagina_siguiente.pack(side="right")

        return pagina

    @staticmethod
    def _configurar_columnas(frame):
        for i, (_, peso) in enumerate(COLUMNAS_ETIQUETAS):
            frame.grid_columnconfigure(i, weight=peso)

    def _lotes_orden_visual(self):
        """Lotes del más reciente al más antiguo (el más nuevo se ve primero)."""
        return list(reversed(self.lotes))

    def _total_etiquetas(self):
        """Total de etiquetas en todo el historial, sin abrir ningún .jsonl
        (usa el conteo que ya viene en el manifiesto liviano)."""
        return sum(lote.get("total_detalle", 0) for lote in self.lotes)

    def _indice_lote(self, ruta):
        """Índice de offsets (byte de inicio de cada línea) de un .jsonl,
        construido una sola vez por archivo y reutilizado durante la sesión."""
        if ruta in self._indices_lote:
            return self._indices_lote[ruta]
        offsets = []
        try:
            with open(ruta, "rb") as f:
                offset = f.tell()
                for linea in f:
                    if linea.strip():
                        offsets.append(offset)
                    offset = f.tell()
        except FileNotFoundError:
            offsets = []
        self._indices_lote[ruta] = offsets
        return offsets

    @staticmethod
    def _leer_lineas_lote(ruta, offsets, indices):
        resultado = []
        if not indices:
            return resultado
        try:
            with open(ruta, "rb") as f:
                for idx in indices:
                    if 0 <= idx < len(offsets):
                        f.seek(offsets[idx])
                        linea = f.readline()
                        if linea:
                            resultado.append(json.loads(linea.decode("utf-8")))
        except FileNotFoundError:
            pass
        return resultado

    def _leer_pagina_historial(self, inicio, fin):
        """Lee solo el rango [inicio, fin) del historial completo, tocando
        únicamente los .jsonl de los lotes que caen dentro de ese rango."""
        resultado = []
        acumulado = 0
        for lote in self._lotes_orden_visual():
            total_lote = lote.get("total_detalle", 0)
            lote_inicio, lote_fin = acumulado, acumulado + total_lote
            acumulado = lote_fin

            if lote_fin <= inicio or lote_inicio >= fin or not total_lote:
                continue

            ruta = lote.get("detalle_path")
            if not ruta:
                continue

            local_inicio = max(0, inicio - lote_inicio)
            local_fin = min(total_lote, fin - lote_inicio)

            offsets = self._indice_lote(ruta)
            items = self._leer_lineas_lote(ruta, offsets, range(local_inicio, local_fin))
            for item in items:
                item = dict(item)
                item["_excel_origen"] = lote.get("nombre_excel")
                item["_fecha_lote"] = lote.get("fecha")
                item["_detalle_path"] = ruta
                resultado.append(item)
        return resultado

    def _mostrar_mensaje_etiquetas(self, texto):
        """Muestra un mensaje de estado (vacío / buscando / cargando) sin
        tocar las filas ya construidas — solo las oculta."""
        for widgets in self._filas_pool_etiquetas:
            widgets["frame"].pack_forget()
        self.lbl_mensaje_etiquetas.configure(text=texto)
        self.lbl_mensaje_etiquetas.pack(pady=30)

    def _mostrar_filas_etiquetas(self, items):
        """Pinta la página actual reutilizando los widgets de fila ya creados
        (en vez de destruir y recrear todo) para evitar el parpadeo/pixelado
        al construir muchos botones de golpe."""
        self.lbl_mensaje_etiquetas.pack_forget()
        for i, item in enumerate(items):
            if i < len(self._filas_pool_etiquetas):
                widgets = self._filas_pool_etiquetas[i]
            else:
                widgets = self._crear_fila_etiqueta_widgets(self.lista_etiquetas_frame)
                self._filas_pool_etiquetas.append(widgets)
            self._actualizar_fila_etiqueta(widgets, item)
            widgets["frame"].pack(fill="x", pady=4)

        for widgets in self._filas_pool_etiquetas[len(items):]:
            widgets["frame"].pack_forget()

    def _refrescar_pagina_etiquetas(self):
        if hasattr(self, "entrada_busqueda"):
            self.entrada_busqueda.delete(0, "end")
        if self._busqueda_after_id:
            self.root.after_cancel(self._busqueda_after_id)
            self._busqueda_after_id = None

        self.modo_busqueda_etiquetas = False
        self.etiquetas_filtradas = []
        self.pagina_etiquetas_actual = 0

        if not self.lotes:
            self._mostrar_mensaje_etiquetas("Aún no has generado ninguna etiqueta.")
            self.lbl_subtitulo_etiquetas.configure(
                text="Genera un lote de etiquetas para poder buscarlas aquí."
            )
            self.lbl_pagina_etiquetas.configure(text="")
            self.btn_pagina_anterior.configure(state="disabled")
            self.btn_pagina_siguiente.configure(state="disabled")
            return

        self.lbl_subtitulo_etiquetas.configure(
            text="Busca por EAN o por norma y descarga el PDF de cada etiqueta."
        )
        self._renderizar_pagina_etiquetas()

    def _renderizar_pagina_etiquetas(self):
        if self.modo_busqueda_etiquetas:
            total = len(self.etiquetas_filtradas)
            mensaje_vacio = "No se encontraron etiquetas para tu búsqueda."
        else:
            total = self._total_etiquetas()
            mensaje_vacio = "Aún no has generado ninguna etiqueta."

        total_paginas = max(1, (total + ETIQUETAS_POR_PAGINA - 1) // ETIQUETAS_POR_PAGINA)
        self.pagina_etiquetas_actual = max(0, min(self.pagina_etiquetas_actual, total_paginas - 1))

        if total == 0:
            self._mostrar_mensaje_etiquetas(mensaje_vacio)
            self.lbl_pagina_etiquetas.configure(text="")
            self.btn_pagina_anterior.configure(state="disabled")
            self.btn_pagina_siguiente.configure(state="disabled")
            return

        inicio = self.pagina_etiquetas_actual * ETIQUETAS_POR_PAGINA
        fin = min(inicio + ETIQUETAS_POR_PAGINA, total)

        self.lbl_pagina_etiquetas.configure(
            text=f"Mostrando {inicio + 1}-{fin} de {total}  ·  Página {self.pagina_etiquetas_actual + 1} de {total_paginas}"
        )
        self.btn_pagina_anterior.configure(state="normal" if self.pagina_etiquetas_actual > 0 else "disabled")
        self.btn_pagina_siguiente.configure(
            state="normal" if self.pagina_etiquetas_actual < total_paginas - 1 else "disabled"
        )

        if self.modo_busqueda_etiquetas:
            self._mostrar_filas_etiquetas(self.etiquetas_filtradas[inicio:fin])
            return

        # Modo historial: la lectura del .jsonl (y la primera indexación del
        # archivo) puede tardar, así que se hace en un hilo aparte para que
        # el cambio de página/pestaña no trabe la interfaz.
        self._peticion_pagina_id += 1
        peticion_id = self._peticion_pagina_id
        self._mostrar_mensaje_etiquetas("Cargando…")
        hilo = threading.Thread(
            target=self._cargar_pagina_historial_en_hilo, args=(inicio, fin, peticion_id), daemon=True
        )
        hilo.start()

    def _cargar_pagina_historial_en_hilo(self, inicio, fin, peticion_id):
        items = self._leer_pagina_historial(inicio, fin)
        self.root.after(0, self._pagina_historial_cargada, items, peticion_id)

    def _pagina_historial_cargada(self, items, peticion_id):
        if peticion_id != self._peticion_pagina_id:
            return  # el usuario ya cambió de página mientras tanto; resultado obsoleto
        self._mostrar_filas_etiquetas(items)

    def _cambiar_pagina_etiquetas(self, delta):
        self.pagina_etiquetas_actual += delta
        self._renderizar_pagina_etiquetas()

    def _crear_fila_etiqueta_widgets(self, master):
        """Crea el esqueleto (vacío) de una fila una sola vez; el contenido
        se rellena después con _actualizar_fila_etiqueta. Reutilizar estos
        widgets entre páginas evita recrear ~7 widgets por fila cada vez
        (que es lo que causaba el parpadeo/pixelado al cambiar de página)."""
        fila = ctk.CTkFrame(
            master, fg_color=STYLE["surface"], corner_radius=8,
            border_width=1, border_color=STYLE["borde"]
        )
        self._configurar_columnas(fila)

        lbl_ean = ctk.CTkLabel(
            fila, text="", font=("Segoe UI", 12, "bold"), text_color=STYLE["texto_oscuro"], anchor="w"
        )
        lbl_ean.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        lbl_marca = ctk.CTkLabel(fila, text="", font=FONT_SMALL, text_color=STYLE["texto_oscuro"], anchor="w")
        lbl_marca.grid(row=0, column=1, sticky="ew", padx=6)
        lbl_norma = ctk.CTkLabel(fila, text="", font=FONT_SMALL, text_color=STYLE["texto_oscuro"], anchor="w")
        lbl_norma.grid(row=0, column=2, sticky="ew", padx=6)
        lbl_estado = ctk.CTkLabel(fila, text="", font=("Segoe UI", 11, "bold"), anchor="w")
        lbl_estado.grid(row=0, column=3, sticky="ew", padx=6)

        btn_preview = ctk.CTkButton(
            fila, text="👁", width=34, height=30, font=FONT_LABEL,
            border_width=1, border_color=STYLE["borde"], corner_radius=6
        )
        btn_preview.grid(row=0, column=4, sticky="e", padx=(6, 0), pady=10)

        btn_descargar = ctk.CTkButton(fila, text="", font=FONT_TINY, height=30, corner_radius=6)
        btn_descargar.grid(row=0, column=5, sticky="e", padx=10, pady=10)

        btn_eliminar = ctk.CTkButton(
            fila, text="🗑", width=34, height=30, font=FONT_LABEL,
            fg_color=STYLE["surface"], hover_color=STYLE["advertencia_suave"],
            text_color=STYLE["advertencia"], border_width=1, border_color=STYLE["borde"],
            corner_radius=6
        )
        btn_eliminar.grid(row=0, column=6, sticky="e", padx=(0, 10), pady=10)

        return {
            "frame": fila, "ean": lbl_ean, "marca": lbl_marca, "norma": lbl_norma,
            "estado": lbl_estado, "preview": btn_preview, "descargar": btn_descargar,
            "eliminar": btn_eliminar,
        }

    def _actualizar_fila_etiqueta(self, widgets, item):
        widgets["ean"].configure(text=item.get("ean") or "—")
        widgets["marca"].configure(text=item.get("marca") or "—")
        widgets["norma"].configure(text=item.get("norma") or "—")

        hay_error = bool(item.get("error"))
        ruta_pdf = item.get("pdf_path")
        tiene_pdf = bool(ruta_pdf) and os.path.exists(ruta_pdf)

        widgets["estado"].configure(
            text="OK" if not hay_error else "Con errores",
            text_color=STYLE["exito"] if not hay_error else STYLE["advertencia"]
        )

        widgets["preview"].configure(
            fg_color=STYLE["surface"] if tiene_pdf else STYLE["borde"],
            hover_color=STYLE["surface_alt"] if tiene_pdf else STYLE["borde"],
            text_color=STYLE["texto_oscuro"] if tiene_pdf else STYLE["texto_secundario"],
            state="normal" if tiene_pdf else "disabled",
            command=(lambda ruta=ruta_pdf: self._previsualizar_pdf(ruta)) if tiene_pdf else None
        )
        widgets["descargar"].configure(
            text="⬇ Descargar PDF" if tiene_pdf else "No disponible",
            fg_color=STYLE["secundario"] if tiene_pdf else STYLE["borde"],
            hover_color=STYLE["secundario_hover"] if tiene_pdf else STYLE["borde"],
            text_color=STYLE["texto_claro"] if tiene_pdf else STYLE["texto_secundario"],
            state="normal" if tiene_pdf else "disabled",
            command=(lambda ruta=ruta_pdf: self._descargar_pdf(ruta)) if tiene_pdf else None
        )
        widgets["eliminar"].configure(command=lambda it=item: self._confirmar_eliminar_etiqueta(it))

    def _confirmar_eliminar_etiqueta(self, item):
        descripcion = item.get("ean") or f"fila {item.get('fila')}"
        norma = item.get("norma") or "—"
        if not messagebox.askyesno(
            "Eliminar etiqueta",
            f"¿Eliminar la etiqueta {descripcion} ({norma})?\n\n"
            "Esto también borrará su PDF si existe. Esta acción no se puede deshacer."
        ):
            return

        if self._eliminar_etiqueta(item):
            self._renderizar_pagina_etiquetas()
        else:
            messagebox.showerror(
                "No se pudo eliminar",
                "La etiqueta ya no se encontró en el historial (puede que haya cambiado mientras tanto)."
            )

    def _eliminar_etiqueta(self, item):
        """Borra una etiqueta puntual: la quita de su .jsonl de lote y
        elimina su PDF si existe. Si el lote se queda sin etiquetas, también
        se elimina del historial."""
        ruta = item.get("_detalle_path")
        fila_numero = item.get("fila")
        if not ruta:
            return False

        lote = next((l for l in self.lotes if l.get("detalle_path") == ruta), None)
        if lote is None:
            return False

        lineas_restantes = []
        eliminado = None
        try:
            with open(ruta, "r", encoding="utf-8") as f:
                for linea in f:
                    linea_limpia = linea.strip()
                    if not linea_limpia:
                        continue
                    data = json.loads(linea_limpia)
                    if eliminado is None and data.get("fila") == fila_numero:
                        eliminado = data
                        continue
                    lineas_restantes.append(linea_limpia)
        except FileNotFoundError:
            return False

        if eliminado is None:
            return False

        if lineas_restantes:
            with open(ruta, "w", encoding="utf-8") as f:
                for linea in lineas_restantes:
                    f.write(linea)
                    f.write("\n")
            lote["total_detalle"] = len(lineas_restantes)
            if not eliminado.get("error"):
                lote["generadas"] = max(0, lote.get("generadas", 0) - 1)
        else:
            try:
                os.remove(ruta)
            except OSError:
                pass
            self.lotes.remove(lote)
            if self.estado_lote is lote:
                self.estado_lote = self.lotes[-1] if self.lotes else None

        self._indices_lote.pop(ruta, None)
        _guardar_manifiesto_lotes(self.lotes)
        self._refrescar_card_ultimo()

        ruta_pdf = eliminado.get("pdf_path")
        if ruta_pdf and os.path.exists(ruta_pdf):
            try:
                os.remove(ruta_pdf)
            except OSError:
                pass

        if self.modo_busqueda_etiquetas:
            self.etiquetas_filtradas = [
                f for f in self.etiquetas_filtradas
                if not (f.get("_detalle_path") == ruta and f.get("fila") == fila_numero)
            ]

        return True

    def _previsualizar_pdf(self, ruta_pdf):
        if not ruta_pdf or not os.path.exists(ruta_pdf):
            messagebox.showwarning(
                "PDF no encontrado",
                "El archivo PDF de esta etiqueta ya no existe en la carpeta original."
            )
            return

        try:
            doc = fitz.open(ruta_pdf)
            pagina = doc.load_page(0)
            pix = pagina.get_pixmap(dpi=200)
            modo = "RGB" if pix.n < 4 else "RGBA"
            imagen = Image.frombytes(modo, (pix.width, pix.height), pix.samples)
            doc.close()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la vista previa:\n{e}")
            return

        max_w, max_h = 520, 640
        escala = min(max_w / imagen.width, max_h / imagen.height, 1.0)
        if escala < 1.0:
            imagen = imagen.resize(
                (max(1, int(imagen.width * escala)), max(1, int(imagen.height * escala))), Image.LANCZOS
            )

        if self._ventana_preview is not None and self._ventana_preview.winfo_exists():
            self._ventana_preview.destroy()

        ventana = ctk.CTkToplevel(self.root)
        self._ventana_preview = ventana
        ventana.title(f"Vista previa · {os.path.basename(ruta_pdf)}")
        ventana.configure(fg_color=STYLE["fondo"])
        ventana.resizable(False, False)
        ventana.transient(self.root)

        ctk_img = ctk.CTkImage(light_image=imagen, size=(imagen.width, imagen.height))
        etiqueta_img = ctk.CTkLabel(ventana, image=ctk_img, text="")
        etiqueta_img.image = ctk_img
        etiqueta_img.pack(padx=20, pady=(20, 10))

        botones = ctk.CTkFrame(ventana, fg_color="transparent")
        botones.pack(fill="x", padx=20, pady=(0, 20))
        ctk.CTkButton(
            botones, text="⬇ Descargar PDF", font=FONT_SMALL, height=34,
            fg_color=STYLE["secundario"], hover_color=STYLE["secundario_hover"],
            text_color=STYLE["texto_claro"], corner_radius=6,
            command=lambda: self._descargar_pdf(ruta_pdf)
        ).pack(side="left")
        ctk.CTkButton(
            botones, text="Cerrar", font=FONT_SMALL, height=34,
            fg_color=STYLE["surface"], hover_color=STYLE["surface_alt"],
            text_color=STYLE["texto_oscuro"], border_width=1, border_color=STYLE["borde"],
            corner_radius=6, command=ventana.destroy
        ).pack(side="right")

        ventana.update_idletasks()
        ventana.grab_set()

    def _descargar_pdf(self, ruta_origen):
        if not ruta_origen or not os.path.exists(ruta_origen):
            messagebox.showwarning(
                "PDF no encontrado",
                "El archivo PDF de esta etiqueta ya no existe en la carpeta original."
            )
            return
        destino = filedialog.asksaveasfilename(
            title="Guardar etiqueta como",
            initialfile=os.path.basename(ruta_origen),
            defaultextension=".pdf",
            filetypes=[("Archivo PDF", "*.pdf")]
        )
        if not destino:
            return
        try:
            shutil.copy(ruta_origen, destino)
            messagebox.showinfo("Descargado", f"Etiqueta guardada en:\n{destino}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _filtrar_etiquetas(self):
        """Se dispara con cada tecla; espera una pausa breve (debounce) antes
        de lanzar la búsqueda real para no escanear el historial en cada golpe
        de tecla."""
        if not self.lotes:
            return
        if self._busqueda_after_id:
            self.root.after_cancel(self._busqueda_after_id)
        consulta = self.entrada_busqueda.get().strip()
        self._busqueda_after_id = self.root.after(350, lambda: self._ejecutar_busqueda(consulta))

    def _ejecutar_busqueda(self, consulta):
        self._busqueda_after_id = None

        if not consulta:
            self.modo_busqueda_etiquetas = False
            self.etiquetas_filtradas = []
            self.pagina_etiquetas_actual = 0
            self._renderizar_pagina_etiquetas()
            return

        self.modo_busqueda_etiquetas = True
        self._mostrar_mensaje_etiquetas("Buscando…")
        self.lbl_pagina_etiquetas.configure(text="")
        self.btn_pagina_anterior.configure(state="disabled")
        self.btn_pagina_siguiente.configure(state="disabled")

        hilo = threading.Thread(target=self._buscar_en_hilo, args=(consulta,), daemon=True)
        hilo.start()

    def _buscar_en_hilo(self, consulta):
        consulta_norm = consulta.upper()
        resultado = []
        for lote in self._lotes_orden_visual():
            ruta = lote.get("detalle_path")
            if not ruta or not os.path.exists(ruta):
                continue
            try:
                with open(ruta, "r", encoding="utf-8") as f:
                    for linea in f:
                        linea = linea.strip()
                        if not linea:
                            continue
                        item = json.loads(linea)
                        if (consulta_norm in (item.get("ean") or "").upper()
                                or consulta_norm in (item.get("norma") or "").upper()):
                            item = dict(item)
                            item["_excel_origen"] = lote.get("nombre_excel")
                            item["_fecha_lote"] = lote.get("fecha")
                            item["_detalle_path"] = ruta
                            resultado.append(item)
            except (FileNotFoundError, json.JSONDecodeError):
                continue
        self.root.after(0, self._busqueda_completada, consulta, resultado)

    def _busqueda_completada(self, consulta, resultado):
        if self.entrada_busqueda.get().strip() != consulta:
            return  # el usuario ya escribió algo más; este resultado quedó obsoleto
        self.etiquetas_filtradas = resultado
        self.pagina_etiquetas_actual = 0
        self._renderizar_pagina_etiquetas()

    # ---------------------------------------------------------------- #
    # Página: Configuración de normas
    # ---------------------------------------------------------------- #
    def _crear_pagina_configuracion(self, master):
        pagina = ctk.CTkFrame(master, fg_color=STYLE["fondo"], corner_radius=0)

        header = ctk.CTkFrame(pagina, fg_color="transparent")
        header.pack(fill="x", padx=30, pady=(26, 10))
        ctk.CTkLabel(
            header, text="⚙️  Configuración de normas", font=FONT_TITLE, text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w")
        ctk.CTkLabel(
            header, text="Administra los campos que lleva cada norma al generar las etiquetas.",
            font=FONT_LABEL, text_color=STYLE["texto_secundario"]
        ).pack(anchor="w", pady=(4, 0))

        cuerpo = ctk.CTkFrame(pagina, fg_color="transparent")
        cuerpo.pack(fill="both", expand=True, padx=30, pady=(10, 24))
        cuerpo.grid_columnconfigure(0, weight=2)
        cuerpo.grid_columnconfigure(1, weight=3)
        cuerpo.grid_rowconfigure(0, weight=1)

        lista_card = ctk.CTkFrame(
            cuerpo, fg_color=STYLE["surface"], corner_radius=14,
            border_width=1, border_color=STYLE["borde"]
        )
        lista_card.grid(row=0, column=0, sticky="nsew", padx=(0, 20))

        lista_header = ctk.CTkFrame(lista_card, fg_color="transparent")
        lista_header.pack(fill="x", padx=16, pady=(16, 8))
        ctk.CTkLabel(
            lista_header, text="Normas configuradas", font=FONT_SUBTITLE, text_color=STYLE["texto_oscuro"]
        ).pack(side="left")
        ctk.CTkButton(
            lista_header, text="+ Nueva", font=FONT_TINY, height=28, width=80,
            fg_color=STYLE["primario"], hover_color=STYLE["primario_hover"],
            text_color=STYLE["texto_oscuro"], corner_radius=6,
            command=self._iniciar_nueva_norma
        ).pack(side="right")

        self.lista_normas_frame = ctk.CTkScrollableFrame(lista_card, fg_color="transparent")
        self.lista_normas_frame.pack(fill="both", expand=True, padx=10, pady=(0, 12))

        self.editor_norma_card = ctk.CTkFrame(
            cuerpo, fg_color=STYLE["surface"], corner_radius=14,
            border_width=1, border_color=STYLE["borde"]
        )
        self.editor_norma_card.grid(row=0, column=1, sticky="nsew")

        return pagina

    def _refrescar_pagina_configuracion(self):
        self._normas_config = configuracion.cargar_config()
        self._norma_seleccionada = None
        self._creando_norma = False
        self._campos_editor = []
        self._orientacion_editor = configuracion.ORIENTACION_DEFECTO
        self._refrescar_lista_normas()
        self._refrescar_editor_norma()

    def _refrescar_lista_normas(self):
        self._limpiar_frame(self.lista_normas_frame)
        normas = configuracion.listar_normas(self._normas_config)

        if not normas:
            ctk.CTkLabel(
                self.lista_normas_frame, text="No hay normas configuradas todavía.",
                font=FONT_SMALL, text_color=STYLE["texto_secundario"],
                wraplength=200, justify="left"
            ).pack(pady=20, padx=10)
            return

        for nombre in normas:
            campos = configuracion.obtener_campos(self._normas_config, nombre)
            seleccionada = nombre == self._norma_seleccionada
            ctk.CTkButton(
                self.lista_normas_frame, text=f"{nombre}\n{len(campos)} campo(s)",
                anchor="w", font=FONT_SMALL, height=48, corner_radius=8,
                fg_color=STYLE["surface_alt"] if seleccionada else "transparent",
                hover_color=STYLE["surface_alt"], text_color=STYLE["texto_oscuro"],
                command=lambda n=nombre: self._seleccionar_norma(n)
            ).pack(fill="x", pady=3)

    def _seleccionar_norma(self, nombre):
        self._norma_seleccionada = nombre
        self._creando_norma = False
        self._campos_editor = configuracion.obtener_campos(self._normas_config, nombre)
        self._orientacion_editor = configuracion.obtener_orientacion(self._normas_config, nombre)
        self._refrescar_lista_normas()
        self._refrescar_editor_norma()

    def _iniciar_nueva_norma(self):
        self._norma_seleccionada = None
        self._creando_norma = True
        self._campos_editor = []
        self._orientacion_editor = configuracion.ORIENTACION_DEFECTO
        self._refrescar_lista_normas()
        self._refrescar_editor_norma()

    def _refrescar_editor_norma(self):
        self._limpiar_frame(self.editor_norma_card)

        if not self._creando_norma and not self._norma_seleccionada:
            ctk.CTkLabel(
                self.editor_norma_card,
                text="Selecciona una norma de la lista para editar sus campos,\n"
                     "o crea una nueva con \"+ Nueva\".",
                font=FONT_LABEL, text_color=STYLE["texto_secundario"], justify="left"
            ).pack(padx=20, pady=40)
            return

        contenido = ctk.CTkFrame(self.editor_norma_card, fg_color="transparent")
        contenido.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(
            contenido, text="Nombre de la norma", font=FONT_SMALL, text_color=STYLE["texto_secundario"]
        ).pack(anchor="w")
        if self._creando_norma:
            self.entrada_nombre_norma = ctk.CTkEntry(
                contenido, placeholder_text="Ej. NOM-004-SE-2021", font=FONT_LABEL, height=36
            )
            self.entrada_nombre_norma.pack(fill="x", pady=(4, 16))
        else:
            ctk.CTkLabel(
                contenido, text=self._norma_seleccionada, font=("Segoe UI", 15, "bold"),
                text_color=STYLE["texto_oscuro"], anchor="w"
            ).pack(fill="x", pady=(4, 16))

        ctk.CTkLabel(
            contenido, text="Orientación de impresión", font=FONT_SMALL,
            text_color=STYLE["texto_secundario"]
        ).pack(anchor="w")
        self.segmented_orientacion = ctk.CTkSegmentedButton(
            contenido, values=["Vertical", "Horizontal"],
            font=FONT_LABEL, height=34,
            fg_color=STYLE["fondo"], selected_color=STYLE["primario"],
            selected_hover_color=STYLE["primario_hover"], unselected_color=STYLE["fondo"],
            unselected_hover_color=STYLE["surface_alt"], text_color=STYLE["texto_oscuro"],
            command=self._cambiar_orientacion_editor
        )
        self.segmented_orientacion.set(
            "Horizontal" if self._orientacion_editor == "horizontal" else "Vertical"
        )
        self.segmented_orientacion.pack(fill="x", pady=(4, 16))

        # Botones y fila de "agregar campo" se anclan abajo (side="bottom")
        # para que sigan visibles aunque la lista de campos tenga muchas
        # filas; la lista de campos, en medio, es la que scrollea.
        botones = ctk.CTkFrame(contenido, fg_color="transparent")
        botones.pack(side="bottom", fill="x")
        ctk.CTkButton(
            botones, text="💾  Guardar cambios", font=FONT_SUBTITLE, height=42,
            fg_color=STYLE["primario"], hover_color=STYLE["primario_hover"],
            text_color=STYLE["texto_oscuro"], corner_radius=10,
            command=self._guardar_norma_editor
        ).pack(side="left")

        if self._creando_norma:
            ctk.CTkButton(
                botones, text="Cancelar", font=FONT_SMALL, height=42, width=110,
                fg_color=STYLE["surface"], hover_color=STYLE["surface_alt"],
                text_color=STYLE["texto_oscuro"], border_width=1, border_color=STYLE["borde"],
                corner_radius=8, command=self._refrescar_pagina_configuracion
            ).pack(side="left", padx=(10, 0))
        else:
            ctk.CTkButton(
                botones, text="🗑  Eliminar norma", font=FONT_SMALL, height=42,
                fg_color=STYLE["surface"], hover_color=STYLE["advertencia_suave"],
                text_color=STYLE["advertencia"], border_width=1, border_color=STYLE["borde"],
                corner_radius=8, command=self._eliminar_norma_editor
            ).pack(side="right")

        agregar_fila = ctk.CTkFrame(contenido, fg_color="transparent")
        agregar_fila.pack(side="bottom", fill="x", pady=(0, 20))
        self.entrada_nuevo_campo = ctk.CTkEntry(
            agregar_fila, placeholder_text="Nombre del campo (ej. TALLA)", font=FONT_LABEL, height=34
        )
        self.entrada_nuevo_campo.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.entrada_nuevo_campo.bind("<Return>", lambda e: self._agregar_campo_editor())
        ctk.CTkButton(
            agregar_fila, text="+ Agregar campo", font=FONT_SMALL, height=34, width=140,
            fg_color=STYLE["secundario"], hover_color=STYLE["secundario_hover"],
            text_color=STYLE["texto_claro"], corner_radius=6,
            command=self._agregar_campo_editor
        ).pack(side="right")

        ctk.CTkLabel(
            contenido, text="Campos que lleva esta etiqueta", font=FONT_SMALL,
            text_color=STYLE["texto_secundario"]
        ).pack(anchor="w")

        self.lista_campos_frame = ctk.CTkScrollableFrame(contenido, fg_color="transparent")
        self.lista_campos_frame.pack(fill="both", expand=True, pady=(6, 10))
        self._renderizar_campos_editor()

    def _renderizar_campos_editor(self):
        self._limpiar_frame(self.lista_campos_frame)
        if not self._campos_editor:
            ctk.CTkLabel(
                self.lista_campos_frame, text="Esta norma todavía no tiene campos.",
                font=FONT_TINY, text_color=STYLE["texto_secundario"]
            ).pack(anchor="w", pady=4)
            return

        for campo in self._campos_editor:
            fila = ctk.CTkFrame(self.lista_campos_frame, fg_color=STYLE["fondo"], corner_radius=6)
            fila.pack(fill="x", pady=2)
            ctk.CTkLabel(
                fila, text=campo, font=FONT_SMALL, text_color=STYLE["texto_oscuro"], anchor="w"
            ).pack(side="left", padx=10, pady=6)
            ctk.CTkButton(
                fila, text="✕", width=26, height=26, font=FONT_TINY,
                fg_color="transparent", hover_color=STYLE["advertencia_suave"],
                text_color=STYLE["texto_secundario"], corner_radius=6,
                command=lambda c=campo: self._quitar_campo_editor(c)
            ).pack(side="right", padx=6, pady=4)

    def _agregar_campo_editor(self):
        campo = self.entrada_nuevo_campo.get().strip().upper()
        if not campo:
            return
        if campo in self._campos_editor:
            messagebox.showwarning("Campo repetido", f"El campo '{campo}' ya está en la lista.")
            return
        self._campos_editor.append(campo)
        self.entrada_nuevo_campo.delete(0, "end")
        self._renderizar_campos_editor()

    def _quitar_campo_editor(self, campo):
        if campo in self._campos_editor:
            self._campos_editor.remove(campo)
        self._renderizar_campos_editor()

    def _cambiar_orientacion_editor(self, valor):
        self._orientacion_editor = "horizontal" if valor == "Horizontal" else "vertical"

    def _guardar_norma_editor(self):
        if self._creando_norma:
            nombre = self.entrada_nombre_norma.get().strip()
            error = configuracion.validar_nombre_norma(nombre, self._normas_config)
            if error:
                messagebox.showerror("Nombre inválido", error)
                return
            if not self._campos_editor:
                messagebox.showwarning("Sin campos", "Agrega al menos un campo antes de guardar.")
                return
            configuracion.agregar_norma(
                self._normas_config, nombre, self._campos_editor, self._orientacion_editor
            )
            norma_guardada = nombre
        else:
            if not self._campos_editor:
                messagebox.showwarning("Sin campos", "Una norma debe tener al menos un campo.")
                return
            configuracion.actualizar_campos_norma(
                self._normas_config, self._norma_seleccionada, self._campos_editor
            )
            configuracion.actualizar_orientacion_norma(
                self._normas_config, self._norma_seleccionada, self._orientacion_editor
            )
            norma_guardada = self._norma_seleccionada

        configuracion.guardar_config(self._normas_config)
        messagebox.showinfo("Guardado", f"La norma '{norma_guardada}' se guardó correctamente.")

        self._norma_seleccionada = norma_guardada
        self._creando_norma = False
        self._refrescar_lista_normas()
        self._refrescar_editor_norma()

    def _eliminar_norma_editor(self):
        if not self._norma_seleccionada:
            return
        if not messagebox.askyesno(
            "Eliminar norma",
            f"¿Eliminar la norma '{self._norma_seleccionada}'?\n\n"
            "Las etiquetas ya generadas con esta norma no se ven afectadas, pero ya no "
            "podrás generar nuevas etiquetas con ella hasta que la vuelvas a crear."
        ):
            return
        configuracion.eliminar_norma(self._normas_config, self._norma_seleccionada)
        configuracion.guardar_config(self._normas_config)
        self._refrescar_pagina_configuracion()


if __name__ == "__main__":
    GenerdorEtiquetas()
