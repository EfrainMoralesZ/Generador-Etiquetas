# -- SISTEMA V&C - GENERADOR DE ETIQUETAS -- #
import os
import json
import shutil
import threading
from datetime import datetime

import customtkinter as ctk
from tkinter import filedialog, messagebox

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:
    DND_FILES = None
    TkinterDnD = None

from armadoEtiqueta import generar_etiquetas_desde_excel, previsualizar_etiquetas_desde_excel

APP_VERSION = "1.0.0"
ESTADO_PATH = os.path.join("data", "estado_app.json")

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

COLUMNAS_ETIQUETAS = [("EAN", 2), ("Marca", 2), ("Norma", 3), ("Estado", 2), ("", 2)]

ctk.set_appearance_mode("light")


def _formato_tamano(num_bytes):
    tamano = float(num_bytes)
    for unidad in ("B", "KB", "MB", "GB"):
        if tamano < 1024 or unidad == "GB":
            return f"{tamano:.0f} {unidad}" if unidad == "B" else f"{tamano:.1f} {unidad}"
        tamano /= 1024
    return f"{tamano:.1f} GB"


def _cargar_estado():
    if not os.path.exists(ESTADO_PATH):
        return None
    try:
        with open(ESTADO_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _guardar_estado(estado):
    os.makedirs(os.path.dirname(ESTADO_PATH), exist_ok=True)
    with open(ESTADO_PATH, "w", encoding="utf-8") as f:
        json.dump(estado, f, ensure_ascii=False, indent=2)


class GenerdorEtiquetas:

    PASOS = ["Archivo cargado", "Analizando datos", "Generando PDF", "Finalizado"]

    def __init__(self):
        self.excel_path = None
        self.resultado_analisis = None
        self.estado_lote = _cargar_estado()
        self.estado_pasos = ["pendiente"] * 4
        self.dnd_activo = False
        self.filas_etiquetas = []

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
        self.pagina_info = self._crear_pagina_info(self.contenedor_paginas)

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
            ("info", "ℹ️", "Información"),
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
            "info": self.pagina_info,
        }
        for pagina in paginas.values():
            pagina.pack_forget()
        paginas[clave].pack(fill="both", expand=True)

        for c, btn in self.btn_nav.items():
            btn.configure(fg_color=STYLE["surface_alt"] if c == clave else "transparent")

        if clave == "etiquetas":
            self._refrescar_pagina_etiquetas()

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

    def _cargar_archivo(self, ruta):
        self.excel_path = ruta
        self.resultado_analisis = None
        self._render_dropzone_archivo(ruta)
        self.btn_generar.configure(state="disabled")

        self.estado_pasos = ["completado", "progreso", "pendiente", "pendiente"]
        self._actualizar_stepper()
        self.progress.set(0)
        self.lbl_progreso.configure(text="")
        self._limpiar_frame(self.banner_resultado)

        self._agregar_actividad("📥", "Archivo cargado", os.path.basename(ruta))
        self._agregar_actividad("🔍", "Analizando datos", "Validando información y aplicando reglas")

        hilo = threading.Thread(target=self._analizar_en_hilo, args=(ruta,), daemon=True)
        hilo.start()

    def _quitar_archivo(self):
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
        self.estado_pasos[1] = "completado"
        self._actualizar_stepper()

        if hasattr(self, "lbl_info_archivo"):
            self.lbl_info_archivo.configure(
                text=f"{resultado['listas']} de {resultado['total_filas']} filas listas"
            )

        if resultado["listas"] > 0:
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

        self.estado_lote = {
            "nombre_excel": os.path.basename(self.excel_path),
            "fecha": datetime.now().strftime("%d/%m/%Y · %H:%M"),
            "output_dir": resultado["output_dir"],
            "json_path": resultado["json_path"],
            "detalle": resultado["detalle"],
        }
        _guardar_estado(self.estado_lote)
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
        self.lista_etiquetas_frame.pack(fill="both", expand=True, padx=24, pady=(0, 24))

        return pagina

    @staticmethod
    def _configurar_columnas(frame):
        for i, (_, peso) in enumerate(COLUMNAS_ETIQUETAS):
            frame.grid_columnconfigure(i, weight=peso)

    def _refrescar_pagina_etiquetas(self):
        self._limpiar_frame(self.lista_etiquetas_frame)
        self.filas_etiquetas = []
        if hasattr(self, "entrada_busqueda"):
            self.entrada_busqueda.delete(0, "end")

        if not self.estado_lote or not self.estado_lote.get("detalle"):
            ctk.CTkLabel(
                self.lista_etiquetas_frame, text="Aún no has generado ninguna etiqueta.",
                font=FONT_LABEL, text_color=STYLE["texto_secundario"]
            ).pack(pady=30)
            self.lbl_subtitulo_etiquetas.configure(
                text="Genera un lote de etiquetas para poder buscarlas aquí."
            )
            return

        self.lbl_subtitulo_etiquetas.configure(
            text="Busca por EAN o por norma y descarga el PDF de cada etiqueta."
        )

        for item in self.estado_lote["detalle"]:
            fila = self._crear_fila_etiqueta(self.lista_etiquetas_frame, item)
            self.filas_etiquetas.append((fila, item))

        self._filtrar_etiquetas()

    def _crear_fila_etiqueta(self, master, item):
        fila = ctk.CTkFrame(
            master, fg_color=STYLE["surface"], corner_radius=8,
            border_width=1, border_color=STYLE["borde"]
        )
        self._configurar_columnas(fila)

        ctk.CTkLabel(
            fila, text=item.get("ean") or "—", font=("Segoe UI", 12, "bold"),
            text_color=STYLE["texto_oscuro"], anchor="w"
        ).grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        ctk.CTkLabel(
            fila, text=item.get("marca") or "—", font=FONT_SMALL,
            text_color=STYLE["texto_oscuro"], anchor="w"
        ).grid(row=0, column=1, sticky="ew", padx=6)
        ctk.CTkLabel(
            fila, text=item.get("norma") or "—", font=FONT_SMALL,
            text_color=STYLE["texto_oscuro"], anchor="w"
        ).grid(row=0, column=2, sticky="ew", padx=6)

        hay_error = bool(item.get("error"))
        ruta_pdf = item.get("pdf_path")
        tiene_pdf = bool(ruta_pdf) and os.path.exists(ruta_pdf)

        estado_texto = "OK" if not hay_error else "Con errores"
        estado_color = STYLE["exito"] if not hay_error else STYLE["advertencia"]
        ctk.CTkLabel(
            fila, text=estado_texto, font=("Segoe UI", 11, "bold"),
            text_color=estado_color, anchor="w"
        ).grid(row=0, column=3, sticky="ew", padx=6)

        ctk.CTkButton(
            fila, text="⬇ Descargar PDF" if tiene_pdf else "No disponible",
            font=FONT_TINY, height=30,
            fg_color=STYLE["secundario"] if tiene_pdf else STYLE["borde"],
            hover_color=STYLE["secundario_hover"] if tiene_pdf else STYLE["borde"],
            text_color=STYLE["texto_claro"] if tiene_pdf else STYLE["texto_secundario"],
            corner_radius=6, state="normal" if tiene_pdf else "disabled",
            command=(lambda ruta=ruta_pdf: self._descargar_pdf(ruta)) if tiene_pdf else None
        ).grid(row=0, column=4, sticky="e", padx=10, pady=10)

        return fila

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
        consulta = self.entrada_busqueda.get().strip().upper()
        for fila, _ in self.filas_etiquetas:
            fila.pack_forget()
        for fila, item in self.filas_etiquetas:
            coincide = (
                not consulta
                or consulta in (item.get("ean") or "").upper()
                or consulta in (item.get("norma") or "").upper()
            )
            if coincide:
                fila.pack(fill="x", pady=4)

    # ---------------------------------------------------------------- #
    # Página: Información
    # ---------------------------------------------------------------- #
    def _crear_pagina_info(self, master):
        pagina = ctk.CTkFrame(master, fg_color=STYLE["fondo"], corner_radius=0)
        contenedor = ctk.CTkFrame(
            pagina, fg_color=STYLE["surface"], corner_radius=14,
            border_width=1, border_color=STYLE["borde"]
        )
        contenedor.pack(fill="x", padx=30, pady=30)

        ctk.CTkLabel(
            contenedor, text="ℹ️  Información", font=FONT_TITLE, text_color=STYLE["texto_oscuro"]
        ).pack(anchor="w", padx=24, pady=(24, 8))

        texto = (
            "Generador de Etiquetas convierte un archivo Excel en etiquetas PDF conforme a las "
            "Normas Oficiales Mexicanas (NOM) configuradas en data/config_etiquetas.json.\n\n"
            "1. Sube o arrastra tu archivo Excel en la pestaña 'Generador'.\n"
            "2. El sistema valida cada fila y detecta la norma según la columna 'CODIGO FORMATO'.\n"
            "3. Al generar, se crea un PDF individual por cada etiqueta.\n"
            "4. En 'Etiquetas generadas' puedes buscar por EAN o norma y descargar el PDF de "
            "cualquier etiqueta de forma individual."
        )
        ctk.CTkLabel(
            contenedor, text=texto, font=FONT_LABEL, text_color=STYLE["texto_secundario"],
            justify="left", anchor="w", wraplength=640
        ).pack(fill="x", padx=24, pady=(0, 24))

        return pagina


if __name__ == "__main__":
    GenerdorEtiquetas()
