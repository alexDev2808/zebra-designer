"""Etiquetas Zebra ZT411 (300 / 600 dpi) desde Excel.

Ejecutar:  python app.py
"""
from __future__ import annotations

import ctypes
import json
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageDraw, ImageTk

import excel_data
import preview
import printing
import theme
import zpl
from theme import C
from zpl import Element

APP_TITLE = "Etiquetas Zebra ZT411"
SETTINGS_PATH = Path.home() / ".zebra_etiquetas.json"
NET_OPTION = "Impresora de red (IP / puerto 9100)…"
PRESETS = {
    "Personalizado": None,
    "4 × 6 pulg (101.6 × 152.4 mm)": (101.6, 152.4),
    "4 × 3 pulg (101.6 × 76.2 mm)": (101.6, 76.2),
    "4 × 2 pulg (101.6 × 50.8 mm)": (101.6, 50.8),
    "4 × 1 pulg (101.6 × 25.4 mm)": (101.6, 25.4),
    "3 × 2 pulg (76.2 × 50.8 mm)": (76.2, 50.8),
    "2 × 1 pulg (50.8 × 25.4 mm)": (50.8, 25.4),
    "100 × 150 mm": (100.0, 150.0),
    "100 × 50 mm": (100.0, 50.0),
    "50 × 25 mm": (50.0, 25.0),
}


def _num(var: tk.Variable, cast=float):
    """Lee un número aceptando coma decimal; lanza ValueError si no es válido."""
    return cast(float(str(var.get()).strip().replace(",", ".")))


def _fmt(v) -> str:
    return f"{v:g}" if isinstance(v, float) else str(v)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.f = max(1.0, self.winfo_fpixels("1i") / 96)  # escala de pantalla (125 %, 150 %…)
        self.handle = int(5 * self.f)                     # radio de las manijas de redimensionado
        self.minsize(int(1100 * self.f), int(680 * self.f))
        self.geometry(f"{int(1400 * self.f)}x{int(860 * self.f)}")
        self.state("zoomed")
        theme.apply(self, self.f)
        self._set_icon()

        self.settings = self._load_settings()
        self.template = zpl.default_template()
        self.template_path: str | None = None
        self.excel_path: str | None = None
        self.headers: list[str] = []
        self.rows: list[dict] = []
        self.visible: list[int] = []
        self.current_row = 0
        self.sel: int | None = None

        self._loading = False
        self._drag = None
        self._render_job = None
        self._scale = 1.0
        self._origin = (0, 0)
        self._boxes: list[tuple[int, int, int, int]] = []
        self._handles: dict[str, tuple[int, int]] = {}
        self._photo = None

        self._build_ui()
        self._restore_session()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _set_icon(self):
        img = Image.new("RGBA", (32, 32), C["primary"])
        d = ImageDraw.Draw(img)
        x = 6
        for w in (2, 1, 3, 1, 2, 1, 1, 3):
            d.rectangle([x, 7, x + w - 1, 21], fill="white")
            x += w + 1
        d.rectangle([6, 24, 26, 25], fill=C["accent"])
        self._icon = ImageTk.PhotoImage(img)
        self.iconphoto(True, self._icon)

    # ================================================================== UI
    def _build_ui(self):
        self._scrollers: list[tk.Canvas] = []
        self.unbind_class("TCombobox", "<MouseWheel>")  # la rueda desplaza el panel, no cambia valores
        self.bind_all("<MouseWheel>", self._on_wheel)
        self._build_header()
        self._build_toolbar()
        self._build_statusbar()

        main = ttk.PanedWindow(self, orient="horizontal")
        main.pack(fill="both", expand=True, padx=10, pady=(10, 8))

        left = ttk.Frame(main, style="Card.TFrame", width=int(450 * self.f))
        left.pack_propagate(False)
        main.add(left, weight=0)
        self.nb = ttk.Notebook(left)
        self.nb.pack(fill="both", expand=True)
        self._build_elements_tab()
        self._build_label_tab()
        self._build_printer_tab()

        right = ttk.PanedWindow(main, orient="vertical")
        main.add(right, weight=1)
        self.right_pane = right
        self._build_preview(right)
        self._build_table(right)

    def _build_header(self):
        h = ttk.Frame(self, style="Header.TFrame", padding=(16, 10))
        h.pack(fill="x")
        theme.draw_logo(h, C["navy"], self.f).pack(side="left", padx=(0, 12))
        titles = ttk.Frame(h, style="Header.TFrame")
        titles.pack(side="left")
        ttk.Label(titles, text="Etiquetas Zebra", style="Header.TLabel").pack(anchor="w")
        ttk.Label(titles, text="Diseño e impresión de etiquetas desde Excel · ZT411 300 / 600 dpi",
                  style="HeaderSub.TLabel").pack(anchor="w")

        self.v_badge = tk.StringVar(value="300 dpi")
        ttk.Label(h, textvariable=self.v_badge, style="Badge.TLabel").pack(side="right", padx=(10, 0))
        ttk.Button(h, text="↻", style="Header.TButton", width=3,
                   command=self._refresh_printers).pack(side="right", padx=(6, 0))
        self.v_dest = tk.StringVar()
        self.dest_cb = ttk.Combobox(h, textvariable=self.v_dest, state="readonly", width=46,
                                    style="Header.TCombobox", font=theme.F_BASE)
        self.dest_cb.pack(side="right")
        self.dest_cb.bind("<<ComboboxSelected>>", lambda e: self._on_dest_select())
        ttk.Label(h, text="IMPRESORA", style="HeaderSub.TLabel").pack(side="right", padx=(0, 8))

    def _build_toolbar(self):
        bar = ttk.Frame(self, style="Toolbar.TFrame", padding=(14, 8))
        bar.pack(fill="x")
        ttk.Separator(self).pack(fill="x")

        def group(title):
            g = ttk.Frame(bar, style="Toolbar.TFrame")
            g.pack(side="left", padx=(0, 18))
            ttk.Label(g, text=title, style="Group.TLabel").pack(anchor="w")
            row = ttk.Frame(g, style="Toolbar.TFrame")
            row.pack(anchor="w", pady=(2, 0))
            return row

        g = group("DATOS")
        ttk.Button(g, text="Abrir Excel…", style="Tool.TButton", command=self.open_excel).pack(side="left")
        ttk.Label(g, text="  Hoja", style="Card.TLabel").pack(side="left")
        self.sheet_cb = ttk.Combobox(g, state="readonly", width=16)
        self.sheet_cb.pack(side="left", padx=(6, 0))
        self.sheet_cb.bind("<<ComboboxSelected>>", lambda e: self.load_sheet(self.sheet_cb.get()))

        g = group("PLANTILLA")
        for text, cmd in (("Nueva", self.new_template), ("Abrir…", self.open_template),
                          ("Guardar", self.save_template), ("Guardar como…", lambda: self.save_template(ask=True))):
            ttk.Button(g, text=text, style="Tool.TButton", command=cmd).pack(side="left", padx=(0, 4))

        g = group("ZPL")
        ttk.Button(g, text="Ver código", style="Tool.TButton", command=self.show_zpl).pack(side="left", padx=(0, 4))
        ttk.Button(g, text="Exportar…", style="Tool.TButton", command=self.export_zpl).pack(side="left")

        acts = ttk.Frame(bar, style="Toolbar.TFrame")
        acts.pack(side="right", anchor="s")
        ttk.Button(acts, text="Imprimir todos", style="Primary.TButton",
                   command=lambda: self.print_rows("all")).pack(side="left", padx=(0, 6))
        ttk.Button(acts, text="Imprimir seleccionados", style="Accent.TButton",
                   command=lambda: self.print_rows("sel")).pack(side="left")

    def _build_statusbar(self):
        sb = ttk.Frame(self, style="Status.TFrame")
        sb.pack(fill="x", side="bottom")
        self.status = tk.StringVar(value="Abra un archivo de Excel para comenzar.")
        self.summary = tk.StringVar()
        ttk.Label(sb, textvariable=self.status, style="Status.TLabel").pack(side="left")
        ttk.Label(sb, textvariable=self.summary, style="StatusMuted.TLabel").pack(side="right")

    @staticmethod
    def _row(parent, r, text, widget, hint=None):
        ttk.Label(parent, text=text, style="Card.TLabel").grid(row=r, column=0, sticky="w", pady=4, padx=(0, 10))
        widget.grid(row=r, column=1, sticky="ew", pady=4)
        if hint:
            ttk.Label(parent, text=hint, style="Muted.TLabel").grid(row=r, column=2, sticky="w", padx=(8, 0))

    @staticmethod
    def _section(parent, r, text, top=14):
        ttk.Label(parent, text=text, style="Section.TLabel").grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(top, 4))

    def _tab(self, title):
        """Pestaña con desplazamiento vertical (el contenido puede ser más alto que la ventana)."""
        outer = ttk.Frame(self.nb, style="Card.TFrame")
        self.nb.add(outer, text=title)
        canvas = tk.Canvas(outer, bg=C["surface"], highlightthickness=0, bd=0, yscrollincrement=20)
        sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        f = ttk.Frame(canvas, style="Card.TFrame", padding=(16, 6, 16, 16))
        f.columnconfigure(1, weight=1)
        win = canvas.create_window(0, 0, window=f, anchor="nw")

        def relayout(_e=None):
            canvas.configure(scrollregion=(0, 0, f.winfo_reqwidth(), f.winfo_reqheight()))
            if f.winfo_reqheight() > canvas.winfo_height() > 1:
                sb.pack(side="right", fill="y")
            else:
                sb.pack_forget()
                canvas.yview_moveto(0)
        f.bind("<Configure>", relayout)
        canvas.bind("<Configure>", lambda e: (canvas.itemconfigure(win, width=e.width), relayout()))
        self._scrollers.append(canvas)
        return f

    def _on_wheel(self, e):
        w = self.winfo_containing(e.x_root, e.y_root)
        if w is None or isinstance(w, tk.Listbox):
            return
        for c in self._scrollers:
            if str(w).startswith(str(c)):
                top, bottom = c.yview()
                if top > 0 or bottom < 1:
                    c.yview_scroll(int(-e.delta / 120) * 2, "units")
                return

    # ------------------------------------------------------------ Elementos
    def _build_elements_tab(self):
        f = self._tab("Elementos")
        self._section(f, 0, "Elementos de la etiqueta", top=6)
        self.el_list = tk.Listbox(f, height=7, exportselection=False)
        theme.style_listbox(self.el_list)
        self.el_list.grid(row=1, column=0, columnspan=3, sticky="ew")
        self.el_list.bind("<<ListboxSelect>>", self._on_list_select)
        b = ttk.Frame(f, style="Card.TFrame")
        b.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        ttk.Button(b, text="+ Texto", style="Primary.TButton",
                   command=lambda: self.add_element("texto")).pack(side="left")
        ttk.Button(b, text="+ Código", style="Primary.TButton",
                   command=lambda: self.add_element("codigo")).pack(side="left", padx=4)
        ttk.Button(b, text="↓", style="Icon.TButton", command=lambda: self.move_element(1)).pack(side="right")
        ttk.Button(b, text="↑", style="Icon.TButton", command=lambda: self.move_element(-1)).pack(side="right", padx=2)
        ttk.Button(b, text="Eliminar", command=self.delete_element).pack(side="right", padx=2)
        ttk.Button(b, text="Duplicar", command=self.duplicate_element).pack(side="right")

        ttk.Separator(f).grid(row=3, column=0, columnspan=3, sticky="ew", pady=(14, 0))
        self.props_title = ttk.Label(f, text="Propiedades", style="Section.TLabel")
        self.props_title.grid(row=4, column=0, columnspan=3, sticky="w", pady=(10, 4))

        p = ttk.Frame(f, style="Card.TFrame")
        p.grid(row=5, column=0, columnspan=3, sticky="nsew")
        p.columnconfigure(1, weight=1)
        self.v_content, self.v_x, self.v_y, self.v_rot = (tk.StringVar() for _ in range(4))
        self.content_entry = ttk.Entry(p, textvariable=self.v_content)
        self._row(p, 0, "Contenido", self.content_entry)
        ins = ttk.Frame(p, style="Card.TFrame")
        self.col_cb = ttk.Combobox(ins, state="readonly", width=16)
        self.col_cb.pack(side="left", fill="x", expand=True)
        ttk.Button(ins, text="Insertar columna", command=self._insert_column).pack(side="left", padx=(6, 0))
        self._row(p, 1, "", ins)
        ttk.Label(p, style="Muted.TLabel", wraplength=360, justify="left",
                  text="Use {Columna} para tomar el valor del Excel; puede combinarlo con texto fijo, "
                       "p. ej.  Lote: {Lote}").grid(row=2, column=0, columnspan=3, sticky="w", pady=(0, 4))
        pos = ttk.Frame(p, style="Card.TFrame")
        ttk.Entry(pos, textvariable=self.v_x, width=7).pack(side="left")
        ttk.Label(pos, text="  Y", style="Card.TLabel").pack(side="left")
        ttk.Entry(pos, textvariable=self.v_y, width=7).pack(side="left", padx=(6, 0))
        ttk.Label(pos, text="  mm", style="Muted.TLabel").pack(side="left")
        self._row(p, 3, "Posición  X", pos)
        self._row(p, 4, "Rotación", ttk.Combobox(p, textvariable=self.v_rot, state="readonly",
                                                 width=8, values=list(zpl.ROTATIONS)))

        # --- Texto
        self.text_frame = tf = ttk.Frame(p, style="Card.TFrame")
        tf.columnconfigure(1, weight=1)
        self.v_font, self.v_bw, self.v_lines, self.v_align = (tk.StringVar() for _ in range(4))
        self._row(tf, 0, "Tamaño letra (mm)", ttk.Spinbox(tf, textvariable=self.v_font, from_=1, to=50,
                                                          increment=0.5, width=8))
        self._row(tf, 1, "Ancho bloque (mm)", ttk.Spinbox(tf, textvariable=self.v_bw, from_=0, to=300,
                                                          increment=1, width=8), "0 = sin ajuste")
        self._row(tf, 2, "Máx. líneas", ttk.Spinbox(tf, textvariable=self.v_lines, from_=1, to=20, width=8))
        self._row(tf, 3, "Alineación", ttk.Combobox(tf, textvariable=self.v_align, state="readonly",
                                                    width=10, values=list(zpl.ALIGNS)))

        # --- Código de barras
        self.bc_frame = bf = ttk.Frame(p, style="Card.TFrame")
        bf.columnconfigure(1, weight=1)
        self.v_btype, self.v_bh, self.v_mod, self.v_fixw = (tk.StringVar() for _ in range(4))
        self.v_wmode = tk.StringVar(value="module")
        self.v_mod_scale = tk.DoubleVar(value=0.33)
        self.v_show = tk.BooleanVar()
        self._row(bf, 0, "Tipo de código", ttk.Combobox(bf, textvariable=self.v_btype, state="readonly",
                                                        width=12, values=zpl.BARCODE_TYPES))
        self._row(bf, 1, "Alto (mm)", ttk.Spinbox(bf, textvariable=self.v_bh, from_=1, to=200,
                                                  increment=0.5, width=8), "no aplica a QR")
        ttk.Label(bf, text="ANCHO DEL CÓDIGO", style="Group.TLabel").grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(10, 2))
        ttk.Radiobutton(bf, text="Por grosor de barra", variable=self.v_wmode, value="module",
                        command=self._on_wmode).grid(row=3, column=0, columnspan=3, sticky="w")
        mod = ttk.Frame(bf, style="Card.TFrame")
        self.mod_entry = ttk.Entry(mod, textvariable=self.v_mod, width=7)
        self.mod_entry.pack(side="left")
        self.mod_scale = ttk.Scale(mod, from_=0.1, to=1.0, variable=self.v_mod_scale,
                                   command=lambda v: self.v_mod.set(f"{float(v):.2f}"))
        self.mod_scale.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self._row(bf, 4, "   Barra fina (mm)", mod)
        ttk.Radiobutton(bf, text="Ancho fijo — el código se ajusta a esta medida", variable=self.v_wmode,
                        value="fixed", command=self._on_wmode).grid(row=5, column=0, columnspan=3, sticky="w",
                                                                     pady=(4, 0))
        self.fixw_spin = ttk.Spinbox(bf, textvariable=self.v_fixw, from_=5, to=300, increment=1, width=8)
        self._row(bf, 6, "   Ancho (mm)", self.fixw_spin)
        self.bc_info = ttk.Label(bf, style="Value.TLabel", justify="left")
        self.bc_info.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(6, 4))
        ttk.Checkbutton(bf, text="Mostrar texto legible debajo", variable=self.v_show).grid(
            row=8, column=0, columnspan=3, sticky="w", pady=2)
        ttk.Label(bf, style="Muted.TLabel", wraplength=360, justify="left", text=(
            "Tip: arrastre las manijas del código en la vista previa para cambiar ancho y alto. "
            "Con ancho fijo, cada registro usa el grosor de barra más grande que cabe en esa medida."
        )).grid(row=9, column=0, columnspan=3, sticky="w", pady=(4, 0))

        for v in (self.v_content, self.v_x, self.v_y, self.v_rot, self.v_font, self.v_bw, self.v_lines,
                  self.v_align, self.v_btype, self.v_bh, self.v_mod, self.v_fixw, self.v_show):
            v.trace_add("write", lambda *a: self._on_prop_change())

    # ------------------------------------------------------------- Etiqueta
    def _build_label_tab(self):
        f = self._tab("Etiqueta")
        self.v_width, self.v_height = tk.StringVar(), tk.StringVar()
        self.v_dpi, self.v_dark, self.v_speed = tk.StringVar(), tk.StringVar(), tk.StringVar()
        self.v_offx, self.v_offy = tk.StringVar(), tk.StringVar()
        self.v_copcol, self.v_copies = tk.StringVar(), tk.StringVar()
        self.v_preset = tk.StringVar(value="Personalizado")

        self._section(f, 0, "Tamaño", top=6)
        preset = ttk.Combobox(f, textvariable=self.v_preset, state="readonly", values=list(PRESETS))
        preset.bind("<<ComboboxSelected>>", lambda e: self._apply_preset())
        self._row(f, 1, "Predefinido", preset)
        self._row(f, 2, "Ancho (mm)", ttk.Spinbox(f, textvariable=self.v_width, from_=10, to=300, increment=1, width=10))
        self._row(f, 3, "Alto (mm)", ttk.Spinbox(f, textvariable=self.v_height, from_=5, to=1000, increment=1, width=10))
        self._row(f, 4, "Resolución", ttk.Combobox(f, textvariable=self.v_dpi, state="readonly", width=8,
                                                   values=[str(d) for d in zpl.DPI_OPTIONS]), "dpi")

        self._section(f, 5, "Calidad de impresión")
        self._row(f, 6, "Oscuridad", ttk.Spinbox(f, textvariable=self.v_dark, from_=-1, to=30, width=8),
                  "0–30 · -1 = impresora")
        self._row(f, 7, "Velocidad", ttk.Spinbox(f, textvariable=self.v_speed, from_=0, to=14, width=8),
                  "pulg/s · 0 = impresora")
        self._row(f, 8, "Desplaz. X (mm)", ttk.Spinbox(f, textvariable=self.v_offx, from_=-50, to=50,
                                                       increment=0.5, width=8))
        self._row(f, 9, "Desplaz. Y (mm)", ttk.Spinbox(f, textvariable=self.v_offy, from_=-50, to=50,
                                                       increment=0.5, width=8))

        self._section(f, 10, "Copias")
        self.copcol_cb = ttk.Combobox(f, textvariable=self.v_copcol, state="readonly", width=18)
        self._row(f, 11, "Columna cantidad", self.copcol_cb, "opcional")
        self._row(f, 12, "Copias por registro", ttk.Spinbox(f, textvariable=self.v_copies, from_=1, to=999, width=8))
        ttk.Label(f, style="Muted.TLabel", wraplength=370, justify="left", text=(
            "Las medidas se guardan en milímetros: la misma plantilla imprime igual en la ZT411 de "
            "300 dpi y en la de 600 dpi. Con columna de cantidad, cada registro se imprime esas veces "
            "(× copias por registro); los registros con 0 se omiten."
        )).grid(row=13, column=0, columnspan=3, sticky="w", pady=(12, 0))

        for v in (self.v_width, self.v_height, self.v_dpi, self.v_dark, self.v_speed,
                  self.v_offx, self.v_offy, self.v_copcol, self.v_copies):
            v.trace_add("write", lambda *a: self._on_template_change())

    def _apply_preset(self):
        size = PRESETS.get(self.v_preset.get())
        if size:
            self.v_width.set(_fmt(size[0]))
            self.v_height.set(_fmt(size[1]))

    # ------------------------------------------------------------ Impresora
    def _build_printer_tab(self):
        f = self._tab("Impresora")
        self.printer_tab = f
        s = self.settings
        self.v_mode = tk.StringVar(value=s.get("mode", "windows"))
        self.v_printer = tk.StringVar(value=s.get("printer", ""))
        self.v_ip = tk.StringVar(value=s.get("ip", ""))
        self.v_port = tk.StringVar(value=str(s.get("port", 9100)))

        self._section(f, 0, "Conexión", top=6)
        ttk.Radiobutton(f, text="Instalada en Windows (USB / red con driver ZDesigner)",
                        variable=self.v_mode, value="windows").grid(row=1, column=0, columnspan=3, sticky="w")
        row = ttk.Frame(f, style="Card.TFrame")
        self.printer_cb = ttk.Combobox(row, textvariable=self.v_printer, state="readonly", width=30)
        self.printer_cb.pack(side="left", fill="x", expand=True)
        self.printer_cb.bind("<<ComboboxSelected>>", lambda e: (self.v_mode.set("windows"), self._dpi_from_printer()))
        ttk.Button(row, text="Actualizar", command=self._refresh_printers).pack(side="left", padx=(6, 0))
        self._row(f, 2, "Impresora", row)
        ttk.Radiobutton(f, text="Directa por red (IP, puerto RAW 9100)",
                        variable=self.v_mode, value="red").grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))
        self._row(f, 4, "Dirección IP", ttk.Entry(f, textvariable=self.v_ip, width=18))
        self._row(f, 5, "Puerto", ttk.Entry(f, textvariable=self.v_port, width=8))

        self._section(f, 6, "Herramientas")
        ttk.Button(f, text="Imprimir prueba (registro actual, 1 copia)", style="Primary.TButton",
                   command=self.print_test).grid(row=7, column=0, columnspan=3, sticky="ew")
        ttk.Button(f, text="Calibrar sensor de etiqueta (~JC)",
                   command=lambda: self._send("~JC\n", "Calibración enviada.")).grid(
            row=8, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        ttk.Button(f, text="Imprimir configuración de la impresora (~WC)",
                   command=lambda: self._send("~WC\n", "Etiqueta de configuración enviada.")).grid(
            row=9, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        ttk.Label(f, style="Muted.TLabel", wraplength=370, justify="left", text=(
            "El ZPL se envía en modo RAW (sin pasar por el driver gráfico). Si la impresora tiene "
            "«300dpi» o «600dpi» en su nombre, la resolución se ajusta automáticamente. "
            "Calibre el sensor al cambiar de rollo o de tamaño de etiqueta."
        )).grid(row=10, column=0, columnspan=3, sticky="w", pady=(12, 0))

        for v in (self.v_mode, self.v_printer, self.v_ip, self.v_port):
            v.trace_add("write", lambda *a: self._sync_dest())
        self._refresh_printers()

    def _refresh_printers(self):
        names = printing.list_printers()
        self.printer_cb["values"] = names
        self.dest_cb["values"] = names + [NET_OPTION]
        if self.v_printer.get() not in names and names:
            zebra = [n for n in names if any(k in n.lower() for k in ("zebra", "zt4", "zdesigner"))]
            self.v_printer.set((zebra or [printing.default_printer() or names[0]])[0])
        self._sync_dest()

    def _sync_dest(self):
        if self.v_mode.get() == "red":
            ip = self.v_ip.get().strip()
            self.v_dest.set(f"Red · {ip}:{self.v_port.get()}" if ip else NET_OPTION)
        else:
            self.v_dest.set(self.v_printer.get())
        self._update_summary()

    def _on_dest_select(self):
        dest = self.v_dest.get()
        if dest == NET_OPTION:
            self.v_mode.set("red")
            self.nb.select(self.printer_tab)
            if not self.v_ip.get().strip():
                self.status.set("Escriba la IP de la impresora en la pestaña «Impresora».")
        else:
            self.v_mode.set("windows")
            self.v_printer.set(dest)
            self._dpi_from_printer()

    def _dpi_from_printer(self):
        """Si el nombre de la impresora indica la resolución (p. ej. 'ZT411-600dpi'), la aplica."""
        name = self.v_printer.get().lower().replace(" ", "")
        for dpi in zpl.DPI_OPTIONS:
            if f"{dpi}dpi" in name and str(dpi) != self.v_dpi.get():
                self.v_dpi.set(str(dpi))
                self.status.set(f"Resolución ajustada a {dpi} dpi según la impresora seleccionada.")

    # ------------------------------------------------------- Vista previa
    def _build_preview(self, parent):
        card = ttk.Frame(parent, style="Card.TFrame")
        parent.add(card, weight=3)
        top = ttk.Frame(card, style="Card.TFrame", padding=(14, 10, 14, 8))
        top.pack(fill="x")
        ttk.Label(top, text="Vista previa", style="Section.TLabel").pack(side="left")
        ttk.Button(top, text="◀", style="Icon.TButton", command=lambda: self.goto_row(-1)).pack(side="left", padx=(16, 2))
        ttk.Button(top, text="▶", style="Icon.TButton", command=lambda: self.goto_row(1)).pack(side="left")
        self.row_lbl = ttk.Label(top, text="Sin datos", style="Value.TLabel")
        self.row_lbl.pack(side="left", padx=10)
        ttk.Label(top, style="Muted.TLabel",
                  text="Arrastre para mover · manijas para redimensionar · flechas 0.5 mm (Shift 0.1 mm)"
                  ).pack(side="right")
        self.canvas = tk.Canvas(card, bg=C["canvas"], highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda e: self.schedule_render())
        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", lambda e: setattr(self, "_drag", None))
        self.canvas.bind("<Motion>", self._on_hover)
        for key, dx, dy in (("Left", -1, 0), ("Right", 1, 0), ("Up", 0, -1), ("Down", 0, 1)):
            self.canvas.bind(f"<{key}>", lambda e, dx=dx, dy=dy: self._nudge(e, dx, dy))

    def _build_table(self, parent):
        card = ttk.Frame(parent, style="Card.TFrame", padding=(14, 10, 14, 12))
        parent.add(card, weight=2)
        top = ttk.Frame(card, style="Card.TFrame")
        top.pack(fill="x", pady=(0, 8))
        ttk.Label(top, text="Datos", style="Section.TLabel").pack(side="left")
        ttk.Label(top, text="Buscar", style="Card.TLabel").pack(side="left", padx=(18, 6))
        self.v_search = tk.StringVar()
        self.v_search.trace_add("write", lambda *a: self.refresh_table())
        ttk.Entry(top, textvariable=self.v_search, width=32).pack(side="left")
        ttk.Button(top, text="Quitar selección",
                   command=lambda: self.tree.selection_remove(self.tree.selection())).pack(side="right")
        ttk.Button(top, text="Seleccionar todo",
                   command=lambda: self.tree.selection_set(self.tree.get_children())).pack(side="right", padx=6)
        body = ttk.Frame(card, style="Card.TFrame")
        body.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(body, show="headings", selectmode="extended")
        self.tree.tag_configure("odd", background=C["surface_alt"])
        ys = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(body, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

    # ============================================== plantilla <-> controles
    def _load_template_into_ui(self):
        t = self.template
        self._loading = True
        try:
            self.v_width.set(_fmt(t.width_mm))
            self.v_height.set(_fmt(t.height_mm))
            self.v_dpi.set(str(t.dpi))
            self.v_dark.set(str(t.darkness))
            self.v_speed.set(str(t.speed))
            self.v_offx.set(_fmt(t.offset_x_mm))
            self.v_offy.set(_fmt(t.offset_y_mm))
            self.v_copcol.set(t.copies_column)
            self.v_copies.set(str(t.copies))
            match = next((k for k, v in PRESETS.items() if v == (t.width_mm, t.height_mm)), "Personalizado")
            self.v_preset.set(match)
        finally:
            self._loading = False
        self.v_badge.set(f"{t.dpi} dpi")
        self.sel = 0 if t.elements else None
        self._refresh_element_list()
        self._load_props()
        self._update_title()
        self._update_summary()
        self.schedule_render()

    def _on_template_change(self):
        if self._loading:
            return
        t = self.template
        try:
            t.width_mm = max(1.0, _num(self.v_width))
            t.height_mm = max(1.0, _num(self.v_height))
            t.dpi = int(self.v_dpi.get() or 300)
            t.darkness = _num(self.v_dark, int)
            t.speed = _num(self.v_speed, int)
            t.offset_x_mm = _num(self.v_offx)
            t.offset_y_mm = _num(self.v_offy)
            t.copies = max(1, _num(self.v_copies, int))
        except ValueError:
            return
        t.copies_column = self.v_copcol.get()
        if PRESETS.get(self.v_preset.get()) != (t.width_mm, t.height_mm):
            self.v_preset.set("Personalizado")
        self.v_badge.set(f"{t.dpi} dpi")
        self._update_summary()
        self.schedule_render()

    @staticmethod
    def _describe(el: Element) -> str:
        kind = "Texto" if el.kind == "texto" else el.barcode_type
        return f"  {kind:<11}  {el.content}"

    def _refresh_element_list(self):
        self.el_list.delete(0, "end")
        for el in self.template.elements:
            self.el_list.insert("end", self._describe(el))
        if self.sel is not None:
            self.el_list.selection_set(self.sel)
            self.el_list.see(self.sel)

    def _load_props(self):
        el = self._el()
        self._loading = True
        try:
            if not el:
                self.props_title.configure(text="Propiedades — agregue un elemento")
                self.v_content.set("")
                self.text_frame.grid_remove()
                self.bc_frame.grid_remove()
                return
            self.props_title.configure(text="Propiedades · " + ("Texto" if el.kind == "texto" else "Código de barras"))
            self.v_content.set(el.content)
            self.v_x.set(_fmt(el.x_mm))
            self.v_y.set(_fmt(el.y_mm))
            self.v_rot.set(el.rotation)
            self.v_font.set(_fmt(el.font_mm))
            self.v_bw.set(_fmt(el.width_mm))
            self.v_lines.set(str(el.max_lines))
            self.v_align.set(el.align)
            self.v_btype.set(el.barcode_type)
            self.v_bh.set(_fmt(el.height_mm))
            self.v_mod.set(_fmt(el.module_mm))
            self.v_mod_scale.set(min(1.0, max(0.1, el.module_mm)))
            self.v_wmode.set("fixed" if el.fixed_width_mm > 0 else "module")
            self.v_fixw.set(_fmt(el.fixed_width_mm) if el.fixed_width_mm > 0 else "")
            self.v_show.set(el.show_text)
            frame, other = (self.text_frame, self.bc_frame) if el.kind == "texto" else (self.bc_frame, self.text_frame)
            other.grid_remove()
            frame.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(4, 0))
        finally:
            self._loading = False
        self._update_width_controls()
        self._update_bc_info()

    def _update_width_controls(self):
        fixed = self.v_wmode.get() == "fixed"
        for w in (self.mod_entry, self.mod_scale):
            w.state(["disabled"] if fixed else ["!disabled"])
        self.fixw_spin.state(["!disabled"] if fixed else ["disabled"])

    def _on_wmode(self):
        el = self._el()
        if not el:
            return
        if self.v_wmode.get() == "fixed" and not self.v_fixw.get().strip():
            data = zpl.fill(el.content, self._current_row())
            w = zpl.barcode_width_mm(el, data, self.template.dpi)
            self._loading = True
            self.v_fixw.set(f"{w:.1f}" if w else "40")
            self._loading = False
        self._update_width_controls()
        self._on_prop_change()

    def _update_bc_info(self):
        el = self._el()
        if not el or el.kind != "codigo":
            return
        data = zpl.fill(el.content, self._current_row())
        parts = []
        for dpi in zpl.DPI_OPTIONS:
            w = zpl.barcode_width_mm(el, data, dpi)
            m = zpl.module_dots(el, data, dpi)
            parts.append(f"{dpi} dpi: {w:.1f} mm (barra {m} pts)" if w else f"{dpi} dpi: datos no válidos")
        self.bc_info.configure(text="Ancho real con el registro actual\n" + "   ·   ".join(parts))

    def _on_prop_change(self):
        el = self._el()
        if self._loading or not el:
            return
        try:
            el.x_mm = max(0.0, _num(self.v_x))
            el.y_mm = max(0.0, _num(self.v_y))
            el.font_mm = max(0.5, _num(self.v_font))
            el.width_mm = max(0.0, _num(self.v_bw))
            el.max_lines = max(1, _num(self.v_lines, int))
            el.height_mm = max(0.5, _num(self.v_bh))
            el.module_mm = max(0.05, _num(self.v_mod))
            el.fixed_width_mm = max(1.0, _num(self.v_fixw)) if self.v_wmode.get() == "fixed" else 0.0
        except ValueError:
            return
        el.content = self.v_content.get()
        el.rotation = self.v_rot.get()
        el.align = self.v_align.get()
        el.barcode_type = self.v_btype.get()
        el.show_text = bool(self.v_show.get())
        self.v_mod_scale.set(min(1.0, max(0.1, el.module_mm)))
        self.el_list.delete(self.sel)
        self.el_list.insert(self.sel, self._describe(el))
        self.el_list.selection_set(self.sel)
        self._update_bc_info()
        self.schedule_render()

    def _el(self) -> Element | None:
        els = self.template.elements
        return els[self.sel] if self.sel is not None and 0 <= self.sel < len(els) else None

    def _select_element(self, i: int | None):
        self.sel = i
        self.el_list.selection_clear(0, "end")
        if i is not None:
            self.el_list.selection_set(i)
            self.el_list.see(i)
        self._load_props()
        self.schedule_render()

    def _on_list_select(self, _e):
        cur = self.el_list.curselection()
        if cur:
            self._select_element(cur[0])

    def _insert_column(self):
        col = self.col_cb.get()
        if col and self._el():
            self.content_entry.insert("insert", "{" + col + "}")
            self.content_entry.focus_set()

    def add_element(self, kind: str):
        col = self.col_cb.get() or (self.headers[0] if self.headers else "Columna")
        el = Element(kind=kind, content="{" + col + "}", x_mm=4, y_mm=4)
        if kind == "codigo":
            el.height_mm = 12
        self.template.elements.append(el)
        self.sel = len(self.template.elements) - 1
        self._refresh_element_list()
        self._select_element(self.sel)
        self.nb.select(0)

    def duplicate_element(self):
        el = self._el()
        if not el:
            return
        copy = Element.from_dict(vars(el))
        copy.x_mm += 3
        copy.y_mm += 3
        self.template.elements.insert(self.sel + 1, copy)
        self.sel += 1
        self._refresh_element_list()
        self._select_element(self.sel)

    def delete_element(self):
        if self._el() is None:
            return
        del self.template.elements[self.sel]
        n = len(self.template.elements)
        self.sel = min(self.sel, n - 1) if n else None
        self._refresh_element_list()
        self._select_element(self.sel)

    def move_element(self, d: int):
        els, i = self.template.elements, self.sel
        if i is None or not (0 <= i + d < len(els)):
            return
        els[i], els[i + d] = els[i + d], els[i]
        self.sel = i + d
        self._refresh_element_list()
        self._select_element(self.sel)

    # ======================================================== vista previa
    def schedule_render(self):
        if self._render_job:
            self.after_cancel(self._render_job)
        self._render_job = self.after(25, self.render_preview)

    def _current_row(self) -> dict:
        return self.rows[self.current_row] if 0 <= self.current_row < len(self.rows) else {}

    def render_preview(self):
        self._render_job = None
        t = self.template
        c = self.canvas
        cw, ch = max(200, c.winfo_width()), max(150, c.winfo_height())
        wd, hd = zpl.mm_to_dots(t.width_mm, t.dpi), zpl.mm_to_dots(t.height_mm, t.dpi)
        if wd <= 0 or hd <= 0:
            return
        m = 50 * self.f
        s = max(0.02, min((cw - 2 * m) / wd, (ch - 2.2 * m) / hd))
        img, self._boxes = preview.render(t, self._current_row(), s)
        self._scale = s
        ox, oy = max(int(m), (cw - img.width) // 2), max(int(m * 0.8), (ch - img.height) // 2)
        self._origin = (ox, oy)
        self._photo = ImageTk.PhotoImage(img)

        c.delete("all")
        c.create_text(16, ch - 10, anchor="sw", fill=C["muted"], font=theme.F_SMALL,
                      text=f"{_fmt(t.width_mm)} × {_fmt(t.height_mm)} mm   ·   {t.dpi} dpi   ·   "
                           f"{wd} × {hd} puntos   ·   vista previa aproximada")
        self._draw_rulers(ox, oy, img.width, img.height, t.dpi * s / 25.4)
        for i, col in ((6, "#C9D1E3"), (3, "#B8C2D9")):
            c.create_rectangle(ox + i - 2, oy + i, ox + img.width + i - 2, oy + img.height + i, fill=col, outline="")
        c.create_image(ox, oy, image=self._photo, anchor="nw")

        self._handles = {}
        el = self._el()
        if el and self.sel < len(self._boxes):
            x0, y0, x1, y1 = (v + o for v, o in zip(self._boxes[self.sel], (ox, oy, ox, oy)))
            c.create_rectangle(x0 - 3, y0 - 3, x1 + 3, y1 + 3, outline=C["primary"], width=2, dash=(5, 3))
            if el.rotation == "0°":
                self._handles["e"] = (x1 + 3, (y0 + y1) // 2)
                if el.kind == "codigo" and el.barcode_type not in ("QR", "DataMatrix"):
                    self._handles["s"] = ((x0 + x1) // 2, y1 + 3)
            for hx, hy in self._handles.values():
                k = self.handle
                c.create_rectangle(hx - k, hy - k, hx + k, hy + k,
                                   fill="white", outline=C["primary"], width=2)

        if self.rows:
            n = zpl.copies_for_row(t, self._current_row())
            self.row_lbl.configure(text=f"Registro {self.current_row + 1} de {len(self.rows)}  ·  {n} copia(s)")
        else:
            self.row_lbl.configure(text="Sin datos — se muestran los marcadores {Columna}")
        self._update_bc_info()

    def _draw_rulers(self, ox, oy, w, h, px_mm):
        c = self.canvas
        step = 1 if px_mm >= 4 else 2 if px_mm >= 2 else 5
        label_every = 10 if px_mm * 10 >= 28 else 20
        font = (theme.FONT, 7)
        for mm in range(0, int(w / px_mm) + 1, step):
            x = ox + mm * px_mm
            size = 9 if mm % label_every == 0 else 5 if mm % 5 == 0 else 3
            c.create_line(x, oy - 6 - size, x, oy - 6, fill=C["ruler"])
            if mm % label_every == 0:
                c.create_text(x + 2, oy - 18, anchor="sw", text=str(mm), fill=C["ruler"], font=font)
        for mm in range(0, int(h / px_mm) + 1, step):
            y = oy + mm * px_mm
            size = 9 if mm % label_every == 0 else 5 if mm % 5 == 0 else 3
            c.create_line(ox - 6 - size, y, ox - 6, y, fill=C["ruler"])
            if mm % label_every == 0 and mm:
                c.create_text(ox - 18, y, anchor="e", text=str(mm), fill=C["ruler"], font=font)

    def goto_row(self, d: int):
        if not self.rows:
            return
        self.current_row = (self.current_row + d) % len(self.rows)
        iid = str(self.current_row)
        if self.tree.exists(iid):
            self.tree.see(iid)
        self.schedule_render()

    def _handle_at(self, x, y) -> str | None:
        return next((k for k, (hx, hy) in self._handles.items()
                     if abs(x - hx) <= self.handle + 3 and abs(y - hy) <= self.handle + 3), None)

    def _element_at(self, x, y) -> int | None:
        px, py = x - self._origin[0], y - self._origin[1]
        return next((i for i in reversed(range(len(self._boxes)))
                     if self._boxes[i][0] <= px <= self._boxes[i][2]
                     and self._boxes[i][1] <= py <= self._boxes[i][3]), None)

    def _on_hover(self, e):
        h = self._handle_at(e.x, e.y)
        cursor = {"e": "sb_h_double_arrow", "s": "sb_v_double_arrow"}.get(h)
        if not cursor:
            cursor = "fleur" if self._element_at(e.x, e.y) is not None else ""
        self.canvas.configure(cursor=cursor)

    def _on_press(self, e):
        self.canvas.focus_set()
        el = self._el()
        handle = self._handle_at(e.x, e.y)
        if handle and el:
            if el.kind == "codigo":
                data = zpl.fill(el.content, self._current_row())
                start = (zpl.barcode_width_mm(el, data, self.template.dpi) or 40) if handle == "e" else el.height_mm
            else:
                x0, _, x1, _ = self._boxes[self.sel]
                start = (x1 - x0) / self._scale * 25.4 / self.template.dpi
            self._drag = ("resize-" + handle, e.x, e.y, start)
            return
        hit = self._element_at(e.x, e.y)
        if hit is None:
            self._drag = None
            return
        self._select_element(hit)
        el = self.template.elements[hit]
        self._drag = ("move", e.x, e.y, (el.x_mm, el.y_mm))

    def _on_drag(self, e):
        el = self._el()
        if not self._drag or not el:
            return
        mode, sx, sy, start = self._drag
        k = 25.4 / (self.template.dpi * self._scale)  # mm por píxel de pantalla
        self._loading = True
        try:
            if mode == "move":
                el.x_mm = max(0.0, round(start[0] + (e.x - sx) * k, 1))
                el.y_mm = max(0.0, round(start[1] + (e.y - sy) * k, 1))
                self.v_x.set(_fmt(el.x_mm))
                self.v_y.set(_fmt(el.y_mm))
            elif mode == "resize-e":
                w = max(5.0, round(start + (e.x - sx) * k, 1))
                if el.kind == "codigo":
                    el.fixed_width_mm = w
                    self.v_wmode.set("fixed")
                    self.v_fixw.set(_fmt(w))
                    self._update_width_controls()
                else:
                    el.width_mm = w
                    self.v_bw.set(_fmt(w))
            elif mode == "resize-s":
                el.height_mm = max(1.0, round(start + (e.y - sy) * k, 1))
                self.v_bh.set(_fmt(el.height_mm))
        finally:
            self._loading = False
        self.schedule_render()

    def _nudge(self, e, dx, dy):
        el = self._el()
        if not el:
            return
        step = 0.1 if e.state & 0x1 else 0.5
        el.x_mm = max(0.0, round(el.x_mm + dx * step, 1))
        el.y_mm = max(0.0, round(el.y_mm + dy * step, 1))
        self._loading = True
        self.v_x.set(_fmt(el.x_mm))
        self.v_y.set(_fmt(el.y_mm))
        self._loading = False
        self.schedule_render()

    # ================================================================ Excel
    def open_excel(self):
        path = filedialog.askopenfilename(title="Abrir base de datos",
                                          filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if path:
            self._open_excel(path)

    def _open_excel(self, path: str, sheet: str | None = None):
        try:
            sheets = excel_data.list_sheets(path)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"No se pudo abrir el Excel:\n{exc}")
            return
        self.excel_path = path
        self.sheet_cb["values"] = sheets
        sheet = sheet if sheet in sheets else sheets[0]
        self.sheet_cb.set(sheet)
        self.load_sheet(sheet)

    def load_sheet(self, sheet: str):
        try:
            self.headers, self.rows = excel_data.read_sheet(self.excel_path, sheet)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"No se pudo leer la hoja:\n{exc}")
            return
        self.current_row = 0
        self.col_cb["values"] = self.headers
        if self.headers:
            self.col_cb.set(self.headers[0])
        self.copcol_cb["values"] = [""] + self.headers
        self.tree["columns"] = ["#"] + [f"c{i}" for i in range(len(self.headers))]
        self.tree.heading("#", text="#")
        self.tree.column("#", width=56, stretch=False, anchor="e")
        for i, h in enumerate(self.headers):
            self.tree.heading(f"c{i}", text=h, anchor="w")
            self.tree.column(f"c{i}", width=140, stretch=True)
        self.refresh_table()
        missing = sorted({m.strip() for el in self.template.elements
                          for m in zpl.PLACEHOLDER.findall(el.content)} - set(self.headers))
        msg = f"{Path(self.excel_path).name} › {sheet}: {len(self.rows)} registros, {len(self.headers)} columnas."
        if missing:
            msg += "   ⚠ La plantilla usa columnas que no existen: " + ", ".join(missing)
        self.status.set(msg)
        self.schedule_render()

    def refresh_table(self):
        q = self.v_search.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        self.visible = [i for i, r in enumerate(self.rows)
                        if not q or any(q in str(v).lower() for v in r.values())]
        for n, i in enumerate(self.visible):
            r = self.rows[i]
            self.tree.insert("", "end", iid=str(i), tags=("odd",) if n % 2 else (),
                             values=[i + 1] + [r.get(h, "") for h in self.headers])
        self._update_summary()

    def _on_tree_select(self, _e):
        sel = self.tree.selection()
        if sel:
            self.current_row = int(sel[-1])
            self.schedule_render()
        self._update_summary()

    def _update_summary(self):
        if not hasattr(self, "tree"):
            return
        sel = self.tree.selection()
        rows = [self.rows[int(i)] for i in sel] if sel else [self.rows[i] for i in self.visible]
        total = sum(zpl.copies_for_row(self.template, r) for r in rows)
        scope = f"{len(sel)} seleccionado(s)" if sel else f"{len(self.visible)} registro(s)"
        dest = self.v_dest.get() if hasattr(self, "v_dest") else ""
        self.summary.set(f"{scope}  ·  {total} etiqueta(s)  ·  {self.template.dpi} dpi  ·  {dest}")

    # ============================================================ plantillas
    def new_template(self):
        if messagebox.askyesno(APP_TITLE, "¿Crear una plantilla nueva? Se perderán los cambios no guardados."):
            self.template = zpl.default_template()
            self.template_path = None
            self._load_template_into_ui()

    def open_template(self):
        path = filedialog.askopenfilename(title="Abrir plantilla", filetypes=[("Plantilla", "*.json")])
        if path:
            self._open_template(path)

    def _open_template(self, path: str) -> bool:
        try:
            self.template = zpl.Template.load(path)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"No se pudo abrir la plantilla:\n{exc}")
            return False
        self.template_path = path
        self._load_template_into_ui()
        return True

    def save_template(self, ask: bool = False):
        path = self.template_path
        if ask or not path:
            path = filedialog.asksaveasfilename(title="Guardar plantilla", defaultextension=".json",
                                                filetypes=[("Plantilla", "*.json")])
            if not path:
                return
        self.template.save(path)
        self.template_path = path
        self._update_title()
        self.status.set(f"Plantilla guardada: {path}")

    def _update_title(self):
        name = Path(self.template_path).name if self.template_path else "sin guardar"
        self.title(f"{APP_TITLE} — {name}")

    # ============================================================= impresión
    def _target_rows(self, which: str) -> list[dict] | None:
        if not self.rows:
            if messagebox.askyesno(APP_TITLE, "No hay datos de Excel cargados. ¿Imprimir la etiqueta sin datos?"):
                return [{}]
            return None
        if which == "sel":
            idx = sorted(int(i) for i in self.tree.selection())
            if not idx:
                messagebox.showinfo(APP_TITLE, "Seleccione uno o más registros en la tabla "
                                               "(Ctrl/Shift + clic o «Seleccionar todo»).")
                return None
        else:
            idx = self.visible
        return [self.rows[i] for i in idx]

    def _destination(self) -> str:
        if self.v_mode.get() == "red":
            return f"{self.v_ip.get()}:{self.v_port.get()}"
        return self.v_printer.get()

    def print_rows(self, which: str):
        if self.v_mode.get() == "windows":
            self._dpi_from_printer()
        rows = self._target_rows(which)
        if not rows:
            return
        data, total = zpl.build_job(self.template, rows)
        if total == 0:
            messagebox.showwarning(APP_TITLE, "No hay etiquetas que imprimir (cantidad 0).")
            return
        if not messagebox.askyesno(APP_TITLE, f"Se imprimirán {total} etiqueta(s) de {len(rows)} registro(s)\n"
                                              f"en: {self._destination()}  ({self.template.dpi} dpi)\n\n¿Continuar?"):
            return
        self._send(data, f"Enviadas {total} etiqueta(s) a {self._destination()}.")

    def print_test(self):
        if self.v_mode.get() == "windows":
            self._dpi_from_printer()
        self._send(zpl.label_zpl(self.template, self._current_row(), 1), "Etiqueta de prueba enviada.")

    def _send(self, data: str, ok_msg: str):
        mode, printer = self.v_mode.get(), self.v_printer.get()
        ip, port = self.v_ip.get().strip(), self.v_port.get().strip()
        if mode == "windows" and not printer:
            messagebox.showwarning(APP_TITLE, "Elija una impresora en el encabezado o en la pestaña «Impresora».")
            return
        if mode == "red" and not ip:
            messagebox.showwarning(APP_TITLE, "Escriba la IP de la impresora en la pestaña «Impresora».")
            self.nb.select(self.printer_tab)
            return
        self.status.set("Enviando a la impresora…")

        def work():
            try:
                raw = data.encode("utf-8")
                if mode == "red":
                    printing.send_tcp(ip, int(port or 9100), raw)
                else:
                    printing.send_windows(printer, raw)
                self.after(0, lambda: self.status.set("✔ " + ok_msg))
            except Exception as exc:
                err = str(exc)
                self.after(0, lambda: (self.status.set("Error al imprimir."),
                                       messagebox.showerror(APP_TITLE, f"Error al imprimir:\n{err}")))

        threading.Thread(target=work, daemon=True).start()

    def show_zpl(self):
        win = tk.Toplevel(self)
        win.title("ZPL del registro actual")
        win.geometry("700x520")
        win.configure(bg=C["surface"])
        txt = tk.Text(win, font=("Consolas", 10), wrap="none", bg=C["navy"], fg="#E4EAFB",
                      insertbackground="white", relief="flat", padx=12, pady=10,
                      selectbackground=C["primary"])
        txt.pack(fill="both", expand=True)
        txt.insert("1.0", zpl.label_zpl(self.template, self._current_row(), 1))
        bar = ttk.Frame(win, style="Card.TFrame", padding=10)
        bar.pack(fill="x")
        ttk.Label(bar, style="Muted.TLabel", text="Puede pegarlo en labelary.com/viewer.html para verificarlo."
                  ).pack(side="left")

        def copy():
            self.clipboard_clear()
            self.clipboard_append(txt.get("1.0", "end-1c"))
            self.status.set("ZPL copiado al portapapeles.")
        ttk.Button(bar, text="Copiar", style="Primary.TButton", command=copy).pack(side="right")

    def export_zpl(self):
        rows = self._target_rows("sel" if self.tree.selection() else "all")
        if not rows:
            return
        data, total = zpl.build_job(self.template, rows)
        path = filedialog.asksaveasfilename(title="Exportar ZPL", defaultextension=".zpl",
                                            filetypes=[("ZPL", "*.zpl"), ("Texto", "*.txt")])
        if path:
            Path(path).write_text(data, encoding="utf-8")
            self.status.set(f"Exportadas {total} etiqueta(s) a {path}")

    # ========================================================= preferencias
    def _load_settings(self) -> dict:
        try:
            return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _restore_session(self):
        s = self.settings
        tpl = s.get("template")
        if not (tpl and Path(tpl).exists() and self._open_template(tpl)):
            self._load_template_into_ui()
        if s.get("excel") and Path(s["excel"]).exists():
            self._open_excel(s["excel"], s.get("sheet"))
        self.after(150, self._initial_layout)

    def _initial_layout(self):
        self.update_idletasks()
        h = self.right_pane.winfo_height()
        if h < 200:  # la ventana aún no termina de dimensionarse
            self.after(150, self._initial_layout)
            return
        self.right_pane.sashpos(0, int(h * 0.62))
        self.right_pane.bind("<Configure>", lambda e: self.after_idle(self._keep_table_visible), add="+")

    def _keep_table_visible(self):
        """Evita que la vista previa ocupe todo el panel y oculte la tabla de datos."""
        h, pos = self.right_pane.winfo_height(), self.right_pane.sashpos(0)
        min_table = int(220 * self.f)
        if h > min_table * 2 and pos > h - min_table:
            self.right_pane.sashpos(0, h - min_table)

    def _on_close(self):
        s = {"mode": self.v_mode.get(), "printer": self.v_printer.get(), "ip": self.v_ip.get(),
             "port": self.v_port.get(), "template": self.template_path, "excel": self.excel_path,
             "sheet": self.sheet_cb.get()}
        try:
            SETTINGS_PATH.write_text(json.dumps(s, indent=2, ensure_ascii=False), encoding="utf-8")
        except OSError:
            pass
        self.destroy()


if __name__ == "__main__":
    try:  # texto nítido en pantallas con escala 125 % / 150 %
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App().mainloop()
