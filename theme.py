"""Tema visual: azul rey con acentos dorados sobre fondos neutros fríos."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

C = {
    "primary": "#1A3FA6",        # azul rey
    "primary_hover": "#2752C7",
    "primary_pressed": "#122F7E",
    "primary_soft": "#E3EAFB",   # fondos seleccionados / chips
    "navy": "#0D1A44",           # encabezado y barra de estado
    "navy_text": "#AFC0EE",
    "accent": "#F2B33D",         # dorado
    "accent_hover": "#F6C562",
    "accent_pressed": "#D99A22",
    "bg": "#EEF1F8",
    "surface": "#FFFFFF",
    "surface_alt": "#F6F8FD",    # filas alternas
    "border": "#D3DAE8",
    "text": "#18213A",
    "muted": "#5F6B84",
    "canvas": "#DDE3F0",
    "ruler": "#8C97B0",
    "danger": "#C62828",
}

FONT = "Segoe UI"
F_BASE = (FONT, 10)
F_SMALL = (FONT, 9)
F_BOLD = (FONT, 10, "bold")
F_TITLE = (f"{FONT} Semibold", 15)
F_SECTION = (f"{FONT} Semibold", 11)


def apply(root: tk.Tk, f: float = 1.0) -> ttk.Style:
    """Aplica el tema. `f` = factor de escala de pantalla (1.0 = 96 dpi)."""
    st = ttk.Style(root)
    st.theme_use("clam")
    root.configure(bg=C["bg"])

    st.configure(".", background=C["bg"], foreground=C["text"], font=F_BASE,
                 bordercolor=C["border"], lightcolor=C["border"], darkcolor=C["border"],
                 troughcolor=C["bg"], focuscolor=C["primary"], selectbackground=C["primary"],
                 selectforeground="white", insertcolor=C["text"])

    # Contenedores
    st.configure("TFrame", background=C["bg"])
    st.configure("Card.TFrame", background=C["surface"])
    st.configure("Header.TFrame", background=C["navy"])
    st.configure("Toolbar.TFrame", background=C["surface"])
    st.configure("Status.TFrame", background=C["navy"])
    st.configure("Card.TLabelframe", background=C["surface"], bordercolor=C["border"],
                 lightcolor=C["surface"], darkcolor=C["surface"], relief="solid", borderwidth=1)
    st.configure("Card.TLabelframe.Label", background=C["surface"], foreground=C["primary"], font=F_BOLD)
    st.configure("TPanedwindow", background=C["bg"])
    st.configure("Sash", sashthickness=6, gripcount=0, background=C["bg"])
    st.configure("TSeparator", background=C["border"])

    # Etiquetas
    st.configure("TLabel", background=C["bg"])
    st.configure("Card.TLabel", background=C["surface"])
    st.configure("Muted.TLabel", background=C["surface"], foreground=C["muted"], font=F_SMALL)
    st.configure("Section.TLabel", background=C["surface"], foreground=C["text"], font=F_SECTION)
    st.configure("Group.TLabel", background=C["surface"], foreground=C["muted"], font=(FONT, 8, "bold"))
    st.configure("Value.TLabel", background=C["primary_soft"], foreground=C["primary"], font=F_SMALL,
                 padding=(8, 3))
    st.configure("Header.TLabel", background=C["navy"], foreground="white", font=F_TITLE)
    st.configure("HeaderSub.TLabel", background=C["navy"], foreground=C["navy_text"], font=F_SMALL)
    st.configure("Badge.TLabel", background=C["accent"], foreground=C["navy"], font=(FONT, 9, "bold"),
                 padding=(10, 3))
    st.configure("Status.TLabel", background=C["navy"], foreground="#E4EAFB", font=F_SMALL, padding=(10, 4))
    st.configure("StatusMuted.TLabel", background=C["navy"], foreground=C["navy_text"], font=F_SMALL,
                 padding=(10, 4))

    # Botones
    common = dict(padding=(12, 5), relief="flat", borderwidth=1, focusthickness=0, width=0)
    st.configure("TButton", background=C["surface"], foreground=C["primary"], bordercolor=C["border"],
                 lightcolor=C["surface"], darkcolor=C["surface"], **common)
    st.map("TButton",
           background=[("pressed", C["primary_soft"]), ("active", C["surface_alt"])],
           bordercolor=[("active", C["primary"]), ("focus", C["primary"])],
           lightcolor=[("active", C["surface_alt"])], darkcolor=[("active", C["surface_alt"])])
    for name, base, hover, pressed, fg in (
            ("Primary", C["primary"], C["primary_hover"], C["primary_pressed"], "white"),
            ("Accent", C["accent"], C["accent_hover"], C["accent_pressed"], C["navy"])):
        st.configure(f"{name}.TButton", background=base, foreground=fg, bordercolor=base,
                     lightcolor=base, darkcolor=base, font=F_BOLD, **common)
        st.map(f"{name}.TButton",
               background=[("pressed", pressed), ("active", hover)],
               bordercolor=[("pressed", pressed), ("active", hover)],
               lightcolor=[("pressed", pressed), ("active", hover)],
               darkcolor=[("pressed", pressed), ("active", hover)])
    st.configure("Tool.TButton", padding=(10, 4))
    st.configure("Icon.TButton", padding=(6, 3), width=3)
    st.configure("Header.TButton", background=C["navy"], foreground="white", bordercolor="#2A3A70",
                 lightcolor=C["navy"], darkcolor=C["navy"], **{**common, "padding": (8, 3)})
    st.map("Header.TButton", background=[("active", "#1A2A5E")], bordercolor=[("active", C["accent"])])

    # Campos
    field = dict(fieldbackground=C["surface"], background=C["surface"], bordercolor=C["border"],
                 lightcolor=C["surface"], darkcolor=C["surface"], padding=(6, 4), arrowcolor=C["primary"])
    for w in ("TEntry", "TSpinbox", "TCombobox"):
        st.configure(w, **field)
        st.map(w, bordercolor=[("focus", C["primary"]), ("hover", "#AEB9D3")],
               lightcolor=[("focus", C["primary"])],
               fieldbackground=[("readonly", C["surface"]), ("disabled", C["bg"])],
               foreground=[("disabled", C["muted"])],
               selectbackground=[("readonly", C["surface"])], selectforeground=[("readonly", C["text"])])
    st.configure("Header.TCombobox", fieldbackground="#1A2A5E", background="#1A2A5E", foreground="white",
                 bordercolor="#2A3A70", lightcolor="#1A2A5E", darkcolor="#1A2A5E", arrowcolor=C["accent"],
                 padding=(8, 4))
    st.map("Header.TCombobox", fieldbackground=[("readonly", "#1A2A5E")],
           foreground=[("readonly", "white")], selectbackground=[("readonly", "#1A2A5E")],
           selectforeground=[("readonly", "white")], bordercolor=[("focus", C["accent"]), ("hover", C["accent"])])
    root.option_add("*TCombobox*Listbox.font", F_BASE)
    root.option_add("*TCombobox*Listbox.selectBackground", C["primary"])
    root.option_add("*TCombobox*Listbox.selectForeground", "white")
    root.option_add("*TCombobox*Listbox.background", C["surface"])

    for w in ("TCheckbutton", "TRadiobutton"):
        st.configure(w, background=C["surface"], foreground=C["text"], indicatorbackground=C["surface"],
                     indicatorforeground=C["primary"], upperbordercolor=C["border"], lowerbordercolor=C["border"])
        st.map(w, indicatorbackground=[("selected", C["primary"]), ("active", C["primary_soft"])],
               indicatorforeground=[("selected", "white")], background=[("active", C["surface"])])
    st.configure("Horizontal.TScale", background=C["primary"], troughcolor=C["primary_soft"],
                 bordercolor=C["surface"], lightcolor=C["primary"], darkcolor=C["primary"])

    # Pestañas
    st.configure("TNotebook", background=C["surface"], borderwidth=0, tabmargins=(8, 8, 8, 0))
    st.configure("TNotebook.Tab", background=C["surface"], foreground=C["muted"], padding=(16, 7),
                 borderwidth=0, font=F_BOLD, lightcolor=C["surface"], bordercolor=C["surface"])
    st.map("TNotebook.Tab",
           background=[("selected", C["primary_soft"]), ("active", C["surface_alt"])],
           foreground=[("selected", C["primary"])],
           lightcolor=[("selected", C["primary_soft"])])

    # Tabla
    st.configure("Treeview", background=C["surface"], fieldbackground=C["surface"], foreground=C["text"],
                 rowheight=int(28 * f), bordercolor=C["border"], lightcolor=C["surface"], darkcolor=C["surface"])
    st.map("Treeview", background=[("selected", C["primary"])], foreground=[("selected", "white")])
    st.configure("Treeview.Heading", background=C["primary"], foreground="white", font=F_BOLD,
                 relief="flat", padding=(8, 6), bordercolor=C["primary_pressed"],
                 lightcolor=C["primary"], darkcolor=C["primary"])
    st.map("Treeview.Heading", background=[("active", C["primary_hover"])])

    for orient in ("Vertical", "Horizontal"):
        st.configure(f"{orient}.TScrollbar", background="#C3CCDF", troughcolor=C["surface"],
                     bordercolor=C["surface"], lightcolor="#C3CCDF", darkcolor="#C3CCDF",
                     arrowcolor=C["muted"], gripcount=0)
        st.map(f"{orient}.TScrollbar", background=[("active", C["primary"])])
    return st


def style_listbox(lb: tk.Listbox) -> None:
    lb.configure(bg=C["surface"], fg=C["text"], font=F_BASE, relief="flat", bd=0,
                 highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["primary"],
                 selectbackground=C["primary"], selectforeground="white", activestyle="none")


def draw_logo(parent, bg: str, f: float = 1.0) -> tk.Canvas:
    """Pequeño logotipo de código de barras para el encabezado."""
    c = tk.Canvas(parent, width=int(34 * f), height=int(34 * f), bg=bg, highlightthickness=0)
    c.create_rectangle(1, 1, 33, 33, fill=C["primary"], outline="")
    x = 7
    for w in (2, 1, 3, 1, 2, 1, 1, 3):
        c.create_rectangle(x, 8, x + w - 1, 22, fill="white", outline="")
        x += w + 1
    c.create_rectangle(7, 25, 27, 26, fill=C["accent"], outline="")
    c.scale("all", 0, 0, f, f)
    return c
