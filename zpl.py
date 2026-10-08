"""Modelo de plantilla y generación de ZPL II para Zebra ZT411 (300 / 600 dpi).

Todas las medidas se guardan en milímetros y se convierten a puntos según el
dpi elegido, así la misma plantilla imprime igual en 300 y en 600 dpi.
"""
from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path

import symbols

BARCODE_TYPES = ["Code 128", "Code 39", "EAN-13", "UPC-A", "QR", "DataMatrix"]
ROTATIONS = {"0°": "N", "90°": "R", "180°": "I", "270°": "B"}
ALIGNS = {"Izquierda": "L", "Centro": "C", "Derecha": "R"}
DPI_OPTIONS = (300, 600)
PLACEHOLDER = re.compile(r"\{([^{}]+)\}")


def mm_to_dots(mm: float, dpi: int) -> int:
    return max(0, int(round(float(mm) * dpi / 25.4)))


@dataclass
class Element:
    kind: str = "texto"            # "texto" | "codigo"
    content: str = ""              # admite {Columna}
    x_mm: float = 2.0
    y_mm: float = 2.0
    rotation: str = "0°"
    # texto
    font_mm: float = 3.0
    width_mm: float = 0.0          # 0 = una sola línea sin ajuste
    max_lines: int = 1
    align: str = "Izquierda"
    # código de barras
    barcode_type: str = "Code 128"
    height_mm: float = 10.0
    module_mm: float = 0.33        # ancho de la barra más fina / módulo QR
    fixed_width_mm: float = 0.0    # >0: el módulo se ajusta para que el código mida este ancho
    show_text: bool = True

    @classmethod
    def from_dict(cls, d: dict) -> "Element":
        known = {f.name for f in fields(cls)}
        return cls(**{k: v for k, v in d.items() if k in known})

    def describe(self) -> str:
        tag = "Texto" if self.kind == "texto" else self.barcode_type
        return f"[{tag}] {self.content}"


@dataclass
class Template:
    width_mm: float = 100.0
    height_mm: float = 50.0
    dpi: int = 300
    darkness: int = -1             # -1 = usar la de la impresora (0-30)
    speed: int = 0                 # 0 = usar la de la impresora (pulg/s)
    offset_x_mm: float = 0.0
    offset_y_mm: float = 0.0
    copies_column: str = ""        # columna del Excel con la cantidad por registro
    copies: int = 1                # multiplicador de copias por registro
    elements: list[Element] = field(default_factory=list)

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "Template":
        known = {f.name for f in fields(cls)} - {"elements"}
        t = cls(**{k: v for k, v in d.items() if k in known})
        t.elements = [Element.from_dict(e) for e in d.get("elements", [])]
        return t

    def save(self, path: str | Path) -> None:
        Path(path).write_text(json.dumps(self.to_dict(), indent=2, ensure_ascii=False), encoding="utf-8")

    @classmethod
    def load(cls, path: str | Path) -> "Template":
        return cls.from_dict(json.loads(Path(path).read_text(encoding="utf-8")))


def default_template() -> Template:
    return Template(elements=[
        Element(kind="texto", content="{Descripcion}", x_mm=4, y_mm=3, font_mm=4.5,
                width_mm=92, max_lines=2),
        Element(kind="codigo", content="{Codigo}", x_mm=4, y_mm=15, barcode_type="Code 128",
                height_mm=14, module_mm=0.33),
        Element(kind="texto", content="Precio: ${Precio}", x_mm=4, y_mm=40, font_mm=5),
        Element(kind="texto", content="Lote: {Lote}", x_mm=62, y_mm=41, font_mm=3.5),
    ])


def fill(text: str, row: dict) -> str:
    """Sustituye {Columna} por el valor del registro (deja el marcador si no existe)."""
    return PLACEHOLDER.sub(lambda m: str(row.get(m.group(1).strip(), m.group(0))), text)


def copies_for_row(t: Template, row: dict) -> int:
    n = 1
    if t.copies_column:
        try:
            n = int(float(str(row.get(t.copies_column, "1")).replace(",", ".") or 0))
        except ValueError:
            n = 1
    return max(0, n) * max(1, int(t.copies))


def module_dots(el: Element, data: str, dpi: int) -> int:
    """Ancho del módulo en puntos. Con ancho fijo se usa el mayor módulo que cabe en él."""
    limit = 10 if el.barcode_type == "QR" else 99
    m = max(1, mm_to_dots(el.module_mm, dpi))
    if el.fixed_width_mm > 0:
        try:
            m = max(1, mm_to_dots(el.fixed_width_mm, dpi) // symbols.symbol_modules(el.barcode_type, data))
        except Exception:
            pass  # datos inválidos: se queda con el módulo manual
    return min(limit, m)


def barcode_width_mm(el: Element, data: str, dpi: int) -> float | None:
    """Ancho real impreso del código (mm) para estos datos, o None si no se puede calcular."""
    try:
        return symbols.symbol_modules(el.barcode_type, data) * module_dots(el, data, dpi) * 25.4 / dpi
    except Exception:
        return None


def _field_data(data: str) -> str:
    # ^FH permite escapar los caracteres reservados de ZPL en hexadecimal.
    data = data.replace("_", "_5F").replace("^", "_5E").replace("~", "_7E")
    return f"^FH_^FD{data}^FS"


def element_zpl(el: Element, row: dict, dpi: int) -> str:
    x, y = mm_to_dots(el.x_mm, dpi), mm_to_dots(el.y_mm, dpi)
    rot = ROTATIONS.get(el.rotation, "N")
    data = fill(el.content, row)
    out = [f"^FO{x},{y}"]

    if el.kind == "texto":
        h = max(10, mm_to_dots(el.font_mm, dpi))
        out.append(f"^A0{rot},{h},{h}")
        if el.width_mm > 0:
            w = mm_to_dots(el.width_mm, dpi)
            out.append(f"^FB{w},{max(1, int(el.max_lines))},0,{ALIGNS.get(el.align, 'L')},0")
            data = data.replace("\r\n", "\n").replace("\n", "\\&")
        else:
            data = data.replace("\r\n", " ").replace("\n", " ")
        out.append(_field_data(data))
        return "".join(out)

    module = module_dots(el, data, dpi)
    h = max(1, mm_to_dots(el.height_mm, dpi))
    hr = "Y" if el.show_text else "N"
    t = el.barcode_type
    if t == "Code 128":
        out += [f"^BY{module}", f"^BC{rot},{h},{hr},N,N,A"]
    elif t == "Code 39":
        out += [f"^BY{module},3", f"^B3{rot},N,{h},{hr},N"]
        data = data.upper()
    elif t == "EAN-13":
        out += [f"^BY{module}", f"^BE{rot},{h},{hr},N"]
        data = symbols.digits(data)[:12]
    elif t == "UPC-A":
        out += [f"^BY{module}", f"^BU{rot},{h},{hr},N,Y"]
        data = symbols.digits(data)[:11]
    elif t == "QR":
        out.append(f"^BQ{rot},2,{module}")
        data = "QA," + data
    elif t == "DataMatrix":
        out.append(f"^BX{rot},{module},200")
    out.append(_field_data(data))
    return "".join(out)


def label_zpl(t: Template, row: dict, copies: int = 1) -> str:
    dpi = int(t.dpi)
    lines = []
    if 0 <= int(t.darkness) <= 30:
        lines.append(f"~SD{int(t.darkness):02d}")
    lines += [
        "^XA",
        "^CI28",  # UTF-8 (acentos, ñ)
        f"^PW{mm_to_dots(t.width_mm, dpi)}",
        f"^LL{mm_to_dots(t.height_mm, dpi)}",
        f"^LH{mm_to_dots(t.offset_x_mm, dpi)},{mm_to_dots(t.offset_y_mm, dpi)}",
    ]
    if int(t.speed) > 0:
        lines.append(f"^PR{int(t.speed)}")
    lines += [element_zpl(el, row, dpi) for el in t.elements]
    lines += [f"^PQ{max(1, copies)}", "^XZ"]
    return "\n".join(lines) + "\n"


def build_job(t: Template, rows: list[dict]) -> tuple[str, int]:
    """Devuelve (ZPL completo, total de etiquetas)."""
    parts, total = [], 0
    for row in rows:
        n = copies_for_row(t, row)
        if n <= 0:
            continue
        parts.append(label_zpl(t, row, n))
        total += n
    return "".join(parts), total
