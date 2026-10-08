"""Vista previa aproximada de la etiqueta con Pillow.

La impresión real la genera la impresora a partir del ZPL; esto solo sirve para
acomodar los elementos. Se calcula en puntos (igual que el ZPL) y se escala.
"""
from __future__ import annotations

import math
from functools import lru_cache

from PIL import Image, ImageDraw, ImageFont

import symbols
import zpl
from zpl import Element, Template, mm_to_dots

_FONTS = ["arialnb.ttf", "ARIALNB.TTF", "arialbd.ttf", "DejaVuSansCondensed-Bold.ttf", "arial.ttf"]
_ROTATE = {"90°": -90, "180°": 180, "270°": 90}


@lru_cache(maxsize=64)
def _font(px: int):
    """Fuente cuya altura total (ascendente + descendente) cabe en `px`, como la celda de ^A0."""
    for name in _FONTS:
        try:
            font = ImageFont.truetype(name, px)
        except OSError:
            continue
        asc, desc = font.getmetrics()
        return ImageFont.truetype(name, max(1, int(px * px / (asc + desc)))) if asc + desc > px else font
    return ImageFont.load_default()


def _wrap(text: str, font, width: float) -> list[str]:
    out = []
    for para in text.split("\n"):
        cur = ""
        for word in para.split(" "):
            cand = word if not cur else f"{cur} {word}"
            if not cur or font.getlength(cand) <= width:
                cur = cand
            else:
                out.append(cur)
                cur = word
        out.append(cur)
    return out


def _render_text(el: Element, data: str, dpi: int, s: float) -> Image.Image:
    h = max(10, mm_to_dots(el.font_mm, dpi))
    px = max(4, round(h * s))
    font = _font(px)
    data = data.replace("\r\n", "\n")
    if el.width_mm > 0:
        wpx = max(1, round(mm_to_dots(el.width_mm, dpi) * s))
        lines = _wrap(data, font, wpx)[: max(1, int(el.max_lines))]
        img = Image.new("RGBA", (wpx, px * len(lines)), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        align = zpl.ALIGNS.get(el.align, "L")
        for i, line in enumerate(lines):
            tw = font.getlength(line)
            x = 0 if align == "L" else (wpx - tw) / 2 if align == "C" else wpx - tw
            d.text((x, i * px), line, font=font, fill="black")
        return img
    line = data.replace("\n", " ")
    img = Image.new("RGBA", (max(1, math.ceil(font.getlength(line))), px), (0, 0, 0, 0))
    ImageDraw.Draw(img).text((0, 0), line, font=font, fill="black")
    return img


def _render_barcode(el: Element, data: str, dpi: int, s: float) -> Image.Image:
    if not data:
        raise ValueError("sin datos")
    module = zpl.module_dots(el, data, dpi)

    if el.barcode_type in ("QR", "DataMatrix"):
        if el.barcode_type == "QR":
            matrix = symbols.qr_matrix(data)
        else:
            # Aproximación: marco en L y patrón del tamaño estimado.
            n = symbols.datamatrix_size(data)
            matrix = [[(c == 0) or (r == n - 1) or (r == 0 and c % 2 == 0) or
                       (c == n - 1 and r % 2 == 1) or ((r * 7 + c * 3 + len(data)) % 5 < 2)
                       for c in range(n)] for r in range(n)]
        cell = module * s
        size = max(1, math.ceil(len(matrix) * cell))
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        for r, row in enumerate(matrix):
            for c, on in enumerate(row):
                if on:
                    d.rectangle([c * cell, r * cell, (c + 1) * cell - 0.01, (r + 1) * cell - 0.01], fill="black")
        return img

    mods, human = symbols.linear_modules(el.barcode_type, data)
    h = max(1, mm_to_dots(el.height_mm, dpi)) * s
    mpx = module * s
    hr_px = round(max(12, module * 7) * s) if el.show_text else 0
    gap = round(module * 2 * s) if el.show_text else 0
    w = max(1, math.ceil(len(mods) * mpx))
    img = Image.new("RGBA", (w, max(1, math.ceil(h + gap + hr_px))), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    i = 0
    while i < len(mods):
        if mods[i] == "1":
            j = i
            while j < len(mods) and mods[j] == "1":
                j += 1
            d.rectangle([round(i * mpx), 0, max(round(i * mpx), round(j * mpx) - 1), h - 1], fill="black")
            i = j
        else:
            i += 1
    if el.show_text:
        if el.barcode_type == "Code 39":
            human = f"*{human}*"
        font = _font(max(4, hr_px))
        tw = font.getlength(human)
        d.text(((w - tw) / 2, h + gap), human, font=font, fill="black")
    return img


def _render_error(msg: str, dpi: int, s: float) -> Image.Image:
    w, h = round(mm_to_dots(30, dpi) * s), round(mm_to_dots(6, dpi) * s)
    img = Image.new("RGBA", (max(w, 20), max(h, 10)), (255, 230, 230, 255))
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, img.width - 1, img.height - 1], outline="red")
    d.text((3, 2), msg, font=_font(max(8, h // 2)), fill="red")
    return img


def render(t: Template, row: dict, s: float) -> tuple[Image.Image, list[tuple[int, int, int, int]]]:
    """Dibuja la etiqueta. `s` = píxeles por punto. Devuelve imagen y bbox de cada elemento."""
    dpi = int(t.dpi)
    W = max(1, round(mm_to_dots(t.width_mm, dpi) * s))
    H = max(1, round(mm_to_dots(t.height_mm, dpi) * s))
    base = Image.new("RGBA", (W, H), "white")
    ox, oy = mm_to_dots(t.offset_x_mm, dpi), mm_to_dots(t.offset_y_mm, dpi)
    boxes = []
    for el in t.elements:
        data = zpl.fill(el.content, row)
        try:
            img = (_render_text if el.kind == "texto" else _render_barcode)(el, data, dpi, s)
        except Exception as exc:  # datos inválidos para el tipo de código, etc.
            img = _render_error(f"{el.barcode_type}: {exc}", dpi, s)
        if el.rotation in _ROTATE:
            img = img.rotate(_ROTATE[el.rotation], expand=True)
        x = round((ox + mm_to_dots(el.x_mm, dpi)) * s)
        y = round((oy + mm_to_dots(el.y_mm, dpi)) * s)
        base.paste(img, (x, y), img)
        boxes.append((x, y, x + img.width, y + img.height))
    ImageDraw.Draw(base).rectangle([0, 0, W - 1, H - 1], outline="#b0b0b0")
    return base.convert("RGB"), boxes
