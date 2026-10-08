"""Geometría de los códigos de barras: patrón de módulos y ancho del símbolo.

Lo usan tanto el generador de ZPL (para ajustar el módulo a un ancho fijo) como la
vista previa, así ambos calculan el mismo tamaño.
"""
from __future__ import annotations

try:
    import barcode as pybarcode
except ImportError:
    pybarcode = None
try:
    import qrcode
    from qrcode.constants import ERROR_CORRECT_Q
except ImportError:
    qrcode = None

LINEAR = ("Code 128", "Code 39", "EAN-13", "UPC-A")
# (capacidad aprox. en caracteres, módulos por lado) para DataMatrix ECC200 cuadrado
_DM_SIZES = [(3, 10), (6, 12), (10, 14), (16, 16), (24, 18), (36, 20), (44, 22), (60, 24),
             (72, 26), (88, 32), (124, 36), (172, 40), (228, 44), (288, 48), (348, 52)]


def digits(s: str) -> str:
    return "".join(c for c in s if c.isdigit())


def linear_modules(btype: str, data: str) -> tuple[str, str]:
    """Patrón '1010…' (sin zona silenciosa) y texto legible."""
    if pybarcode is None:
        raise RuntimeError("instale python-barcode")
    if not data:
        raise ValueError("sin datos")
    if btype == "Code 128":
        code = pybarcode.get_barcode_class("code128")(data)
    elif btype == "Code 39":
        # python-barcode usa barra ancha = 3 módulos, igual que ^BY n,3 en ZPL
        code = pybarcode.get_barcode_class("code39")(data.upper(), add_checksum=False)
    elif btype == "EAN-13":
        d = digits(data)
        if len(d) < 12:
            raise ValueError("EAN-13 requiere 12 o 13 dígitos")
        code = pybarcode.get_barcode_class("ean13")(d[:12])
    elif btype == "UPC-A":
        d = digits(data)
        if len(d) < 11:
            raise ValueError("UPC-A requiere 11 o 12 dígitos")
        code = pybarcode.get_barcode_class("upca")(d[:11])
    else:
        raise ValueError(btype)
    return "".join(code.build()), code.get_fullcode()


def qr_matrix(data: str) -> list[list[bool]]:
    if qrcode is None:
        raise RuntimeError("instale qrcode")
    if not data:
        raise ValueError("sin datos")
    qr = qrcode.QRCode(border=0, error_correction=ERROR_CORRECT_Q)  # ZPL: ^FDQA, = nivel Q
    qr.add_data(data)
    qr.make(fit=True)
    return qr.get_matrix()


def datamatrix_size(data: str) -> int:
    return next((n for cap, n in _DM_SIZES if len(data) <= cap), 52)


def symbol_modules(btype: str, data: str) -> int:
    """Ancho del símbolo en módulos."""
    if btype in LINEAR:
        return len(linear_modules(btype, data)[0])
    if btype == "QR":
        return len(qr_matrix(data))
    if btype == "DataMatrix":
        return datamatrix_size(data)
    raise ValueError(btype)
