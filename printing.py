"""Envío de ZPL a la impresora: cola de Windows (RAW) o red TCP 9100."""
from __future__ import annotations

import socket

try:
    import win32print
except ImportError:  # pywin32 no instalado o sistema no Windows
    win32print = None


def list_printers() -> list[str]:
    if win32print is None:
        return []
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    return [p[2] for p in win32print.EnumPrinters(flags)]


def default_printer() -> str:
    if win32print is None:
        return ""
    try:
        return win32print.GetDefaultPrinter()
    except Exception:
        return ""


def send_windows(printer: str, data: bytes, job_name: str = "Etiquetas Zebra") -> None:
    if win32print is None:
        raise RuntimeError("pywin32 no está instalado (pip install pywin32).")
    handle = win32print.OpenPrinter(printer)
    try:
        win32print.StartDocPrinter(handle, 1, (job_name, None, "RAW"))
        try:
            win32print.StartPagePrinter(handle)
            win32print.WritePrinter(handle, data)
            win32print.EndPagePrinter(handle)
        finally:
            win32print.EndDocPrinter(handle)
    finally:
        win32print.ClosePrinter(handle)


def send_tcp(host: str, port: int, data: bytes, timeout: float = 10) -> None:
    with socket.create_connection((host, int(port)), timeout=timeout) as s:
        s.sendall(data)
