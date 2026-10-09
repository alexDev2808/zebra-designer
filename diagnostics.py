"""Comprobación de impresoras Zebra y del driver ZDesigner.

- Cola de Windows: existencia, driver y versión, estado (sin conexión, pausa, error).
- Impresora en red (o cola con puerto IP_x.x.x.x): conexión al puerto 9100 y estado
  real con ~HS (sin papel, pausa, cabezal abierto, sin ribbon).
- USB: Zebra conectada (VID 0A5F) cuyo dispositivo no tiene driver.
- Obtención e instalación del driver configurado en driver.json.
"""
from __future__ import annotations

import hashlib
import json
import shutil
import socket
import subprocess
import sys
import tempfile
import time
import urllib.request
from dataclasses import dataclass
from pathlib import Path

try:
    import win32print
except ImportError:
    win32print = None

OK, INFO, WARN, ERROR = "ok", "info", "warn", "error"
SEVERITY = {OK: 0, INFO: 1, WARN: 2, ERROR: 3}

APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
CONFIG_PATH = APP_DIR / "driver.json"
DEFAULT_CONFIG = {
    "version": "10.6.26.28275",
    "installer": "",
    "sha256": "",
    "pagina_oficial": "https://www.zebra.com/us/en/support-downloads/printers/industrial/zt411.html",
}

# Banderas de estado y atributos de la cola (winspool.h)
_STATUS = {
    0x00000001: (WARN, "La cola de impresión está en pausa en Windows"),
    0x00000002: (ERROR, "La impresora reporta un error"),
    0x00000008: (ERROR, "Atasco de papel"),
    0x00000010: (WARN, "Sin papel / etiquetas"),
    0x00000080: (ERROR, "La impresora está sin conexión"),
    0x00001000: (ERROR, "La impresora no está disponible"),
    0x00100000: (WARN, "La impresora requiere intervención del usuario"),
    0x00400000: (ERROR, "Tapa o cabezal abierto"),
}
_ATTR_WORK_OFFLINE = 0x400


@dataclass
class Check:
    level: str
    title: str
    detail: str = ""
    needs_driver: bool = False


def load_config() -> dict:
    cfg = dict(DEFAULT_CONFIG)
    try:
        cfg.update(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
    except (OSError, ValueError):
        pass
    return cfg


def _ver_str(v: int) -> str:
    return f"{v >> 48}.{(v >> 32) & 0xFFFF}.{(v >> 16) & 0xFFFF}.{v & 0xFFFF}"


def _ver_tuple(s: str) -> tuple[int, ...]:
    try:
        return tuple(int(p) for p in s.split("."))
    except ValueError:
        return (0,)


def installed_drivers() -> dict[str, str]:
    """{nombre del driver: versión} de los drivers de impresora instalados."""
    if win32print is None:
        return {}
    try:
        return {d["Name"]: _ver_str(d["DriverVersion"]) for d in win32print.EnumPrinterDrivers(None, None, 6)}
    except Exception:
        return {}


def printer_info(name: str) -> dict | None:
    if win32print is None:
        return None
    try:
        h = win32print.OpenPrinter(name)
    except Exception:
        return None
    try:
        return win32print.GetPrinter(h, 2)
    except Exception:
        return None
    finally:
        win32print.ClosePrinter(h)


def zebra_usb_devices() -> list[dict]:
    """Dispositivos USB de Zebra presentes (vendor 0A5F) con su estado en el Administrador de dispositivos."""
    ps = ("Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue | "
          "Where-Object { $_.InstanceId -like 'USB*VID_0A5F*' } | "
          "Select-Object Status, Class, FriendlyName, InstanceId | ConvertTo-Json -Compress")
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                             capture_output=True, text=True, timeout=20,
                             creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)).stdout.strip()
        if not out:
            return []
        data = json.loads(out)
        return data if isinstance(data, list) else [data]
    except Exception:
        return []


def query_host_status(ip: str, port: int = 9100, timeout: float = 2.0) -> dict | None:
    """Envía ~HS y devuelve las banderas principales, o None si no hubo respuesta válida."""
    with socket.create_connection((ip, int(port)), timeout=timeout) as s:
        s.sendall(b"~HS\r\n")
        buf, end = b"", time.time() + timeout
        while buf.count(b"\x03") < 3 and time.time() < end:
            try:
                chunk = s.recv(1024)
            except socket.timeout:
                break
            if not chunk:
                break
            buf += chunk
    parts = [p.strip(b"\x02\r\n ").decode("ascii", "ignore") for p in buf.split(b"\x03") if p.strip()]
    if len(parts) < 2:
        return None
    f1, f2 = parts[0].split(","), parts[1].split(",")
    if len(f1) < 3 or len(f2) < 4:
        return None
    return {"paper_out": f1[1] == "1", "paused": f1[2] == "1",
            "head_open": f2[2] == "1", "ribbon_out": f2[3] == "1"}


def check_network(ip: str, port: int, label: str | None = None) -> list[Check]:
    label = label or f"{ip}:{port}"
    try:
        st = query_host_status(ip, port)
    except OSError as exc:
        return [Check(ERROR, f"Sin conexión con {label}",
                      f"No se pudo conectar al puerto {port} ({exc}). Revise que la impresora esté "
                      "encendida, el cable de red y la dirección IP.")]
    checks = [Check(OK, f"La impresora responde en {label}")]
    if st is None:
        return checks
    if st["head_open"]:
        checks.append(Check(ERROR, "Cabezal abierto", "Cierre el cabezal de impresión."))
    if st["paper_out"]:
        checks.append(Check(WARN, "Sin papel / etiquetas", "Cargue el rollo y, si cambió de tamaño, calibre."))
    if st["ribbon_out"]:
        checks.append(Check(WARN, "Sin ribbon", "Coloque el ribbon (solo aplica en transferencia térmica)."))
    if st["paused"]:
        checks.append(Check(WARN, "La impresora está en pausa", "Presione el botón de pausa en la impresora."))
    if len(checks) == 1:
        checks[0].detail = "Lista para imprimir: con papel, sin pausa y cabezal cerrado."
    return checks


def _is_zebra_driver(drv: str) -> bool:
    low = drv.lower()
    return any(k in low for k in ("zdesigner", "zebra", "seagull", "generic / text only"))


def _driver_checks(name: str, info: dict, drivers: dict[str, str], cfg: dict) -> list[Check]:
    drv = info.get("pDriverName", "")
    ver = drivers.get(drv, "")
    shared = name.startswith("\\\\")
    low = drv.lower()
    rec = cfg["version"]
    where = (f" La cola es compartida desde {name.split(chr(92))[2]}: el driver se actualiza en ese "
             "servidor." if shared else "")
    if "zdesigner" in low:
        if ver and _ver_tuple(ver)[0] < _ver_tuple(rec)[0]:
            return [Check(WARN, f"Driver ZDesigner antiguo ({ver})",
                          f"Se recomienda la versión {rec} para la ZT411.{where}", needs_driver=not shared)]
        if ver and _ver_tuple(ver) < _ver_tuple(rec):
            return [Check(INFO, f"Driver ZDesigner {ver}",
                          f"Existe una versión más reciente ({rec}).{where}", needs_driver=not shared)]
        return [Check(OK, f"Driver ZDesigner {ver or ''}".strip(), "Driver oficial de Zebra (ZPL).")]
    if "generic / text only" in low:
        return [Check(OK, "Driver «Generic / Text Only»", "Funciona para enviar ZPL en modo RAW.")]
    if "seagull" in low or "zebra" in low:
        return [Check(INFO, f"Driver de terceros: {drv} {ver}".strip(),
                      "Funciona en modo RAW (BarTender/Seagull). El driver recomendado por Zebra es "
                      f"ZDesigner {rec}.")]
    return [Check(WARN, f"La impresora no parece ser Zebra ({drv})",
                  "Esta impresora no entiende ZPL: imprimiría el código como texto. Elija una ZT411.")]


def check_windows_printer(name: str, cfg: dict, drivers: dict[str, str]) -> list[Check]:
    if not name:
        return [Check(WARN, "No hay impresora seleccionada", "Elíjala en el encabezado.")]
    info = printer_info(name)
    if info is None:
        return [Check(ERROR, f"No se encontró la impresora «{name}»",
                      "Puede que se haya eliminado o que el servidor de impresión no esté disponible. "
                      "Actualice la lista (↻) o instale la impresora.",
                      needs_driver=not name.startswith("\\\\") and not any("zdesigner" in d.lower() for d in drivers))]
    checks = _driver_checks(name, info, drivers, cfg)
    if info.get("Attributes", 0) & _ATTR_WORK_OFFLINE:
        checks.append(Check(ERROR, "Windows marca la impresora «sin conexión»",
                            "La impresora USB está apagada o desconectada, o la cola tiene activado "
                            "«Usar impresora sin conexión».",))
    status = info.get("Status", 0)
    for flag, (level, text) in _STATUS.items():
        if status & flag:
            checks.append(Check(level, text))
    port = info.get("pPortName", "")
    # Cola TCP/IP estándar de una Zebra: se consulta el estado real. A otras marcas no se les
    # envía nada (imprimirían el comando ~HS como texto).
    if port.upper().startswith("IP_") and _is_zebra_driver(info.get("pDriverName", "")):
        checks += check_network(port[3:], 9100, f"{port[3:]} (puerto de la cola)")
    return checks


def check_usb(drivers: dict[str, str]) -> list[Check]:
    checks = []
    for dev in zebra_usb_devices():
        name = dev.get("FriendlyName") or "Zebra"
        if dev.get("Status") == "OK":
            checks.append(Check(OK, f"Zebra conectada por USB: {name}"))
        else:
            checks.append(Check(ERROR, f"Zebra USB sin driver: {name}",
                                "Windows detecta la impresora pero no tiene el driver instalado.",
                                needs_driver=True))
    return checks


def run(mode: str, printer: str, ip: str, port: str | int, scan_usb: bool = True) -> list[Check]:
    if sys.platform != "win32":
        return [Check(INFO, "Comprobación disponible solo en Windows")]
    cfg = load_config()
    drivers = installed_drivers()
    if mode == "red":
        checks = (check_network(ip, int(port or 9100)) if ip else
                  [Check(WARN, "Falta la IP de la impresora", "Escríbala en la pestaña «Impresora».")])
        checks.insert(0, Check(INFO, "Conexión directa por red", "No requiere driver de Windows."))
    else:
        checks = check_windows_printer(printer, cfg, drivers)
    if scan_usb:
        checks += check_usb(drivers)
    return checks


def worst(checks: list[Check]) -> str:
    return max((c.level for c in checks), key=SEVERITY.get, default=OK)


# --------------------------------------------------------------- instalación
def obtain_installer(cfg: dict, progress=None) -> Path:
    """Descarga (http/https) o copia (ruta local / UNC) el instalador y verifica su SHA-256."""
    src = cfg["installer"].strip()
    if not src:
        raise ValueError("No hay instalador configurado en driver.json")
    dest_dir = Path(tempfile.gettempdir()) / "EtiquetasZebra"
    dest_dir.mkdir(exist_ok=True)
    dest = dest_dir / Path(src.split("?")[0].replace("\\", "/")).name
    if src.lower().startswith(("http://", "https://")):
        with urllib.request.urlopen(src, timeout=30) as r, open(dest, "wb") as f:
            total, done = int(r.headers.get("Content-Length") or 0), 0
            while chunk := r.read(256 * 1024):
                f.write(chunk)
                done += len(chunk)
                if progress:
                    progress(done, total)
    else:
        shutil.copyfile(src, dest)
        if progress:
            progress(1, 1)
    expected = cfg.get("sha256", "").strip().lower()
    if expected:
        h = hashlib.sha256()
        with open(dest, "rb") as f:
            for block in iter(lambda: f.read(1024 * 1024), b""):
                h.update(block)
        if h.hexdigest() != expected:
            dest.unlink(missing_ok=True)
            raise ValueError("El instalador descargado no coincide con la huella SHA-256 esperada. "
                             "Se descartó por seguridad.")
    return dest


def run_installer(path: Path) -> int:
    """Ejecuta el instalador con elevación (UAC) y espera a que termine. Devuelve el código de salida."""
    import pywintypes
    import win32event
    import win32process
    from win32com.shell import shell, shellcon
    try:
        info = shell.ShellExecuteEx(fMask=shellcon.SEE_MASK_NOCLOSEPROCESS, lpVerb="runas",
                                    lpFile=str(path), nShow=1)
    except pywintypes.error as exc:
        if exc.winerror == 1223:
            raise RuntimeError("Se canceló la solicitud de permisos de administrador.") from None
        raise
    handle = info["hProcess"]
    win32event.WaitForSingleObject(handle, win32event.INFINITE)
    return win32process.GetExitCodeProcess(handle)
