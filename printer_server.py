"""
Baraka POS Printer Server - Windows Edition
=============================================
Drop-in replacement for the Raspberry Pi server.py.
Uses FastAPI + python-escpos (same as original).

Auto-discovers printers on startup and updates the PRINTERS dict.
Adds WSA bridge (adb reverse) for Android access.
Prints test receipt on each printer at startup.
"""

import os
import sys
import uuid
import base64
import binascii
import json
import socket
import time
import logging
from typing import Optional
from datetime import datetime

from fastapi import FastAPI, UploadFile, HTTPException, Query, Request
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.exceptions import RequestValidationError
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
from PIL import Image, ImageDraw, ImageFont

# ESC/POS library
from escpos.printer import Network
from escpos.exceptions import Error as EscposError

# Local modules
from printer_discovery import PrinterDiscovery
from wsa_bridge import WSABridge
from print_queue import PrintQueue

# Load environment variables
load_dotenv()

# ─── Configuration ───────────────────────────────────────────
UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", "uploads")
ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "bmp", "gif"}
MAX_CONTENT_LENGTH = int(os.getenv("MAX_UPLOAD_SIZE_MB", "20")) * 1024 * 1024
SERVER_HOST = os.getenv("SERVER_HOST", "0.0.0.0")
SERVER_PORT = int(os.getenv("SERVER_PORT", "3006"))
SCAN_ON_STARTUP = os.getenv("SCAN_ON_STARTUP", "true").lower() in ("true", "1", "yes")
TEST_PRINT_ON_STARTUP = os.getenv("TEST_PRINT_ON_STARTUP", "true").lower() in ("true", "1", "yes")
WSA_BRIDGE_ENABLED = os.getenv("WSA_BRIDGE_ENABLED", "true").lower() in ("true", "1", "yes")
WSA_ADB_PORT = int(os.getenv("WSA_ADB_PORT", "58526"))
PRINTER_REGISTRY = os.getenv("PRINTER_REGISTRY", "printer_registry.json")
MIN_FEED_BEFORE_CUT = int(os.getenv("MIN_FEED_BEFORE_CUT", "4"))
QUEUE_MAX_RETRIES = int(os.getenv("QUEUE_MAX_RETRIES", "3"))
QUEUE_RETRY_BASE_DELAY = float(os.getenv("QUEUE_RETRY_BASE_DELAY", "1.0"))
QUEUE_JOB_HISTORY_SIZE = int(os.getenv("QUEUE_JOB_HISTORY_SIZE", "100"))

# Printer configurations - auto-populated by discovery
PRINTERS = {}

# Connection pool - reuse connections instead of creating new ones
_printer_connections = {}

# Print job queue (replaces CUPS spooler)
print_queue = PrintQueue(
    max_retries=QUEUE_MAX_RETRIES,
    retry_base_delay=QUEUE_RETRY_BASE_DELAY,
    history_size=QUEUE_JOB_HISTORY_SIZE,
)

# ─── Registry persistence ────────────────────────────────────
def load_registry():
    """Load printer registry from file, renumbering from printer_1."""
    global PRINTERS
    reg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), PRINTER_REGISTRY)
    if os.path.exists(reg_path):
        try:
            with open(reg_path, "r") as f:
                data = json.load(f)
            entries = sorted(data.values(), key=lambda x: x.get("name", ""))
            for i, info in enumerate(entries, start=1):
                PRINTERS[f"printer_{i}"] = {
                    "host": info.get("last_ip", info.get("host", "")),
                    "port": info.get("port", 9100),
                    "mac": info.get("mac", "unknown"),
                }
            logging.info(f"Loaded {len(PRINTERS)} printer(s) from registry")
        except Exception as e:
            logging.warning(f"Could not load registry: {e}")


def save_registry():
    """Save printer registry to file."""
    reg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), PRINTER_REGISTRY)
    data = {}
    for name, config in PRINTERS.items():
        data[name] = {
            "name": name,
            "last_ip": config["host"],
            "port": config["port"],
            "mac": config.get("mac", "unknown"),
            "last_seen": datetime.now().isoformat(),
        }
    try:
        with open(reg_path, "w") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        logging.warning(f"Could not save registry: {e}")


# ─── App setup ───────────────────────────────────────────────
app = FastAPI(title="Baraka Printer Server", version="2.0-windows")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# ─── Error handlers ──────────────────────────────────────────
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    if isinstance(exc, (HTTPException, RequestValidationError, EscposError)):
        raise exc
    print(f"[ERROR] Unhandled exception: {type(exc).__name__}: {str(exc)}")
    return JSONResponse(status_code=500, content={
        "success": False, "error": "Internal server error",
        "detail": str(exc), "type": type(exc).__name__
    })

@app.exception_handler(EscposError)
async def escpos_exception_handler(request: Request, exc: EscposError):
    print(f"[PRINTER ERROR] {str(exc)}")
    return JSONResponse(status_code=500, content={
        "success": False, "error": "Printer error",
        "detail": str(exc), "type": "PrinterError"
    })

@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    return JSONResponse(status_code=422, content={
        "success": False, "error": "Validation error",
        "detail": exc.errors(), "type": "ValidationError"
    })

@app.exception_handler(HTTPException)
async def http_exception_handler(request: Request, exc: HTTPException):
    return JSONResponse(status_code=exc.status_code, content={
        "success": False, "error": exc.detail, "type": "HTTPException"
    })


# ─── Helpers ─────────────────────────────────────────────────
def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def _init_printer(p: Network):
    """Send ESC/POS reset to clear any leftover state before printing."""
    p._raw(b'\x1b\x40')              # ESC @ -- hardware reset to factory defaults
    p._raw(b'\x1b\x45\x00')          # ESC E 0 -- emphasis OFF
    p._raw(b'\x1b\x47\x00')          # ESC G 0 -- double-strike OFF


def _connect_printer(printer_name: str) -> Network:
    """Get or create a printer connection. Raises RuntimeError on failure (safe for worker threads)."""
    if printer_name not in PRINTERS:
        raise RuntimeError(f"Unknown printer: {printer_name}. Available: {list(PRINTERS.keys())}")

    if printer_name in _printer_connections:
        conn = _printer_connections[printer_name]
        try:
            if conn.device is not None:
                conn.device.getpeername()
                return conn
        except (OSError, AttributeError):
            pass
        try:
            conn.close()
        except Exception:
            pass
        _printer_connections.pop(printer_name, None)

    config = PRINTERS[printer_name]
    try:
        printer = Network(config["host"], port=config["port"])
        _printer_connections[printer_name] = printer
        return printer
    except Exception as e:
        _printer_connections.pop(printer_name, None)
        raise RuntimeError(f"Cannot connect to {printer_name} ({config['host']}:{config['port']}): {e}")


def get_printer(printer_name: str) -> Network:
    """HTTP-handler wrapper around _connect_printer."""
    try:
        return _connect_printer(printer_name)
    except RuntimeError as e:
        raise HTTPException(status_code=500, detail=str(e))


def evict_printer_connection(printer_name: str):
    """Remove a broken connection from the pool so the next call gets a fresh one."""
    conn = _printer_connections.pop(printer_name, None)
    if conn:
        try:
            conn.close()
        except Exception:
            pass


def _prepare_image_for_thermal(filepath: str, paper_width: int) -> str:
    """
    Optimize an image for thermal printing using Pillow.
    Handles: RGBA/transparency, resize to paper width, gamma correction
    to eliminate gray shadows, and Floyd-Steinberg dithering to 1-bit.
    """
    from PIL import ImageEnhance, ImageOps

    img = Image.open(filepath)
    try:
        if img.mode == "RGBA" or "transparency" in img.info:
            background = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode == "RGBA":
                background.paste(img, mask=img.split()[3])
            else:
                background.paste(img)
            img = background
        elif img.mode != "RGB":
            img = img.convert("RGB")

        if img.width > paper_width:
            ratio = paper_width / img.width
            new_height = int(img.height * ratio)
            img = img.resize((paper_width, new_height), Image.Resampling.LANCZOS)

        img = ImageOps.autocontrast(img, cutoff=0.5)
        img = ImageEnhance.Contrast(img).enhance(1.3)
        img = ImageEnhance.Sharpness(img).enhance(1.2)

        bw = img.convert("L")

        gamma = 0.9
        lut = [min(255, int(255 * ((i / 255.0) ** gamma))) for i in range(256)]
        bw = bw.point(lut)

        bw = bw.point(lambda x: 255 if x > 245 else x)

        img = bw.convert("1", dither=Image.Dither.FLOYDSTEINBERG)

        img.save(filepath, "PNG")
    finally:
        img.close()

    return filepath


def close_printer_connection(p):
    """No-op - keep connections alive for speed."""
    pass

def get_local_ip():
    try:
        import subprocess, platform
        flags = subprocess.CREATE_NO_WINDOW if platform.system() == "Windows" else 0
        result = subprocess.run(["ipconfig"] if platform.system() == "Windows" else ["ip", "route"],
                                capture_output=True, text=True, creationflags=flags)
        import re
        for ip in re.findall(r"(\d+\.\d+\.\d+\.\d+)", result.stdout):
            if ip.startswith(("192.168.", "10.")) or re.match(r"172\.(1[6-9]|2\d|3[01])\.", ip):
                if not ip.startswith("172.30.") and not ip.endswith(".255") and not ip.endswith(".0"):
                    return ip
    except Exception:
        pass
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


# ─── Endpoints ───────────────────────────────────────────────

@app.get("/")
@app.get("/health")
def health():
    return {
        "ok": True,
        "status": "running",
        "message": "Thermal Printer API with python-escpos (Windows)",
        "printers": list(PRINTERS.keys()),
        "version": "2.0-windows",
    }

@app.get("/printers")
def get_printers():
    result = {}
    for name, config in PRINTERS.items():
        result[name] = {
            "host": config["host"],
            "port": config["port"],
            "mac": config.get("mac", "unknown"),
        }
    return {"printers": result}


def _merge_discovered_printers(found):
    """
    Merge discovered printers into PRINTERS dict.
    Matches by MAC first (handles IP changes), then by IP.
    Returns (new_count, updated_count).
    """
    new_count = 0
    updated_count = 0

    for ip, info in found.items():
        found_mac = info.get("mac") or "unknown"
        found_port = info.get("port", 9100)
        matched_name = None

        # Priority 1: match by MAC address (handles DHCP IP changes)
        if found_mac != "unknown":
            for name, config in PRINTERS.items():
                existing_mac = config.get("mac", "unknown")
                if existing_mac != "unknown" and existing_mac == found_mac:
                    matched_name = name
                    break

        # Priority 2: match by IP address
        if not matched_name:
            for name, config in PRINTERS.items():
                if config["host"] == ip:
                    matched_name = name
                    break

        if matched_name:
            config = PRINTERS[matched_name]
            old_ip = config["host"]
            if old_ip != ip:
                logging.info(
                    f"  {matched_name}: IP changed {old_ip} -> {ip} "
                    f"(MAC: {found_mac})"
                )
                config["host"] = ip
                evict_printer_connection(matched_name)
                updated_count += 1
            if found_mac != "unknown":
                config["mac"] = found_mac
            config["port"] = found_port
        else:
            existing_nums = []
            for n in PRINTERS:
                if n.startswith("printer_"):
                    try:
                        existing_nums.append(int(n.split("_")[1]))
                    except (ValueError, IndexError):
                        pass
            next_num = max(existing_nums, default=0) + 1
            printer_name = f"printer_{next_num}"

            PRINTERS[printer_name] = {
                "host": ip,
                "port": found_port,
                "mac": found_mac,
            }
            new_count += 1
            logging.info(f"  Discovered: {printer_name} @ {ip} (MAC: {found_mac})")

    return new_count, updated_count


@app.post("/printers/discover")
def discover_printers():
    """Re-scan network for printers."""
    discovery = PrinterDiscovery()
    found = discovery.scan()

    new_count, updated_count = _merge_discovered_printers(found)

    save_registry()
    return {
        "success": True,
        "discovered": new_count,
        "updated": updated_count,
        "total": len(PRINTERS),
        "printers": list(PRINTERS.keys()),
    }


# ─── Print Text ──────────────────────────────────────────────

@app.post("/print-text")
@app.post("/print/text")
async def print_text(
    text: str = Query(..., description="Text to print"),
    printer: str = Query("printer_1", description="Printer name"),
    printer_name: str = Query(None, description="Printer name (backward compatibility)"),
    lines_after: int = Query(0, description="Feed lines before cut"),
    cut: bool = Query(True, description="Auto cut after printing"),
    bold: bool = Query(False, description="Bold text"),
    underline: int = Query(0, description="Underline mode"),
    width: int = Query(1, description="Width multiplier (1-8)"),
    height: int = Query(1, description="Height multiplier (1-8)"),
    align: str = Query("left", description="Alignment"),
    invert: bool = Query(False, description="Invert colors")
):
    if printer_name:
        printer = printer_name
    if printer not in PRINTERS:
        raise HTTPException(status_code=400, detail=f"Unknown printer: {printer}. Available: {list(PRINTERS.keys())}")

    _text, _printer, _cut, _lines_after = text, printer, cut, lines_after
    _bold, _underline, _width, _height, _align, _invert = bold, underline, width, height, align, invert

    def execute():
        p = _connect_printer(_printer)
        try:
            _init_printer(p)
            use_custom_size = _width != 1 or _height != 1
            p.set(align=_align, bold=_bold, underline=_underline, invert=_invert,
                  width=_width, height=_height, custom_size=use_custom_size)
            p.text(_text)
            if not _text.endswith('\n'):
                p.text('\n')
            p.set()
            feed_lines = _lines_after if _lines_after > 0 else MIN_FEED_BEFORE_CUT
            if _cut:
                p.text('\n' * feed_lines)
                p.cut(feed=False)
            elif _lines_after > 0:
                p.text('\n' * _lines_after)
        except Exception:
            evict_printer_connection(_printer)
            raise

    job_id = print_queue.submit(printer, "text", execute, {
        "text": text[:100], "bold": bold, "align": align, "cut": cut,
    })
    return JSONResponse(status_code=202, content={
        "success": True, "job_id": job_id, "queued": True,
        "message": f"Text print job queued for {printer}", "printer": printer,
    })


# ─── Print Image (file upload) ───────────────────────────────

@app.post("/print-image")
@app.post("/print/image")
async def print_image(
    image: UploadFile,
    printer: str = Query("printer_1", description="Printer name"),
    printer_name: str = Query(None, description="Printer name (backward compatibility)"),
    lines_after: int = Query(0, description="Feed lines before cut"),
    cut: bool = Query(True, description="Auto cut after printing"),
    center: bool = Query(True, description="Center image"),
    paper_width: int = Query(510, description="Paper width in pixels")
):
    if printer_name:
        printer = printer_name
    if printer not in PRINTERS:
        raise HTTPException(status_code=400, detail=f"Unknown printer: {printer}. Available: {list(PRINTERS.keys())}")
    if not image.filename:
        raise HTTPException(status_code=400, detail="No image provided")
    if not allowed_file(image.filename):
        raise HTTPException(status_code=400, detail=f"Invalid image type. Allowed: {ALLOWED_EXTENSIONS}")

    filename = secure_filename(image.filename)
    unique_filename = f"{uuid.uuid4()}_{filename}"
    filepath = os.path.join(UPLOAD_FOLDER, unique_filename)

    content = await image.read()
    if len(content) > MAX_CONTENT_LENGTH:
        raise HTTPException(status_code=413, detail="File too large")

    with open(filepath, "wb") as f:
        f.write(content)

    _prepare_image_for_thermal(filepath, paper_width)

    _printer, _filepath, _center, _cut, _lines_after = printer, filepath, center, cut, lines_after

    def execute():
        try:
            p = _connect_printer(_printer)
            try:
                _init_printer(p)
                if _center:
                    p.set(align='center')
                p.image(_filepath)
                if _center:
                    p.set(align='left')
                feed_lines = _lines_after if _lines_after > 0 else MIN_FEED_BEFORE_CUT
                if _cut:
                    p.text('\n' * feed_lines)
                    p.cut(feed=False)
                elif _lines_after > 0:
                    p.text('\n' * _lines_after)
            except Exception:
                evict_printer_connection(_printer)
                raise
        finally:
            if os.path.exists(_filepath):
                try:
                    os.remove(_filepath)
                except Exception:
                    pass

    job_id = print_queue.submit(printer, "image", execute, {
        "filename": filename, "center": center, "cut": cut,
    })
    return JSONResponse(status_code=202, content={
        "success": True, "job_id": job_id, "queued": True,
        "message": f"Image print job queued for {printer}", "printer": printer,
        "filename": filename,
    })


# ─── Print QR Code ───────────────────────────────────────────

@app.post("/print/qr")
async def print_qr(
    text: str = Query(..., description="Text to encode"),
    printer: str = Query("printer_1", description="Printer name"),
    size: int = Query(3, description="QR code size (1-8)"),
    lines_after: int = Query(0, description="Feed lines before cut"),
    cut: bool = Query(True, description="Auto cut"),
    center: bool = Query(True, description="Center QR code")
):
    if printer not in PRINTERS:
        raise HTTPException(status_code=400, detail=f"Unknown printer: {printer}. Available: {list(PRINTERS.keys())}")

    _text, _printer, _size, _cut, _center, _lines_after = text, printer, size, cut, center, lines_after

    def execute():
        p = _connect_printer(_printer)
        try:
            _init_printer(p)
            if _center:
                p.set(align='center')
            p.qr(_text, size=_size)
            if _center:
                p.set(align='left')
            feed_lines = _lines_after if _lines_after > 0 else MIN_FEED_BEFORE_CUT
            if _cut:
                p.text('\n' * feed_lines)
                p.cut(feed=False)
            elif _lines_after > 0:
                p.text('\n' * _lines_after)
        except Exception:
            evict_printer_connection(_printer)
            raise

    job_id = print_queue.submit(printer, "qr", execute, {
        "text": text[:100], "size": size, "cut": cut,
    })
    return JSONResponse(status_code=202, content={
        "success": True, "job_id": job_id, "queued": True,
        "message": f"QR print job queued for {printer}", "printer": printer,
    })


# ─── Print Barcode ───────────────────────────────────────────

@app.post("/print/barcode")
async def print_barcode(
    code: str = Query(..., description="Barcode data"),
    printer: str = Query("printer_1", description="Printer name"),
    barcode_type: str = Query("CODE39", description="Barcode type"),
    height: int = Query(64, description="Barcode height"),
    width: int = Query(2, description="Barcode width"),
    lines_after: int = Query(0, description="Feed lines before cut"),
    cut: bool = Query(True, description="Auto cut"),
    center: bool = Query(True, description="Center barcode")
):
    if printer not in PRINTERS:
        raise HTTPException(status_code=400, detail=f"Unknown printer: {printer}. Available: {list(PRINTERS.keys())}")

    _code, _printer, _barcode_type = code, printer, barcode_type
    _height, _width, _cut, _center, _lines_after = height, width, cut, center, lines_after

    def execute():
        p = _connect_printer(_printer)
        try:
            _init_printer(p)
            if _center:
                p.set(align='center')
            p.barcode(_code, _barcode_type, height=_height, width=_width, pos='BELOW', font='A')
            if _center:
                p.set(align='left')
            feed_lines = _lines_after if _lines_after > 0 else MIN_FEED_BEFORE_CUT
            if _cut:
                p.text('\n' * feed_lines)
                p.cut(feed=False)
            elif _lines_after > 0:
                p.text('\n' * _lines_after)
        except Exception:
            evict_printer_connection(_printer)
            raise

    job_id = print_queue.submit(printer, "barcode", execute, {
        "code": code, "barcode_type": barcode_type, "cut": cut,
    })
    return JSONResponse(status_code=202, content={
        "success": True, "job_id": job_id, "queued": True,
        "message": f"Barcode print job queued for {printer}", "printer": printer,
    })


# ─── Job Queue Management ────────────────────────────────────

@app.get("/jobs")
def list_jobs(printer: str = Query(None, description="Filter by printer name")):
    """List all active and recent print jobs."""
    jobs = print_queue.get_queue(printer_name=printer)
    return {"success": True, "jobs": jobs, "count": len(jobs)}


@app.get("/jobs/{job_id}")
def get_job(job_id: str):
    """Get status of a specific print job."""
    job = print_queue.get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail=f"Job {job_id} not found")
    return {"success": True, "job": job}


@app.delete("/jobs/{job_id}")
def cancel_job(job_id: str):
    """Cancel a pending print job."""
    cancelled = print_queue.cancel_job(job_id)
    if not cancelled:
        raise HTTPException(status_code=400, detail=f"Job {job_id} cannot be cancelled (not pending or not found)")
    return {"success": True, "message": f"Job {job_id} cancelled"}


@app.get("/queue/status")
def queue_status():
    """Get queue health: pending counts per printer, worker status."""
    status = print_queue.get_status()
    return {"success": True, **status}


# ─── Cut ─────────────────────────────────────────────────────

@app.api_route("/cut", methods=["GET", "POST"])
async def cut_paper(
    printer: str = Query("printer_1"),
    printer_name: str = Query(None),
    lines_before: int = Query(0),
    feed: int = Query(None),
    mode: str = Query("partial")
):
    if printer_name:
        printer = printer_name
    if feed is not None:
        lines_before = feed
    cut_mode = "PART" if mode.lower() in ("partial", "part") else "FULL"
    try:
        p = get_printer(printer)
        _init_printer(p)
        feed_lines = lines_before if lines_before > 0 else MIN_FEED_BEFORE_CUT
        p.text('\n' * feed_lines)
        p.cut(mode=cut_mode, feed=False)
        close_printer_connection(p)
        return {"success": True, "message": f"Paper cut on {printer}", "printer": printer}
    except EscposError as e:
        evict_printer_connection(printer)
        raise HTTPException(status_code=500, detail=f"Printer error: {str(e)}")


# ─── Beep ────────────────────────────────────────────────────

@app.get("/beep")
@app.post("/beep")
async def beep(
    printer: str = Query("printer_1"),
    printer_name: str = Query(None),
    count: int = Query(1),
    duration: int = Query(1),
    beep_time: int = Query(None, alias="time")
):
    if printer_name:
        printer = printer_name
    if beep_time is not None:
        duration = beep_time
    try:
        p = get_printer(printer)
        _init_printer(p)
        count = max(1, min(9, count))
        duration = max(1, min(9, duration))
        try:
            p.buzzer(times=count, duration=duration)
        except Exception:
            p._raw(b'\x1b\x42' + bytes([count]) + bytes([duration]))
        close_printer_connection(p)
        return {"success": True, "message": f"Beep sent to {printer}", "printer": printer}
    except EscposError as e:
        evict_printer_connection(printer)
        raise HTTPException(status_code=500, detail=f"Printer error: {str(e)}")


# ─── Print Raw ───────────────────────────────────────────────

@app.post("/print-raw")
async def print_raw(
    printer: str = Query("printer_1"),
    printer_name: str = Query(None),
    base64_data: str = Query(None, alias="base64"),
    hex_data: str = Query(None, alias="hex")
):
    if printer_name:
        printer = printer_name
    if printer not in PRINTERS:
        raise HTTPException(status_code=400, detail=f"Unknown printer: {printer}. Available: {list(PRINTERS.keys())}")
    if not base64_data and not hex_data:
        raise HTTPException(status_code=400, detail="Provide 'base64' or 'hex' parameter")

    if base64_data:
        data = base64.b64decode(base64_data)
    else:
        data = binascii.unhexlify(hex_data.strip())

    _printer, _data = printer, data

    def execute():
        p = _connect_printer(_printer)
        try:
            p._raw(_data)
        except Exception:
            evict_printer_connection(_printer)
            raise

    job_id = print_queue.submit(printer, "raw", execute, {"bytes": len(data)})
    return JSONResponse(status_code=202, content={
        "success": True, "job_id": job_id, "queued": True,
        "message": f"Raw print job queued for {printer}", "printer": printer,
        "bytes": len(data),
    })


# ─── Cash Drawer ─────────────────────────────────────────────

@app.api_route("/drawer", methods=["GET", "POST"])
async def drawer(
    printer: str = Query("printer_1"),
    printer_name: str = Query(None),
    pin: int = Query(0),
    t1: int = Query(100),
    t2: int = Query(100)
):
    if printer_name:
        printer = printer_name
    try:
        p = get_printer(printer)
        _init_printer(p)
        pin_val = 0 if pin == 0 else 1
        t1_val = max(0, min(255, t1))
        t2_val = max(0, min(255, t2))
        p._raw(b'\x1b\x70' + bytes([pin_val, t1_val, t2_val]))
        close_printer_connection(p)
        return {"success": True, "message": f"Cash drawer pulse sent to {printer}", "printer": printer}
    except EscposError as e:
        evict_printer_connection(printer)
        raise HTTPException(status_code=500, detail=f"Printer error: {str(e)}")


# ─── Feed ────────────────────────────────────────────────────

@app.api_route("/feed", methods=["GET", "POST"])
async def feed(
    printer: str = Query("printer_1"),
    printer_name: str = Query(None),
    lines: int = Query(3)
):
    if printer_name:
        printer = printer_name
    try:
        p = get_printer(printer)
        _init_printer(p)
        lines_val = max(0, min(255, lines))
        p._raw(b'\x1b\x64' + bytes([lines_val]))
        close_printer_connection(p)
        return {"success": True, "message": f"Fed {lines_val} lines on {printer}", "printer": printer}
    except EscposError as e:
        evict_printer_connection(printer)
        raise HTTPException(status_code=500, detail=f"Printer error: {str(e)}")


# ─── Startup ─────────────────────────────────────────────────

def print_test_receipts():
    """Print a test receipt on every discovered printer."""
    local_ip = get_local_ip()
    import platform

    for name, config in PRINTERS.items():
        print(f"  Sending test print to {name} @ {config['host']}...")
        try:
            p = Network(config["host"], port=config["port"])

            # Generate test receipt image
            img = generate_test_image(local_ip, SERVER_PORT, name, config["host"], config.get("mac", "unknown"))
            filepath = os.path.join(UPLOAD_FOLDER, f"test_{name}.png")
            img.save(filepath)

            p.set(align='center')
            _prepare_image_for_thermal(filepath, 576)
            p.image(filepath)
            p.set(align='left')
            p.text('\n\n\n')
            p.cut()
            p.buzzer(times=3, duration=5)

            print(f"  [OK] {name} - test print sent")

            # Cleanup
            try:
                os.remove(filepath)
            except:
                pass
            try:
                p.close()
            except:
                pass

        except Exception as e:
            print(f"  [FAIL] {name} - {e}")

        time.sleep(2)


def generate_test_image(server_ip, server_port, printer_name, printer_ip, printer_mac):
    """Generate a professional test receipt image."""
    width = 576
    height = 900
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)

    def load_font(size, bold=False):
        paths = [
            "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/calibrib.ttf" if bold else "C:/Windows/Fonts/calibri.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ]
        for fp in paths:
            if os.path.exists(fp):
                try:
                    return ImageFont.truetype(fp, size)
                except:
                    pass
        return ImageFont.load_default()

    title_font = load_font(28, True)
    header_font = load_font(20, True)
    normal_font = load_font(16)
    small_font = load_font(14)
    pad = 20
    y = 20

    # Title
    for text, font in [("BARAKA POS", title_font), ("PRINTER SERVER TEST", header_font)]:
        bbox = draw.textbbox((0, 0), text, font=font)
        draw.text(((width - (bbox[2] - bbox[0])) // 2, y), text, fill="black", font=font)
        y += bbox[3] - bbox[1] + 15

    draw.line([(pad, y), (width - pad, y)], fill="black", width=2)
    y += 20

    # Server info
    draw.text((pad, y), "SERVER INFORMATION", fill="black", font=header_font)
    y += 30
    import platform as plat
    for label, value in [
        ("Server IP:", server_ip),
        ("Server Port:", str(server_port)),
        ("URL:", f"http://{server_ip}:{server_port}"),
        ("Platform:", f"{plat.system()} {plat.release()}"),
        ("Hostname:", plat.node()),
    ]:
        draw.text((pad, y), label, fill="black", font=normal_font)
        draw.text((pad + 160, y), value, fill="black", font=normal_font)
        y += 25

    y += 10
    draw.line([(pad, y), (width - pad, y)], fill="black", width=2)
    y += 20

    # Printer info
    draw.text((pad, y), "PRINTER INFORMATION", fill="black", font=header_font)
    y += 30
    for label, value in [
        ("Printer Name:", printer_name),
        ("Printer IP:", printer_ip),
        ("Printer MAC:", printer_mac),
        ("Port:", "9100"),
    ]:
        draw.text((pad, y), label, fill="black", font=normal_font)
        draw.text((pad + 160, y), value, fill="black", font=normal_font)
        y += 25

    y += 10
    draw.line([(pad, y), (width - pad, y)], fill="black", width=2)
    y += 20

    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    bbox = draw.textbbox((0, 0), ts, font=small_font)
    draw.text(((width - (bbox[2] - bbox[0])) // 2, y), ts, fill="black", font=small_font)
    y += 30

    draw.line([(pad, y), (width - pad, y)], fill="black", width=2)
    y += 25

    for msg in ["Printer is working correctly!", "Ready to accept print jobs."]:
        bbox = draw.textbbox((0, 0), msg, font=normal_font)
        draw.text(((width - (bbox[2] - bbox[0])) // 2, y), msg, fill="black", font=normal_font)
        y += 25

    return img.crop((0, 0, width, y + 10))


# ─── Main ────────────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    import sys

    # When running via pythonw.exe (hidden mode), stdout/stderr are None.
    # Redirect ALL output to a log file before anything else runs.
    log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "server.log")
    _hidden_mode = sys.stdout is None or not hasattr(sys.stdout, "write")
    if _hidden_mode:
        _log_fh = open(log_file, "a", encoding="utf-8", buffering=1)
        sys.stdout = _log_fh
        sys.stderr = _log_fh

    # Remove any stale handlers that were auto-created with stream=None
    root_logger = logging.getLogger()
    for h in root_logger.handlers[:]:
        root_logger.removeHandler(h)

    log_handler = logging.StreamHandler(sys.stdout)
    log_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    root_logger.setLevel(logging.INFO)
    root_logger.addHandler(log_handler)

    if not _hidden_mode:
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        root_logger.addHandler(file_handler)

    print("=" * 50)
    print("Baraka Printer Server - Windows Edition v2.0")
    print("=" * 50)

    # Load existing registry
    load_registry()

    # Auto-discover printers
    if SCAN_ON_STARTUP:
        print(f"\nScanning network for printers...")
        discovery = PrinterDiscovery()
        found = discovery.scan()

        new_count, updated_count = _merge_discovered_printers(found)

        if not PRINTERS:
            print("  No printers found. Add them via POST /printers/discover")
        else:
            if updated_count:
                print(f"  Updated {updated_count} printer IP(s)")
            if new_count:
                print(f"  Found {new_count} new printer(s)")
            print(f"  Total printers: {len(PRINTERS)}")
            save_registry()

    print(f"\nAvailable printers: {list(PRINTERS.keys())}")
    for name, config in PRINTERS.items():
        print(f"  {name}: {config['host']}:{config['port']} (MAC: {config.get('mac', 'unknown')})")

    # Test print
    if TEST_PRINT_ON_STARTUP and PRINTERS:
        print(f"\nPrinting test receipts...")
        print_test_receipts()

    # WSA Bridge
    if WSA_BRIDGE_ENABLED:
        bridge = WSABridge(adb_port=WSA_ADB_PORT, server_port=SERVER_PORT)
        bridge.setup()

    print(f"\nStarting server on {SERVER_HOST}:{SERVER_PORT}")
    print(f"  Local:   http://localhost:{SERVER_PORT}")
    print(f"  Network: http://{get_local_ip()}:{SERVER_PORT}")
    print(f"  Health:  http://localhost:{SERVER_PORT}/health")
    print(f"\nFeatures: Text, Images, QR codes, Barcodes, Cash drawer, Feed")
    print(f"Print queue: max_retries={QUEUE_MAX_RETRIES}, retry_delay={QUEUE_RETRY_BASE_DELAY}s")
    print(f"No CUPS, No PPD, Direct network printing!\n")

    uvicorn.run(app, host=SERVER_HOST, port=SERVER_PORT)
