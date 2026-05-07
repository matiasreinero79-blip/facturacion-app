"""
pdf_reader.py — Extracts invoice fields from PDFs and images.

Supports:
  - PDF (pdfplumber → PyPDF2 fallback)
  - Images PNG/JPG via pytesseract OCR (requires Tesseract installed)

Argentine AFIP invoice fields supported:
  numero_factura, tipo_factura, cuit, empresa, importe,
  fecha_emision, numero_cuenta, numero_cliente
"""

import re
import os
from pathlib import Path
from datetime import datetime
from typing import Optional, Tuple

try:
    import pdfplumber
    _HAS_PDFPLUMBER = True
except ImportError:
    _HAS_PDFPLUMBER = False

try:
    import PyPDF2
    _HAS_PYPDF2 = True
except ImportError:
    _HAS_PYPDF2 = False

try:
    from PIL import Image
    _HAS_PIL = True
except ImportError:
    _HAS_PIL = False

try:
    import pytesseract
    # Auto-detect Tesseract on Windows common install paths
    _TESSERACT_CANDIDATES = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        r"C:\Users\{}\AppData\Local\Programs\Tesseract-OCR\tesseract.exe".format(
            os.environ.get("USERNAME", "")
        ),
    ]
    for _tpath in _TESSERACT_CANDIDATES:
        if os.path.exists(_tpath):
            pytesseract.pytesseract.tesseract_cmd = _tpath
            break
    _HAS_TESSERACT = True
except ImportError:
    _HAS_TESSERACT = False


# ─── Field extraction patterns ───────────────────────────────────────────────
#
# Argentine AFIP invoices have a very specific format:
#   - Type is shown as a big letter (A/B/C/E/M) in a box
#   - Invoice number: XXXX-XXXXXXXX (4 digits - 8 digits)
#   - CUIT: XX-XXXXXXXX-X
#
_PATTERNS = {
    # ── Tipo de factura (AFIP: A, B, C, E, M) ───────────────────────────────
    "tipo_factura": [
        r"(?:factura\s+tipo|tipo\s+(?:de\s+)?factura|comprobante\s+tipo)[:\s]+([ABCEMX])\b",
        r"(?:^|\s)FACTURA\s+([ABCEMX])\b",
        # Letter alone in a prominent position (AFIP layout)
        r"^\s*([ABCEM])\s*$",
    ],

    # ── Número de factura (AFIP: XXXX-XXXXXXXX) ──────────────────────────────
    "numero_factura": [
        # AFIP standard: 0001-00000123
        r"(?:n[°º]?\s*(?:de\s+)?(?:comprobante|factura)|comprobante\s+n[°º]?)[:\s]+(\d{4}-\d{6,8})",
        r"\b(\d{4}-\d{6,8})\b",
        # Generic invoice number
        r"(?:factura\s*(?:n[°º]|nro|num|#)\.?\s*)([A-Z0-9][\w\-]+)",
        r"(?:invoice\s*(?:n[°º]|no|#)\.?\s*)([A-Z0-9][\w\-]+)",
        r"(?:comprobante\s*(?:n[°º]|nro|#)\.?\s*)([A-Z0-9][\w\-]+)",
        r"(?:N[°º]\s*(?:de\s+)?factura\s*[:\-]?\s*)([A-Z0-9][\w\-]+)",
    ],

    # ── CUIT (XX-XXXXXXXX-X) ─────────────────────────────────────────────────
    "cuit": [
        r"(?:c\.?u\.?i\.?t\.?\s*[:\-]?\s*)(\d{2}-\d{7,8}-\d)",
        r"(?:cuit\s*[:\-]?\s*)(\d{2}-\d{7,8}-\d)",
        # Without label — raw CUIT format
        r"\b(\d{2}-\d{7,8}-\d)\b",
    ],

    # ── Importe ──────────────────────────────────────────────────────────────
    "importe": [
        r"(?:total\s+(?:a\s+pagar|general|importe)\s*[:\$]?\s*)\$?\s*([\d.,]+)",
        r"(?:importe\s+total\s*[:\$]?\s*)\$?\s*([\d.,]+)",
        r"(?:amount\s+due\s*[:\$]?\s*)\$?\s*([\d.,]+)",
        r"(?:total\s*[:\$]\s*)\$?\s*([\d.,]+)",
        r"\$\s*([\d.,]+)",
    ],

    # ── Fecha de emisión ─────────────────────────────────────────────────────
    "fecha_emision": [
        r"(?:fecha\s+(?:de\s+)?emisi[oó]n\s*[:\-]?\s*)(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})",
        r"(?:fecha\s*[:\-]?\s*)(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})",
        r"(?:date\s*[:\-]?\s*)(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})",
    ],

    # ── Empresa / Razón Social ────────────────────────────────────────────────
    "empresa": [
        r"(?:(?:proveedor|emisor|raz[oó]n\s+social|from)\s*[:\-]?\s*)([A-ZÁÉÍÓÚÑ][A-Za-záéíóúñ\s,\.]+(?:S\.?A\.?|S\.?R\.?L\.?|S\.?A\.?S\.?|INC\.?)?)",
    ],

    # ── Número de cuenta ─────────────────────────────────────────────────────
    "numero_cuenta": [
        r"(?:n[°º]?\s*(?:de\s+)?cuenta\s*[:\-]?\s*)([A-Z0-9][\w\-]+)",
        r"(?:account\s*(?:n[°º]|no|#)\.?\s*[:\-]?\s*)([A-Z0-9][\w\-]+)",
    ],

    # ── Número de cliente ─────────────────────────────────────────────────────
    "numero_cliente": [
        r"(?:n[°º]?\s*(?:de\s+)?cliente\s*[:\-]?\s*)([A-Z0-9][\w\-]+)",
        r"(?:customer\s*(?:id|n[°º]|no|#)\.?\s*[:\-]?\s*)([A-Z0-9][\w\-]+)",
    ],
}

_DATE_FORMATS = [
    "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
    "%d/%m/%y", "%d-%m-%y", "%d.%m.%y",
    "%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d",
]

_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"}


def _normalize_date(raw: str) -> str:
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(raw.strip(), fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return raw.strip()


def _extract_text_from_pdf(pdf_path: str) -> str:
    text = ""
    if _HAS_PDFPLUMBER:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text() or ""
                    text += page_text + "\n"
            if text.strip():
                return text
        except Exception:
            pass

    if _HAS_PYPDF2:
        try:
            with open(pdf_path, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                for page in reader.pages:
                    text += (page.extract_text() or "") + "\n"
        except Exception:
            pass

    return text


def _extract_text_from_image(image_path: str) -> Tuple[str, bool]:
    """
    Returns (text, ocr_used).
    ocr_used=True means tesseract was invoked successfully.
    """
    if not (_HAS_PIL and _HAS_TESSERACT):
        return "", False
    try:
        img = Image.open(image_path)
        # Upscale small images for better OCR accuracy
        w, h = img.size
        if w < 1200:
            scale = 1200 / w
            img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
        text = pytesseract.image_to_string(img, lang="spa+eng")
        return text, True
    except Exception:
        return "", False


def _extract_text(file_path: str) -> Tuple[str, bool]:
    """
    Dispatch to PDF or image extractor.
    Returns (text, ocr_used).
    """
    ext = Path(file_path).suffix.lower()
    if ext in _IMAGE_EXTENSIONS:
        return _extract_text_from_image(file_path)
    else:
        text = _extract_text_from_pdf(file_path)
        return text, False


def _detect_tipo_factura(text: str) -> str:
    """
    Special detection for AFIP invoice type (A/B/C/E/M).
    Looks for a single prominent letter in the document.
    """
    # Try patterns first
    for pattern in _PATTERNS["tipo_factura"]:
        m = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if m:
            return m.group(1).upper()
    return ""


def parse_invoice(file_path: str) -> dict:
    """
    Attempt to extract invoice fields from a PDF or image file.
    Always returns a dict with all expected keys; undetected fields are empty strings.
    Also returns metadata about extraction confidence.
    """
    path = Path(file_path)
    is_image = path.suffix.lower() in _IMAGE_EXTENSIONS

    result = {
        "numero_cuenta":  "",
        "numero_cliente": "",
        "empresa":        "",
        "numero_factura": "",
        "cuit":           "",
        "tipo_factura":   "",
        "importe":        "",
        "fecha_emision":  "",
        "fecha_envio":    "",
        "estado_pago":    "Pendiente",
        "archivo":        path.name,
        "ruta":           str(path.resolve()),
        "fecha_carga":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        # Metadata (not saved to Excel, used by UI)
        "_ocr_used":      False,
        "_is_image":      is_image,
        "_fields_found":  0,
    }

    text, ocr_used = _extract_text(file_path)
    result["_ocr_used"] = ocr_used

    if not text.strip():
        return result

    # Extract tipo_factura with dedicated logic
    result["tipo_factura"] = _detect_tipo_factura(text)

    # Extract remaining fields with pattern matching
    extractable = [k for k in _PATTERNS if k != "tipo_factura"]
    fields_found = 1 if result["tipo_factura"] else 0

    for field in extractable:
        patterns = _PATTERNS[field]
        for pattern in patterns:
            m = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
            if m:
                result[field] = m.group(1).strip()
                fields_found += 1
                break

    result["_fields_found"] = fields_found

    # Clean importe: keep only digits, comma, dot
    if result["importe"]:
        result["importe"] = re.sub(r"[^\d.,]", "", result["importe"])

    # Normalise date
    if result["fecha_emision"]:
        result["fecha_emision"] = _normalize_date(result["fecha_emision"])

    # Clean CUIT: ensure format XX-XXXXXXXX-X
    if result["cuit"]:
        digits = re.sub(r"\D", "", result["cuit"])
        if len(digits) == 11:
            result["cuit"] = f"{digits[:2]}-{digits[2:10]}-{digits[10]}"

    return result


def tesseract_available() -> bool:
    """Check whether Tesseract OCR is usable."""
    if not (_HAS_PIL and _HAS_TESSERACT):
        return False
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def supported_extensions() -> list:
    """Return list of supported file extensions."""
    exts = ["pdf"]
    if tesseract_available():
        exts += ["jpg", "jpeg", "png", "bmp", "tiff", "webp"]
    return exts
