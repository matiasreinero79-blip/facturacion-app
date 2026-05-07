import re
from pathlib import Path
from datetime import datetime
from typing import Optional

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


# Patterns ordered from most-specific to least-specific
_PATTERNS = {
    "numero_factura": [
        r"(?:factura\s*(?:n[°º]|nro|num|#)\.?\s*)([A-Z0-9][\w\-]+)",
        r"(?:invoice\s*(?:n[°º]|no|#)\.?\s*)([A-Z0-9][\w\-]+)",
        r"(?:comprobante\s*(?:n[°º]|nro|#)\.?\s*)([A-Z0-9][\w\-]+)",
        r"(?:N[°º]\s*(?:de\s+)?factura\s*[:\-]?\s*)([A-Z0-9][\w\-]+)",
    ],
    "importe": [
        r"(?:total\s+(?:a\s+pagar|general|importe)\s*[:\$]?\s*)([\d.,]+)",
        r"(?:importe\s+total\s*[:\$]?\s*)([\d.,]+)",
        r"(?:amount\s+due\s*[:\$]?\s*)([\d.,]+)",
        r"(?:total\s*[:\$]\s*)([\d.,]+)",
        r"\$\s*([\d.,]+)",
    ],
    "fecha_emision": [
        r"(?:fecha\s+(?:de\s+)?emisi[oó]n\s*[:\-]?\s*)(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})",
        r"(?:fecha\s*[:\-]?\s*)(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})",
        r"(?:date\s*[:\-]?\s*)(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})",
    ],
    "empresa": [
        r"(?:(?:proveedor|emisor|razón\s+social|from)\s*[:\-]?\s*)([A-ZÁÉÍÓÚÑ][A-Za-záéíóúñ\s,\.]+(?:S\.?A\.?|S\.?R\.?L\.?|S\.?A\.?S\.?|INC\.?)?)",
    ],
    "numero_cuenta": [
        r"(?:n[°º]?\s*(?:de\s+)?cuenta\s*[:\-]?\s*)([A-Z0-9][\w\-]+)",
        r"(?:account\s*(?:n[°º]|no|#)\.?\s*[:\-]?\s*)([A-Z0-9][\w\-]+)",
    ],
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


def _normalize_date(raw: str) -> str:
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(raw.strip(), fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return raw.strip()


def _extract_text(pdf_path: str) -> str:
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


def parse_invoice(pdf_path: str) -> dict:
    """
    Attempt to extract invoice fields from a PDF.
    Always returns a dict with all expected keys; undetected fields are empty strings.
    """
    path = Path(pdf_path)
    result = {
        "numero_cuenta": "",
        "numero_cliente": "",
        "empresa": "",
        "numero_factura": "",
        "importe": "",
        "fecha_emision": "",
        "fecha_envio": "",
        "estado_pago": "Pendiente",
        "archivo": path.name,
        "ruta": str(path.resolve()),
        "fecha_carga": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    text = _extract_text(pdf_path)
    if not text.strip():
        return result

    for field, patterns in _PATTERNS.items():
        for pattern in patterns:
            m = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
            if m:
                result[field] = m.group(1).strip()
                break

    # Clean importe: keep only digits, comma, dot
    if result["importe"]:
        result["importe"] = re.sub(r"[^\d.,]", "", result["importe"])

    # Normalise date
    if result["fecha_emision"]:
        result["fecha_emision"] = _normalize_date(result["fecha_emision"])

    return result
