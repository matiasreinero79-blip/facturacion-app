"""
pdf_reader.py — Extrae datos de facturas argentinas desde PDF o imagen.

Funciones públicas:
    extract_invoice_data(file_path, debug=False) -> dict
    parse_invoice(file_path) -> dict          ← alias para compatibilidad
    tesseract_available() -> bool
    supported_extensions() -> list[str]

Helpers exportados:
    normalize_ocr_text(text) -> str
    parse_argentinian_money(value) -> str
    parse_argentinian_date(value) -> str
"""

from __future__ import annotations

import os
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple

# ── Dependencias opcionales ──────────────────────────────────────────────────

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

    _TESSERACT_CANDIDATES = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.join(
            os.environ.get("LOCALAPPDATA", ""),
            r"Programs\Tesseract-OCR\tesseract.exe",
        ),
    ]
    for _tp in _TESSERACT_CANDIDATES:
        if os.path.isfile(_tp):
            pytesseract.pytesseract.tesseract_cmd = _tp
            break

    _HAS_TESSERACT = True
except ImportError:
    _HAS_TESSERACT = False

_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"}

# ── Meses en español ─────────────────────────────────────────────────────────

_MESES = {
    "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
    "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
    "septiembre": "09", "octubre": "10", "noviembre": "11", "diciembre": "12",
}

# ── Números que NO son importes ──────────────────────────────────────────────
# CAE tiene 14 dígitos, códigos de barra 13+, teléfonos ~10 dígitos
_FAKE_IMPORTE_MIN = 0.01
_FAKE_IMPORTE_MAX = 9_999_999.99   # más de 10 M → probablemente un código
_IMPORTE_DIGITS_MAX = 10           # ignorar si el número crudo tiene > 10 dígitos


# ═══════════════════════════════════════════════════════════════════════════════
#  Helpers de normalización
# ═══════════════════════════════════════════════════════════════════════════════

def normalize_ocr_text(text: str) -> str:
    """
    Limpia el texto extraído de un PDF/imagen:
    - Unifica saltos de línea
    - Reemplaza tabulaciones por espacios
    - Colapsa espacios múltiples dentro de líneas
    - Elimina líneas completamente vacías consecutivas
    """
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\t", " ")
    lines = []
    for line in text.split("\n"):
        line = re.sub(r" {2,}", " ", line).strip()
        lines.append(line)
    # Colapsar múltiples líneas vacías en una sola
    result = re.sub(r"\n{3,}", "\n\n", "\n".join(lines))
    return result


def parse_argentinian_money(value: str) -> str:
    """
    Convierte un string con formato argentino a decimal punto inglés.

    Ejemplos:
        '418.352,22'  → '418352.22'
        '202.706,74'  → '202706.74'
        '1.234'       → '1234.00'   (sin centavos)
        '123,45'      → '123.45'    (sólo centavos, sin miles)
        '$  1.234,56' → '1234.56'
    """
    if not value:
        return ""

    # Quitar símbolo de moneda y espacios
    v = re.sub(r"[\$\s]", "", value)

    # Formato argentino: punto como miles, coma como decimal
    # Detectar si hay coma decimal: X.XXX,XX
    if re.search(r"\d\.\d{3},\d{2}$", v):
        v = v.replace(".", "").replace(",", ".")
    # Sólo coma sin puntos de miles: 123,45
    elif re.match(r"^\d+,\d{1,2}$", v):
        v = v.replace(",", ".")
    # Sólo puntos (miles sin decimales): 1.234
    elif re.match(r"^\d{1,3}(\.\d{3})+$", v):
        v = v.replace(".", "")
    else:
        # Fallback: quitar todo excepto dígitos y punto/coma final
        v = re.sub(r"[^\d,\.]", "", v)
        v = v.replace(",", ".")

    try:
        num = float(v)
    except ValueError:
        return ""

    # Rechazar valores imposibles como importe de factura
    raw_digits = re.sub(r"\D", "", value)
    if len(raw_digits) > _IMPORTE_DIGITS_MAX:
        return ""
    if num < _FAKE_IMPORTE_MIN or num > _FAKE_IMPORTE_MAX:
        return ""

    return f"{num:.2f}"


def parse_argentinian_date(value: str) -> str:
    """
    Convierte fecha argentina a formato ISO YYYY-MM-DD.

    Soporta:
        '20/04/2026'           → '2026-04-20'
        '20-04-2026'           → '2026-04-20'
        '20 de abril de 2026'  → '2026-04-20'
        '20 de Abril 2026'     → '2026-04-20'
        '2026-04-20'           → '2026-04-20'  (ya en ISO)
    """
    if not value:
        return ""

    value = value.strip()

    # Ya en ISO
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", value)
    if m:
        return value

    # DD/MM/YYYY o DD-MM-YYYY o DD.MM.YYYY
    m = re.match(r"^(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{4})$", value)
    if m:
        d, mo, y = m.groups()
        return f"{y}-{mo.zfill(2)}-{d.zfill(2)}"

    # "20 de abril de 2026" o "20 de Abril 2026"
    m = re.match(
        r"^(\d{1,2})\s+de\s+([A-Za-záéíóúñÁÉÍÓÚÑ]+)\s+(?:de\s+)?(\d{4})$",
        value,
        re.IGNORECASE,
    )
    if m:
        d, mes_str, y = m.groups()
        mo = _MESES.get(mes_str.lower(), "")
        if mo:
            return f"{y}-{mo}-{d.zfill(2)}"

    return value  # devolver como vino si no reconocemos el formato


# ═══════════════════════════════════════════════════════════════════════════════
#  Extracción de texto
# ═══════════════════════════════════════════════════════════════════════════════

def _extract_text_pdf(path: str) -> str:
    text = ""
    if _HAS_PDFPLUMBER:
        try:
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    text += (page.extract_text() or "") + "\n"
            if text.strip():
                return text
        except Exception:
            pass

    if _HAS_PYPDF2:
        try:
            with open(path, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                for page in reader.pages:
                    text += (page.extract_text() or "") + "\n"
        except Exception:
            pass

    return text


def _extract_text_image(path: str) -> str:
    if not (_HAS_PIL and _HAS_TESSERACT):
        return ""
    try:
        img = Image.open(path).convert("RGB")
        w, h = img.size
        if w < 1400:
            scale = 1400 / w
            img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
        text = pytesseract.image_to_string(img, lang="spa+eng", config="--psm 6")
        return text
    except Exception:
        return ""


def _get_text(file_path: str) -> Tuple[str, bool]:
    """Devuelve (texto_normalizado, ocr_usado)."""
    ext = Path(file_path).suffix.lower()
    if ext in _IMAGE_EXTENSIONS:
        raw = _extract_text_image(file_path)
        return normalize_ocr_text(raw), True
    else:
        raw = _extract_text_pdf(file_path)
        return normalize_ocr_text(raw), False


# ═══════════════════════════════════════════════════════════════════════════════
#  Búsqueda de campos
# ═══════════════════════════════════════════════════════════════════════════════

def _search(patterns: list[str], text: str, flags: int = re.IGNORECASE | re.MULTILINE) -> str:
    """Aplica una lista de patrones y devuelve el primer match del grupo 1."""
    for pattern in patterns:
        m = re.search(pattern, text, flags)
        if m:
            return m.group(1).strip()
    return ""


def _find_importe(text: str) -> str:
    """
    Busca importe con prioridad:
      1. 'TOTAL A PAGAR'
      2. 'TOTAL FACTURA'
      3. 'IMPORTE TOTAL'
      4. 'TOTAL' genérico

    Usa parse_argentinian_money para validar y convertir.
    Ignora valores con demasiados dígitos (CAE, códigos de barra).
    """
    # Patrón de número argentino: acepta 1.234.567,89 o 1234567,89 o 1234.56
    _NUM = r"[\$\s]*([\d\.]+,\d{2}|\d+\.\d{3}(?:\.\d{3})*|\d+,\d{2}|\d+\.\d{2})"

    priority_labels = [
        r"total\s+a\s+pagar",
        r"total\s+a\s+abonar",
        r"importe\s+a\s+pagar",
        r"total\s+factura",
        r"importe\s+factura",
        r"importe\s+total",
        r"total\s+general",
        r"amount\s+due",
        r"total",
    ]

    for label in priority_labels:
        pattern = rf"(?:{label})\s*[:\$]?\s*{_NUM}"
        for m in re.finditer(pattern, text, re.IGNORECASE | re.MULTILINE):
            candidate = parse_argentinian_money(m.group(1))
            if candidate:
                return candidate

    return ""


def _find_numero_factura(text: str) -> str:
    """
    Busca número de comprobante AFIP: XXXX-XXXXXXXX
    Prioriza el formato con etiqueta, luego bare.
    """
    patterns = [
        # Con etiqueta
        r"(?:n[°º]?\s*(?:de\s+)?(?:comprobante|factura)|comprobante\s*n[°º]?|factura\s*n[°º]?)\s*[:\-]?\s*(\d{4}-\d{6,8})",
        # Sin etiqueta — número puro AFIP
        r"\b(\d{4}-\d{6,8})\b",
        # Fallback genérico
        r"(?:factura|invoice|n[°º])\s*[:\-\s]+([A-Z0-9][\w\-]{4,})",
    ]
    return _search(patterns, text)


def _find_tipo_factura(text: str) -> str:
    """Detecta tipo AFIP (A/B/C/E/M)."""
    patterns = [
        r"(?:factura\s+tipo|tipo\s+(?:de\s+)?(?:comprobante|factura))\s*[:\-]?\s*([ABCEMX])\b",
        r"(?:^|\b)FACTURA\s+([ABCEM])\b",
        r"(?:comprobante\s+)?([ABCEM])\s*[-–]\s*\d{4}-\d{6}",
    ]
    v = _search(patterns, text)
    return v.upper() if v else ""


def _find_cuit(text: str, label_hints: list[str] | None = None) -> str:
    """
    Busca CUIT con formato XX-XXXXXXXX-X.
    Si se dan `label_hints` (p.ej. ['proveedor','emisor']), prioriza esos contextos.
    """
    _CUIT_RE = r"(\d{2}[-\s]\d{7,8}[-\s]\d)"

    if label_hints:
        for hint in label_hints:
            m = re.search(
                rf"(?:{hint})[^\n]{{0,80}}{_CUIT_RE}",
                text,
                re.IGNORECASE | re.DOTALL,
            )
            if m:
                raw = re.sub(r"\s", "-", m.group(1))
                return _normalize_cuit(raw)

    # Todos los CUIT del documento — devolver el primero
    for m in re.finditer(_CUIT_RE, text):
        raw = re.sub(r"\s", "-", m.group(1))
        c = _normalize_cuit(raw)
        if c:
            return c
    return ""


def _normalize_cuit(raw: str) -> str:
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 11:
        return f"{digits[:2]}-{digits[2:10]}-{digits[10]}"
    return ""


def _find_fecha(text: str, label_patterns: list[str]) -> str:
    """Busca fecha usando etiquetas dadas y la convierte a ISO."""
    _DATE_BARE = r"(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{4})"
    _DATE_TEXTUAL = r"(\d{1,2}\s+de\s+[A-Za-záéíóúñ]+\s+(?:de\s+)?\d{4})"

    for label in label_patterns:
        for date_re in [_DATE_TEXTUAL, _DATE_BARE]:
            m = re.search(
                rf"(?:{label})\s*[:\-]?\s*{date_re}",
                text,
                re.IGNORECASE | re.MULTILINE,
            )
            if m:
                return parse_argentinian_date(m.group(1))

    # Buscar cualquier fecha textual en español
    for date_re in [_DATE_TEXTUAL, _DATE_BARE]:
        m = re.search(date_re, text, re.IGNORECASE)
        if m:
            return parse_argentinian_date(m.group(1))

    return ""


def _find_empresa(text: str) -> str:
    patterns = [
        r"(?:proveedor|emisor|raz[oó]n\s+social|from)\s*[:\-]?\s*([A-ZÁÉÍÓÚÑ][^\n]{3,60})",
        # Primera línea no vacía del documento suele ser el nombre de la empresa
    ]
    v = _search(patterns, text)
    if v:
        return v.strip().rstrip(",.:;")

    # Heurística: primera línea no vacía
    for line in text.split("\n"):
        line = line.strip()
        if len(line) > 3 and not re.match(r"^\d", line):
            return line
    return ""


def _find_cliente(text: str) -> str:
    patterns = [
        r"(?:cliente|titular|asociado|asegurado|sr\.?|sra\.?)\s*[:\-]?\s*([A-ZÁÉÍÓÚÑ][A-Za-záéíóúñ\s,\.]{5,60})",
        r"(?:nombre\s+(?:y\s+apellido|completo))\s*[:\-]?\s*([A-ZÁÉÍÓÚÑ][A-Za-záéíóúñ\s]{5,50})",
        r"(?:a\s+nombre\s+de)\s*[:\-]?\s*([A-ZÁÉÍÓÚÑ][A-Za-záéíóúñ\s]{5,50})",
    ]
    return _search(patterns, text)


def _find_numero_cliente(text: str) -> str:
    patterns = [
        r"(?:n[°º]?\s*(?:de\s+)?(?:cliente|asociado|afiliado|socio|beneficiario|p[oó]liza))\s*[:\-]?\s*([A-Z0-9][\w\-\.]{2,20})",
        r"(?:n[°º]\s*asociado|n[°º]\s*afiliado)\s*[:\-]?\s*([A-Z0-9][\w\-\.]{2,20})",
        r"(?:customer\s*(?:id|n[°º]|no|#))\s*[:\-]?\s*([A-Z0-9][\w\-]{2,20})",
    ]
    return _search(patterns, text)


def _find_numero_cuenta(text: str) -> str:
    patterns = [
        r"(?:n[°º]?\s*(?:de\s+)?cuenta)\s*[:\-]?\s*([A-Z0-9][\w\-]{2,20})",
        r"(?:account\s*(?:n[°º]|no|#))\s*[:\-]?\s*([A-Z0-9][\w\-]{2,20})",
    ]
    return _search(patterns, text)


def _find_dni_cliente(text: str) -> str:
    """Busca DNI/CUIL del cliente (número de 7-8 dígitos sin guiones)."""
    patterns = [
        r"(?:dni|d\.n\.i\.?)\s*[:\-]?\s*(\d{7,8})",
        r"(?:documento)\s*[:\-]?\s*(\d{7,8})",
    ]
    return _search(patterns, text)


# ═══════════════════════════════════════════════════════════════════════════════
#  Función principal
# ═══════════════════════════════════════════════════════════════════════════════

def extract_invoice_data(file_path: str, debug: bool = False) -> dict:
    """
    Extrae campos de una factura argentina (PDF o imagen).

    Siempre devuelve un dict con todas las claves.
    Campos no detectados quedan como cadena vacía.

    Args:
        file_path: Ruta al archivo PDF o imagen.
        debug:     Si True, imprime en consola el texto OCR y los campos hallados.

    Returns:
        dict con campos + metadatos (_ocr_used, _fields_found).
    """
    path = Path(file_path)
    is_image = path.suffix.lower() in _IMAGE_EXTENSIONS

    result: dict = {
        "numero_cuenta":  "",
        "numero_cliente": "",
        "empresa":        "",
        "numero_factura": "",
        "cuit":           "",
        "tipo_factura":   "",
        "importe":        "",
        "fecha_emision":  "",
        "fecha_envio":    "",   # vencimiento / fecha de envío
        "estado_pago":    "Pendiente",
        "archivo":        path.name,
        "ruta":           str(path.resolve()),
        "fecha_carga":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        # Metadatos (no se guardan en Excel)
        "_ocr_used":     False,
        "_is_image":     is_image,
        "_fields_found": 0,
        "_raw_text":     "",
    }

    text, ocr_used = _get_text(file_path)
    result["_ocr_used"] = ocr_used
    result["_raw_text"] = text

    if debug:
        print("\n" + "═" * 70)
        print(f"[DEBUG] Archivo: {path.name}")
        print(f"[DEBUG] OCR usado: {ocr_used}")
        print("[DEBUG] Texto extraído:")
        print("─" * 70)
        print(text[:3000])
        print("─" * 70)

    if not text.strip():
        if debug:
            print("[DEBUG] Sin texto extraído. Todos los campos vacíos.")
        return result

    # ── Extracción de campos ─────────────────────────────────────────────────

    result["tipo_factura"]   = _find_tipo_factura(text)
    result["numero_factura"] = _find_numero_factura(text)
    result["cuit"]           = _find_cuit(
        text,
        label_hints=["proveedor", "emisor", "c\.?u\.?i\.?t\.?\s*emisor", "cuit\s*proveedor"]
    )
    result["empresa"]        = _find_empresa(text)
    result["numero_cliente"] = _find_numero_cliente(text)
    result["numero_cuenta"]  = _find_numero_cuenta(text)

    # Cliente: primero buscar por etiqueta, luego DNI como referencia
    result["empresa"] = _find_empresa(text)
    cliente = _find_cliente(text)
    if not cliente:
        cliente = _find_dni_cliente(text)
    result["numero_cuenta"] = cliente if cliente else result["numero_cuenta"]

    result["fecha_emision"]  = _find_fecha(
        text,
        label_patterns=[
            r"fecha\s+de\s+emisi[oó]n",
            r"fecha\s+emisi[oó]n",
            r"fecha\s+factura",
            r"fecha",
        ],
    )
    result["fecha_envio"] = _find_fecha(
        text,
        label_patterns=[
            r"(?:fecha\s+de\s+)?vencimiento",
            r"vto\.?",
            r"fecha\s+de\s+pago",
            r"fecha\s+l[íi]mite",
        ],
    )

    result["importe"] = _find_importe(text)

    # ── Contar campos encontrados ────────────────────────────────────────────
    tracked = [
        "tipo_factura", "numero_factura", "cuit", "empresa",
        "numero_cliente", "importe", "fecha_emision", "fecha_envio",
    ]
    result["_fields_found"] = sum(1 for k in tracked if result.get(k))

    if debug:
        print("[DEBUG] Campos extraídos:")
        for k in tracked:
            print(f"  {k:20s}: {result.get(k, '')!r}")
        print(f"  {'importe':20s}: {result['importe']!r}")
        print(f"  Total campos: {result['_fields_found']}")
        print("═" * 70 + "\n")

    return result


# ── Alias de compatibilidad ───────────────────────────────────────────────────

def parse_invoice(file_path: str, debug: bool = False) -> dict:
    """Alias de extract_invoice_data para compatibilidad con código existente."""
    return extract_invoice_data(file_path, debug=debug)


# ── Utilidades ────────────────────────────────────────────────────────────────

def tesseract_available() -> bool:
    if not (_HAS_PIL and _HAS_TESSERACT):
        return False
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def supported_extensions() -> list:
    exts = ["pdf"]
    if tesseract_available():
        exts += ["jpg", "jpeg", "png", "bmp", "tiff", "webp"]
    return exts


# ── CLI de prueba ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python pdf_reader.py <archivo.pdf|imagen.jpg>")
        sys.exit(1)
    result = extract_invoice_data(sys.argv[1], debug=True)
    print("\nResultado final:")
    for k, v in result.items():
        if not k.startswith("_"):
            print(f"  {k:20s}: {v}")
