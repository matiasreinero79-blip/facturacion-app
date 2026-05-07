"""
pdf_reader.py — Parser de facturas argentinas.
Extracts: empresa, numero_factura, cuit, tipo_factura, importe,
          fecha_emision, fecha_envio, numero_cliente, numero_cuenta.
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
    for _tp in [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Programs\Tesseract-OCR\tesseract.exe"),
    ]:
        if os.path.isfile(_tp):
            pytesseract.pytesseract.tesseract_cmd = _tp
            break
    _HAS_TESSERACT = True
except ImportError:
    _HAS_TESSERACT = False

_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"}

_MESES = {
    "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
    "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
    "septiembre": "09", "octubre": "10", "noviembre": "11", "diciembre": "12",
}


# ═══════════════════════════════════════════════════════════════════════════════
#  Limpieza OCR
# ═══════════════════════════════════════════════════════════════════════════════

def normalize_ocr_text(text: str) -> str:
    """Limpia artefactos comunes del OCR antes de aplicar regex."""
    # Unificar saltos de línea
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\t", " ")

    lines = []
    for line in text.split("\n"):
        # Colapsar espacios internos múltiples
        line = re.sub(r" {2,}", " ", line).strip()
        lines.append(line)

    # Colapsar 3+ líneas vacías consecutivas en 1
    text = re.sub(r"\n{3,}", "\n\n", "\n".join(lines))
    return text


# ═══════════════════════════════════════════════════════════════════════════════
#  Conversión de formatos argentinos
# ═══════════════════════════════════════════════════════════════════════════════

def parse_argentinian_money(raw: str) -> str:
    """
    Convierte importe argentino a float string con punto decimal.
    Rechaza:
      - Strings con más de 10 dígitos (CAE, código de barra)
      - Valores > 9.999.999,99 (imposible como factura)
      - Strings que sean sólo 1-3 dígitos (demasiado cortos para ser importe)
    """
    if not raw:
        return ""

    raw = raw.strip()

    # Rechazar si hay demasiados dígitos (CAE tiene 14, código de barra 13+)
    only_digits = re.sub(r"\D", "", raw)
    if len(only_digits) > 10:
        return ""

    # Quitar $, espacios, letras sueltas
    v = re.sub(r"[^\d,\.]", "", raw)

    if not v:
        return ""

    # Formato argentino: 418.352,22
    if re.match(r"^\d{1,3}(\.\d{3})+,\d{2}$", v):
        v = v.replace(".", "").replace(",", ".")
    # 1.234 (miles sin centavos)
    elif re.match(r"^\d{1,3}(\.\d{3})+$", v):
        v = v.replace(".", "")
    # 123,45 (coma decimal, sin miles)
    elif re.match(r"^\d+,\d{2}$", v):
        v = v.replace(",", ".")
    # 123.45 (punto decimal)
    elif re.match(r"^\d+\.\d{2}$", v):
        pass  # ya está bien
    # Solo dígitos
    elif re.match(r"^\d+$", v):
        pass
    else:
        # Último intento: reemplazar coma por punto
        v = v.replace(",", ".")

    try:
        num = float(v)
    except ValueError:
        return ""

    # Validar rango razonable para una factura
    if num < 0.01 or num > 9_999_999.99:
        return ""

    # Rechazar si tiene muy pocos dígitos (sería "por" o similar)
    if len(only_digits) < 2:
        return ""

    return f"{num:.2f}"


def parse_argentinian_date(raw: str) -> str:
    """Convierte fecha argentina a ISO YYYY-MM-DD."""
    if not raw:
        return ""
    raw = raw.strip()

    # Ya en ISO
    if re.match(r"^\d{4}-\d{2}-\d{2}$", raw):
        return raw

    # DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY
    m = re.match(r"^(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{4})$", raw)
    if m:
        d, mo, y = m.groups()
        return f"{y}-{mo.zfill(2)}-{d.zfill(2)}"

    # "20 de abril de 2026" o "20 de Abril 2026"
    m = re.match(
        r"^(\d{1,2})\s+de\s+([A-Za-záéíóúñÁÉÍÓÚÑ]+)\s+(?:de\s+)?(\d{4})$",
        raw,
        re.IGNORECASE,
    )
    if m:
        d, mes_str, y = m.groups()
        mo = _MESES.get(mes_str.lower())
        if mo:
            return f"{y}-{mo}-{d.zfill(2)}"

    return raw


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
        return pytesseract.image_to_string(img, lang="spa+eng", config="--psm 6")
    except Exception:
        return ""


def _get_text(file_path: str) -> Tuple[str, bool]:
    ext = Path(file_path).suffix.lower()
    if ext in _IMAGE_EXTENSIONS:
        raw = _extract_text_image(file_path)
        return normalize_ocr_text(raw), True
    else:
        raw = _extract_text_pdf(file_path)
        return normalize_ocr_text(raw), False


# ═══════════════════════════════════════════════════════════════════════════════
#  Extracción de campos — reglas estrictas
# ═══════════════════════════════════════════════════════════════════════════════

# Patrón de número argentino (punto-miles, coma-decimal)
_MONEY_PAT = r"(\d{1,3}(?:\.\d{3})*,\d{2}|\d{1,3}(?:\.\d{3})+|\d+,\d{2}|\d+\.\d{2})"


def _find_importe(text: str) -> str:
    """
    Busca el importe MAYOR entre los candidatos con etiquetas de alta prioridad.
    Prioridad de etiquetas:
     1. TOTAL A PAGAR / SALDO / TOTAL A ABONAR
     2. TOTAL FACTURA / IMPORTE FACTURA
     3. IMPORTE TOTAL / TOTAL GENERAL
     4. TOTAL: (genérico — solo si tiene al menos 6 dígitos en el número)
    Rango válido: entre $1.000 y $9.999.999.
    Ignora candidatos con más de 10 dígitos (CAE, códigos de barra).
    """
    # Número argentino: acepta 1.234.567,89 | 418.352,22 | 12345,67 | 1234.56
    _NUM = r"(\d{1,3}(?:\.\d{3})*,\d{2}|\d{1,3}(?:\.\d{3})+|\d{4,},\d{2}|\d{4,}\.\d{2})"

    priority_groups = [
        # Grupo 1 — máxima prioridad
        [
            rf"total\s+a\s+(?:pagar|abonar)\s*[:\$]?\s*\$?\s*{_NUM}",
            rf"saldo\s+(?:a\s+pagar|deudor|pendiente)\s*[:\$]?\s*\$?\s*{_NUM}",
            rf"importe\s+a\s+pagar\s*[:\$]?\s*\$?\s*{_NUM}",
        ],
        # Grupo 2
        [
            rf"total\s+factura\s*[:\$]?\s*\$?\s*{_NUM}",
            rf"importe\s+(?:de\s+)?factura\s*[:\$]?\s*\$?\s*{_NUM}",
        ],
        # Grupo 3
        [
            rf"importe\s+total\s*[:\$]?\s*\$?\s*{_NUM}",
            rf"total\s+general\s*[:\$]?\s*\$?\s*{_NUM}",
            rf"amount\s+due\s*[:\$]?\s*\$?\s*{_NUM}",
        ],
        # Grupo 4 — genérico (requiere >= 6 dígitos en el número para evitar falsos)
        [
            rf"(?<!\w)total\s*[:\$]\s*\$?\s*(\d{{1,3}}(?:\.\d{{3}})*,\d{{2}}|\d{{6,}},\d{{2}})",
        ],
    ]

    for group in priority_groups:
        best = ""
        best_val = 0.0
        for pat in group:
            for m in re.finditer(pat, text, re.IGNORECASE | re.MULTILINE):
                candidate = parse_argentinian_money(m.group(1))
                if candidate:
                    try:
                        v = float(candidate)
                        if v > best_val:
                            best_val = v
                            best = candidate
                    except ValueError:
                        pass
        if best:
            return best
    return ""


def _find_numero_factura(text: str) -> str:
    """
    Busca número de comprobante AFIP: XXXX-XXXXXXXX.
    Requiere al menos 4 dígitos en parte izquierda y 6-8 en parte derecha.
    Prioriza líneas con etiqueta explícita.
    """
    # Con etiqueta
    label_patterns = [
        r"(?:factura\s*n[°º]?|comprobante\s*n[°º]?|n[°º]\.?\s*(?:de\s+)?(?:factura|comprobante))\s*[:\-]?\s*(\d{3,4}-\d{6,8})",
        r"(?:factura|comprobante)\s*[:\-]?\s*(\d{4}-\d{6,8})",
    ]
    for pat in label_patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1).strip()

    # Sin etiqueta: patrón AFIP puro (4 dígitos exactos - 6,7 u 8 dígitos)
    # Usar \b para asegurar que no sea parte de un número más largo
    for pat in [
        r"(?<![\d])([0-9]{4}-[0-9]{8})(?![\d])",
        r"(?<![\d])([0-9]{4}-[0-9]{7})(?![\d])",
        r"(?<![\d])([0-9]{4}-[0-9]{6})(?![\d])",
    ]:
        m = re.search(pat, text)
        if m:
            return m.group(1).strip()
    return ""


def _find_tipo_factura(text: str) -> str:
    patterns = [
        r"factura\s+(?:tipo\s+)?([ABCEMX])\b",
        r"tipo\s+(?:de\s+)?(?:comprobante|factura)\s*[:\-]?\s*([ABCEMX])\b",
        r"\bFACTURA\s+([ABCEM])\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1).upper()
    return ""


def _normalize_cuit(raw: str) -> str:
    d = re.sub(r"\D", "", raw)
    if len(d) == 11:
        return f"{d[:2]}-{d[2:10]}-{d[10]}"
    return ""


def _find_cuit(text: str) -> str:
    """
    Busca CUIT con formato XX-XXXXXXXX-X.
    Primero busca con etiqueta de proveedor/emisor, luego cualquier CUIT.
    """
    CUIT_RE = r"(\d{2}[-\s]\d{7,8}[-\s]\d)\b"

    # Con etiqueta
    for label in [r"cuit\s*(?:del?\s*)?(?:proveedor|emisor)", r"c\.?u\.?i\.?t\.?"]:
        m = re.search(rf"{label}\s*[:\-]?\s*{CUIT_RE}", text, re.IGNORECASE)
        if m:
            c = _normalize_cuit(m.group(1))
            if c:
                return c

    # Sin etiqueta — primer CUIT válido en el documento
    for m in re.finditer(CUIT_RE, text):
        c = _normalize_cuit(m.group(1))
        if c:
            return c
    return ""


def _find_empresa(text: str) -> str:
    """
    Busca el nombre del emisor/proveedor.
    Estrategia:
     1. Etiqueta explícita (proveedor, emisor, razón social)
     2. Líneas en MAYÚSCULAS que parezcan nombre de empresa (mín. 5 chars)
     3. Primera línea no vacía que no sea numérica ni keyword
    """
    # Etiqueta explícita
    label_patterns = [
        r"(?:proveedor|emisor|raz[oó]n\s+social)\s*[:\-]?\s*([A-ZÁÉÍÓÚÑ][^\n]{4,70})",
        r"(?:empresa|compañ[íi]a)\s*[:\-]?\s*([A-ZÁÉÍÓÚÑ][^\n]{4,70})",
    ]
    for pat in label_patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1).strip().rstrip(".,;:")

    # Líneas en MAYÚSCULAS que parezcan nombre de empresa
    # (solo letras, espacios, puntos — sin dígitos — mín 5 chars)
    _SKIP_WORDS = re.compile(
        r"^(cuit|fecha|total|factura|n[°º]|importe|vencimiento|periodo|"
        r"pagina|p[aá]gina|tel[eé]fono|tel\.|mail|e-mail|web|www|http)",
        re.IGNORECASE,
    )
    for line in text.split("\n"):
        line = line.strip()
        # Línea mayúscula, min 5 chars, sin dígitos, no es keyword
        if (
            len(line) >= 5
            and line == line.upper()              # toda en mayúsculas
            and re.match(r"^[A-ZÁÉÍÓÚÑ\s\.\,\&\-]+$", line)  # solo letras y puntuación
            and not _SKIP_WORDS.match(line)
        ):
            return line

    # Fallback: primera línea significativa
    for line in text.split("\n"):
        line = line.strip()
        if (
            len(line) >= 5
            and not re.match(r"^[\d\s\.\-\/\$\%\+]+$", line)
            and not re.match(r"^\d{4}-\d", line)
            and not _SKIP_WORDS.match(line)
        ):
            return line
    return ""


def _find_numero_cliente(text: str) -> str:
    """
    Busca número de cliente / asociado / afiliado.
    Mínimo 5 dígitos o formato NNNNN-N para evitar palabras sueltas.
    """
    patterns = [
        # Swiss Medical: "Asociado Nro: 3024521-1" o "N° Asociado: 3024521-1"
        r"(?:asociado\s*(?:n[°º]?|nro\.?|num\.?|:)|n[°º]\.?\s*asociado)\s*[:\-]?\s*(\d{5,10}-\d{1,4}|\d{6,10})",
        # Afiliado
        r"(?:afiliado\s*(?:n[°º]?|nro\.?|num\.?)|n[°º]\.?\s*afiliado)\s*[:\-]?\s*(\d{5,10}-\d{1,4}|\d{6,10})",
        # Cliente (mínimo 5 dígitos)
        r"(?:n[°º]?\.?\s*(?:de\s+)?cliente)\s*[:\-]?\s*(\d{5,10}[\-\.]?\d{0,4})",
        # Póliza / beneficiario (mínimo 5 dígitos)
        r"(?:p[oó]liza|beneficiario)\s*[:\-]?\s*(\d{5,10}[\-\.]?\d{0,4})",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = m.group(1).strip().rstrip("-")
            # Mínimo 5 caracteres numéricos totales
            if len(re.sub(r"\D", "", val)) >= 5:
                return val
    return ""


def _find_numero_cuenta(text: str) -> str:
    """Busca número de cuenta. Mínimo 5 dígitos para evitar palabras sueltas."""
    patterns = [
        r"(?:n[°º]?\.?\s*(?:de\s+)?cuenta)\s*[:\-]?\s*(\d{5,10}[\-\.]?\d{0,4})",
        r"(?:cuenta\s*n[°º]?)\s*[:\-]?\s*(\d{5,10}[\-\.]?\d{0,4})",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = m.group(1).strip().rstrip("-")
            if len(re.sub(r"\D", "", val)) >= 5:
                return val
    return ""


def _find_fecha(text: str, labels: list[str]) -> str:
    """Busca fecha con las etiquetas dadas. Retorna ISO YYYY-MM-DD."""
    DATE_NUM = r"(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{4})"
    DATE_TXT = r"(\d{1,2}\s+de\s+[A-Za-záéíóúñ]+\s+(?:de\s+)?\d{4})"

    for label in labels:
        for date_re in [DATE_TXT, DATE_NUM]:
            m = re.search(rf"(?:{label})\s*[:\-]?\s*{date_re}", text, re.IGNORECASE)
            if m:
                return parse_argentinian_date(m.group(1))

    # Fallback: cualquier fecha numérica en el texto
    m = re.search(DATE_NUM, text)
    if m:
        return parse_argentinian_date(m.group(1))
    return ""


def _find_cliente_nombre(text: str) -> str:
    """Busca nombre completo del cliente (persona física o jurídica)."""
    patterns = [
        r"(?:cliente|titular|asegurado|sr\.?|sra\.?)\s*[:\-]?\s*([A-ZÁÉÍÓÚÑ][A-Za-záéíóúñ\s]{6,50})",
        r"(?:a\s+nombre\s+de|nombre\s+y\s+apellido)\s*[:\-]?\s*([A-ZÁÉÍÓÚÑ][A-Za-záéíóúñ\s]{6,50})",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1).strip().rstrip(".,;:")
    return ""


# ═══════════════════════════════════════════════════════════════════════════════
#  Función principal
# ═══════════════════════════════════════════════════════════════════════════════

def extract_invoice_data(file_path: str, debug: bool = False) -> dict:
    """
    Extrae campos de una factura argentina (PDF o imagen).
    Campos no detectados quedan como cadena vacía.
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
        "fecha_envio":    "",
        "estado_pago":    "Pendiente",
        "archivo":        path.name,
        "ruta":           str(path.resolve()),
        "fecha_carga":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "_ocr_used":      False,
        "_is_image":      is_image,
        "_fields_found":  0,
        "_raw_text":      "",
    }

    text, ocr_used = _get_text(file_path)
    result["_ocr_used"] = ocr_used
    result["_raw_text"] = text

    if debug:
        sep = "═" * 70
        print(f"\n{sep}\n[DEBUG] {path.name}  |  OCR={ocr_used}")
        print("─" * 70)
        print(text[:4000])
        print("─" * 70)

    if not text.strip():
        return result

    # ── Extraer campos ───────────────────────────────────────────────────────
    result["tipo_factura"]   = _find_tipo_factura(text)
    result["numero_factura"] = _find_numero_factura(text)
    result["cuit"]           = _find_cuit(text)
    result["empresa"]        = _find_empresa(text)
    result["numero_cliente"] = _find_numero_cliente(text)
    result["numero_cuenta"]  = _find_numero_cuenta(text)
    result["importe"]        = _find_importe(text)

    result["fecha_emision"]  = _find_fecha(text, [
        r"fecha\s+de\s+emisi[oó]n",
        r"fecha\s+emisi[oó]n",
        r"fecha\s+factura",
        r"fecha\s+de\s+factura",
        r"fecha",
    ])
    result["fecha_envio"] = _find_fecha(text, [
        r"(?:fecha\s+de\s+)?vencimiento",
        r"vto\.?",
        r"fecha\s+(?:de\s+)?pago",
        r"fecha\s+l[íi]mite",
        r"abonar\s+antes\s+del",
    ])

    # Calcular campos encontrados
    tracked = ["tipo_factura", "numero_factura", "cuit", "empresa",
               "numero_cliente", "importe", "fecha_emision", "fecha_envio"]
    result["_fields_found"] = sum(1 for k in tracked if result.get(k))

    if debug:
        print("[DEBUG] Campos detectados:")
        for k in tracked + ["numero_cuenta"]:
            val = result.get(k, "")
            status = "✅" if val else "❌"
            print(f"  {status} {k:<22} {val!r}")
        print(f"\n  Total campos: {result['_fields_found']}/8")
        print("═" * 70 + "\n")

    return result


# ── Alias de compatibilidad ───────────────────────────────────────────────────

def parse_invoice(file_path: str, debug: bool = False) -> dict:
    """Alias para compatibilidad con código existente."""
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


# ── CLI de debug ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python modules/pdf_reader.py <factura.pdf>")
        sys.exit(1)
    data = extract_invoice_data(sys.argv[1], debug=True)
    print("\nResultado final:")
    for k, v in data.items():
        if not k.startswith("_"):
            print(f"  {k:<22} {v}")
