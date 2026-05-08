"""
pdf_reader.py — Extracción de facturas argentinas.

Flujo:
  1. Extraer texto del PDF/imagen con pdfplumber / PyPDF2 / pytesseract.
  2. Si hay OPENAI_API_KEY disponible → extract_with_ai(text).
  3. Si no → fallback regex conservador (deja vacío si no está seguro).

Funciones públicas:
    extract_invoice_data(file_path, api_key=None, debug=False) -> dict
    parse_invoice(file_path, api_key=None, debug=False) -> dict   # alias
    extract_with_ai(text, api_key) -> dict
    tesseract_available() -> bool
    supported_extensions() -> list[str]
"""

from __future__ import annotations

import json
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

try:
    import openai as _openai_module
    _HAS_OPENAI = True
except ImportError:
    _HAS_OPENAI = False

_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"}

_MESES = {
    "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
    "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
    "septiembre": "09", "octubre": "10", "noviembre": "11", "diciembre": "12",
}

# Campos esperados en la respuesta (con sus defaults)
_EMPTY_RESULT = {
    "empresa":           "",
    "numero_factura":    "",
    "cuit":              "",
    "tipo_factura":      "",
    "cliente":           "",
    "numero_cliente":    "",
    "numero_cuenta":     "",
    "fecha_emision":     "",
    "fecha_vencimiento": "",
    "importe":           "",
    "total_factura":     "",
    "total_a_pagar":     "",
    "estado_pago":       "Pendiente",
}

# ═══════════════════════════════════════════════════════════════════════════════
#  Extracción de texto
# ═══════════════════════════════════════════════════════════════════════════════

def normalize_ocr_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\t", " ")
    lines = [re.sub(r" {2,}", " ", l).strip() for l in text.split("\n")]
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines))


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
    raw = _extract_text_pdf(file_path)
    return normalize_ocr_text(raw), False


# ═══════════════════════════════════════════════════════════════════════════════
#  Extracción con IA (OpenAI)
# ═══════════════════════════════════════════════════════════════════════════════

_AI_PROMPT = """\
Sos un asistente experto en extraer datos de facturas de cualquier país y empresa.
Analizá el texto y devolvé SOLO un JSON válido con estas claves:

- empresa: razón social o nombre comercial del EMISOR de la factura (quien factura, no quien paga).
  Buscar cerca del logo, encabezado, razón social o junto al CUIT/VAT del emisor.
  Ejemplos: "Swiss Medical", "Telecentro", "Edesur", "Google LLC", "Amazon".
  NO usar textos descriptivos, párrafos, ni frases largas. Si no es claro, dejar vacío.

- numero_factura: número de comprobante, invoice number, folio (ej: 1346-19688880, INV-2024-001)

- cuit: CUIT o CUIL del EMISOR (quien factura). Formato XX-XXXXXXXX-X.
  Solo del emisor, NO del cliente. Ignorar DNI o CUIT del receptor.

- tipo_factura: letra del comprobante según AFIP: A, B, C, E o M.
  Buscar "FACTURA A", "Factura tipo B", "Comprobante C", etc. Si no aparece, dejar vacío.

- cliente: nombre completo del cliente, titular, bill-to, sold-to, receptor.

- numero_cliente: número de cliente, asociado, afiliado, subscriber, customer ID, account holder,
  member number, número de abonado. Puede tener letras y números.

- numero_cuenta: número de cuenta, account number, número de servicio, customer account,
  service number, número de suministro. Diferente al número de factura.

- fecha_emision: fecha de emisión en formato YYYY-MM-DD
- fecha_vencimiento: fecha de vencimiento / due date en formato YYYY-MM-DD
- importe: importe base (sin intereses), como número con punto decimal
- total_factura: total de la factura
- total_a_pagar: total a pagar / amount due / balance due
- estado_pago: siempre "Pendiente"

Reglas estrictas:
- Si un campo no aparece claramente, usá string vacío "".
- No inventes datos. Si no estás seguro, dejá vacío.
- Los importes: solo números con punto decimal (ej: 1234.56), sin símbolos.
- Las fechas: formato YYYY-MM-DD.
- Respondé SOLO el JSON, sin texto adicional, sin markdown.
"""


def extract_with_ai(text: str, api_key: str) -> dict:
    """
    Llama a OpenAI GPT-4o-mini para extraer campos de la factura.
    Lanza excepción detallada si falla para que el llamador pueda loguearla.
    Devuelve dict vacío solo si api_key es inválida o no hay openai instalado.
    """
    if not _HAS_OPENAI:
        raise RuntimeError("La librería 'openai' no está instalada. Corré : pip install openai")

    api_key = (api_key or "").strip()
    if not api_key:
        raise ValueError("OPENAI_API_KEY está vacía. Configúrala en ⚙️ Configuración.")

    # Truncar texto a 6000 chars para no exceder tokens
    text_to_send = text.strip()[:6000]
    if not text_to_send:
        raise ValueError("El texto extraído del PDF está vacío. El PDF puede ser una imagen sin OCR.")

    client = _openai_module.OpenAI(api_key=api_key)

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": _AI_PROMPT},
            {"role": "user", "content": f"Texto de la factura:\n\n{text_to_send}"},
        ],
        temperature=0,
        max_tokens=1000,
        response_format={"type": "json_object"},  # fuerza JSON válido (gpt-4o-mini lo soporta)
    )

    raw = (response.choices[0].message.content or "").strip()

    if not raw:
        raise ValueError("OpenAI devolvió una respuesta vacía.")

    data = json.loads(raw)   # si el JSON es inválido, lanza JSONDecodeError

    # Normalizar: asegurar que todas las claves esperadas existen como strings
    result = dict(_EMPTY_RESULT)
    for k in _EMPTY_RESULT:
        val = data.get(k)
        if val is not None and str(val).strip():
            result[k] = str(val).strip()
    return result


# ═══════════════════════════════════════════════════════════════════════════════
#  Fallback regex conservador
# ═══════════════════════════════════════════════════════════════════════════════

def _safe_money(raw: str) -> str:
    """Convierte importe argentino. Rechaza si no tiene al menos 4 dígitos o >10 dígitos."""
    if not raw:
        return ""
    digits = re.sub(r"\D", "", raw)
    if len(digits) < 4 or len(digits) > 10:
        return ""
    v = re.sub(r"[^\d,\.]", "", raw)
    if re.match(r"^\d{1,3}(\.\d{3})*,\d{2}$", v):
        v = v.replace(".", "").replace(",", ".")
    elif re.match(r"^\d+,\d{2}$", v):
        v = v.replace(",", ".")
    elif re.match(r"^\d{1,3}(\.\d{3})+$", v):
        v = v.replace(".", "")
    else:
        v = v.replace(",", ".")
    try:
        num = float(v)
        if num < 100 or num > 9_999_999:
            return ""
        return f"{num:.2f}"
    except ValueError:
        return ""


def _safe_date(raw: str) -> str:
    if not raw:
        return ""
    raw = raw.strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}$", raw):
        return raw
    m = re.match(r"^(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{4})$", raw)
    if m:
        d, mo, y = m.groups()
        return f"{y}-{mo.zfill(2)}-{d.zfill(2)}"
    m = re.match(r"^(\d{1,2})\s+de\s+([A-Za-záéíóúñ]+)\s+(?:de\s+)?(\d{4})$", raw, re.IGNORECASE)
    if m:
        d, mes, y = m.groups()
        mo = _MESES.get(mes.lower(), "")
        if mo:
            return f"{y}-{mo}-{d.zfill(2)}"
    return raw


_NUM_ARG = r"(\d{1,3}(?:\.\d{3})*,\d{2}|\d{4,},\d{2}|\d{4,}\.\d{2})"
_DATE    = r"(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{4}|\d{1,2}\s+de\s+[A-Za-záéíóúñ]+\s+(?:de\s+)?\d{4})"


def _regex_fallback(text: str) -> dict:
    """
    Extracción regex conservadora.
    Solo completa un campo si hay alta confianza. Deja vacío si hay duda.
    """
    result = dict(_EMPTY_RESULT)

    def _debug_context(match: re.Match, label: str):
        if debug:
            start = max(0, match.start() - 60)
            print(f"[DEBUG] {label} context: ...{text[start:match.end()+30]}...")

    # ── numero_factura: patrón AFIP + formatos internacionales ────────────────
    # AFIP: XXXX-XXXXXXXX
    m = re.search(r"(?<!\d)(\d{4}-\d{6,8})(?!\d)", text)
    if m:
        result["numero_factura"] = m.group(1)
    else:
        # Internacional: Invoice No, Bill Number, Comprobante N°, Folio, etc.
        m = re.search(
            r"(?:invoice\s*(?:no\.?|number|#|n[°º])?|"
            r"bill\s*(?:no\.?|number|#)?|comprobante\s*(?:n[°º]?|nro\.?)?|"
            r"n[°º]\s*(?:de\s+)?factura|factura\s*n[°º]?|folio|receipt\s*(?:no\.?|#)?)"
            r"\s*[:\-#]?\s*([A-Z0-9]{1,6}[-/]?[A-Z0-9]{2,12})\b",
            text, re.IGNORECASE,
        )
        if m:
            val = m.group(1)
            # Descartar si es pura fecha o muy corto
            if len(re.sub(r"\D", "", val)) >= 3 and not re.match(r"^\d{1,2}[/\-]\d{1,2}$", val):
                result["numero_factura"] = val

    # ── tipo_factura ───────────────────────────────────────────────────────
    # Buscar letra A/B/C/E/M junto a keyword de comprobante
    m = re.search(
        r"(?:factura|comprobante|invoice|tipo|código|codigo)\s+(?:tipo\s+)?([ABCEMX])\b",
        text, re.IGNORECASE,
    )
    if not m:
        # Segundo intento: "Tipo: A" o "Código comprobante: B"
        m = re.search(
            r"(?:tipo|código|codigo)\s*[:\-]?\s*(?:comprobante\s*[:\-]?\s*)?([ABCEMX])\b",
            text, re.IGNORECASE,
        )
    if m:
        result["tipo_factura"] = m.group(1).upper()

    # ── CUIT: XX-XXXXXXXX-X — solo si hay keyword CUIT/CUIL cerca ─────────────
    m = re.search(
        r"(?:CUIT|CUIL)\s*[:\-]?\s*(\d{2}[-\s]\d{7,8}[-\s]\d)\b",
        text, re.IGNORECASE,
    )
    if m:
        d = re.sub(r"\D", "", m.group(1))
        if len(d) == 11:
            result["cuit"] = f"{d[:2]}-{d[2:10]}-{d[10]}"

    # ── empresa: validación estricta de nombre comercial real ─────────────────
    _GARBAGE_SUB = re.compile(
        r"FREPOLREF|FEPOLREF|FEPO|POLREF|CODIGO|ASEGURADO|RIESGO|OBJETO|"
        r"SEGURO|SERVICIO|DETALLE|DESCRIPCION|CONCEPTO|CONDICION|BARCODE|"
        r"COMPROBANTE|DUPLICADO|ORIGINAL|PERIODO|CUOTAS|AFILIADO",
        re.IGNORECASE,
    )
    _GARBAGE_EXACT = re.compile(
        r"^(CUIT|CUIL|FECHA|TOTAL|FACTURA|IMPORTE|VENCIMIENTO|PAGINA|PAG|"
        r"TEL|FAX|IVA|NRO|NUM|SON|PESOS|DOLARES|DEBE|HABER|SUBTOTAL|"
        r"SALDO|RECIBO|CLIENTE|ASOCIADO|REF|COD|CAE|MIL|POR|WEB|"
        r"HTTP|WWW|EMAIL|HABER|Y|DE|LA|EL|LOS|LAS|DEL|AL)$",
        re.IGNORECASE,
    )
    _skip_prefix = re.compile(
        r"^(cuit|cuil|fecha|total|factura|n[°º]|importe|vencimiento|periodo|pagina|tel)",
        re.IGNORECASE,
    )

    def _looks_like_ocr_noise(word: str) -> bool:
        if _GARBAGE_SUB.search(word):
            return True
        if _GARBAGE_EXACT.match(word):
            return True
        if re.search(r"[BCDFGHJKLMNPQRSTVWXYZ]{4,}", word, re.IGNORECASE):
            return True
        return False

    def _is_valid_empresa(line: str) -> bool:
        words = line.split()
        if not (1 <= len(words) <= 4):
            return False
        if len(line) < 3:
            return False
        if line != line.upper():
            return False
        if not re.match(r"^[A-ZÁÉÍÓÚÑ\s\.\,\&\-]+$", line):
            return False
        if _skip_prefix.match(line):
            return False
        if any(len(w) < 2 for w in words):
            return False
        if any(_looks_like_ocr_noise(w) for w in words):
            return False
        return True

    empresa_found = ""
    lines_list = text.split("\n")
    for i, line in enumerate(lines_list):
        if re.search(r"(?:CUIT|CUIL)\s*[:\-]?\s*\d{2}", line, re.IGNORECASE):
            for j in range(max(0, i - 3), i):
                candidate = lines_list[j].strip()
                if _is_valid_empresa(candidate):
                    empresa_found = candidate
                    break
            break
    if not empresa_found:
        for line in lines_list:
            candidate = line.strip()
            if _is_valid_empresa(candidate):
                empresa_found = candidate
                break
    result["empresa"] = empresa_found

    # ── numero_cliente: sinónimos español + inglés ────────────────────────────
    m = re.search(
        r"(?:n[°º]?\s*de\s+)?(?:cliente|asociado|afiliado|abonado|socio|suscriptor|"
        r"customer(?:\s*id|\s*no\.?|\s*number|\s*#)?|subscriber(?:\s*id|\s*no\.?)?|"
        r"account\s*holder|member(?:\s*id|\s*no\.?|\s*number)?|policy\s*(?:holder|no\.?)|titular)"
        r"(?:\s*(?:n[°º]|nro\.?|num\.?|id|#|no\.?))?\s*[:\-]?\s*([A-Z0-9]{3,20}(?:[-\.][0-9]{1,8})?)",
        text, re.IGNORECASE,
    )
    if m:
        val = m.group(1).rstrip("-")
        if len(re.sub(r"\D", "", val)) >= 3:
            result["numero_cliente"] = val
        if debug:
            start = max(0, m.start() - 60)
            print(f"[DEBUG] numero_cliente context: ...{text[start:m.end()+30]}...")

    # ── numero_cuenta: sinónimos español + inglés ─────────────────────────────
    m = re.search(
        r"(?:n[°º]?\s*(?:de\s+)?)?(?:cuenta|account(?:\s*(?:number|no\.?|id|#))?|"
        r"n[°º]\s*cuenta|customer\s*account|service\s*(?:account|number|no\.?)|"
        r"n[°º]\s*(?:de\s+)?(?:servicio|suministro)|suministro)"
        r"(?:\s*(?:n[°º]|nro\.?|num\.?|#|no\.?))?\s*[:\-#]?\s*([A-Z0-9]{3,20}(?:[-\.][0-9]{1,8})?)",
        text, re.IGNORECASE,
    )
    if m:
        val = m.group(1).rstrip("-")
        if val != result.get("numero_factura", "") and len(re.sub(r"\D", "", val)) >= 3:
            result["numero_cuenta"] = val
        if debug:
            start = max(0, m.start() - 60)
            print(f"[DEBUG] numero_cuenta context: ...{text[start:m.end()+30]}...")


    # ── fecha_emision — solo con keyword explícita, nunca "fecha" suelto ──────
    m = re.search(
        rf"(?:fecha\s+(?:de\s+)?(?:emisi[oó]n|factura|comprobante))\s*[:\-]?\s*{_DATE}",
        text, re.IGNORECASE,
    )
    if m:
        result["fecha_emision"] = _safe_date(m.group(1))

    # ── fecha_vencimiento ────────────────────────────────────────────────────
    m = re.search(
        rf"(?:vencimiento|vto\.?|fecha\s+de\s+pago)\s*[:\-]?\s*{_DATE}",
        text, re.IGNORECASE,
    )
    if m:
        result["fecha_vencimiento"] = _safe_date(m.group(1))
        result["fecha_envio"] = result["fecha_vencimiento"]  # alias

    # ── total_a_pagar (mayor prioridad) ─────────────────────────────────────
    best_pay = ""
    best_val = 0.0
    for pat in [
        rf"total\s+a\s+(?:pagar|abonar)\s*[:\$]?\s*\$?\s*{_NUM_ARG}",
        rf"saldo\s+(?:a\s+pagar|deudor)\s*[:\$]?\s*\$?\s*{_NUM_ARG}",
    ]:
        for m in re.finditer(pat, text, re.IGNORECASE):
            v = _safe_money(m.group(1))
            if v and float(v) > best_val:
                best_val = float(v)
                best_pay = v
    result["total_a_pagar"] = best_pay

    # ── total_factura ────────────────────────────────────────────────────────
    best_fac = ""
    best_fac_val = 0.0
    for pat in [
        rf"total\s+factura\s*[:\$]?\s*\$?\s*{_NUM_ARG}",
        rf"importe\s+(?:total|factura)\s*[:\$]?\s*\$?\s*{_NUM_ARG}",
    ]:
        for m in re.finditer(pat, text, re.IGNORECASE):
            v = _safe_money(m.group(1))
            if v and float(v) > best_fac_val:
                best_fac_val = float(v)
                best_fac = v
    result["total_factura"] = best_fac

    # importe = total_a_pagar si existe, si no total_factura
    result["importe"] = best_pay or best_fac

    return result


# ═══════════════════════════════════════════════════════════════════════════════
#  Función principal
# ═══════════════════════════════════════════════════════════════════════════════

def extract_invoice_data(file_path: str, api_key: str = "", debug: bool = False) -> dict:
    """
    Extrae campos de una factura argentina.
    1. Extrae texto (PDF o imagen OCR).
    2. Si api_key disponible → OpenAI GPT (error visible en consola/UI).
    3. Si no → regex conservador.
    """
    path = Path(file_path)
    is_image = path.suffix.lower() in _IMAGE_EXTENSIONS

    meta = {
        "archivo":       path.name,
        "ruta":          str(path.resolve()),
        "fecha_carga":   datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "_ocr_used":     False,
        "_is_image":     is_image,
        "_ai_used":      False,
        "_fields_found": 0,
        "_raw_text":     "",
    }

    text, ocr_used = _get_text(file_path)
    meta["_ocr_used"] = ocr_used
    meta["_raw_text"] = text

    if debug:
        print(f"\n{'═'*70}\n[DEBUG] {path.name} | OCR={ocr_used}")
        print("─" * 70)
        print(text[:4000])
        print("─" * 70)

    if not text.strip():
        result = dict(_EMPTY_RESULT)
        result.update(meta)
        return result

    # ── Resolución de la clave API ───────────────────────────────────────────
    effective_key = (api_key or "").strip() or os.environ.get("OPENAI_API_KEY", "")

    # ── Extracción ───────────────────────────────────────────────────────────
    fields: dict = {}
    if effective_key and _HAS_OPENAI:
        try:
            fields = extract_with_ai(text, effective_key)
            if fields:
                meta["_ai_used"] = True
                if debug:
                    print("[DEBUG] Modo: OpenAI GPT ✅")
        except Exception as exc:
            meta["_ai_error"] = str(exc)
            if debug:
                print(f"[DEBUG] OpenAI error: {exc}")

    if not fields:
        fields = _regex_fallback(text)
        if debug:
            print("[DEBUG] Modo: regex fallback")

    # Campos que siempre vienen de metadatos del archivo, no del texto
    fields.setdefault("estado_pago", "Pendiente")
    fields["fecha_envio"] = fields.get("fecha_vencimiento", "")

    # Contar campos encontrados
    tracked = ["empresa", "numero_factura", "cuit", "cliente",
               "numero_cliente", "importe", "fecha_emision", "fecha_vencimiento"]
    meta["_fields_found"] = sum(1 for k in tracked if fields.get(k))

    if debug:
        print("[DEBUG] Campos detectados:")
        for k in tracked + ["total_a_pagar", "total_factura"]:
            v = fields.get(k, "")
            print(f"  {'✅' if v else '❌'} {k:<22} {v!r}")
        print(f"\n  _ai_used: {meta['_ai_used']}  |  campos: {meta['_fields_found']}/8")
        print("═" * 70 + "\n")

    result = dict(_EMPTY_RESULT)
    result.update(fields)
    result.update(meta)
    return result


# ── Alias de compatibilidad ───────────────────────────────────────────────────

def parse_invoice(file_path: str, api_key: str = "", debug: bool = False) -> dict:
    """Alias de extract_invoice_data para compatibilidad con código existente."""
    return extract_invoice_data(file_path, api_key=api_key, debug=debug)


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
        print("Uso: python modules/pdf_reader.py <factura.pdf> [openai_api_key]")
        sys.exit(1)
    key = sys.argv[2] if len(sys.argv) > 2 else ""
    data = extract_invoice_data(sys.argv[1], api_key=key, debug=True)
    print("\nResultado final:")
    for k, v in data.items():
        if not k.startswith("_"):
            print(f"  {k:<22} {v}")
