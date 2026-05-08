"""
Patch: agrega _ocr_short flag en pdf_reader.py y aviso en app_web.py.
Ejecutar con: python _patch_ocr.py
"""
import pathlib, re, sys

# ─── pdf_reader.py ──────────────────────────────────────────────────────────
pr = pathlib.Path("modules/pdf_reader.py")
src = pr.read_text(encoding="utf-8")

OLD_PR = (
    '    text, ocr_used = _get_text(file_path)\n'
    '    meta["_ocr_used"] = ocr_used\n'
    '    meta["_raw_text"] = text\n'
    '\n'
    '    if debug:\n'
)
NEW_PR = (
    '    text, ocr_used = _get_text(file_path)\n'
    '    meta["_ocr_used"] = ocr_used\n'
    '    meta["_raw_text"] = text\n'
    '\n'
    '    # Detectar documento ilegible (imagen borrosa, scan de baja calidad, etc.)\n'
    '    _OCR_MIN_CHARS = 80\n'
    '    meta["_ocr_short"] = len(text.strip()) < _OCR_MIN_CHARS\n'
    '\n'
    '    if debug:\n'
)
if OLD_PR not in src:
    print("SKIP pdf_reader.py (ya patcheado)")
else:
    src2 = src.replace(OLD_PR, NEW_PR, 1)
    pr.write_text(src2, encoding="utf-8")
    print("OK pdf_reader.py")

# ─── app_web.py ─────────────────────────────────────────────────────────────
aw = pathlib.Path("app_web.py")
src = aw.read_text(encoding="utf-8")

OLD_AW = "        parsed = _parse_once(name, data_bytes)\n\n        # \u2500\u2500 Invoice form \u2500\u2500"
NEW_AW = (
    "        parsed = _parse_once(name, data_bytes)\n"
    "\n"
    "        # \u2500\u2500 Aviso OCR insuficiente \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\n"
    "        if parsed.get(\"_ocr_short\"):\n"
    "            st.error(\n"
    "                \"\U0001f4c4 **No se pudo leer bien el documento.**\\n\\n\"\n"
    "                \"Sub\u00ed el PDF original o una foto m\u00e1s n\u00edtida, derecha y recortada.\",\n"
    "                icon=\"\u26a0\ufe0f\",\n"
    "            )\n"
    "\n"
    "        # \u2500\u2500 Invoice form \u2500\u2500"
)
if OLD_AW not in src:
    print("ERROR: bloque app_web no encontrado"); sys.exit(1)
src2 = src.replace(OLD_AW, NEW_AW, 1)
aw.write_text(src2, encoding="utf-8")
print("OK app_web.py")
print("Ambos archivos actualizados correctamente.")
