"""
Streamlit web interface — Sistema de Gestión de Facturas
Run with: streamlit run app_web.py
"""

import os
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Dict, List

import pandas as pd
import streamlit as st

from config import load_config, save_config
from modules.email_sender import EmailSender
from modules.excel_manager import COLUMNS, ExcelManager
from modules.file_manager import FileManager
from modules.pdf_reader import parse_invoice

# ═══════════════════════════════════════════════════════════════════════════════
#  Page config  (must be the very first Streamlit call)
# ═══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Gestión de Facturas",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ═══════════════════════════════════════════════════════════════════════════════
#  CSS
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown(
    """
<style>
/* ── App header ──────────────────────────────────────────────────── */
.app-header {
    background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
    padding: 22px 32px; border-radius: 12px; margin-bottom: 24px; color: #fff;
}
.app-header h1 { margin: 0; font-size: 1.75rem; font-weight: 700; }
.app-header p  { margin: 5px 0 0; opacity: .82; font-size: .95rem; }

/* ── Section card ────────────────────────────────────────────────── */
.card {
    background: #fff; border-radius: 10px; padding: 20px 24px;
    margin-bottom: 14px; border: 1px solid #e2e8f0;
    box-shadow: 0 2px 8px rgba(0,0,0,.05);
}

/* ── Progress pill ───────────────────────────────────────────────── */
.step-pill {
    display: inline-block; background: #2E75B6; color: #fff;
    padding: 5px 18px; border-radius: 20px; font-size: .88rem;
    font-weight: 600; margin-bottom: 10px;
}

/* ── Status badges ───────────────────────────────────────────────── */
.badge-paga {
    background: #d4edda; color: #155724;
    padding: 3px 12px; border-radius: 12px; font-size: .82rem; font-weight: 600;
}
.badge-pendiente {
    background: #fff3cd; color: #856404;
    padding: 3px 12px; border-radius: 12px; font-size: .82rem; font-weight: 600;
}

/* ── Empty state ─────────────────────────────────────────────────── */
.empty-state {
    text-align: center; padding: 48px 0; color: #94a3b8; font-size: 1.05rem;
}

/* Streamlit tab tweaks */
.stTabs [data-baseweb="tab"] { font-size: .97rem; padding: 10px 22px; }
</style>
""",
    unsafe_allow_html=True,
)

# ═══════════════════════════════════════════════════════════════════════════════
#  Session state
# ═══════════════════════════════════════════════════════════════════════════════

_DEFAULTS: Dict = {
    "cfg":           None,          # loaded lazily
    "file_queue":    [],            # [{"name": str, "data": bytes}, …]
    "proc_idx":      0,
    "proc_active":   False,
    "proc_results":  {"added": 0, "duplicates": 0, "skipped": 0},
    "parsed_cache":  {},            # filename → dict from parse_invoice
    "flash":         None,          # ("success"|"warning"|"error", text)
}

for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

if st.session_state.cfg is None:
    st.session_state.cfg = load_config()


# ─── Helpers ────────────────────────────────────────────────────────────────

def _cfg() -> dict:
    return st.session_state.cfg


def _excel() -> ExcelManager:
    return ExcelManager(_cfg().get("excel_filename", "control_facturas.xlsx"))


def _files() -> FileManager:
    return FileManager(_cfg().get("facturas_base_dir", "Facturas"))


def _flash(kind: str, msg: str) -> None:
    st.session_state.flash = (kind, msg)


def _show_flash() -> None:
    f = st.session_state.flash
    if f:
        getattr(st, f[0])(f[1])
        st.session_state.flash = None


def _save_pdf_bytes(name: str, data: bytes, fecha_emision: str) -> str:
    """
    Write *data* to Facturas/YYYY-MM/<name>.
    Returns the absolute path string (or '' on failure).
    """
    try:
        fm = _files()
        fm.ensure_base_dir()
        dest_dir  = fm.month_folder(fecha_emision)
        dest_path = dest_dir / name

        if dest_path.exists():
            stem, suf = Path(name).stem, Path(name).suffix
            counter = 1
            while dest_path.exists():
                dest_path = dest_dir / f"{stem}_{counter}{suf}"
                counter += 1

        dest_path.write_bytes(data)
        return str(dest_path.resolve())
    except Exception:
        return ""


def _api_key() -> str:
    """Resolve OpenAI API key: st.secrets → config → env var."""
    # 1. Streamlit Cloud secrets (highest priority)
    try:
        secret = (st.secrets.get("OPENAI_API_KEY") or "").strip()
        if secret:
            return secret
    except Exception:
        pass
    # 2. Config file (⚙️ Configuración tab)
    from_cfg = _cfg().get("openai_api_key", "").strip()
    if from_cfg:
        return from_cfg
    # 3. Environment variable
    return os.environ.get("OPENAI_API_KEY", "")


import re as _re

def _clean_empresa(value: str) -> str:
    """
    Último filtro antes de mostrar empresa/proveedor en el formulario.
    Si el valor contiene basura OCR conocida → devuelve vacío.
    Mejor vacío que dato incorrecto.
    """
    v = (value or "").strip()
    if not v:
        return ""

    # Substrings que nunca pertenecen a un nombre comercial
    _BAD_SUB = _re.compile(
        r"FREPOLREF|FEPOLREF|FEPO|POLREF|FREPOL|FREPO",
        _re.IGNORECASE,
    )
    # Palabras exactas que no son nombres de empresa
    _BAD_WORD = _re.compile(
        r"^(REF|MIL|POR|COD|CODIGO|CAE|NRO|NUM|IVA|WEB|FAX|"
        r"RIESGO|OBJETO|SEGURO|ASEGURADO|CLIENTE|AFILIADO|"
        r"COMPROBANTE|DETALLE|DESCRIPCION|CONCEPTO|CONDICION)$",
        _re.IGNORECASE,
    )

    # Rechazar si contiene substring de basura OCR
    if _BAD_SUB.search(v):
        return ""

    words = v.split()

    # Rechazar si alguna palabra exacta es un token inválido
    if any(_BAD_WORD.match(w) for w in words):
        return ""

    # Rechazar si una sola palabra tiene 4+ consonantes seguidas (ruido OCR)
    # Ej: FPLRF, XKTRZ, BRTPQ → basura. "SWISS", "EDESUR" → OK
    if _re.search(r"[BCDFGHJKLMNPQRSTVWXYZ]{4,}", v, _re.IGNORECASE):
        return ""

    return v


def _parse_once(name: str, data: bytes) -> dict:
    """Parse a PDF and cache the result in session state."""
    if name not in st.session_state.parsed_cache:
        # Detect extension from name for image support
        suffix = Path(name).suffix.lower() or ".pdf"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(data)
            tmp_path = tmp.name
        try:
            result = parse_invoice(tmp_path, api_key=_api_key())
        except Exception:
            result = {
                "archivo":     name,
                "estado_pago": "Pendiente",
                "fecha_carga": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
        finally:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
        st.session_state.parsed_cache[name] = result
    return st.session_state.parsed_cache[name]


def _excel_download_button(label: str = "📥  Descargar Excel") -> None:
    path = Path(_cfg().get("excel_filename", "control_facturas.xlsx"))
    if path.exists():
        st.download_button(
            label=label,
            data=path.read_bytes(),
            file_name=path.name,
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            ),
            use_container_width=True,
        )


# ═══════════════════════════════════════════════════════════════════════════════
#  Header
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown(
    """
<div class="app-header">
  <h1>🧾 Sistema de Gestión de Facturas</h1>
  <p>Automatización · importación · control · notificación por email</p>
</div>
""",
    unsafe_allow_html=True,
)

_show_flash()

# ═══════════════════════════════════════════════════════════════════════════════
#  Tabs
# ═══════════════════════════════════════════════════════════════════════════════

tab_import, tab_control, tab_email, tab_config = st.tabs(
    [
        "📁  Importar Facturas",
        "📊  Control de Facturas",
        "📧  Email a Finanzas",
        "⚙️  Configuración",
    ]
)

# ═══════════════════════════════════════════════════════════════════════════════
#  TAB 1 — IMPORTAR
# ═══════════════════════════════════════════════════════════════════════════════

with tab_import:

    # ── A) Upload view ────────────────────────────────────────────────────────
    if not st.session_state.proc_active:
        st.subheader("Cargar archivos PDF")

        st.info(
            "Selecciona uno o varios PDF. La app extrae los datos automáticamente "
            "y te mostrará un formulario por cada factura para revisar y completar.",
            icon="ℹ️",
        )

        uploaded = st.file_uploader(
            "Arrastra los PDFs aquí o haz clic para seleccionar",
            type=["pdf"],
            accept_multiple_files=True,
            label_visibility="collapsed",
        )

        if uploaded:
            st.write(f"**{len(uploaded)} archivo(s) listo(s):**")
            for f in uploaded:
                st.markdown(f"&nbsp;&nbsp;📄 `{f.name}`")

            st.divider()
            col_btn, col_note = st.columns([1, 3])
            with col_btn:
                start = st.button(
                    "▶  Procesar Facturas",
                    type="primary",
                    use_container_width=True,
                )
            with col_note:
                st.caption(
                    "Próximamente: importar desde  "
                    "Email IMAP · Portal Web · Carpeta compartida"
                )

            if start:
                st.session_state.file_queue   = [
                    {"name": f.name, "data": f.read()} for f in uploaded
                ]
                st.session_state.proc_idx     = 0
                st.session_state.proc_active  = True
                st.session_state.proc_results = {"added": 0, "duplicates": 0, "skipped": 0}
                st.session_state.parsed_cache = {}
                st.rerun()

        else:
            st.markdown(
                '<div class="empty-state">⬆️  Arrastra archivos PDF aquí para comenzar</div>',
                unsafe_allow_html=True,
            )

    # ── B) Done view ──────────────────────────────────────────────────────────
    elif st.session_state.proc_idx >= len(st.session_state.file_queue):
        r = st.session_state.proc_results
        st.success(
            f"✅  Procesamiento completado — "
            f"**{r['added']}** guardadas · "
            f"**{r['duplicates']}** duplicadas · "
            f"**{r['skipped']}** saltadas"
        )

        col_a, col_b, _ = st.columns([1, 1, 3])
        with col_a:
            if st.button("🔄  Procesar más facturas", type="primary", use_container_width=True):
                st.session_state.proc_active = False
                st.session_state.file_queue  = []
                st.rerun()
        with col_b:
            _excel_download_button("📊  Descargar Excel")

    # ── C) Processing view — one form per invoice ─────────────────────────────
    else:
        queue = st.session_state.file_queue
        idx   = st.session_state.proc_idx
        total = len(queue)
        r     = st.session_state.proc_results

        file_info = queue[idx]
        name, data_bytes = file_info["name"], file_info["data"]

        # Progress
        st.markdown(
            f'<div class="step-pill">Factura {idx + 1} de {total}</div>',
            unsafe_allow_html=True,
        )
        st.progress(idx / total)

        done_so_far = r["added"] + r["duplicates"] + r["skipped"]
        if done_so_far:
            st.caption(
                f"Hasta ahora → ✅ {r['added']} guardadas  "
                f"⚠️ {r['duplicates']} duplicadas  "
                f"⏭ {r['skipped']} saltadas"
            )

        st.markdown(f"**Archivo:** `{name}`")
        st.divider()

        parsed = _parse_once(name, data_bytes)

        # ── Invoice form ───────────────────────────────────────────────────
        # Indicador modo IA vs regex
        ai_key = _api_key()
        ai_error = parsed.get("_ai_error", "")
        if ai_key:
            if parsed.get("_ai_used"):
                st.success("🤖 Datos extraídos con **IA (OpenAI)**", icon="✨")
            elif ai_error:
                st.error(f"❌ **Error de OpenAI:** {ai_error}\n\n_Usando regex como fallback._", icon="🚨")
            else:
                st.info("⚙️ Extracción por **regex**.", icon="ℹ️")
        else:
            st.warning(
                "🔑 No hay OPENAI_API_KEY configurada — usando regex. "
                "Configurá la key en **⚙️ Configuración** para mejor precisión.",
                icon="⚠️",
            )

        with st.form("invoice_form", clear_on_submit=False):
            st.markdown("#### Datos de la factura")
            st.warning(
                "⚠️ **Revisá los campos antes de guardar.** "
                "Los campos que no se pudieron detectar con certeza quedan vacíos para completar manualmente.",
                icon="📋",
            )
            st.caption("Campos obligatorios: Número de Factura, Fecha y Total a Pagar.")

            col1, col2 = st.columns(2)
            with col1:
                empresa        = st.text_input("Empresa / Proveedor",
                                               value=_clean_empresa(parsed.get("empresa", "")))
                numero_factura = st.text_input("Número de Factura *",
                                               value=parsed.get("numero_factura", ""))
                cuit           = st.text_input("CUIT Emisor",
                                               value=parsed.get("cuit", ""))
                tipo_factura   = st.selectbox(
                                               "Tipo de Factura",
                                               ["", "A", "B", "C", "E", "M"],
                                               index=["","A","B","C","E","M"].index(
                                                   parsed.get("tipo_factura", "") if parsed.get("tipo_factura","") in ["A","B","C","E","M"] else ""
                                               ))
            with col2:
                cliente        = st.text_input("Cliente / Titular",
                                               value=parsed.get("cliente", ""))
                numero_cliente = st.text_input("N° Cliente / Asociado",
                                               value=parsed.get("numero_cliente", ""))
                numero_cuenta  = st.text_input("Número de Cuenta",
                                               value=parsed.get("numero_cuenta", ""))
                estado_pago    = st.selectbox("Estado de Pago",
                                              ["Pendiente", "Paga"], index=0)

            col3, col4 = st.columns(2)
            with col3:
                fecha_emision  = st.text_input("Fecha de Emisión (YYYY-MM-DD)",
                                               value=parsed.get("fecha_emision", ""))
                fecha_envio    = st.text_input("Vencimiento (YYYY-MM-DD)",
                                               value=parsed.get("fecha_vencimiento",
                                                               parsed.get("fecha_envio", "")))
            with col4:
                total_factura  = st.text_input("Total Factura",
                                               value=parsed.get("total_factura", ""))
                importe        = st.text_input("Total a Pagar",
                                               value=parsed.get("total_a_pagar",
                                                               parsed.get("importe", "")))

            col_save, col_skip, _ = st.columns([1, 1, 4])
            with col_save:
                save_btn = st.form_submit_button(
                    "💾  Guardar", type="primary", use_container_width=True
                )
            with col_skip:
                skip_btn = st.form_submit_button(
                    "⏭  Saltar", use_container_width=True
                )

            if save_btn:
                invoice: Dict = {
                    "numero_cuenta":  numero_cuenta.strip(),
                    "numero_cliente": numero_cliente.strip(),
                    "empresa":        empresa.strip(),
                    "numero_factura": numero_factura.strip(),
                    "cuit":           cuit.strip(),
                    "tipo_factura":   tipo_factura,
                    "importe":        importe.strip(),
                    "fecha_emision":  fecha_emision.strip(),
                    "fecha_envio":    fecha_envio.strip(),
                    "estado_pago":    estado_pago,
                    "archivo":        name,
                    "ruta":           "",
                    "fecha_carga":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }

                # Save PDF to Facturas/YYYY-MM/
                saved_path = _save_pdf_bytes(name, data_bytes, fecha_emision.strip())
                if saved_path:
                    invoice["ruta"]    = saved_path
                    invoice["archivo"] = Path(saved_path).name

                ok = _excel().add_invoice(invoice)
                if ok:
                    st.session_state.proc_results["added"] += 1
                else:
                    st.session_state.proc_results["duplicates"] += 1
                    _flash(
                        "warning",
                        f"⚠️ Duplicado: N° {numero_factura} de {empresa} "
                        "ya existe en el registro y fue omitida.",
                    )

                st.session_state.proc_idx += 1
                st.rerun()

            elif skip_btn:
                st.session_state.proc_results["skipped"] += 1
                st.session_state.proc_idx += 1
                st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
#  TAB 2 — CONTROL DE FACTURAS
# ═══════════════════════════════════════════════════════════════════════════════

with tab_control:
    em       = _excel()
    all_inv  = em.get_all_invoices()
    total_n  = len(all_inv)
    pending_n = sum(1 for i in all_inv
                    if i.get("estado_pago", "").lower() == "pendiente")
    paid_n   = total_n - pending_n

    # ── Metrics + download ───────────────────────────────────────────────────
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total facturas", total_n)
    m2.metric("Pendientes",     pending_n)
    m3.metric("Pagadas",        paid_n)
    with m4:
        _excel_download_button()

    st.divider()

    if not all_inv:
        st.markdown(
            '<div class="empty-state">'
            "📂  No hay facturas registradas todavía.<br>"
            "Importa archivos PDF en la pestaña <strong>📁 Importar Facturas</strong>."
            "</div>",
            unsafe_allow_html=True,
        )
    else:
        # ── Filters ─────────────────────────────────────────────────────────
        fc1, fc2, fc3 = st.columns([2, 4, 1])
        with fc1:
            estado_filter = st.selectbox(
                "Estado", ["Todos", "Pendiente", "Paga"],
                label_visibility="collapsed",
            )
        with fc2:
            search = st.text_input(
                "Buscar", placeholder="Buscar empresa, N° factura…",
                label_visibility="collapsed",
            )
        with fc3:
            st.button("↺  Actualizar", use_container_width=True)

        # Build DataFrame with display headers
        header_map  = {k: h for k, h in COLUMNS}
        keys        = [k for k, _ in COLUMNS]
        headers     = [h for _, h in COLUMNS]

        df = pd.DataFrame(all_inv)[keys]
        df.columns = headers

        # Apply filters
        view_df = df.copy()
        if estado_filter != "Todos":
            view_df = view_df[
                view_df["Estado de Pago"].str.lower() == estado_filter.lower()
            ]
        if search.strip():
            mask = view_df.apply(
                lambda row: row.astype(str)
                              .str.contains(search.strip(), case=False, na=False)
                              .any(),
                axis=1,
            )
            view_df = view_df[mask]

        st.caption(
            f"Mostrando **{len(view_df)}** de **{total_n}** facturas  "
            "— edita la columna *Estado de Pago* y presiona **Guardar cambios**"
        )

        # ── Editable table ───────────────────────────────────────────────────
        non_editable_headers = [h for k, h in COLUMNS if k != "estado_pago"]

        edited_df = st.data_editor(
            view_df.reset_index(drop=True),
            column_config={
                "Estado de Pago": st.column_config.SelectboxColumn(
                    "Estado de Pago",
                    options=["Pendiente", "Paga"],
                    width="medium",
                    required=True,
                ),
                "Ruta del Archivo": st.column_config.TextColumn(width="large"),
                "Importe": st.column_config.TextColumn(width="small"),
                "Fecha de Emisión": st.column_config.TextColumn(width="medium"),
                "Fecha de Envío":   st.column_config.TextColumn(width="medium"),
                "Fecha de Carga":   st.column_config.TextColumn(width="medium"),
            },
            disabled=non_editable_headers,
            use_container_width=True,
            hide_index=True,
            key="invoice_editor",
        )

        # ── Save status changes ──────────────────────────────────────────────
        col_save_btn, _ = st.columns([1, 5])
        with col_save_btn:
            if st.button("💾  Guardar cambios", use_container_width=True):
                orig_map = {
                    (str(row["Empresa"]), str(row["Número de Factura"])):
                        str(row["Estado de Pago"])
                    for _, row in view_df.iterrows()
                }
                edit_map = {
                    (str(row["Empresa"]), str(row["Número de Factura"])):
                        str(row["Estado de Pago"])
                    for _, row in edited_df.iterrows()
                }
                changed = [
                    (emp, num, new_s)
                    for (emp, num), new_s in edit_map.items()
                    if orig_map.get((emp, num)) != new_s
                ]
                if changed:
                    em2 = _excel()
                    for emp, num, new_s in changed:
                        em2.update_invoice_status(emp, num, new_s)
                    _flash("success",
                           f"✅ Estado actualizado para {len(changed)} factura(s).")
                    st.rerun()
                else:
                    st.info("No se detectaron cambios en el estado de pago.")

# ═══════════════════════════════════════════════════════════════════════════════
#  TAB 3 — EMAIL A FINANZAS
# ═══════════════════════════════════════════════════════════════════════════════

with tab_email:
    em       = _excel()
    pending  = em.get_pending_invoices()
    ecfg     = _cfg().get("email", {})
    is_ok    = all(
        ecfg.get(k, "").strip()
        for k in ("remitente", "finanzas", "password", "smtp_server")
    )

    col_left, col_right = st.columns([3, 2])

    with col_left:
        st.subheader("Facturas pendientes de pago")
        if not pending:
            st.success("✅  No hay facturas pendientes.")
        else:
            st.warning(f"⚠️  **{len(pending)}** factura(s) pendiente(s)")
            show_keys = ["empresa", "numero_factura", "importe",
                         "fecha_emision", "fecha_envio"]
            hdr       = {k: h for k, h in COLUMNS}
            pdf_df    = (
                pd.DataFrame(pending)
                  .reindex(columns=[c for c in show_keys
                                    if c in pd.DataFrame(pending).columns])
                  .rename(columns=hdr)
            )
            st.dataframe(pdf_df, use_container_width=True, hide_index=True)

    with col_right:
        st.subheader("Enviar notificación")

        if not is_ok:
            st.error(
                "El email no está configurado.\n\n"
                "Ve a **⚙️ Configuración** e ingresa los datos SMTP."
            )
        elif not pending:
            st.info("Nada que enviar.")
        else:
            attach_pdfs = st.checkbox(
                "Adjuntar PDFs de las facturas",
                value=False,
                help="Adjunta los archivos PDF al email. "
                     "Puede aumentar considerablemente el tamaño del mensaje.",
            )

            if st.button(
                "📤  Enviar Email a Finanzas",
                type="primary",
                use_container_width=True,
            ):
                with st.spinner("Enviando…"):
                    ok, msg = EmailSender(ecfg).send(
                        pending,
                        _cfg().get("excel_filename", "control_facturas.xlsx"),
                        attach_pdfs=attach_pdfs,
                    )
                (st.success if ok else st.error)(
                    f"{'✅' if ok else '❌'}  {msg}"
                )

        st.divider()
        st.caption(
            f"**Remitente:** {ecfg.get('remitente') or '—'}  \n"
            f"**Destinatario:** {ecfg.get('finanzas') or '—'}  \n"
            f"**SMTP:** {ecfg.get('smtp_server') or '—'}"
            f":{ecfg.get('smtp_port') or '—'}"
        )

# ═══════════════════════════════════════════════════════════════════════════════
#  TAB 4 — CONFIGURACIÓN
# ═══════════════════════════════════════════════════════════════════════════════

with tab_config:

    # ── OpenAI API Key ───────────────────────────────────────────────────────
    st.subheader("🤖 Inteligencia Artificial (OpenAI)")
    st.markdown(
        "La app usa **GPT-4o-mini** para extraer automáticamente los datos de las facturas. "
        "Sin API key, usa un parser regex básico."
    )

    with st.form("ai_config"):
        ai_key_input = st.text_input(
            "OPENAI_API_KEY",
            value=_cfg().get("openai_api_key", ""),
            type="password",
            placeholder="sk-...",
            help="Obtené tu key en https://platform.openai.com/api-keys",
        )
        ai_col1, ai_col2, _ = st.columns([1, 1, 3])
        with ai_col1:
            save_ai = st.form_submit_button("💾  Guardar key", type="primary", use_container_width=True)
        with ai_col2:
            test_ai = st.form_submit_button("🔌  Probar conexión", use_container_width=True)

        if save_ai:
            st.session_state.cfg["openai_api_key"] = ai_key_input.strip()
            save_config(st.session_state.cfg)
            # Limpiar caché de parseo para que se re-procese con la nueva key
            st.session_state.parsed_cache = {}
            st.success("✅  API key guardada. El caché de extracción fue limpiado.")

        if test_ai:
            if not ai_key_input.strip():
                st.error("Ingresá una API key primero.")
            else:
                try:
                    import openai as _oa
                    _oa.OpenAI(api_key=ai_key_input.strip()).models.list()
                    st.success("✅  Conexión con OpenAI exitosa.")
                except Exception as e:
                    st.error(f"❌  Error: {e}")

    # Estado actual
    current_key = _cfg().get("openai_api_key", "") or os.environ.get("OPENAI_API_KEY", "")
    if current_key:
        st.caption(f"✅ Key activa: `{current_key[:8]}...{current_key[-4:]}`")
    else:
        st.caption("⚠️ Sin API key — usando regex fallback.")

    st.divider()

    # ── Email settings ───────────────────────────────────────────────────────
    st.subheader("Configuración de Email (SMTP)")

    ecfg = _cfg().get("email", {})

    with st.form("email_config"):
        c1, c2 = st.columns(2)
        with c1:
            rem    = st.text_input("Email Remitente",
                                   value=ecfg.get("remitente", ""),
                                   placeholder="tu@email.com")
            fin    = st.text_input("Email Finanzas (destinatario)",
                                   value=ecfg.get("finanzas", ""),
                                   placeholder="finanzas@empresa.com")
            pwd    = st.text_input("Contraseña / App Password",
                                   value=ecfg.get("password", ""),
                                   type="password")
        with c2:
            srv    = st.text_input("Servidor SMTP",
                                   value=ecfg.get("smtp_server", "smtp.gmail.com"))
            port   = st.number_input("Puerto SMTP",
                                     value=int(ecfg.get("smtp_port", 587)),
                                     min_value=1, max_value=65535, step=1)
            tls    = st.checkbox("Usar TLS / STARTTLS",
                                 value=ecfg.get("use_tls", True))

        btn_s, btn_t, _ = st.columns([1, 1, 4])
        with btn_s:
            save_email = st.form_submit_button(
                "💾  Guardar", type="primary", use_container_width=True
            )
        with btn_t:
            test_conn = st.form_submit_button(
                "🔌  Probar conexión", use_container_width=True
            )

        new_ecfg = {
            "remitente":   rem.strip(),
            "finanzas":    fin.strip(),
            "password":    pwd,
            "smtp_server": srv.strip(),
            "smtp_port":   int(port),
            "use_tls":     tls,
        }

        if save_email:
            st.session_state.cfg["email"] = new_ecfg
            save_config(st.session_state.cfg)
            st.success("✅  Configuración guardada.")

        if test_conn:
            with st.spinner("Probando conexión SMTP…"):
                ok, msg = EmailSender(new_ecfg).test_connection()
            (st.success if ok else st.error)(f"{'✅' if ok else '❌'}  {msg}")

    st.divider()

    # ── Paths settings ───────────────────────────────────────────────────────
    st.subheader("Rutas del sistema")

    with st.form("paths_config"):
        p1, p2 = st.columns(2)
        with p1:
            new_base = st.text_input(
                "Carpeta de facturas",
                value=_cfg().get("facturas_base_dir", "Facturas"),
                help="Ruta relativa o absoluta donde se guardarán los PDFs.",
            )
        with p2:
            new_xl = st.text_input(
                "Archivo Excel",
                value=_cfg().get("excel_filename", "control_facturas.xlsx"),
                help="Nombre (o ruta) del archivo Excel de control.",
            )

        if st.form_submit_button("💾  Guardar rutas"):
            st.session_state.cfg["facturas_base_dir"] = new_base.strip() or "Facturas"
            st.session_state.cfg["excel_filename"]     = new_xl.strip() or "control_facturas.xlsx"
            save_config(st.session_state.cfg)
            st.success("✅  Rutas guardadas.")

    st.divider()

    # ── Help notes ───────────────────────────────────────────────────────────
    with st.expander("📋  Ayuda — configuración de email por proveedor"):
        st.markdown(
            """
**Gmail**
- Servidor: `smtp.gmail.com` · Puerto: `587` · TLS: ✅
- Usa una **App Password** (no tu contraseña normal).
  Actívala en: *Cuenta Google → Seguridad → Verificación en dos pasos → Contraseñas de aplicación*.

**Outlook / Microsoft 365**
- Servidor: `smtp.office365.com` · Puerto: `587` · TLS: ✅

**Yahoo Mail**
- Servidor: `smtp.mail.yahoo.com` · Puerto: `587` · TLS: ✅
- Requiere también una App Password desde la configuración de seguridad de la cuenta.

**Servidor propio / corporativo**
- Consulta a tu equipo de IT los datos SMTP.
- El campo *Contraseña* puede dejarse vacío si el servidor no requiere autenticación.
"""
        )
