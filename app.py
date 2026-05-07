import os
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional

from config import load_config, save_config
from modules.excel_manager import COLUMNS, ExcelManager
from modules.email_sender import EmailSender
from modules.file_manager import FileManager
from modules.importers.folder_importer import FolderImporter
from modules.pdf_reader import parse_invoice
from modules.utils import mono_font, open_path, preferred_themes, system_font

# ─── Column indices (derived from COLUMNS so they never go stale) ────────────
_KEY_IDX            = {key: i for i, (key, _) in enumerate(COLUMNS)}
_IDX_EMPRESA        = _KEY_IDX["empresa"]
_IDX_NUMERO_FACTURA = _KEY_IDX["numero_factura"]
_IDX_ESTADO_PAGO    = _KEY_IDX["estado_pago"]
_IDX_RUTA           = _KEY_IDX["ruta"]


# ═══════════════════════════════════════════════════════════════════════════════
#  Invoice entry dialog
# ═══════════════════════════════════════════════════════════════════════════════

class InvoiceDialog(tk.Toplevel):
    """Modal form for reviewing / completing invoice data before saving."""

    _FIELDS = [
        ("numero_cuenta",  "Número de Cuenta:"),
        ("numero_cliente", "Número de Cliente:"),
        ("empresa",        "Empresa: *"),
        ("numero_factura", "Número de Factura: *"),
        ("importe",        "Importe:"),
        ("fecha_emision",  "Fecha de Emisión (YYYY-MM-DD):"),
        ("fecha_envio",    "Fecha de Envío   (YYYY-MM-DD):"),
    ]

    def __init__(self, parent: tk.Tk, data: dict, title: str = "Datos de Factura"):
        super().__init__(parent)
        self.title(title)
        self.resizable(True, True)
        self.geometry("540x560")
        self.transient(parent)
        self.grab_set()
        self.result: Optional[dict] = None
        self._extra = {
            "estado_pago": data.get("estado_pago", "Pendiente"),
            "archivo":     data.get("archivo", ""),
            "ruta":        data.get("ruta", ""),
            "fecha_carga": data.get("fecha_carga",
                                    datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        }

        root = ttk.Frame(self, padding=20)
        root.pack(fill="both", expand=True)

        ttk.Label(root, text=title, font=system_font(12, bold=True)).pack(
            anchor="w", pady=(0, 12))

        # File info banner
        if data.get("archivo"):
            banner = ttk.Frame(root)
            banner.pack(fill="x", pady=(0, 10))
            ttk.Label(banner, text=f"\U0001f4c4  {data['archivo']}",
                      foreground="#1F4E79",
                      font=system_font(9)).pack(anchor="w", padx=8, pady=4)

        # Form grid
        form = ttk.Frame(root)
        form.pack(fill="both", expand=True)
        form.columnconfigure(1, weight=1)

        self._vars: Dict[str, tk.StringVar] = {}
        for row, (key, label) in enumerate(self._FIELDS):
            ttk.Label(form, text=label, anchor="w").grid(
                row=row, column=0, sticky="w", padx=(0, 10), pady=5)
            var = tk.StringVar(value=str(data.get(key) or ""))
            ttk.Entry(form, textvariable=var).grid(
                row=row, column=1, sticky="ew", pady=5)
            self._vars[key] = var

        # Estado de pago
        estado_row = len(self._FIELDS)
        ttk.Label(form, text="Estado de Pago:").grid(
            row=estado_row, column=0, sticky="w", padx=(0, 10), pady=5)
        self._estado_var = tk.StringVar(value=self._extra["estado_pago"])
        ttk.Combobox(form, textvariable=self._estado_var,
                     values=["Pendiente", "Paga"],
                     state="readonly", width=14).grid(
            row=estado_row, column=1, sticky="w", pady=5)

        ttk.Label(form,
                  text="* campos requeridos para control de duplicados",
                  foreground="gray",
                  font=system_font(8)).grid(
            row=estado_row + 1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # Buttons
        btn = ttk.Frame(root)
        btn.pack(fill="x", pady=(14, 0))
        ttk.Button(btn, text="Cancelar (saltar)", command=self.destroy).pack(
            side="right", padx=(8, 0))
        ttk.Button(btn, text="Guardar", command=self._save).pack(side="right")

        self.wait_window()

    def _save(self) -> None:
        self.result = {k: v.get().strip() for k, v in self._vars.items()}
        self.result["estado_pago"] = self._estado_var.get()
        self.result.update(self._extra)
        self.destroy()


# ═══════════════════════════════════════════════════════════════════════════════
#  Settings dialog
# ═══════════════════════════════════════════════════════════════════════════════

class SettingsDialog(tk.Toplevel):

    _FIELDS = [
        ("remitente",   "Email Remitente:",          False),
        ("finanzas",    "Email Finanzas:",            False),
        ("password",    "Contraseña / App Password:", True),
        ("smtp_server", "Servidor SMTP:",             False),
        ("smtp_port",   "Puerto SMTP:",               False),
    ]

    def __init__(self, parent: tk.Tk, app_config: dict):
        super().__init__(parent)
        self.title("Configuración de Email")
        self.geometry("460x380")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.result: Optional[dict] = None
        self._app_config = app_config
        email_cfg = app_config.get("email", {})

        root = ttk.Frame(self, padding=24)
        root.pack(fill="both", expand=True)

        ttk.Label(root, text="Configuración de Email",
                  font=system_font(12, bold=True)).pack(anchor="w", pady=(0, 14))

        form = ttk.Frame(root)
        form.pack(fill="both", expand=True)
        form.columnconfigure(1, weight=1)

        self._vars: Dict[str, tk.StringVar] = {}
        for row, (key, label, secret) in enumerate(self._FIELDS):
            ttk.Label(form, text=label, anchor="w").grid(
                row=row, column=0, sticky="w", padx=(0, 12), pady=6)
            var = tk.StringVar(value=str(email_cfg.get(key, "") or ""))
            ttk.Entry(form, textvariable=var,
                      show=("*" if secret else "")).grid(
                row=row, column=1, sticky="ew", pady=6)
            self._vars[key] = var

        tls_row = len(self._FIELDS)
        self._tls_var = tk.BooleanVar(value=email_cfg.get("use_tls", True))
        ttk.Checkbutton(form, text="Usar TLS / STARTTLS",
                        variable=self._tls_var).grid(
            row=tls_row, column=1, sticky="w", pady=6)

        btn = ttk.Frame(root)
        btn.pack(fill="x", pady=(16, 0))
        ttk.Button(btn, text="Probar conexión", command=self._test).pack(side="left")
        ttk.Button(btn, text="Cancelar", command=self.destroy).pack(
            side="right", padx=(8, 0))
        ttk.Button(btn, text="Guardar", command=self._save).pack(side="right")

        self.wait_window()

    def _collect(self) -> dict:
        cfg = {k: v.get().strip() for k, v in self._vars.items()}
        cfg["smtp_port"] = int(cfg.get("smtp_port") or 587)
        cfg["use_tls"]   = self._tls_var.get()
        return cfg

    def _test(self) -> None:
        self.config(cursor="wait")
        self.update()
        ok, msg = EmailSender(self._collect()).test_connection()
        self.config(cursor="")
        (messagebox.showinfo if ok else messagebox.showerror)(
            "Resultado", msg, parent=self)

    def _save(self) -> None:
        self.result = self._collect()
        self.destroy()


# ═══════════════════════════════════════════════════════════════════════════════
#  Main application
# ═══════════════════════════════════════════════════════════════════════════════

class FacturacionApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Sistema de Gestión de Facturas")
        self.geometry("1150x720")
        self.minsize(900, 580)

        self._cfg       = load_config()
        self._file_mgr  = FileManager(self._cfg.get("facturas_base_dir", "Facturas"))
        self._excel_mgr = ExcelManager(self._cfg.get("excel_filename",
                                                      "control_facturas.xlsx"))

        # Processing queue state
        self._queue:   List[Path] = []
        self._total:   int = 0
        self._results: Dict[str, int] = {}

        self._selected_folder = tk.StringVar()
        self._status_var      = tk.StringVar(value="Listo.")
        self._found_pdfs:  List[Path] = []

        self._setup_styles()
        self._build_ui()
        self._refresh_table()

    # ─── styles ─────────────────────────────────────────────────────────────

    def _setup_styles(self) -> None:
        s = ttk.Style(self)

        # Pick the best available theme for this OS
        for theme in preferred_themes():
            try:
                s.theme_use(theme)
                break
            except tk.TclError:
                continue

        # On macOS the "aqua" theme ignores most configure() calls — that is
        # intentional: the OS renders native widgets automatically.
        # On Windows and Linux we apply a custom colour scheme.
        if sys.platform != "darwin":
            bg = "#eef2f7"
            self.configure(bg=bg)
            s.configure("TFrame",           background=bg)
            s.configure("TLabel",           background=bg)
            s.configure("TLabelframe",      background=bg)
            s.configure("TLabelframe.Label",background=bg)
            s.configure("TNotebook",        background=bg)
            s.configure("Status.TLabel",    background="#dde5ed")
        else:
            # On macOS we only override what doesn't break the native look
            pass

        s.configure("TLabel",           font=system_font(10))
        s.configure("TLabelframe.Label",font=system_font(10, bold=True),
                                         foreground="#1F4E79")
        s.configure("TNotebook.Tab",    font=system_font(10), padding=(14, 6))
        s.configure("Status.TLabel",    font=system_font(9),  foreground="#555")
        s.configure("Accent.TButton",   font=system_font(10, bold=True))

    # ─── UI construction ────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self._build_header()
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=(8, 4))
        self._build_import_tab(nb)
        self._build_view_tab(nb)
        self._build_statusbar()

    def _build_header(self) -> None:
        # Use a plain tk.Frame so the background colour works on all platforms
        hdr = tk.Frame(self, bg="#1F4E79", pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Sistema de Gestión de Facturas",
                 font=system_font(15, bold=True),
                 fg="white", bg="#1F4E79").pack(side="left", padx=20)
        ttk.Button(hdr, text="Configuración Email",
                   command=self._open_settings).pack(side="right", padx=12, pady=2)

    def _build_import_tab(self, nb: ttk.Notebook) -> None:
        tab = ttk.Frame(nb, padding=16)
        nb.add(tab, text="  Importar Facturas  ")

        # ── Source selection ──────────────────────────────────────────────
        src = ttk.LabelFrame(tab, text="1.  Seleccionar carpeta de facturas",
                             padding=12)
        src.pack(fill="x", pady=(0, 10))

        row = ttk.Frame(src)
        row.pack(fill="x")
        row.columnconfigure(1, weight=1)

        ttk.Label(row, text="Carpeta:").grid(row=0, column=0, sticky="w",
                                             padx=(0, 8))
        ttk.Entry(row, textvariable=self._selected_folder).grid(
            row=0, column=1, sticky="ew", padx=(0, 8))
        ttk.Button(row, text="Examinar…",
                   command=self._browse).grid(row=0, column=2, padx=(0, 6))
        ttk.Button(row, text="Buscar PDFs",
                   command=self._find_pdfs,
                   style="Accent.TButton").grid(row=0, column=3)

        ttk.Label(src,
                  text="Próximamente: Email IMAP  |  Portal Web  |"
                       "  Carpeta compartida en red",
                  foreground="#888",
                  font=system_font(8)).pack(anchor="w", pady=(6, 0))

        # ── Found-PDF list ────────────────────────────────────────────────
        lst = ttk.LabelFrame(
            tab,
            text="2.  Facturas encontradas  "
                 "(Ctrl+clic o Shift+clic para seleccionar varias)",
            padding=8,
        )
        lst.pack(fill="both", expand=True, pady=(0, 10))

        lf = ttk.Frame(lst)
        lf.pack(fill="both", expand=True)

        sb = ttk.Scrollbar(lf, orient="vertical")
        sb.pack(side="right", fill="y")
        self._listbox = tk.Listbox(
            lf,
            selectmode="extended",
            height=10,
            font=mono_font(9),
            activestyle="dotbox",
            yscrollcommand=sb.set,
        )
        self._listbox.pack(fill="both", expand=True)
        sb.config(command=self._listbox.yview)

        self._found_lbl = ttk.Label(lst, text="0 archivos encontrados",
                                    foreground="gray", font=system_font(9))
        self._found_lbl.pack(anchor="w", pady=(4, 0))

        # ── Action buttons ────────────────────────────────────────────────
        btns = ttk.Frame(tab)
        btns.pack(fill="x")

        ttk.Button(btns, text="Procesar Todas",
                   command=self._process_all,
                   style="Accent.TButton").pack(side="left", padx=(0, 6))
        ttk.Button(btns, text="Procesar Seleccionadas",
                   command=self._process_selected).pack(side="left", padx=(0, 6))

        ttk.Button(btns, text="Enviar Email a Finanzas",
                   command=self._send_email).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="Abrir Excel",
                   command=self._open_excel).pack(side="right")

        self._progress = ttk.Progressbar(tab, mode="determinate")
        self._progress.pack(fill="x", pady=(8, 0))

    def _build_view_tab(self, nb: ttk.Notebook) -> None:
        tab = ttk.Frame(nb, padding=16)
        nb.add(tab, text="  Control de Facturas  ")

        # ── Toolbar ───────────────────────────────────────────────────────
        tb = ttk.Frame(tab)
        tb.pack(fill="x", pady=(0, 8))

        ttk.Button(tb, text="Actualizar",
                   command=self._refresh_table).pack(side="left", padx=(0, 8))
        ttk.Button(tb, text="Abrir Excel",
                   command=self._open_excel).pack(side="left")

        ttk.Label(tb, text="Estado:").pack(side="left", padx=(16, 4))
        self._filter_var = tk.StringVar(value="Todos")
        cb = ttk.Combobox(tb, textvariable=self._filter_var,
                          values=["Todos", "Pendiente", "Paga"],
                          state="readonly", width=12)
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda _: self._refresh_table())

        self._count_lbl = ttk.Label(tb, text="", foreground="gray")
        self._count_lbl.pack(side="right")

        # ── Treeview ──────────────────────────────────────────────────────
        tf = ttk.Frame(tab)
        tf.pack(fill="both", expand=True)

        keys = [k for k, _ in COLUMNS]
        col_widths = {
            "numero_cuenta": 115, "numero_cliente": 115, "empresa": 160,
            "numero_factura": 130, "importe": 95, "fecha_emision": 105,
            "fecha_envio": 105, "estado_pago": 85, "archivo": 190,
            "ruta": 220, "fecha_carga": 130,
        }

        self._tree = ttk.Treeview(tf, columns=keys, show="headings", height=20)
        for key, header in COLUMNS:
            self._tree.heading(key, text=header, anchor="w",
                               command=lambda k=key: self._sort_tree(k))
            self._tree.column(key, width=col_widths.get(key, 120), anchor="w")

        vsb = ttk.Scrollbar(tf, orient="vertical",   command=self._tree.yview)
        hsb = ttk.Scrollbar(tf, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right",  fill="y")
        hsb.pack(side="bottom", fill="x")
        self._tree.pack(fill="both", expand=True)

        self._tree.tag_configure("paga",      background="#e8f5e9")
        self._tree.tag_configure("pendiente", background="#fffde7")

        # Context menu — right-click on Windows/Linux; Ctrl+click on macOS
        self._ctx = tk.Menu(self, tearoff=0)
        self._ctx.add_command(label="Marcar como Paga",
                               command=lambda: self._set_status_selected("Paga"))
        self._ctx.add_command(label="Marcar como Pendiente",
                               command=lambda: self._set_status_selected("Pendiente"))
        self._ctx.add_separator()
        self._ctx.add_command(label="Abrir PDF", command=self._open_pdf)

        self._tree.bind("<Button-3>",         self._show_ctx)   # Windows & Linux
        self._tree.bind("<Control-Button-1>", self._show_ctx)   # macOS trackpad

    def _build_statusbar(self) -> None:
        bar = tk.Frame(self, bg="#dde5ed", pady=4)
        bar.pack(fill="x", side="bottom")
        ttk.Label(bar, textvariable=self._status_var,
                  style="Status.TLabel").pack(side="left", padx=12)

    # ─── import-tab actions ─────────────────────────────────────────────────

    def _browse(self) -> None:
        folder = filedialog.askdirectory(title="Seleccionar carpeta con facturas PDF")
        if folder:
            self._selected_folder.set(folder)
            self._find_pdfs()

    def _find_pdfs(self) -> None:
        folder = self._selected_folder.get().strip()
        if not folder:
            messagebox.showwarning("Aviso", "Selecciona una carpeta primero.")
            return
        imp = FolderImporter(folder)
        self._found_pdfs = imp.get_invoices()
        self._listbox.delete(0, "end")
        for p in self._found_pdfs:
            self._listbox.insert("end", p.name)
        n = len(self._found_pdfs)
        self._found_lbl.config(text=f"{n} archivo(s) encontrado(s)")
        self._set_status(f"{n} PDFs encontrados en: {folder}")

    def _process_all(self) -> None:
        if not self._found_pdfs:
            messagebox.showwarning("Aviso",
                "Primero selecciona una carpeta y presiona 'Buscar PDFs'.")
            return
        self._start_queue(list(self._found_pdfs))

    def _process_selected(self) -> None:
        sel = self._listbox.curselection()
        if not sel:
            messagebox.showwarning("Aviso",
                "Selecciona al menos un archivo en la lista.")
            return
        self._start_queue([self._found_pdfs[i] for i in sel])

    def _start_queue(self, pdfs: List[Path]) -> None:
        self._queue   = list(pdfs)
        self._total   = len(pdfs)
        self._results = {"added": 0, "duplicates": 0, "skipped": 0}
        self._next_in_queue()

    def _next_in_queue(self) -> None:
        done = sum(self._results.values())

        if not self._queue:
            self._progress["value"] = 0
            r = self._results
            messagebox.showinfo(
                "Procesamiento completado",
                f"Facturas procesadas:  {self._total}\n\n"
                f"  Agregadas al Excel:      {r['added']}\n"
                f"  Duplicadas (omitidas):   {r['duplicates']}\n"
                f"  Canceladas / sin datos:  {r['skipped']}",
            )
            self._refresh_table()
            self._set_status("Procesamiento completado.")
            return

        self._progress["value"] = (done / self._total) * 100
        pdf_path = self._queue.pop(0)
        self._set_status(
            f"Procesando {pdf_path.name}  ({done + 1}/{self._total})…")

        try:
            data = parse_invoice(str(pdf_path))
        except Exception:
            data = {
                "archivo":     pdf_path.name,
                "ruta":        str(pdf_path.resolve()),
                "fecha_carga": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }

        dialog = InvoiceDialog(self, data, title=f"Factura: {pdf_path.name}")

        if dialog.result is None:
            self._results["skipped"] += 1
            self.after(1, self._next_in_queue)
            return

        invoice = dialog.result

        # Organise PDF into Facturas/YYYY-MM/
        try:
            dest = self._file_mgr.copy_invoice(
                str(pdf_path), invoice.get("fecha_emision", ""))
            invoice["ruta"]    = str(dest.resolve())
            invoice["archivo"] = dest.name
        except Exception as e:
            self._set_status(f"Aviso al copiar archivo: {e}")

        ok = self._excel_mgr.add_invoice(invoice)
        if ok:
            self._results["added"] += 1
        else:
            self._results["duplicates"] += 1
            messagebox.showwarning(
                "Duplicado detectado",
                f"La factura  N° {invoice.get('numero_factura', '—')}  "
                f"de  {invoice.get('empresa', '—')}  ya existe y fue omitida.",
                parent=self,
            )

        self.after(1, self._next_in_queue)

    # ─── view-tab actions ────────────────────────────────────────────────────

    def _refresh_table(self) -> None:
        for item in self._tree.get_children():
            self._tree.delete(item)

        all_inv = self._excel_mgr.get_all_invoices()
        filt    = self._filter_var.get() if hasattr(self, "_filter_var") else "Todos"
        shown   = 0
        pending = 0

        for inv in all_inv:
            estado = inv.get("estado_pago", "")
            if estado.lower() == "pendiente":
                pending += 1
            if filt != "Todos" and estado.lower() != filt.lower():
                continue
            vals = [inv.get(k, "") for k, _ in COLUMNS]
            tag  = "paga" if estado.lower() == "paga" else "pendiente"
            self._tree.insert("", "end", values=vals, tags=(tag,))
            shown += 1

        if hasattr(self, "_count_lbl"):
            self._count_lbl.config(
                text=(f"Total: {len(all_inv)}  |  "
                      f"Pendientes: {pending}  |  "
                      f"Mostrando: {shown}"))

    def _sort_tree(self, col_key: str) -> None:
        """Toggle-sort the treeview by the clicked column header."""
        items = [(self._tree.set(child, col_key), child)
                 for child in self._tree.get_children("")]
        reverse = getattr(self, f"_sort_rev_{col_key}", False)
        items.sort(reverse=reverse)
        for idx, (_, child) in enumerate(items):
            self._tree.move(child, "", idx)
        setattr(self, f"_sort_rev_{col_key}", not reverse)

    def _set_status_selected(self, new_status: str) -> None:
        for item in self._tree.selection():
            vals = self._tree.item(item, "values")
            if vals:
                self._excel_mgr.update_invoice_status(
                    vals[_IDX_EMPRESA], vals[_IDX_NUMERO_FACTURA], new_status)
        self._refresh_table()
        self._set_status(f"Estado actualizado → {new_status}")

    def _open_pdf(self) -> None:
        sel = self._tree.selection()
        if not sel:
            return
        ruta = self._tree.item(sel[0], "values")[_IDX_RUTA]
        try:
            open_path(ruta)
        except FileNotFoundError:
            messagebox.showwarning("Archivo no encontrado",
                                   f"No se encontró el archivo:\n{ruta}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el PDF:\n{e}")

    def _show_ctx(self, event: tk.Event) -> None:
        row = self._tree.identify_row(event.y)
        if row:
            self._tree.selection_set(row)
            self._ctx.post(event.x_root, event.y_root)

    # ─── shared actions ──────────────────────────────────────────────────────

    def _open_excel(self) -> None:
        path = Path(self._cfg.get("excel_filename", "control_facturas.xlsx"))
        if not path.exists():
            messagebox.showinfo("Aviso",
                "El archivo Excel no existe todavía.\n"
                "Procesa algunas facturas primero.")
            return
        try:
            open_path(str(path.resolve()))
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el Excel:\n{e}")

    def _send_email(self) -> None:
        pending = self._excel_mgr.get_pending_invoices()
        if not pending:
            messagebox.showinfo("Sin pendientes",
                "No hay facturas con estado Pendiente en el registro.")
            return

        email_cfg = self._cfg.get("email", {})
        if not all(email_cfg.get(k, "").strip()
                   for k in ("remitente", "finanzas", "password")):
            if messagebox.askyesno("Configuración incompleta",
                    "Los datos de email no están configurados.\n"
                    "¿Deseas configurarlos ahora?"):
                self._open_settings()
            return

        attach = messagebox.askyesno(
            "¿Adjuntar PDFs?",
            f"Se enviarán {len(pending)} facturas pendientes.\n\n"
            "¿Deseas adjuntar los archivos PDF al email?\n"
            "(Puede aumentar el tamaño del mensaje considerablemente)",
        )

        self.config(cursor="wait")
        self.update()
        try:
            sender = EmailSender(email_cfg)
            excel  = self._cfg.get("excel_filename", "control_facturas.xlsx")
            ok, msg = sender.send(pending, excel, attach_pdfs=attach)
        finally:
            self.config(cursor="")

        (messagebox.showinfo if ok else messagebox.showerror)(
            "Email" if ok else "Error al enviar email", msg)
        self._set_status(msg)

    def _open_settings(self) -> None:
        dlg = SettingsDialog(self, self._cfg)
        if dlg.result:
            self._cfg["email"] = dlg.result
            save_config(self._cfg)
            messagebox.showinfo("Guardado", "Configuración guardada correctamente.")

    # ─── helpers ─────────────────────────────────────────────────────────────

    def _set_status(self, msg: str) -> None:
        self._status_var.set(msg)
        self.update_idletasks()
