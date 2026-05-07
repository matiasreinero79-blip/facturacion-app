from pathlib import Path
from datetime import datetime
from typing import Dict, List

import openpyxl
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side
)
from openpyxl.utils import get_column_letter

# Canonical column order: (field_key, display_header)
COLUMNS: List[tuple] = [
    ("numero_cuenta",   "Número de Cuenta"),
    ("numero_cliente",  "Número de Cliente"),
    ("empresa",         "Empresa"),
    ("numero_factura",  "Número de Factura"),
    ("importe",         "Importe"),
    ("fecha_emision",   "Fecha de Emisión"),
    ("fecha_envio",     "Fecha de Envío"),
    ("estado_pago",     "Estado de Pago"),
    ("archivo",         "Archivo PDF"),
    ("ruta",            "Ruta del Archivo"),
    ("fecha_carga",     "Fecha de Carga"),
]

_COL_WIDTHS = [18, 18, 28, 22, 14, 16, 16, 14, 32, 55, 22]

_FILL_HEADER   = PatternFill("solid", fgColor="1F4E79")
_FILL_EVEN     = PatternFill("solid", fgColor="DCE6F1")
_FILL_ODD      = PatternFill("solid", fgColor="FFFFFF")
_FILL_PAGA     = PatternFill("solid", fgColor="C6EFCE")
_FILL_PENDIENTE= PatternFill("solid", fgColor="FFEB9C")

_FONT_HEADER    = Font(color="FFFFFF", bold=True, size=11, name="Arial")
_FONT_PAGA      = Font(color="276221", name="Arial")
_FONT_PENDIENTE = Font(color="9C5700", name="Arial")

_THIN = Side(style="thin", color="BFBFBF")
_BORDER= Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

_ESTADO_COL_IDX = next(i for i, (k, _) in enumerate(COLUMNS) if k == "estado_pago")


class ExcelManager:

    def __init__(self, excel_path: str = "control_facturas.xlsx"):
        self.excel_path = Path(excel_path)

    # ─── public API ─────────────────────────────────────────────────────────

    def add_invoice(self, data: Dict) -> bool:
        """
        Append *data* as a new row.
        Returns False (and does nothing) when a duplicate is detected.
        """
        empresa = (data.get("empresa") or "").strip()
        numero  = (data.get("numero_factura") or "").strip()

        if empresa and numero and self.invoice_exists(empresa, numero):
            return False

        if not data.get("fecha_carga"):
            data["fecha_carga"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        wb = self._load_or_create()
        ws = wb.active
        row_idx = self._next_row(ws)
        self._write_row(ws, row_idx, data)
        wb.save(str(self.excel_path))
        return True

    def invoice_exists(self, empresa: str, numero_factura: str) -> bool:
        """Check by empresa + numero_factura (case-insensitive)."""
        emp_lo = empresa.strip().lower()
        num_lo = numero_factura.strip().lower()
        for inv in self.get_all_invoices():
            if (inv.get("empresa", "").lower() == emp_lo and
                    inv.get("numero_factura", "").lower() == num_lo):
                return True
        return False

    def get_all_invoices(self) -> List[Dict]:
        if not self.excel_path.exists():
            return []
        wb = openpyxl.load_workbook(str(self.excel_path), read_only=True, data_only=True)
        ws = wb.active
        invoices = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(c is not None for c in row):
                continue
            inv = {}
            for idx, (key, _) in enumerate(COLUMNS):
                val = row[idx] if idx < len(row) else None
                inv[key] = "" if val is None else str(val)
            invoices.append(inv)
        wb.close()
        return invoices

    def get_pending_invoices(self) -> List[Dict]:
        return [i for i in self.get_all_invoices()
                if i.get("estado_pago", "").lower() == "pendiente"]

    def update_invoice_status(self, empresa: str, numero_factura: str, new_status: str) -> bool:
        if not self.excel_path.exists():
            return False
        wb = openpyxl.load_workbook(str(self.excel_path))
        ws = wb.active

        emp_col = next(i + 1 for i, (k, _) in enumerate(COLUMNS) if k == "empresa")
        num_col = next(i + 1 for i, (k, _) in enumerate(COLUMNS) if k == "numero_factura")
        est_col = _ESTADO_COL_IDX + 1

        found = False
        for row in ws.iter_rows(min_row=2):
            emp_val = str(row[emp_col - 1].value or "").strip().lower()
            num_val = str(row[num_col - 1].value or "").strip().lower()
            if emp_val == empresa.strip().lower() and num_val == numero_factura.strip().lower():
                cell = row[est_col - 1]
                cell.value = new_status
                if new_status.lower() == "paga":
                    cell.fill = _FILL_PAGA
                    cell.font = _FONT_PAGA
                else:
                    cell.fill = _FILL_PENDIENTE
                    cell.font = _FONT_PENDIENTE
                found = True
                break

        if found:
            wb.save(str(self.excel_path))
        return found

    # ─── internals ──────────────────────────────────────────────────────────

    def _load_or_create(self) -> openpyxl.Workbook:
        if self.excel_path.exists():
            return openpyxl.load_workbook(str(self.excel_path))
        return self._new_workbook()

    def _new_workbook(self) -> openpyxl.Workbook:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Facturas"
        ws.freeze_panes = "A2"

        for col_idx, (_, header) in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.fill   = _FILL_HEADER
            cell.font   = _FONT_HEADER
            cell.border = _BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        ws.row_dimensions[1].height = 28
        for i, w in enumerate(_COL_WIDTHS, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

        return wb

    def _next_row(self, ws) -> int:
        row = ws.max_row + 1
        # If the sheet only has the header, start at row 2
        if row == 2 and ws.cell(2, 1).value is None:
            row = 2
        return row

    def _write_row(self, ws, row_idx: int, data: Dict) -> None:
        row_fill = _FILL_EVEN if row_idx % 2 == 0 else _FILL_ODD

        for col_idx, (key, _) in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=data.get(key, "") or "")
            cell.border    = _BORDER
            cell.alignment = Alignment(vertical="center")
            cell.fill      = row_fill

        # Colour the estado_pago cell
        estado_cell = ws.cell(row=row_idx, column=_ESTADO_COL_IDX + 1)
        if (data.get("estado_pago") or "").lower() == "paga":
            estado_cell.fill = _FILL_PAGA
            estado_cell.font = _FONT_PAGA
        else:
            estado_cell.fill = _FILL_PENDIENTE
            estado_cell.font = _FONT_PENDIENTE
