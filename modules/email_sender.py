import smtplib
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Dict, List, Tuple


class EmailSender:
    """Send invoice-summary emails to the finance team via SMTP."""

    def __init__(self, config: dict):
        self.config = config

    # ─── public ─────────────────────────────────────────────────────────────

    def send(
        self,
        pending_invoices: List[Dict],
        excel_path: str,
        attach_pdfs: bool = False,
    ) -> Tuple[bool, str]:
        """Return (success, message)."""
        if not pending_invoices:
            return False, "No hay facturas pendientes para enviar."

        ok, err = self._validate_config()
        if not ok:
            return False, err

        pdf_paths = []
        if attach_pdfs:
            pdf_paths = [i.get("ruta", "") for i in pending_invoices if i.get("ruta")]

        msg = self._build_message(pending_invoices, excel_path, pdf_paths)
        return self._smtp_send(msg)

    def test_connection(self) -> Tuple[bool, str]:
        ok, err = self._validate_config()
        if not ok:
            return False, err
        try:
            with self._open_smtp() as server:
                pass
            return True, "Conexión SMTP exitosa."
        except Exception as e:
            return False, str(e)

    # ─── private ────────────────────────────────────────────────────────────

    def _validate_config(self) -> Tuple[bool, str]:
        for key in ("remitente", "finanzas", "password", "smtp_server"):
            if not self.config.get(key, "").strip():
                return False, f"Falta configuración de email: '{key}'."
        return True, ""

    def _open_smtp(self) -> smtplib.SMTP:
        server = smtplib.SMTP(
            self.config["smtp_server"],
            int(self.config.get("smtp_port", 587)),
            timeout=30,
        )
        if self.config.get("use_tls", True):
            server.starttls()
        server.login(self.config["remitente"], self.config["password"])
        return server

    def _smtp_send(self, msg: MIMEMultipart) -> Tuple[bool, str]:
        try:
            with self._open_smtp() as server:
                server.send_message(msg)
            dest = self.config["finanzas"]
            return True, f"Email enviado exitosamente a {dest}."
        except smtplib.SMTPAuthenticationError:
            return False, "Error de autenticación. Verifica tu email y contraseña / app password."
        except smtplib.SMTPConnectError:
            host = self.config.get("smtp_server")
            port = self.config.get("smtp_port")
            return False, f"No se pudo conectar a {host}:{port}."
        except Exception as e:
            return False, f"Error al enviar email: {e}"

    def _build_message(
        self,
        invoices: List[Dict],
        excel_path: str,
        pdf_paths: List[str],
    ) -> MIMEMultipart:
        msg = MIMEMultipart()
        msg["From"]    = self.config["remitente"]
        msg["To"]      = self.config["finanzas"]
        msg["Subject"] = "Facturas pendientes de pago"

        msg.attach(MIMEText(self._html_body(invoices), "html", "utf-8"))
        self._attach_file(msg, excel_path)
        for path in pdf_paths:
            self._attach_file(msg, path)

        return msg

    def _html_body(self, invoices: List[Dict]) -> str:
        rows = "".join(
            f"""<tr>
              <td>{i.get('empresa','')}</td>
              <td>{i.get('numero_factura','')}</td>
              <td>{i.get('importe','')}</td>
              <td>{i.get('fecha_emision','')}</td>
              <td>{i.get('fecha_envio','')}</td>
              <td style="font-size:11px;color:#555;">{i.get('ruta','')}</td>
            </tr>"""
            for i in invoices
        )
        now = datetime.now().strftime("%d/%m/%Y %H:%M")
        return f"""<!DOCTYPE html>
<html><body style="font-family:Arial,Helvetica,sans-serif;color:#333;max-width:900px;">
<h2 style="color:#1F4E79;border-bottom:2px solid #1F4E79;padding-bottom:6px;">
  Resumen de Facturas Pendientes de Pago
</h2>
<p>Fecha de envío: <strong>{now}</strong> &nbsp;|&nbsp;
   Total de facturas: <strong>{len(invoices)}</strong></p>
<table border="0" cellpadding="8" cellspacing="0"
       style="border-collapse:collapse;width:100%;font-size:13px;">
  <thead>
    <tr style="background:#1F4E79;color:#fff;">
      <th align="left">Empresa</th>
      <th align="left">N° Factura</th>
      <th align="left">Importe</th>
      <th align="left">Fecha Emisión</th>
      <th align="left">Fecha Envío</th>
      <th align="left">Ruta</th>
    </tr>
  </thead>
  <tbody>
    {"".join(
        f'<tr style="background:{"#f0f6ff" if idx%2==0 else "#fff"};">'
        + f'<td>{i.get("empresa","")}</td>'
        + f'<td>{i.get("numero_factura","")}</td>'
        + f'<td>{i.get("importe","")}</td>'
        + f'<td>{i.get("fecha_emision","")}</td>'
        + f'<td>{i.get("fecha_envio","")}</td>'
        + f'<td style="font-size:11px;color:#555;">{i.get("ruta","")}</td>'
        + '</tr>'
        for idx, i in enumerate(invoices)
    )}
  </tbody>
</table>
<p style="margin-top:18px;">Se adjunta el archivo <em>control_facturas.xlsx</em> con el detalle completo.</p>
<p style="color:#aaa;font-size:11px;margin-top:24px;">
  Mensaje generado automáticamente por el Sistema de Gestión de Facturas.
</p>
</body></html>"""

    @staticmethod
    def _attach_file(msg: MIMEMultipart, file_path: str) -> None:
        path = Path(file_path)
        if not path.exists():
            return
        with open(path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f'attachment; filename="{path.name}"',
        )
        msg.attach(part)
