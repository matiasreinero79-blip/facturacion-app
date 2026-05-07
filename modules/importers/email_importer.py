"""
Future importer: download invoice PDFs from an IMAP mailbox.

Implementation outline:
  1. Connect with imaplib using stored credentials.
  2. Search for unread emails with PDF attachments (or a specific subject/sender filter).
  3. Download each attachment to a temp directory.
  4. Mark messages as read / move to a processed folder.
  5. Return the list of local paths.
"""

from pathlib import Path
from typing import List

from .base_importer import BaseImporter


class EmailImporter(BaseImporter):

    def __init__(self, email_config: dict = None):
        self.email_config = email_config or {}

    def get_name(self) -> str:
        return "Email IMAP (próximamente)"

    def get_invoices(self) -> List[Path]:
        raise NotImplementedError("Email import is not yet implemented.")

    def is_available(self) -> bool:
        return False
