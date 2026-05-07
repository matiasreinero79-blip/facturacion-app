"""
Future importer: download invoice PDFs from a supplier web portal.

Implementation outline:
  1. Use requests + BeautifulSoup (or Selenium for JS-heavy portals).
  2. Authenticate with stored credentials.
  3. Navigate to the invoice list page.
  4. Download each PDF to a temp directory.
  5. Return the list of local paths.
"""

from pathlib import Path
from typing import List

from .base_importer import BaseImporter


class WebPortalImporter(BaseImporter):

    def __init__(self, portal_config: dict = None):
        self.portal_config = portal_config or {}

    def get_name(self) -> str:
        return "Portal Web (próximamente)"

    def get_invoices(self) -> List[Path]:
        raise NotImplementedError("Web portal import is not yet implemented.")

    def is_available(self) -> bool:
        return False
