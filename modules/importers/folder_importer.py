from pathlib import Path
from typing import List

from .base_importer import BaseImporter


class FolderImporter(BaseImporter):
    """Import invoice PDFs from a local folder."""

    def __init__(self, folder_path: str = None):
        self.folder_path = Path(folder_path) if folder_path else None

    def get_name(self) -> str:
        return "Carpeta Local"

    def set_folder(self, folder_path: str) -> None:
        self.folder_path = Path(folder_path)

    def get_invoices(self) -> List[Path]:
        if not self.is_available():
            return []
        pdfs = list(self.folder_path.glob("*.pdf"))
        pdfs += [p for p in self.folder_path.glob("*.PDF") if p not in pdfs]
        return sorted(pdfs, key=lambda p: p.name.lower())

    def is_available(self) -> bool:
        return self.folder_path is not None and self.folder_path.is_dir()
