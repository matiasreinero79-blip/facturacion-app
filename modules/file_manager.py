import shutil
from pathlib import Path
from datetime import datetime


class FileManager:
    """Organise invoice PDFs into Facturas/YYYY-MM sub-folders."""

    def __init__(self, base_dir: str = "Facturas"):
        self.base_dir = Path(base_dir)

    def ensure_base_dir(self) -> None:
        self.base_dir.mkdir(parents=True, exist_ok=True)

    def month_folder(self, date_str: str = "") -> Path:
        """Return (and create) the YYYY-MM folder for the given date string."""
        year, month = self._parse_ym(date_str)
        folder = self.base_dir / f"{year:04d}-{month:02d}"
        folder.mkdir(parents=True, exist_ok=True)
        return folder

    def copy_invoice(self, source_path: str, date_str: str = "") -> Path:
        """
        Copy *source_path* into the appropriate month folder.
        Returns the destination Path (may differ from source name if a conflict
        was resolved by appending a counter).
        """
        self.ensure_base_dir()
        src = Path(source_path)
        dest_dir = self.month_folder(date_str)
        dest = dest_dir / src.name

        # Resolve name collision (different file, same name)
        if dest.exists() and dest.resolve() != src.resolve():
            counter = 1
            while dest.exists():
                dest = dest_dir / f"{src.stem}_{counter}{src.suffix}"
                counter += 1

        # Don't copy if source IS destination
        if src.resolve() != dest.resolve():
            shutil.copy2(str(src), str(dest))

        return dest

    # ─── helpers ────────────────────────────────────────────────────────────

    def _parse_ym(self, date_str: str):
        """Return (year, month) from a date string, defaulting to today."""
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
            try:
                dt = datetime.strptime(date_str.strip(), fmt)
                return dt.year, dt.month
            except (ValueError, AttributeError):
                continue
        now = datetime.now()
        return now.year, now.month
