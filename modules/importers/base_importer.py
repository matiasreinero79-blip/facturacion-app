from abc import ABC, abstractmethod
from pathlib import Path
from typing import List


class BaseImporter(ABC):
    """Contract every invoice source must implement."""

    @abstractmethod
    def get_name(self) -> str:
        """Human-readable name of this source."""

    @abstractmethod
    def get_invoices(self) -> List[Path]:
        """
        Collect PDFs from the source and return their local paths.
        Implementations may download files to a temp directory first.
        """

    @abstractmethod
    def is_available(self) -> bool:
        """Return True if the importer is configured and ready to use."""
