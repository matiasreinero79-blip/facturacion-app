"""Cross-platform helpers used across the application."""

import os
import subprocess
import sys
from pathlib import Path


def open_path(path: str) -> None:
    """
    Open *path* with the OS default application.

    - Windows : os.startfile
    - macOS   : open
    - Linux   : xdg-open
    """
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"File not found: {path}")

    if sys.platform == "win32":
        os.startfile(str(p))                            # Windows
    elif sys.platform == "darwin":
        subprocess.run(["open", str(p)], check=False)  # macOS
    else:
        subprocess.run(["xdg-open", str(p)], check=False)  # Linux / BSD


def system_font(size: int = 10, bold: bool = False) -> tuple:
    """
    Return a (family, size [, 'bold']) tuple that looks native on each OS.

    Windows  → Segoe UI
    macOS    → Helvetica Neue
    Linux    → DejaVu Sans
    """
    if sys.platform == "win32":
        family = "Segoe UI"
    elif sys.platform == "darwin":
        family = "Helvetica Neue"
    else:
        family = "DejaVu Sans"

    return (family, size, "bold") if bold else (family, size)


def mono_font(size: int = 9) -> tuple:
    """
    Monospace font tuple that exists on every major OS.

    Windows → Consolas
    macOS   → Menlo
    Linux   → DejaVu Sans Mono
    """
    if sys.platform == "win32":
        family = "Consolas"
    elif sys.platform == "darwin":
        family = "Menlo"
    else:
        family = "DejaVu Sans Mono"

    return (family, size)


def preferred_themes() -> tuple:
    """
    Return ttk theme names to try, most-preferred first.
    Falls back gracefully if a theme is not available.
    """
    if sys.platform == "darwin":
        return ("aqua", "clam", "alt", "default")
    if sys.platform == "win32":
        return ("vista", "winnative", "clam", "default")
    return ("clam", "alt", "default")
