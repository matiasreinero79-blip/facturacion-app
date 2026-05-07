import sys
import os
from pathlib import Path


def get_base_dir() -> Path:
    """Return the directory where the app (or exe) lives."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def main():
    base = get_base_dir()
    os.chdir(base)

    from app import FacturacionApp
    app = FacturacionApp()
    app.mainloop()


if __name__ == "__main__":
    main()
