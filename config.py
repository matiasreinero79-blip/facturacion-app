import json
import os
from pathlib import Path

CONFIG_FILE = "config.json"

DEFAULT_CONFIG = {
    "facturas_base_dir": "Facturas",
    "excel_filename": "control_facturas.xlsx",
    "email": {
        "remitente": "",
        "finanzas": "",
        "password": "",
        "smtp_server": "smtp.gmail.com",
        "smtp_port": 587,
        "use_tls": True,
    },
}


def load_config() -> dict:
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                stored = json.load(f)
            # Merge with defaults so new keys are always present
            result = DEFAULT_CONFIG.copy()
            result.update(stored)
            result["email"] = {**DEFAULT_CONFIG["email"], **stored.get("email", {})}
            return result
        except Exception:
            pass
    return {k: (v.copy() if isinstance(v, dict) else v) for k, v in DEFAULT_CONFIG.items()}


def save_config(config: dict) -> None:
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4, ensure_ascii=False)
