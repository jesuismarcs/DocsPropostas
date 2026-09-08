"""Preferências locais, nunca incluídas no repositório."""

import json
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parents[2] / ".env")

CONFIG_FILE = Path(
    os.getenv(
        "DOCSPROPOSTAS_CONFIG",
        str(Path(__file__).resolve().parents[2] / "local" / "config.json"),
    )
)


def load_config():
    defaults = {"theme": "light", "output_path": "", "template_path": ""}
    try:
        data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            defaults.update(data)
    except (OSError, ValueError):
        pass
    return defaults


def save_config(config):
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_FILE.write_text(
        json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8"
    )
