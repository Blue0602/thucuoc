from __future__ import annotations
from pathlib import Path
from typing import Any, Dict
import yaml


def load_config(config_path: str | Path = "config/settings.yaml") -> Dict[str, Any]:
    path = Path(config_path)
    if not path.exists():
        raise FileNotFoundError(f"Không tìm thấy file cấu hình: {path}")
    with path.open("r", encoding="utf-8") as f:
        return yaml.safe_load(f)
