from __future__ import annotations

import sys
from pathlib import Path

APP_DIR_NAME = "TransareApp"


def bundle_root() -> Path:
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass).resolve()
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_path(*parts: str) -> Path:
    return bundle_root().joinpath(*parts)


def default_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        documents_dir = Path.home() / "Documents"
        root = documents_dir if documents_dir.exists() else Path.home()
        app_dir = root / APP_DIR_NAME
        app_dir.mkdir(parents=True, exist_ok=True)
        return app_dir
    return Path.cwd().resolve()
