"""Excel-Reifecheck – Automatisierter Excel-Linter für saubere Tabellenkalkulationen."""

import os as _os
from datetime import datetime as _dt
from pathlib import Path as _Path
try:
    from dotenv import load_dotenv as _load_dotenv
except ImportError:  # noqa: BLE001 - dotenv ist ein Extra
    # python-dotenv gehört zum ``web``-Extra. Der Analysekern soll sich auch
    # dort importieren lassen, wo nur openpyxl verfügbar ist — im Browser
    # (Pyodide) gibt es weder .env-Dateien noch das Paket.
    def _load_dotenv(*_args, **_kwargs) -> bool:
        return False

_load_dotenv()

__version__ = _os.getenv("APP_VERSION", "0.0.0")

# Letzter Build-Zeitstempel: neuster mtime aller .py-Dateien im Package
_pkg_dir = _Path(__file__).parent
_newest = max(
    (f.stat().st_mtime for f in _pkg_dir.rglob("*.py")),
    default=0,
)
BUILD_TIMESTAMP = _dt.fromtimestamp(_newest).strftime("%d.%m.%Y %H:%M") if _newest else "unbekannt"
