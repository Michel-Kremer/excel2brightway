"""
settings.py
===========
Zuletzt verwendete Einstellungen (Projekt, Verknuepfungs-Datenbanken, Pfade)
fuer das Fenster - als JSON im Benutzerordner:
    ~/.excel2brightway/settings.json   (Windows: %USERPROFILE%\\.excel2brightway)
Fehlt oder ist die Datei kaputt, gelten die Standardwerte.
"""

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path

SETTINGS_DIR = Path.home() / ".excel2brightway"
SETTINGS_FILE = SETTINGS_DIR / "settings.json"


@dataclass
class Settings:
    project: str = ""
    link_databases: list = field(default_factory=list)
    location_priority: list = field(default_factory=lambda: ["GLO", "RER", "RoW"])
    last_paths: list = field(default_factory=list)
    tests: list = field(default_factory=lambda: ["format", "fluss", "verweis"])
    bypass_format: bool = False

    @classmethod
    def load(cls, path=SETTINGS_FILE):
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return cls()
        known = {k: v for k, v in data.items() if k in cls.__dataclass_fields__}
        try:
            return cls(**known)
        except TypeError:
            return cls()

    def save(self, path=SETTINGS_FILE):
        path = Path(path)
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(json.dumps(asdict(self), indent=2, ensure_ascii=False), encoding="utf-8")
        except OSError:
            pass  # Einstellungen sind Komfort, kein Muss
