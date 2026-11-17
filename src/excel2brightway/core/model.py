"""
Datenstrukturen, die zwischen den Kernmodulen ausgetauscht werden.
"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

# Schweregrade einer Meldung. FEHLER blockiert das Schreiben (Format-Fehler
# nur, solange der Format-Test nicht ueberbrueckt wird).
FEHLER = "FEHLER"
WARNUNG = "WARNUNG"
HINWEIS = "HINWEIS"

# Namen der Tests, einzeln oder zusammen ausloesbar.
TEST_FORMAT = "format"
TEST_FLUSS = "fluss"
TEST_VERWEIS = "verweis"
ALL_TESTS = (TEST_FORMAT, TEST_FLUSS, TEST_VERWEIS)
TEST_LABELS = {TEST_FORMAT: "Format", TEST_FLUSS: "Fluss", TEST_VERWEIS: "Verweis"}


@dataclass(frozen=True)
class Record:
    """Ein moeglicher Verknuepfungs-Ziel-Datensatz (Aktivitaet oder biosphere-Fluss)."""
    database: str
    code: str
    name: str
    product: Optional[str] = None
    location: Optional[str] = None
    unit: Optional[str] = None
    categories: tuple = ()
    kind: str = "technosphere"  # 'technosphere' | 'biosphere'
    origin: str = "brightway"  # 'brightway' | 'datei'

    @property
    def key(self) -> tuple:
        return (self.database, self.code)


@dataclass
class Candidate:
    """Ein Vorschlag fuer einen nicht eindeutig aufgeloesten Fluss."""
    record: Record
    score: float
    reason: str = ""


@dataclass
class ExchangeRow:
    """Eine exchange-Zeile in einer Excel-Datei, mit Zellposition."""
    file: Path
    sheet: str
    row: int  # 1-basiert
    columns: dict  # Spaltenname (lowercase) -> Spaltennummer (1-basiert)
    values: dict  # Spaltenname (lowercase) -> bereinigter Zellwert
    activity_index: int  # Index der Aktivitaet in WorkbookModel.activities

    @property
    def type(self):
        value = self.values.get("type")
        return str(value).strip().lower() if value is not None else None

    def get(self, key, default=None):
        value = self.values.get(key)
        return default if value is None else value


@dataclass
class ActivityBlock:
    """Eine Aktivitaet ('Activity'-Block) in einer Excel-Datei."""
    file: Path
    sheet: str
    row: int  # Zeile der 'Activity'-Markierung
    name: Optional[str]
    database: Optional[str]
    fields: dict = field(default_factory=dict)  # Metadaten (lowercase key -> Wert)
    field_rows: dict = field(default_factory=dict)  # Metadaten-Key -> Zeile
    parameters: list = field(default_factory=list)  # Liste von Parameter-Dicts
    parameter_group: Optional[str] = None
    exchange_header_row: Optional[int] = None
    exchange_columns: dict = field(default_factory=dict)
    exchanges: list = field(default_factory=list)  # Liste von ExchangeRow

    @property
    def code(self):
        code = self.fields.get("code")
        return str(code) if code is not None else self.name

    @property
    def unit(self):
        return self.fields.get("unit")

    @property
    def location(self):
        return self.fields.get("location")

    @property
    def production_rows(self) -> list:
        return [e for e in self.exchanges if e.type == "production"]

    @property
    def reference_product(self):
        product = self.fields.get("reference product")
        if product:
            return product
        prod = self.production_rows
        if len(prod) == 1:
            return prod[0].get("reference product") or prod[0].get("name") or self.name
        return self.name


@dataclass
class WorkbookModel:
    """Eine eingelesene Excel-Datei im Brightway-Excel-Format."""
    path: Path
    database: Optional[str] = None
    database_location: Optional[tuple] = None  # (sheet, row)
    database_metadata: dict = field(default_factory=dict)
    project_parameters: list = field(default_factory=list)
    database_parameters: list = field(default_factory=list)
    activities: list = field(default_factory=list)  # Liste von ActivityBlock
    issues: list = field(default_factory=list)  # Format-Meldungen beim Einlesen
    cutoffs: dict = field(default_factory=dict)  # Blatt -> cutoff
    readable: bool = True

    def exchange_rows(self):
        for activity in self.activities:
            yield from activity.exchanges


@dataclass
class Issue:
    """Eine Meldung eines Tests (Format, Fluss oder Verweis)."""
    test: str
    severity: str
    kind: str
    message: str
    file: Optional[Path] = None
    sheet: Optional[str] = None
    row: Optional[int] = None
    column: Optional[int] = None
    activity: Optional[str] = None
    exchange: Optional[ExchangeRow] = None
    candidates: list = field(default_factory=list)  # Liste von Candidate

    @property
    def blocking(self) -> bool:
        return self.severity == FEHLER

    def location_label(self) -> str:
        parts = []
        if self.file is not None:
            parts.append(Path(self.file).name)
        if self.sheet:
            parts.append(self.sheet)
        if self.row:
            parts.append(f"Zeile {self.row}")
        return " / ".join(parts)
