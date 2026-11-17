"""
backend.py
==========
Zugriff auf die Datenbanken eines Brightway-Projekts - nur lesend fuer die
Tests. Die Kernlogik spricht ausschliesslich ueber diese Schnittstelle mit
Brightway; in Tests wird MemoryBackend verwendet.

Schnittstelle (duck typing):
  project                      Name des Projekts (oder None)
  databases() -> list[str]     Datenbanken im Projekt
  records(db) -> list[Record]  alle Aktivitaeten/Fluesse einer Datenbank
  inbound_references(db)       Verweise ANDERER Datenbanken auf `db`
                               -> list[InboundReference]
"""

from dataclasses import dataclass

from .model import Record
from .text import normalize_db_name

BIOSPHERE_NAME_HINT = "biosphere"
BIOSPHERE_NODE_TYPES = {"emission", "natural resource", "economic", "social", "inventory indicator"}


def is_biosphere_like(db_name) -> bool:
    """Datenbank sieht nach biosphere aus ('biosphere3', 'ecoinvent-3.12-biosphere', ...)."""
    return BIOSPHERE_NAME_HINT in normalize_db_name(db_name)


@dataclass(frozen=True)
class InboundReference:
    """Ein exchange in einer anderen Datenbank, der auf eine Aktivitaet in `target_database` zeigt."""
    source_database: str
    source_code: str
    source_name: str
    target_database: str
    target_code: str


class MemoryBackend:
    """Backend im Speicher (fuer Tests und Trockenlaeufe)."""

    def __init__(self, databases=None, references=None, project="test"):
        self.project = project
        self._records = {name: list(records) for name, records in (databases or {}).items()}
        self._references = list(references or [])

    def databases(self):
        return sorted(self._records)

    def records(self, database):
        return list(self._records.get(database, []))

    def inbound_references(self, database):
        return [r for r in self._references if r.target_database == database]


class BrightwayBackend:
    """
    Liest ein Brightway-Projekt ueber bw2data (lazy importiert). Aktiviert das
    Projekt nur, wenn es existiert - bw2data wuerde ein unbekanntes Projekt
    sonst stillschweigend neu anlegen.
    """

    def __init__(self, project):
        import bw2data as bd

        existing = {p.name for p in bd.projects}
        if project not in existing:
            raise ValueError(
                f"Brightway-Projekt '{project}' existiert nicht. Vorhanden: {', '.join(sorted(existing))}"
            )
        self.project = project
        self._bd = bd
        self._records_cache = {}

    def _activate(self):
        if self._bd.projects.current != self.project:
            self._bd.projects.set_current(self.project)

    def databases(self):
        self._activate()
        return sorted(self._bd.databases)

    def records(self, database):
        if database in self._records_cache:
            return self._records_cache[database]
        self._activate()
        biosphere_db = is_biosphere_like(database)
        records = []
        for node in self._bd.Database(database):
            node_type = node.get("type")
            kind = "biosphere" if (node_type in BIOSPHERE_NODE_TYPES or (biosphere_db and node_type != "process")) else "technosphere"
            records.append(Record(
                database=database,
                code=node.get("code"),
                name=node.get("name"),
                product=node.get("reference product"),
                location=node.get("location"),
                unit=node.get("unit"),
                categories=tuple(node.get("categories") or ()),
                kind=kind,
                origin="brightway",
            ))
        self._records_cache[database] = records
        return records

    def inbound_references(self, database):
        self._activate()
        from bw2data.backends.schema import ActivityDataset, ExchangeDataset

        query = (
            ExchangeDataset
            .select(ExchangeDataset.output_database, ExchangeDataset.output_code, ExchangeDataset.input_code)
            .where((ExchangeDataset.input_database == database) & (ExchangeDataset.output_database != database))
        )
        rows = list(query.tuples())
        names = {}
        for source_db, source_code, _ in rows:
            key = (source_db, source_code)
            if key in names:
                continue
            try:
                names[key] = ActivityDataset.get(
                    (ActivityDataset.database == source_db) & (ActivityDataset.code == source_code)
                ).name
            except Exception:
                names[key] = source_code
        return [
            InboundReference(
                source_database=source_db, source_code=source_code,
                source_name=names.get((source_db, source_code), source_code),
                target_database=database, target_code=target_code,
            )
            for source_db, source_code, target_code in rows
        ]


def list_projects():
    """Namen aller Brightway-Projekte (leer, wenn bw2data fehlt)."""
    try:
        import bw2data as bd
    except ImportError:
        return []
    return sorted(p.name for p in bd.projects)
