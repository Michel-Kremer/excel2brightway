"""
resolve.py
==========
Fluss-Test: ordnet jeder technosphere-/biosphere-Zeile genau einen Datensatz
(database, code) zu - oder meldet, warum das nicht eindeutig geht, zusammen
mit Kandidaten fuer das Pruefblatt.

Wo gesucht wird (Spalte 'database' der exchange-Zeile hat Vorrang):
  * 'database' nennt eine Datenbank, die in den eingelesenen Dateien definiert
    ist -> nur diese Dateien. Eine gleichnamige Datenbank im Projekt wird
    beim Schreiben durch die Dateien ersetzt; Aktivitaeten, die nur noch dort
    stehen, werden geloescht und sind daher keine Treffer.
  * 'database' nennt eine andere Datenbank -> nur dort. Das Label wird
    normalisiert verglichen ('ecoinvent 3.12' ~ 'ecoinvent-3.12-cutoff').
  * 'database' leer -> technosphere: erst alle Dateien, dann die
    Verknuepfungs-Datenbanken (ResolveConfig.link_databases, sonst alle
    Datenbanken des Projekts). biosphere: die biosphere-Datenbanken.

Eindeutigkeit (technosphere):
  * Treffer ueber 'reference product' (und 'name', falls beide angegeben),
    sonst ueber 'name' als Produkt bzw. Aktivitaetsname.
  * 'location' angegeben -> muss exakt passen.
  * Ohne 'location' und mehrere Standorte -> Rangfolge
    ResolveConfig.location_priority (Standard GLO, RER, RoW). Greift keine,
    ist der Fluss mehrdeutig - es wird NICHT alphabetisch geraten.
  * Treffer in mehreren Datenbanken sind immer mehrdeutig.
  * Bleiben mehrere Datensaetze am selben Standort, wird ein eindeutiger
    Markt ('market for ...') genommen, sonst mehrdeutig.
  * Die Einheit muss nach Normalisierung passen.
"""

from dataclasses import dataclass, field
from typing import Optional

from .backend import is_biosphere_like
from .model import FEHLER, TEST_FLUSS, Candidate, Issue, Record
from .search import MARKET_PREFIXES, SearchIndex, merge_candidates, score_record
from .text import normalize, normalize_db_name, split_categories, tokens, units_compatible

STATUS_OK = "ok"

STATUS_LABELS = {
    "nicht_gefunden": "nicht gefunden",
    "mehrdeutig": "mehrdeutig",
    "standort_nicht_gefunden": "Standort nicht gefunden",
    "kategorie_nicht_gefunden": "Kategorie nicht gefunden",
    "einheit_abweichend": "Einheit passt nicht",
    "datenbank_unbekannt": "Datenbank unbekannt",
    "kein_projekt": "kein Projekt gewaehlt",
    # Fassungen derselben Datei (versions.py)
    "andere_version": "andere Version",
    "variante_ohne_version": "Variante ohne Versionsnummer",
    "version_unklar": "Fassung unklar",
}


@dataclass
class ResolveConfig:
    link_databases: Optional[list] = None  # None/leer = alle Datenbanken des Projekts
    location_priority: tuple = ("GLO", "RER", "RoW")
    candidate_limit: int = 5
    min_score: float = 0.2  # schwaechere Suchtreffer werden nicht vorgeschlagen


@dataclass
class Resolution:
    exchange: object  # ExchangeRow
    status: str
    record: Optional[Record] = None
    candidates: list = field(default_factory=list)
    message: str = ""

    @property
    def ok(self) -> bool:
        return self.status == STATUS_OK


class _DbIndex:
    def __init__(self, records):
        self.records = records
        self.by_product = {}
        self.by_name = {}
        for record in records:
            self.by_product.setdefault(normalize(record.product or record.name), []).append(record)
            self.by_name.setdefault(normalize(record.name), []).append(record)
        self._search = None

    @property
    def search(self) -> SearchIndex:
        if self._search is None:
            self._search = SearchIndex(self.records)
        return self._search


def _split_biosphere_name(name):
    """'Methan (air)' -> ('Methan', 'air'); ohne Suffix -> (name, None)."""
    text = str(name).strip()
    if text.endswith(")") and "(" in text:
        base, _, suffix = text[:-1].rpartition("(")
        if base.strip() and suffix.strip() and "(" not in suffix:
            return base.strip(), suffix.strip().lower()
    return text, None


def file_records(models) -> dict:
    """Aktivitaeten der eingelesenen Dateien als Record je Datenbank."""
    result = {}
    for model in models:
        if not model.database:
            continue
        for activity in model.activities:
            if not activity.name:
                continue
            result.setdefault(model.database, []).append(Record(
                database=model.database,
                code=str(activity.code),
                name=activity.name,
                product=activity.reference_product,
                location=activity.location,
                unit=activity.unit,
                kind="technosphere",
                origin="datei",
            ))
    return result


class Resolver:
    def __init__(self, models, backend=None, config=None):
        self.backend = backend
        self.config = config or ResolveConfig()
        self._file_records = file_records(models)
        self.file_databases = sorted(self._file_records)
        self.project_databases = backend.databases() if backend is not None else []
        self.all_databases = sorted(set(self.file_databases) | set(self.project_databases))
        self._indexes = {}
        self.warnings = []
        if self.config.link_databases and backend is not None:
            missing = [d for d in self.config.link_databases if d not in self.project_databases]
            if missing:
                self.warnings.append(
                    f"Verknuepfungs-Datenbank(en) nicht im Projekt '{backend.project}': {', '.join(missing)}"
                )

    # ------------------------------------------------------------------ Indizes
    def _index(self, database) -> _DbIndex:
        if database not in self._indexes:
            if database in self._file_records:
                # Die Dateien ersetzen die Datenbank beim Schreiben vollstaendig; was nur noch in
                # Brightway steht, wird geloescht und darf kein Treffer sein (sonst "mehrdeutig"
                # nach code-Aenderung bzw. Verweis ins Leere nach dem Schreiben).
                records = list(self._file_records[database])
            elif self.backend is not None and database in self.project_databases:
                records = list(self.backend.records(database))
            else:
                records = []
            self._indexes[database] = _DbIndex(records)
        return self._indexes[database]

    def _allowed(self, kind):
        base = [d for d in (self.config.link_databases or self.project_databases) if d in self.project_databases]
        if kind == "biosphere":
            bio = [d for d in base if is_biosphere_like(d)]
            if not bio:
                bio = [d for d in self.project_databases if is_biosphere_like(d)]
            return bio
        return [d for d in base if not is_biosphere_like(d)]

    # ------------------------------------------------------- Datenbank-Labels
    def target_databases(self, label, kind):
        """
        Datenbanken, in denen ein Fluss gesucht wird. Gibt (stufen, problem)
        zurueck; `stufen` ist eine Liste von Datenbank-Listen, die nacheinander
        versucht werden (Dateien vor Projekt), `problem` None, 'datenbank_unbekannt'
        oder 'kein_projekt'.
        """
        if label:
            lowered = {d.lower(): d for d in self.all_databases}
            if str(label).lower() in lowered:
                return [[lowered[str(label).lower()]]], None
            pool = self.all_databases
            if kind == "biosphere":
                pool = [d for d in pool if is_biosphere_like(d)] or pool
            else:
                pool = [d for d in pool if not is_biosphere_like(d)] or pool
            norm_label = normalize_db_name(label)
            matches = [d for d in pool if norm_label and norm_label in normalize_db_name(d)]
            if matches:
                files = [d for d in matches if d in self._file_records]
                others = [d for d in matches if d not in self._file_records]
                return [s for s in (files, others) if s], None
            if kind == "biosphere":
                allowed = self._allowed("biosphere")
                if allowed:
                    # Label wie 'biosphere3' bei einem Projekt mit 'ecoinvent-3.12-biosphere'
                    return [allowed], None
            if self.backend is None:
                return [], "kein_projekt"
            return [], "datenbank_unbekannt"

        if kind == "biosphere":
            allowed = self._allowed("biosphere")
            if not allowed:
                return [], ("kein_projekt" if self.backend is None else "nicht_gefunden")
            return [allowed], None

        stages = []
        if self.file_databases:
            stages.append(list(self.file_databases))
        allowed = [d for d in self._allowed("technosphere") if d not in self._file_records]
        if allowed:
            stages.append(allowed)
        if not stages:
            return [], ("kein_projekt" if self.backend is None else "nicht_gefunden")
        return stages, None

    # ------------------------------------------------------- exakter Abgleich
    def _collect(self, databases, attribute, key):
        result = []
        for database in databases:
            result.extend(getattr(self._index(database), attribute).get(key, []))
        return result

    def _exact_technosphere(self, databases, product, name, location, unit):
        candidates = []
        if product:
            candidates = self._collect(databases, "by_product", normalize(product))
            if candidates and name:
                narrowed = [c for c in candidates if normalize(c.name) == normalize(name)]
                if narrowed:
                    candidates = narrowed
        if not candidates and name:
            candidates = self._collect(databases, "by_product", normalize(name))
            if not candidates:
                candidates = self._collect(databases, "by_name", normalize(name))
        candidates = [c for c in candidates if c.kind == "technosphere"]
        if not candidates:
            return "nicht_gefunden", None, []

        if location:
            same = [c for c in candidates if (c.location or "").lower() == str(location).lower()]
            if not same:
                return "standort_nicht_gefunden", None, candidates
            candidates = same
        elif len({c.location for c in candidates}) > 1:
            for preferred in self.config.location_priority:
                picked = [c for c in candidates if c.location == preferred]
                if picked:
                    candidates = picked
                    break
            else:
                return "mehrdeutig", None, candidates

        if len({c.database for c in candidates}) > 1:
            # derselbe Fluss in mehreren Datenbanken: nie still eine waehlen
            return "mehrdeutig", None, candidates
        if len(candidates) > 1:
            markets = [c for c in candidates if normalize(c.name).startswith(MARKET_PREFIXES)]
            if len(markets) == 1:
                candidates = markets
            else:
                return "mehrdeutig", None, candidates

        record = candidates[0]
        if not units_compatible(unit, record.unit):
            return "einheit_abweichend", record, [record]
        return "ok", record, []

    def _exact_biosphere(self, databases, name, categories, unit):
        base, hint = _split_biosphere_name(name)
        candidates = [c for c in self._collect(databases, "by_name", normalize(base)) if c.kind == "biosphere"]
        if not candidates:
            return "nicht_gefunden", None, []

        if categories:
            wanted = tuple(c.lower() for c in categories)
            exact = [c for c in candidates if tuple(x.lower() for x in c.categories) == wanted]
            if exact:
                candidates = exact
            elif len(wanted) > 1:
                return "kategorie_nicht_gefunden", None, candidates
            else:
                top = [c for c in candidates if c.categories and c.categories[0].lower() == wanted[0]]
                if not top:
                    return "kategorie_nicht_gefunden", None, candidates
                candidates = top
        elif hint:
            top = [c for c in candidates if c.categories and c.categories[0].lower() == hint]
            if top:
                candidates = top

        if len(candidates) > 1:
            shortest = min(len(c.categories) for c in candidates)
            remaining = [c for c in candidates if len(c.categories) == shortest]
            if len(remaining) == 1:
                candidates = remaining
            else:
                return "mehrdeutig", None, candidates

        record = candidates[0]
        if not units_compatible(unit, record.unit):
            return "einheit_abweichend", record, [record]
        return "ok", record, []

    # ------------------------------------------------------------- Kandidaten
    def candidates_for(self, exchange, databases, pool=(), search_text=None):
        kind = exchange.type
        product = exchange.get("reference product")
        name = exchange.get("name")
        unit = exchange.get("unit")
        location = exchange.get("location")
        categories = split_categories(exchange.values.get("categories"))
        query = search_text or product or name
        if kind == "biosphere" and not search_text:
            query = _split_biosphere_name(name or "")[0]
        limit = self.config.candidate_limit
        query_tokens = tokens(query)

        lists = []
        if pool:
            scored = []
            for record in pool:
                score, reason = score_record(
                    record, query_tokens, unit=unit, location=location, categories=categories,
                    location_priority=self.config.location_priority,
                )
                # exakte Namenstreffer vor reinen Suchtreffern einordnen
                scored.append(Candidate(record=record, score=score + 0.2, reason=reason))
            lists.append(scored)
        for database in databases:
            found = self._index(database).search.search(
                query, unit=unit, location=location, categories=categories, kind=kind,
                limit=limit, location_priority=self.config.location_priority,
            )
            # exakte Namenstreffer (pool) immer zeigen, reine Suchtreffer erst ab min_score
            lists.append([c for c in found if c.score >= self.config.min_score])
        return merge_candidates(lists, limit=limit)

    # ---------------------------------------------------------------- Aufloesen
    def resolve(self, exchange, search_text=None) -> Resolution:
        kind = exchange.type
        label = exchange.get("database")
        flow = exchange.get("reference product") or exchange.get("name") or "?"

        stages, problem = self.target_databases(label, kind)
        all_targets = [d for stage in stages for d in stage]

        if problem == "kein_projekt":
            candidates = []
            if self.file_databases and kind == "technosphere":
                candidates = self.candidates_for(exchange, self.file_databases, search_text=search_text)
            return Resolution(
                exchange, "kein_projekt", candidates=candidates,
                message=f"'{flow}' ist nicht in den Dateien definiert und es ist kein Brightway-Projekt "
                        f"gewaehlt - externe Fluesse koennen so nicht geprueft werden.",
            )
        if problem == "datenbank_unbekannt":
            search_dbs = (self.file_databases if kind == "technosphere" else []) + [
                d for d in self._allowed(kind) if d not in self._file_records
            ]
            return Resolution(
                exchange, "datenbank_unbekannt",
                candidates=self.candidates_for(exchange, search_dbs, search_text=search_text),
                message=f"Datenbank '{label}' gibt es weder in den Dateien noch im Projekt "
                        f"'{self.backend.project}'.",
            )

        status, record, pool = "nicht_gefunden", None, []
        for stage in stages:
            if kind == "biosphere":
                status, record, pool = self._exact_biosphere(
                    stage, exchange.get("name") or "", split_categories(exchange.values.get("categories")),
                    exchange.get("unit"),
                )
            else:
                status, record, pool = self._exact_technosphere(
                    stage, exchange.get("reference product"), exchange.get("name"),
                    exchange.get("location"), exchange.get("unit"),
                )
            if status != "nicht_gefunden":
                break

        if status == STATUS_OK and not search_text:
            return Resolution(exchange, STATUS_OK, record=record)

        if status == "nicht_gefunden" and self.backend is None and not label:
            return Resolution(
                exchange, "kein_projekt",
                candidates=self.candidates_for(exchange, all_targets, search_text=search_text),
                message=f"'{flow}' ist nicht in den Dateien definiert und es ist kein Brightway-Projekt "
                        f"gewaehlt - externe Fluesse koennen so nicht geprueft werden.",
            )

        if status == STATUS_OK and search_text:
            # Eigene Suche bei eigentlich aufgeloestem Fluss: nur Kandidaten liefern
            status = "mehrdeutig"
            pool = [record]

        candidates = self.candidates_for(exchange, all_targets, pool=pool, search_text=search_text)
        return Resolution(
            exchange, status, record=record, candidates=candidates,
            message=self._message(status, flow, exchange, record, pool, all_targets),
        )

    def _message(self, status, flow, exchange, record, pool, databases):
        where = ", ".join(databases) if databases else "keiner Datenbank"
        location = exchange.get("location")
        if status == "nicht_gefunden":
            return f"'{flow}' wurde in {where} nicht gefunden."
        if status == "mehrdeutig":
            locations = sorted({c.location or "-" for c in pool}) if pool else []
            detail = f" (Standorte: {', '.join(locations)})" if locations else ""
            return f"'{flow}' passt auf {len(pool)} Datensaetze{detail}; bitte auswaehlen."
        if status == "standort_nicht_gefunden":
            locations = sorted({c.location or "-" for c in pool})
            return f"'{flow}' gibt es nicht fuer location '{location}' (vorhanden: {', '.join(locations)})."
        if status == "kategorie_nicht_gefunden":
            cats = sorted({"::".join(c.categories) for c in pool})
            return (
                f"'{flow}' gibt es nicht in Kategorie '{exchange.values.get('categories')}' "
                f"(vorhanden: {', '.join(cats[:6])}{' ...' if len(cats) > 6 else ''})."
            )
        if status == "einheit_abweichend":
            return (
                f"Einheit '{exchange.get('unit')}' passt nicht zu '{record.unit}' von "
                f"'{record.name}' ({record.database})."
            )
        return f"'{flow}': {status}"


def resolution_issue(resolution: Resolution) -> Issue:
    """Wandelt eine nicht aufgeloeste Resolution in eine Meldung des Fluss-Tests."""
    exchange = resolution.exchange
    return Issue(
        test=TEST_FLUSS, severity=FEHLER, kind=resolution.status, message=resolution.message,
        file=exchange.file, sheet=exchange.sheet, row=exchange.row,
        column=exchange.columns.get("name") or exchange.columns.get("reference product"),
        exchange=exchange, candidates=list(resolution.candidates),
    )
