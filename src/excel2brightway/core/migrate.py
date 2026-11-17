"""
migrate.py
==========
Datenbank/Version umstellen (Grundform): Alle Fluesse, die auf eine
Quell-Datenbank zeigen (z.B. ecoinvent-3.10-cutoff), werden auf eine
Ziel-Datenbank (z.B. ecoinvent-3.12-cutoff) umgestellt.

Voraussetzung: Die Pruefung gegen die Quelle ist komplett ohne Fehler - jeder
Fluss ist eindeutig einem Datensatz der Quelle zugeordnet.

Je Fluss:
  * Im Ziel wird exakt nach demselben Datensatz gesucht (reference product,
    name, location bzw. bei biosphere name und categories des aufgeloesten
    Quell-Datensatzes, nicht des Excel-Textes).
  * eindeutiger Treffer -> 'unveraendert' (nur database) oder 'umbenannt'
  * sonst -> 'offen': database wird auf das Ziel gesetzt, der Fluss landet
    beim anschliessenden Pruefen im Pruefblatt (mit Kandidaten aus dem Ziel).

Ausgabe: neue Dateien in einem Unterordner '<ziel>/' neben den Originalen
(die Originale bleiben unveraendert; ein eigener Ordner verhindert, dass alte
und neue Fassung zusammen eingelesen werden), je mit Blatt
'_ex2bw_Migration' als Protokoll. Danach laeuft die Pruefung gegen das Ziel.

Geplant: Migrationspakete (randonneur/randonneur_data) vor dem exakten
Abgleich anwenden - dokumentierte Umbenennungen, Aufteilungen und
Einheitenwechsel zwischen ecoinvent-Versionen.
"""

import re
import shutil
from dataclasses import dataclass, field, replace
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

from .apply import _ensure_column
from .checks import CheckOptions, run_checks
from .model import ALL_TESTS
from .review_sheet import save_workbook
from .text import normalize
from .workbook import read_workbook

MIGRATION_SHEET = "_ex2bw_Migration"
_UNSAFE = re.compile(r"[^A-Za-z0-9._-]+")


class MigrationBlockedError(RuntimeError):
    pass


@dataclass
class MigrationResult:
    files: list = field(default_factory=list)  # neue Dateien
    unchanged: int = 0
    renamed: int = 0
    open: int = 0
    log: list = field(default_factory=list)  # (datei, blatt, zeile, ergebnis, alt, neu)
    report: object = None  # CheckReport gegen das Ziel

    def summary(self) -> str:
        return (
            f"{len(self.files)} Datei(en) umgestellt: {self.unchanged} unveraendert, {self.renamed} umbenannt, "
            f"{self.open} offen (siehe Pruefblatt)."
        )


def safe_name(text) -> str:
    return _UNSAFE.sub("_", str(text)).strip("_") or "ziel"


class _TargetIndex:
    def __init__(self, records):
        self.tech = {}
        self.tech_product_location = {}
        self.bio = {}
        for record in records:
            if record.kind == "biosphere":
                key = (normalize(record.name), tuple(c.lower() for c in record.categories))
                self.bio.setdefault(key, []).append(record)
            else:
                key = (normalize(record.product or record.name), normalize(record.name), record.location)
                self.tech.setdefault(key, []).append(record)
                pl = (normalize(record.product or record.name), record.location)
                self.tech_product_location.setdefault(pl, []).append(record)

    def find(self, source):
        if source.kind == "biosphere":
            found = self.bio.get((normalize(source.name), tuple(c.lower() for c in source.categories)), [])
        else:
            found = self.tech.get((normalize(source.product or source.name), normalize(source.name), source.location), [])
            if not found:
                found = self.tech_product_location.get((normalize(source.product or source.name), source.location), [])
        return found[0] if len(found) == 1 else None


def migrate_files(paths, source_backend, source_database, target_backend, target_database,
                  source_biosphere=None, target_biosphere=None, options: CheckOptions = None,
                  output_folder=None) -> MigrationResult:
    """
    Stellt die ausgewaehlten Dateien von `source_database` auf `target_database`
    um (optional zusaetzlich biosphere: `source_biosphere` -> `target_biosphere`).
    """
    options = options or CheckOptions()
    check = run_checks(paths, backend=source_backend, options=CheckOptions(
        tests=ALL_TESTS, bypass_format=options.bypass_format, write_review=True,
        resolve_config=options.resolve_config,
    ))
    blocking = [i for i in check.blocking_issues() if i.test != "verweis"]
    if blocking or check.errors:
        raise MigrationBlockedError(
            "Umstellen geht erst, wenn alle Fluesse korrekt sind: "
            f"{len(blocking)} Fehler offen (siehe Pruefblatt)."
        )
    if target_database not in target_backend.databases():
        raise MigrationBlockedError(f"Ziel-Datenbank '{target_database}' gibt es im Zielprojekt nicht.")

    mapping = {source_database: target_database}
    if source_biosphere and target_biosphere:
        mapping[source_biosphere] = target_biosphere
    indexes = {target: _TargetIndex(target_backend.records(target)) for target in mapping.values()}

    result = MigrationResult()
    for source_path in check.selected:
        model = check.models.get(source_path)
        if model is None or not model.readable:
            continue
        folder = Path(output_folder) if output_folder else source_path.parent / safe_name(target_database)
        folder.mkdir(parents=True, exist_ok=True)
        target_path = folder / source_path.name
        shutil.copy2(source_path, target_path)

        target_model = read_workbook(target_path)
        positions = {(e.sheet, e.row): e for e in target_model.exchange_rows()}
        wb = openpyxl.load_workbook(target_path)
        log_rows = []
        for (path, sheet, row), resolution in check.resolutions.items():
            if path != source_path:
                continue
            record = resolution.record
            if record.database not in mapping:
                continue
            target_db = mapping[record.database]
            exchange = positions.get((sheet, row))
            if exchange is None:
                continue
            activity = target_model.activities[exchange.activity_index]
            ws = wb[sheet]
            cutoff = target_model.cutoffs.get(sheet)
            found = indexes[target_db].find(record)
            old = f"{record.database} | {record.name} | {record.location or '-'}"

            new_values = {"database": target_db}
            renamed = found is not None and (
                normalize(found.name) != normalize(record.name)
                or normalize(found.product) != normalize(record.product)
                or (found.location or None) != (record.location or None)
            )
            if renamed:
                # Excel-Text nur anfassen, wenn sich der Datensatz im Ziel wirklich unterscheidet
                new_values["name"] = found.name
                if record.kind != "biosphere":
                    new_values["location"] = found.location
                    if "reference product" in activity.exchange_columns or (found.product and found.product != found.name):
                        new_values["reference product"] = found.product or found.name
            for column_name, value in new_values.items():
                if exchange.values.get(column_name) == value:
                    continue
                col = _ensure_column(ws, activity, column_name, cutoff)
                if col is None:
                    continue
                ws.cell(row=row, column=col, value=value)

            if found is None:
                outcome, new = "offen", f"{target_db} | (im Pruefblatt waehlen)"
                result.open += 1
            elif renamed:
                outcome, new = "umbenannt", f"{found.database} | {found.name} | {found.location or '-'}"
                result.renamed += 1
            else:
                outcome, new = "unveraendert", f"{found.database} | {found.name} | {found.location or '-'}"
                result.unchanged += 1
            log_rows.append([sheet, row, activity.name, outcome, old, new])
            result.log.append((target_path.name, sheet, row, outcome, old, new))

        if MIGRATION_SHEET in wb.sheetnames:
            del wb[MIGRATION_SHEET]
        ws_log = wb.create_sheet(MIGRATION_SHEET)
        ws_log.cell(row=1, column=1, value="skip")
        ws_log.cell(row=1, column=2, value=(
            f"ex2bw-Umstellung {source_database} -> {target_database}"
            + (f", {source_biosphere} -> {target_biosphere}" if len(mapping) > 1 else "")
        )).font = Font(bold=True)
        for col, name in enumerate(["Blatt", "Zeile", "Aktivitaet", "Ergebnis", "alt", "neu"], start=1):
            ws_log.cell(row=2, column=col, value=name).font = Font(bold=True)
        for values in log_rows:
            ws_log.append(values)
        save_workbook(wb, target_path)
        wb.close()
        result.files.append(target_path)

    if result.files:
        link = [target_database] + ([target_biosphere] if target_biosphere else [])
        result.report = run_checks(result.files, backend=target_backend, options=CheckOptions(
            tests=ALL_TESTS, bypass_format=options.bypass_format, write_review=True,
            resolve_config=replace(options.resolve_config, link_databases=link),
        ))
    return result

