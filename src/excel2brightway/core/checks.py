"""
checks.py
=========
Fuehrt die Tests aus - einzeln oder zusammen - und schreibt die Ergebnisse:

  format   Format-Test (format_check.py), ueberbrueckbar
  fluss    Fluss-Test (resolve.py): jede technosphere-/biosphere-Zeile muss
           genau einem Datensatz zugeordnet werden koennen
  verweis  Verweis-Test (references.py): was wuerde das Schreiben in Brightway
           loeschen, und zeigen andere Datenbanken darauf?

Auswahl: einzelne Dateien und/oder Ordner. Als Kontext (interne Verweise,
weitere Dateien derselben Datenbank) werden immer alle *.xlsx der
betroffenen Ordner gelesen; Meldungen gibt es nur fuer die ausgewaehlten
Dateien.

Ausgabe:
  * Pruefblatt in jeder ausgewaehlten Datei mit Meldungen (review_sheet.py)
  * Fehler-Log 'ex2bw.log' im (ersten) Ordner: alle Meldungen plus alles,
    was nicht in die Excel gehoert (gesperrte Dateien, Ausnahmen, ...)
"""

import logging
import traceback
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from .format_check import NOT_BYPASSABLE, check_format
from .model import ALL_TESTS, FEHLER, HINWEIS, TEST_FLUSS, TEST_FORMAT, TEST_LABELS, TEST_VERWEIS, WARNUNG, Issue
from .references import check_references
from .resolve import ResolveConfig, Resolver, resolution_issue
from .review_sheet import WorkbookLockedError, read_review, strip_review, write_review
from .versions import resolve_versions
from .workbook import list_excel_files, read_workbook

LOG_FILENAME = "ex2bw.log"
RESOLVABLE_TYPES = ("technosphere", "biosphere")

logger = logging.getLogger("excel2brightway")


@dataclass
class CheckOptions:
    tests: tuple = ALL_TESTS
    bypass_format: bool = False
    write_review: bool = True
    write_log: bool = True
    resolve_config: ResolveConfig = field(default_factory=ResolveConfig)


@dataclass
class CheckReport:
    selected: list
    models: dict  # Path -> WorkbookModel (alle gelesenen Dateien)
    issues: list
    resolutions: dict  # (Path, sheet, row) -> Resolution (nur aufgeloeste)
    options: CheckOptions
    errors: list = field(default_factory=list)  # sonstige Fehler (Text)
    log_path: Path = None
    ignored: dict = field(default_factory=dict)  # Pfad -> gezaehlte Fassung (versions.py)

    def counts(self) -> dict:
        result = {FEHLER: 0, WARNUNG: 0, HINWEIS: 0}
        for issue in self.issues:
            result[issue.severity] = result.get(issue.severity, 0) + 1
        return result

    def blocking_issues(self) -> list:
        return [
            i for i in self.issues
            if i.severity == FEHLER
            and not (self.options.bypass_format and i.test == TEST_FORMAT and i.kind not in NOT_BYPASSABLE)
        ]

    @property
    def ok(self) -> bool:
        return not self.blocking_issues() and not self.errors

    def issues_for(self, path) -> list:
        return [i for i in self.issues if i.file is not None and Path(i.file) == Path(path)]

    def summary(self) -> str:
        counts = self.counts()
        tests = ", ".join(TEST_LABELS[t] for t in self.options.tests)
        status = "bereit zum Schreiben" if self.ok else "nicht bereit"
        bypass = " (Format-Test ueberbrueckt)" if self.options.bypass_format else ""
        lines = [
            f"Tests: {tests}{bypass}",
            f"{len(self.selected)} Datei(en): {counts[FEHLER]} Fehler, {counts[WARNUNG]} Warnung(en), "
            f"{counts[HINWEIS]} Hinweis(e) -> {status}",
        ]
        if self.errors:
            lines.append(f"{len(self.errors)} sonstige(r) Fehler, siehe {self.log_path or LOG_FILENAME}")
        return "\n".join(lines)


def expand_selection(paths) -> list:
    """Dateien und Ordner -> Liste von Excel-Dateien (Ordner: oberste Ebene, ohne ~$-Dateien)."""
    files = []
    for path in paths:
        path = Path(path)
        if path.is_dir():
            files.extend(list_excel_files(path))
        else:
            files.append(path)
    unique = []
    for f in files:
        resolved = f.resolve()
        if resolved not in unique:
            unique.append(resolved)
    return unique


def context_files(selected) -> list:
    """Alle Excel-Dateien der Ordner der ausgewaehlten Dateien (inkl. der Auswahl)."""
    files = list(selected)
    for folder in sorted({p.parent for p in selected}):
        for f in list_excel_files(folder):
            f = f.resolve()
            if f not in files:
                files.append(f)
    return files


def _write_log(path, report, started):
    lines = [
        f"excel2brightway - Pruefung vom {started:%Y-%m-%d %H:%M:%S}",
        report.summary(),
        "",
        "Dateien:",
    ]
    lines.extend(f"  {p}" for p in report.selected)
    if report.errors:
        lines += ["", "Sonstige Fehler:"]
        lines.extend(f"  {e}" for e in report.errors)
    lines += ["", "Meldungen:"]
    for issue in report.issues:
        where = issue.location_label() or "-"
        lines.append(f"  [{TEST_LABELS.get(issue.test, issue.test)}] {issue.severity} {issue.kind} @ {where}: {issue.message}")
    Path(path).write_text("\n".join(lines) + "\n", encoding="utf-8")


def run_checks(paths, backend=None, options: CheckOptions = None) -> CheckReport:
    """
    Fuehrt die gewaehlten Tests aus. `backend` (BrightwayBackend/MemoryBackend)
    wird fuer Fluss- und Verweis-Test gebraucht; ohne Backend werden nur
    Verweise innerhalb der Dateien geprueft und externe Fluesse gemeldet.
    """
    options = options or CheckOptions()
    started = datetime.now()
    selected = expand_selection(paths)
    report = CheckReport(selected=selected, models={}, issues=[], resolutions={}, options=options)

    for path in context_files(selected):
        report.models[path] = read_workbook(path)

    # Mehrere Fassungen derselben Datei: nur eine zaehlt (versions.py)
    given = [Path(p) for p in paths]
    single_selection = len(given) == 1 and not given[0].is_dir()
    version_issues, report.ignored = resolve_versions(list(report.models.values()), selected, single_selection)
    report.issues.extend(version_issues)
    for path in report.ignored:
        report.models.pop(path, None)
    selected = report.selected = [p for p in selected if p not in report.ignored]

    selected_models = [report.models[p] for p in selected if p in report.models]
    all_models = list(report.models.values())

    previous = {}
    for model in selected_models:
        if model.readable:
            try:
                previous[model.path] = read_review(model.path)
            except Exception as exc:
                report.errors.append(f"{model.path.name}: altes Pruefblatt nicht lesbar ({exc})")

    # Format ------------------------------------------------------------------
    if TEST_FORMAT in options.tests:
        report.issues.extend(check_format(all_models, scope=set(selected)))
    else:
        for model in selected_models:
            report.issues.extend(i for i in model.issues if i.kind.startswith("datei_"))

    # Fluss -------------------------------------------------------------------
    if TEST_FLUSS in options.tests:
        try:
            resolver = Resolver(all_models, backend=backend, config=options.resolve_config)
        except Exception as exc:
            report.errors.append(f"Brightway-Projekt nicht lesbar: {type(exc).__name__}: {exc}")
            logger.debug(traceback.format_exc())
            resolver = Resolver(all_models, backend=None, config=options.resolve_config)
        for text in resolver.warnings:
            report.issues.append(Issue(test=TEST_FLUSS, severity=WARNUNG, kind="konfiguration", message=text))
        for model in selected_models:
            kept = previous.get(model.path)
            searches = {}
            if kept is not None:
                searches = {pos: r.search for pos, r in kept.by_position().items() if r.search}
            for activity in model.activities:
                for exchange in activity.exchanges:
                    if exchange.type not in RESOLVABLE_TYPES:
                        continue
                    if not (exchange.get("name") or exchange.get("reference product")):
                        continue  # meldet der Format-Test
                    try:
                        resolution = resolver.resolve(exchange, search_text=searches.get((exchange.sheet, exchange.row)))
                    except Exception as exc:
                        report.errors.append(
                            f"{model.path.name}/{exchange.sheet} Zeile {exchange.row}: Abgleich fehlgeschlagen "
                            f"({type(exc).__name__}: {exc})"
                        )
                        logger.debug(traceback.format_exc())
                        continue
                    if resolution.ok:
                        report.resolutions[(model.path, exchange.sheet, exchange.row)] = resolution
                    else:
                        issue = resolution_issue(resolution)
                        issue.activity = activity.name
                        report.issues.append(issue)

    # Verweis -----------------------------------------------------------------
    if TEST_VERWEIS in options.tests:
        try:
            report.issues.extend(check_references(selected_models, all_models, backend))
        except Exception as exc:
            report.errors.append(f"Verweis-Test fehlgeschlagen: {type(exc).__name__}: {exc}")
            logger.debug(traceback.format_exc())

    # Pruefblatt --------------------------------------------------------------
    if options.write_review:
        for model in selected_models:
            if not model.readable:
                continue
            file_issues = report.issues_for(model.path)
            try:
                if file_issues:
                    write_review(model.path, file_issues, previous.get(model.path))
                elif previous.get(model.path) is not None:
                    strip_review(model.path)
            except WorkbookLockedError as exc:
                report.errors.append(str(exc))
            except Exception as exc:
                report.errors.append(f"{model.path.name}: Pruefblatt nicht geschrieben ({type(exc).__name__}: {exc})")
                logger.debug(traceback.format_exc())

    if options.write_log and selected:
        report.log_path = selected[0].parent / LOG_FILENAME
        try:
            _write_log(report.log_path, report, started)
        except OSError as exc:
            report.errors.append(f"Fehler-Log nicht geschrieben: {exc}")
    return report
