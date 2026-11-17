"""
cli.py
======
Konsolenbefehle (siehe pyproject.toml). Fuer die normale Arbeit gibt es das
Fenster (ex2bw-gui); die Konsole bleibt fuer Tests und Automatisierung.

  ex2bw-check   PFAD...   Tests ausfuehren, Pruefblatt + ex2bw.log schreiben
  ex2bw-apply   DATEI...  Auswahl aus dem Pruefblatt uebernehmen
  ex2bw-write   PFAD...   erneut pruefen und nach Brightway schreiben
  ex2bw-migrate PFAD...   Datenbank/Version umstellen (Grundform)
  ex2bw-gui               Fenster starten

PFAD = Excel-Dateien und/oder Ordner (dann alle *.xlsx darin).

Rueckgabewerte: 0 = alles in Ordnung, 1 = Meldungen/blockiert, 2 = Fehler.
"""

import argparse
import sys
from pathlib import Path

from .core.apply import DataChangedError, apply_review, summarize
from .core.checks import CheckOptions, run_checks
from .core.model import ALL_TESTS
from .core.resolve import ResolveConfig
from .core.review_sheet import WorkbookLockedError


def _common(parser, tests=True):
    parser.add_argument("paths", nargs="+", type=Path, help="Excel-Dateien und/oder Ordner")
    parser.add_argument("--project", default=None, help="Brightway-Projekt (fuer Fluss- und Verweis-Test)")
    parser.add_argument(
        "--databases", default=None,
        help="Verknuepfungs-Datenbanken, kommagetrennt (Standard: alle Datenbanken des Projekts)",
    )
    parser.add_argument(
        "--locations", default="GLO,RER,RoW",
        help="Standort-Rangfolge ohne Angabe in der Excel, kommagetrennt (Standard: GLO,RER,RoW)",
    )
    parser.add_argument(
        "--format-ueberbruecken", action="store_true",
        help="Format-Fehler blockieren nicht (Zeilen ohne Pflichtangaben werden beim Schreiben ausgelassen)",
    )
    if tests:
        parser.add_argument(
            "--tests", default=",".join(ALL_TESTS),
            help=f"Auszufuehrende Tests, kommagetrennt (Standard: {','.join(ALL_TESTS)})",
        )


def _split(value):
    return [v.strip() for v in value.split(",") if v.strip()] if value else []


def _backend(project):
    if not project:
        return None
    from .core.backend import BrightwayBackend

    return BrightwayBackend(project)


def _options(args, tests=None):
    selected = tuple(tests or _split(getattr(args, "tests", None)) or ALL_TESTS)
    unknown = [t for t in selected if t not in ALL_TESTS]
    if unknown:
        raise SystemExit(f"Unbekannte(r) Test(s): {', '.join(unknown)} (erlaubt: {', '.join(ALL_TESTS)})")
    return CheckOptions(
        tests=selected,
        bypass_format=args.format_ueberbruecken,
        resolve_config=ResolveConfig(
            link_databases=_split(args.databases) or None,
            location_priority=tuple(_split(args.locations)) or ("GLO", "RER", "RoW"),
        ),
    )


def _print_report(report):
    print(report.summary())
    for issue in report.issues:
        print(f"  [{issue.test}] {issue.severity} {issue.location_label()}: {issue.message}")
    for error in report.errors:
        print(f"  FEHLER: {error}")
    if report.log_path:
        print(f"Fehler-Log: {report.log_path}")


def check(argv=None):
    parser = argparse.ArgumentParser(description="excel2brightway: Tests ausfuehren und Pruefblatt schreiben")
    _common(parser)
    parser.add_argument("--kein-pruefblatt", action="store_true", help="Excel-Dateien nicht veraendern")
    args = parser.parse_args(argv)
    options = _options(args)
    options.write_review = not args.kein_pruefblatt
    try:
        report = run_checks(args.paths, backend=_backend(args.project), options=options)
    except Exception as exc:
        print(f"FEHLER: {exc}")
        return 2
    _print_report(report)
    return 0 if report.ok else 1


def apply(argv=None):
    parser = argparse.ArgumentParser(description="excel2brightway: Auswahl aus dem Pruefblatt uebernehmen")
    parser.add_argument("paths", nargs="+", type=Path, help="Excel-Dateien mit Pruefblatt")
    args = parser.parse_args(argv)
    status = 0
    for path in args.paths:
        try:
            print(summarize(apply_review(path)))
        except (DataChangedError, WorkbookLockedError) as exc:
            print(f"FEHLER: {exc}")
            status = 1
    return status


def _confirm(question):
    return input(f"{question} [j/N]: ").strip().lower() in ("j", "ja", "y", "yes")


def write(argv=None):
    parser = argparse.ArgumentParser(description="excel2brightway: erneut pruefen und nach Brightway schreiben")
    _common(parser, tests=False)
    parser.add_argument("--loeschen-erlauben", action="store_true",
                        help="Aktivitaeten loeschen, die in Brightway, aber in keiner Datei stehen")
    parser.add_argument("--verweise-brechen-erlauben", action="store_true",
                        help="auch schreiben, wenn Verweise aus anderen Datenbanken ins Leere laufen")
    parser.add_argument("--ja", action="store_true", help="ohne Rueckfrage schreiben")
    args = parser.parse_args(argv)
    if not args.project:
        print("FEHLER: --project ist zum Schreiben noetig.")
        return 2

    from .core.write import WriteBlockedError, describe_plan, execute_write, plan_write

    backend = _backend(args.project)
    report, plans = plan_write(
        args.paths, backend, _options(args, tests=ALL_TESTS),
        allow_broken_references=args.verweise_brechen_erlauben,
    )
    _print_report(report)
    if not plans:
        print("Nicht geschrieben - bitte zuerst die Meldungen im Pruefblatt beheben.")
        return 1
    for plan in plans:
        print(describe_plan(plan))
    if not args.ja and not _confirm(f"In Projekt '{args.project}' schreiben?"):
        print("Abgebrochen.")
        return 1
    try:
        written = execute_write(
            plans, args.project, allow_delete=args.loeschen_erlauben,
            allow_broken_references=args.verweise_brechen_erlauben,
        )
    except WriteBlockedError as exc:
        print(f"Nicht geschrieben: {exc}")
        return 1
    except Exception as exc:
        print(f"FEHLER beim Schreiben: {type(exc).__name__}: {exc}")
        return 2
    for database, count in written:
        print(f"{database}: {count} Aktivitaet(en) geschrieben.")
    return 0


def migrate(argv=None):
    parser = argparse.ArgumentParser(description="excel2brightway: Datenbank/Version umstellen (Grundform)")
    _common(parser, tests=False)
    parser.add_argument("--source", required=True, help="Quell-Datenbank, z.B. ecoinvent-3.10-cutoff")
    parser.add_argument("--target", required=True, help="Ziel-Datenbank, z.B. ecoinvent-3.12-cutoff")
    parser.add_argument("--target-project", default=None, help="Projekt der Ziel-Datenbank (Standard: --project)")
    parser.add_argument("--source-biosphere", default=None)
    parser.add_argument("--target-biosphere", default=None)
    args = parser.parse_args(argv)
    if not args.project:
        print("FEHLER: --project ist noetig.")
        return 2
    from .core.migrate import MigrationBlockedError, migrate_files

    source_backend = _backend(args.project)
    target_backend = _backend(args.target_project) if args.target_project else source_backend
    try:
        result = migrate_files(
            args.paths, source_backend, args.source, target_backend, args.target,
            source_biosphere=args.source_biosphere, target_biosphere=args.target_biosphere,
            options=_options(args, tests=ALL_TESTS),
        )
    except MigrationBlockedError as exc:
        print(f"Nicht umgestellt: {exc}")
        return 1
    print(result.summary())
    for path in result.files:
        print(f"  neue Datei: {path}")
    if result.report is not None:
        _print_report(result.report)
    return 0 if result.report is None or result.report.ok else 1


def gui(argv=None):
    from .gui import main

    return main()


if __name__ == "__main__":  # pragma: no cover
    sys.exit(check())
