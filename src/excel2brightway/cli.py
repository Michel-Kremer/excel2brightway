"""
cli.py
======
Konsolenbefehle (siehe pyproject.toml [project.scripts]):

  ex2bw-run             interaktives Menue: [1] Testen/Abgleichen  [2] Laden in Brightway
  ex2bw-check           Stufe 1: Testen & Abgleichen (ehem. test_and_match.py)
  ex2bw-load            Stufe 2: Laden in Brightway (ehem. load_to_brightway.py)
  ex2bw-tidy-registry   flow_registry.yaml aufraeumen (ehem. tidy_flow_registry.py CLI)

Alle nehmen optional --workspace PATH entgegen; ohne das wird die
Umgebungsvariable EX2BW_WORKSPACE, sonst das aktuelle Arbeitsverzeichnis
verwendet (siehe config.resolve_workspace). Verhalten (Menue-Prompts,
"Fortfahren? [j/N]"-Sicherheitsabfrage vor dem Schreiben nach Brightway,
verkettete Check->Load-Frage) entspricht dem der fruehreren Einzelskripte.

Diese Funktionen geben bewusst nichts zurueck (siehe pyproject.toml
[project.scripts]): der generierte Konsolenbefehl ruft sie als
`sys.exit(func())` auf - ein Rueckgabewert ausser None/int wuerde dabei
auf stderr ausgegeben und einen Fehler-Exitcode ausloesen. Wer die
Ergebnisse programmatisch braucht, nutzt matcher.run()/
brightway_writer.write_to_brightway() direkt statt cli.py.
"""

import argparse
from pathlib import Path

from . import brightway_writer
from . import matcher
from . import tidy_flow_registry
from .config import resolve_workspace


def _workspace_argument(parser: argparse.ArgumentParser):
    parser.add_argument(
        "--workspace", type=Path, default=None,
        help="Workspace-Wurzel (Default: $EX2BW_WORKSPACE, sonst aktuelles Arbeitsverzeichnis)",
    )


def check(argv=None):
    """Stufe 1: Testen & Abgleichen."""
    parser = argparse.ArgumentParser(description="Stufe 1: Testen & Abgleichen")
    _workspace_argument(parser)
    parser.add_argument(
        "--no-ecoinvent", dest="ecoinvent_fallback", action="store_false",
        help="Keinen automatischen Abgleich gegen ecoinvent fuer unresolved Flows versuchen",
    )
    args = parser.parse_args(argv)

    ws = resolve_workspace(args.workspace)
    ws.ensure()

    parsed_activities, unresolved = matcher.run(
        clusters=ws.root, registry=ws.registry, out=ws.unresolved,
        ecoinvent_fallback=args.ecoinvent_fallback, resolved_dir=ws.resolved_dir,
    )

    if not unresolved:
        print(f"\nAlles aufgeloest. Naechster Schritt: ex2bw-load (liest {ws.resolved_dir}).")
        choice = input("Jetzt direkt nach Brightway laden? [j/N]: ").strip().lower()
        if choice in ("j", "ja", "y", "yes"):
            load(["--workspace", str(ws.root)])


def load(argv=None):
    """Stufe 2: Laden in Brightway."""
    parser = argparse.ArgumentParser(description="Stufe 2: Laden in Brightway")
    _workspace_argument(parser)
    args = parser.parse_args(argv)

    ws = resolve_workspace(args.workspace)

    if not ws.resolved_dir.exists():
        print(f"{ws.resolved_dir} existiert nicht - zuerst 'ex2bw-check' ausfuehren.")
        return

    files = sorted(ws.resolved_dir.glob("*.yaml"))
    if not files:
        print(f"Keine resolved-Dateien in {ws.resolved_dir} gefunden - zuerst 'ex2bw-check' ausfuehren.")
        return

    print("Verfuegbare resolved-Dateien:")
    for i, path in enumerate(files, 1):
        print(f"  [{i}] {path.name}")

    choice = input("Welche laden? (Standard: alle, sonst z.B. '1,3'): ").strip()
    if not choice:
        selected = files
    else:
        indices = [int(part) for part in choice.split(",") if part.strip()]
        selected = [files[i - 1] for i in indices]

    if not selected:
        print("Keine Auswahl, nichts zu tun.")
        return

    print("\nFolgende Dateien werden nach Brightway geschrieben:")
    for path in selected:
        print(f"  - {path.name}")

    confirm = input("Fortfahren? [j/N]: ").strip().lower()
    if confirm not in ("j", "ja", "y", "yes"):
        print("Abgebrochen.")
        return

    brightway_writer.write_to_brightway(selected)


def tidy_registry(argv=None):
    """flow_registry.yaml aufraeumen (Dedupe + unit-Backfill)."""
    parser = argparse.ArgumentParser(description="flow_registry.yaml aufraeumen (Dedupe + unit-Backfill)")
    _workspace_argument(parser)
    parser.add_argument(
        "--no-backfill", dest="backfill", action="store_false",
        help="Nur Dedupe, kein bw2data-Zugriff",
    )
    args = parser.parse_args(argv)

    ws = resolve_workspace(args.workspace)
    ws.ensure()
    tidy_flow_registry.tidy_registry(ws.registry, ws.root, backfill=args.backfill)


def run(argv=None):
    """Interaktives Menue: [1] Testen/Abgleichen  [2] Laden in Brightway."""
    parser = argparse.ArgumentParser(description="Menue: [1] Testen/Abgleichen  [2] Laden in Brightway")
    _workspace_argument(parser)
    args = parser.parse_args(argv)

    forwarded = ["--workspace", str(args.workspace)] if args.workspace else []

    choice = input("Was moechtest du tun? [1] Testen/Abgleichen  [2] Laden in Brightway (Standard: 1): ").strip()

    if choice in ("", "1"):
        check(forwarded)
        return
    if choice == "2":
        load(forwarded)
        return

    raise ValueError(f"Unbekannte Auswahl: '{choice}' (erwartet '1' oder '2')")


if __name__ == "__main__":
    run()
