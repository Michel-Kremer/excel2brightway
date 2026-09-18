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

ex2bw-load/ex2bw-run nehmen zusaetzlich optional --project NAME entgegen
(Ziel-Brightway-Projekt). Ohne --project wird interaktiv gefragt, Default
ist das aktuell aktivierte bw2data-Projekt.

ex2bw-check fragt (sofern der ecoinvent-Abgleich nicht per --no-ecoinvent
deaktiviert ist und bw2data verfuegbar ist) zusaetzlich einmal nach dem
Brightway-Projekt, das die ecoinvent-/biosphere-Datenbank fuer den Abgleich
enthaelt (siehe ecoinvent_matcher.py - dort ist nichts hardcodiert). Wird
danach direkt nach Brightway geladen, dient dasselbe Projekt als Default
fuer die --project-Abfrage in ex2bw-load, es wird also nicht zweimal
gefragt.

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

    # Projekt fuer den ecoinvent-Abgleich einmal vorab fragen (statt hardcodiert,
    # siehe ecoinvent_matcher.py) - nur wenn der Abgleich ueberhaupt versucht wird
    # und bw2data verfuegbar ist; sonst wie zuvor stillschweigend uebersprungen.
    ecoinvent_project = None
    if args.ecoinvent_fallback:
        try:
            import bw2data  # noqa: F401 -- nur um Verfuegbarkeit/Erreichbarkeit zu pruefen
        except ImportError:
            pass
        else:
            try:
                ecoinvent_project = _choose_project(
                    prompt_label="Welches Projekt enthaelt die ecoinvent-/biosphere-Datenbank fuer den Abgleich"
                )
            except Exception as exc:
                print(f"WARNUNG: Projektauswahl fuer ecoinvent-Abgleich fehlgeschlagen ({exc}), wird uebersprungen.")

    parsed_activities, unresolved = matcher.run(
        clusters=ws.root, registry=ws.registry, out=ws.unresolved,
        ecoinvent_fallback=args.ecoinvent_fallback, resolved_dir=ws.resolved_dir,
        warnings_out=ws.load_warnings,
        ecoinvent_project=ecoinvent_project,
        ecoinvent_prompt=_ecoinvent_db_prompt if ecoinvent_project else None,
    )

    if not unresolved:
        print(f"\nAlles aufgeloest. Naechster Schritt: ex2bw-load (liest {ws.resolved_dir}).")
        choice = input("Jetzt direkt nach Brightway laden? [j/N]: ").strip().lower()
        if choice in ("j", "ja", "y", "yes"):
            # Bereits gewaehltes ecoinvent-Projekt als Default fuer den Schreib-Schritt
            # mitgeben (oft dasselbe Projekt), damit nicht zweimal gefragt wird.
            forwarded = ["--workspace", str(ws.root)]
            if ecoinvent_project:
                forwarded += ["--project", ecoinvent_project]
            load(forwarded)


def _project_argument(parser: argparse.ArgumentParser):
    parser.add_argument(
        "--project", default=None,
        help="Ziel-Brightway-Projekt (Default: aktuell aktiviertes Projekt, interaktive Auswahl moeglich)",
    )


def _choose_project(preselected=None, prompt_label="Welches Projekt") -> str:
    """
    Bestimmt ein Brightway-Projekt (Ziel zum Schreiben ODER Quelle fuer den
    ecoinvent-Abgleich, siehe Aufrufer). preselected (z.B. aus --project)
    wird ohne Rueckfrage uebernommen. Sonst interaktive Auswahl mit dem
    aktuell aktivierten Projekt als Default. `prompt_label` passt den
    Fragetext an den jeweiligen Zweck an.
    """
    import bw2data as bd

    current = bd.projects.current
    if preselected:
        return preselected

    available = sorted(p.name for p in bd.projects)
    print(f"\nVerfuegbare Brightway-Projekte (aktuell aktiv: '{current}'):")
    for i, name in enumerate(available, 1):
        marker = "  <- aktuell aktiv" if name == current else ""
        print(f"  [{i}] {name}{marker}")

    choice = input(f"{prompt_label}? (Standard: aktuell aktiviertes '{current}'): ").strip()
    if not choice:
        return current
    if choice.isdigit():
        return available[int(choice) - 1]
    return choice


def _ecoinvent_db_prompt(label, candidates):
    """
    Rueckfrage-Callback fuer matcher.run(ecoinvent_prompt=...): wird nur
    aufgerufen, wenn die 'database'-Spalte einer Exchange-Zeile (`label`,
    kann None sein) nicht eindeutig einer Datenbank im zuvor gewaehlten
    ecoinvent-Projekt zugeordnet werden konnte (siehe
    ecoinvent_matcher._resolve_database). Wird pro Label nur einmal
    aufgerufen (Ergebnis wird dort gecacht).
    """
    if label:
        print(
            f"\nKonnte die im Excel angegebene Datenbank '{label}' nicht eindeutig "
            f"einer Datenbank im gewaehlten Projekt zuordnen."
        )
    else:
        print(
            "\nFuer den ecoinvent-Abgleich wird eine Datenbank im gewaehlten Projekt "
            "benoetigt (keine 'database'-Angabe im Excel gefunden)."
        )

    if not candidates:
        print("(Projekt enthaelt keine Datenbanken.)")
        return None

    print("Verfuegbare Datenbanken:")
    for i, name in enumerate(candidates, 1):
        print(f"  [{i}] {name}")

    choice = input("Welche verwenden? (leer = fuer dieses Label ueberspringen): ").strip()
    if not choice:
        return None
    if choice.isdigit() and 1 <= int(choice) <= len(candidates):
        return candidates[int(choice) - 1]
    return choice


def load(argv=None):
    """Stufe 2: Laden in Brightway."""
    parser = argparse.ArgumentParser(description="Stufe 2: Laden in Brightway")
    _workspace_argument(parser)
    _project_argument(parser)
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

    project_name = _choose_project(args.project)

    print("\nFolgende Dateien werden nach Brightway geschrieben:")
    for path in selected:
        print(f"  - {path.name}")
    print(f"Ziel-Projekt: '{project_name}'")

    confirm = input("Fortfahren? [j/N]: ").strip().lower()
    if confirm not in ("j", "ja", "y", "yes"):
        print("Abgebrochen.")
        return

    brightway_writer.write_to_brightway(selected, project_name=project_name)


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
    _project_argument(parser)
    args = parser.parse_args(argv)

    forwarded = ["--workspace", str(args.workspace)] if args.workspace else []
    forwarded_load = forwarded + (["--project", args.project] if args.project else [])

    choice = input("Was moechtest du tun? [1] Testen/Abgleichen  [2] Laden in Brightway (Standard: 1): ").strip()

    if choice in ("", "1"):
        check(forwarded)
        return
    if choice == "2":
        load(forwarded_load)
        return

    raise ValueError(f"Unbekannte Auswahl: '{choice}' (erwartet '1' oder '2')")


if __name__ == "__main__":
    run()
