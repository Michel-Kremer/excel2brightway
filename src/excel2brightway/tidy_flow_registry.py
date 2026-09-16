"""
tidy_flow_registry.py
=====================
Pflege der flow_registry.yaml. Zwei Aufraeumschritte:

  1. Dedupe: Eintraege, die auf denselben (database, code) zeigen, aber
     unter mehreren Namen stehen (z.B. 'Yttrium oxide' und
     'market for yttrium oxide'), werden auf einen Namen reduziert -
     bevorzugt auf den, der tatsaechlich in einer *.xlsx im Workspace
     vorkommt, sonst auf den laengsten (spezifischsten) Namen.

  2. Backfill: fehlende 'unit' (Technosphere) bzw. 'unit' + 'categories'
     (Biosphere) werden ueber bw2data aus dem Ziel-Projekt nachgetragen
     (read-only). Ist bw2data / das Projekt nicht erreichbar, wird dieser
     Schritt uebersprungen und die betroffenen Eintraege bleiben
     unveraendert.

Vor dem Schreiben wird flow_registry.yaml nach
flow_registry.yaml.bak_<YYYYMMDD_HHMMSS> gesichert.

Fuer den normalen Gebrauch siehe den Konsolenbefehl ex2bw-tidy-registry
(cli.py). Niedrigschwelliger Direktaufruf ohne Workspace-Konzept:

    python -m excel2brightway.tidy_flow_registry --registry flow_registry.yaml --excel .
"""

import argparse
import shutil
from datetime import datetime
from pathlib import Path

import yaml

from .matcher import normalize

PROJECT_NAME = "ecoinvent-3.12-cutoff"

HEADER = """\
# ============================================================
# flow_registry.yaml - einmal pflegen, in allen Cluster-Dateien wiederverwenden
# Nur Flows eintragen, die NICHT ueber den internen Namensabgleich
# (name-Feld einer anderen Aktivitaet) gefunden werden.
#
# Auto-gepflegt von matcher.py (ecoinvent-Abgleich) und einmalig
# aufgeraeumt von tidy_flow_registry.py ({date}).
# ============================================================
"""


def _excel_flow_names(excel_dir: Path) -> set:
    """Alle normalisierten technosphere-/biosphere-Flow-Namen aus *.xlsx in excel_dir."""
    try:
        from . import excel_loader
    except Exception as exc:  # pragma: no cover - nur Diagnose
        print(f"WARNUNG: excel_loader nicht importierbar ({exc}); Dedupe ohne Excel-Bezug.")
        return set()

    names = set()
    try:
        activities, _internal_index, _project_params, _database_params = excel_loader.load_clusters(excel_dir)
    except Exception as exc:
        print(f"WARNUNG: Konnte {excel_dir} nicht laden ({exc}); Dedupe ohne Excel-Bezug.")
        return set()

    for act in activities:
        for section in ("in", "emit"):
            names.update(normalize(name) for name in (act.get(section) or {}))
    return names


def _dedupe_section(section_name: str, entries: dict, excel_names: set) -> dict:
    """Reduziert Eintraege mit identischem code auf einen Namen."""
    by_code = {}
    no_code = {}
    for name, entry in entries.items():
        entry = entry or {}
        code = entry.get("code")
        if code:
            by_code.setdefault(code, []).append((name, entry))
        else:
            no_code[name] = entry

    kept = {}
    for code, group in by_code.items():
        if len(group) == 1:
            name, entry = group[0]
            kept[name] = entry
            continue

        # Kandidat, der in den Excel-Dateien vorkommt, hat Vorrang;
        # sonst der laengste (spezifischste) Name.
        preferred = [g for g in group if normalize(g[0]) in excel_names]
        pool = preferred or group
        winner_name, winner_entry = max(pool, key=lambda g: len(g[0]))

        dropped = [n for n, _ in group if n != winner_name]
        print(
            f"  [{section_name}] '{winner_name}' behalten; "
            f"entfernt (gleicher code {code}): {', '.join(repr(d) for d in dropped)}"
        )
        kept[winner_name] = winner_entry

    kept.update(no_code)
    # Stabile, lesbare Reihenfolge
    return dict(sorted(kept.items(), key=lambda kv: kv[0].lower()))


def _backfill(registry: dict) -> int:
    """Traegt fehlende unit/categories aus dem Brightway-Projekt nach. Gibt Anzahl Aenderungen zurueck."""
    try:
        import bw2data as bd
    except ImportError:
        print("WARNUNG: bw2data nicht installiert - Backfill uebersprungen.")
        return 0

    try:
        if bd.projects.current != PROJECT_NAME:
            bd.projects.set_current(PROJECT_NAME)
    except Exception as exc:
        print(f"WARNUNG: Projekt '{PROJECT_NAME}' nicht erreichbar ({exc}) - Backfill uebersprungen.")
        return 0

    changed = 0
    not_found = []
    for section in ("technosphere", "biosphere"):
        for name, entry in (registry.get(section) or {}).items():
            database, code = entry.get("database"), entry.get("code")
            if not (database and code):
                continue
            needs_unit = not entry.get("unit")
            needs_cats = section == "biosphere" and not entry.get("categories")
            if not (needs_unit or needs_cats):
                continue

            try:
                node = bd.get_activity((database, code))
            except Exception:
                not_found.append(f"{section}/{name} ({database}/{code})")
                continue

            if needs_unit and node.get("unit"):
                entry["unit"] = node.get("unit")
                changed += 1
            if needs_cats and node.get("categories"):
                entry["categories"] = list(node.get("categories"))
                changed += 1

    if not_found:
        print("  In Brightway nicht gefunden (unveraendert gelassen):")
        for item in not_found:
            print(f"    - {item}")
    return changed


def tidy_registry(registry: Path, excel_dir: Path, backfill: bool = True):
    """
    Bibliotheksfunktion hinter ex2bw-tidy-registry: dedupe + optionaler
    unit/categories-Backfill fuer `registry`, mit `excel_dir` als Bezug
    fuer den Dedupe-Vorrang (siehe _dedupe_section). Schreibt vorher ein
    Backup, danach die aufgeraeumte flow_registry.yaml.
    """
    raw = yaml.safe_load(registry.read_text(encoding="utf-8")) or {}
    reg = {section: dict(raw.get(section) or {}) for section in ("technosphere", "biosphere")}

    before = {s: len(reg[s]) for s in reg}
    print(f"Geladen: {before['technosphere']} technosphere, {before['biosphere']} biosphere.")

    excel_names = _excel_flow_names(excel_dir)
    print(f"{len(excel_names)} Flow-Namen aus {excel_dir} fuer den Dedupe-Bezug.")

    print("Dedupe:")
    for section in ("technosphere", "biosphere"):
        reg[section] = _dedupe_section(section, reg[section], excel_names)

    after = {s: len(reg[s]) for s in reg}
    print(f"Nach Dedupe: {after['technosphere']} technosphere, {after['biosphere']} biosphere.")

    if backfill:
        print("Backfill unit/categories aus Brightway:")
        n = _backfill(reg)
        print(f"  {n} Feld(er) nachgetragen.")

    backup = registry.with_name(f"{registry.name}.bak_{datetime.now():%Y%m%d_%H%M%S}")
    shutil.copy2(registry, backup)
    print(f"Backup: {backup.name}")

    body = yaml.dump(reg, allow_unicode=True, sort_keys=False, default_flow_style=False)
    header = HEADER.format(date=f"{datetime.now():%Y-%m-%d}")
    registry.write_text(header + "\n" + body, encoding="utf-8")
    print(f"Geschrieben: {registry}")


def main():
    parser = argparse.ArgumentParser(description="flow_registry.yaml aufraeumen (Dedupe + unit-Backfill)")
    parser.add_argument("--registry", type=Path, default=Path("flow_registry.yaml"))
    parser.add_argument("--excel", type=Path, default=Path("."))
    parser.add_argument("--no-backfill", dest="backfill", action="store_false",
                        help="Nur Dedupe, kein bw2data-Zugriff")
    args = parser.parse_args()

    tidy_registry(args.registry, args.excel, backfill=args.backfill)


if __name__ == "__main__":
    main()
