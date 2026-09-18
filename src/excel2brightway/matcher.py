"""
matcher.py
==========
Testen & Abgleichen (Stufe 1 von 2): Liest LCI-Cluster-Excel-Dateien im
nativen Brightway2-Excel-Format (ueber bw2io, siehe excel_loader.py) und
eine zentrale flow_registry.yaml ein, loest technosphere-/biosphere-
Referenzen auf:

  1. intern    -> gegen den "name" einer anderen Aktivitaet (ueber alle
                  geladenen Cluster-Dateien hinweg, nicht nur die eigene)
  2. extern    -> gegen die flow_registry.yaml
  3. ecoinvent -> automatischer Abgleich gegen eine lokale ecoinvent-
                  Datenbank (siehe ecoinvent_matcher.py), read-only

Jede erfolgreich aufgeloeste Aktivitaet referenziert danach jeden Flow
konkret ueber (database, code). Aktivitaeten, die vollstaendig aufgeloest
sind, werden pro Ziel-Datenbank nach resolved/<database>.yaml geschrieben.
Nicht aufloesbare Flows landen mit Fuzzy-Vorschlaegen in unresolved.yaml.

Dieses Modul schreibt NICHT nach Brightway - das uebernimmt Stufe 2
(siehe brightway_writer.py). Fuer den normalen Gebrauch siehe die
Konsolenbefehle in cli.py (ex2bw-check/ex2bw-load); `run()` ist die
Bibliotheksfunktion dahinter, `main()`/das CLI unten ist ein niedrigschwelliger,
workspace-unabhaengiger Direktaufruf fuer Skripte:

    python -m excel2brightway.matcher --clusters . --registry output_excel2brightway/flow_registry.yaml
"""

import argparse
import difflib
import re
from pathlib import Path

import yaml


AMOUNT_UNIT_RE = re.compile(r"^\s*(-?[\d.]+)\s*([^\d\s].*)\s*$")


def normalize(name: str) -> str:
    """Normalisiert Flow-/Aktivitaetsnamen fuer den Abgleich (lowercase, getrimmt)."""
    return " ".join(str(name).strip().lower().split())


def parse_amount_unit(value):
    """
    Akzeptiert entweder:
      - String "0.19 kg"
      - Dict {amount, unit, comment?, location?, database?}
    Gibt ein normalisiertes Dict zurueck.
    """
    if isinstance(value, dict):
        result = dict(value)
        result.setdefault("comment", None)
        result.setdefault("location", None)
        result.setdefault("database", None)
        return result

    if isinstance(value, str):
        match = AMOUNT_UNIT_RE.match(value)
        if not match:
            raise ValueError(f"Kann Menge/Einheit nicht parsen: '{value}'")
        amount, unit = match.groups()
        return {
            "amount": float(amount),
            "unit": unit.strip(),
            "comment": None,
            "location": None,
            "database": None,
        }

    raise ValueError(f"Unerwarteter Exchange-Wert: {value!r}")


def load_registry(path: Path) -> dict:
    """Laedt flow_registry.yaml und normalisiert die Keys."""
    with open(path, encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}

    registry = {"technosphere": {}, "biosphere": {}}
    for section in ("technosphere", "biosphere"):
        for name, entry in (raw.get(section) or {}).items():
            registry[section][normalize(name)] = {"display_name": name, **(entry or {})}
    return registry


UNIT_ALIASES = {
    "kg": "kilogram", "kilogram": "kilogram", "kilograms": "kilogram",
    "g": "gram", "gram": "gram", "grams": "gram",
    "t": "metric ton", "ton": "metric ton", "tonne": "metric ton", "metric ton": "metric ton",
    "kwh": "kilowatt hour", "kilowatt hour": "kilowatt hour", "kilowatt-hour": "kilowatt hour",
    "mj": "megajoule", "megajoule": "megajoule", "mega joule": "megajoule",
    "kj": "kilojoule", "kilojoule": "kilojoule", "kilo joule": "kilojoule",
    "m3": "cubic meter", "m³": "cubic meter", "cubic meter": "cubic meter", "cubic metre": "cubic meter",
    "km": "kilometer", "kilometer": "kilometer", "kilometre": "kilometer",
    "l": "litre", "liter": "litre", "litre": "litre",
    "unit": "unit", "piece": "unit", "item": "unit",
    "tkm": "ton kilometer", "ton kilometer": "ton kilometer", "metric ton*km": "ton kilometer",
    "hour": "hour", "h": "hour",
    "m2": "square meter", "m²": "square meter", "square meter": "square meter", "square metre": "square meter",
    "m2a": "square meter-year", "m2*a": "square meter-year", "square meter-year": "square meter-year",
}


def normalize_unit(unit):
    """Normalisiert gaengige Einheiten-Schreibweisen fuer den Vergleich (kg==kilogram, kWh==kilowatt hour, ...)."""
    if not unit:
        return None
    key = str(unit).strip().lower()
    return UNIT_ALIASES.get(key, key)


def build_internal_index(activities: list) -> dict:
    """
    Baut den internen Namensindex ueber eine Aktivitaetenliste (inkl.
    Duplikat-Warnung). Erwartet je Aktivitaet die Felder 'name', 'id',
    'unit' und '_database'. Wird vom Excel-Loader (excel_loader.py)
    genutzt, damit die Duplikat-Erkennung an einer Stelle sitzt.
    """
    internal_index = {}
    for act in activities:
        key = normalize(act["name"])
        if key in internal_index:
            existing = internal_index[key]
            print(
                f"WARNUNG: Doppelter Aktivitaetsname '{act['name']}' "
                f"({existing['database']}/{existing['code']} vs. {act['_database']}/{act['id']})"
            )
        internal_index[key] = {"database": act["_database"], "code": act["id"], "unit": act.get("unit")}
    return internal_index


def resolve_flow(name: str, internal_index: dict, registry_section: dict):
    """
    Versucht, einen Flow-Namen aufzuloesen:
      1. gegen interne Aktivitaeten (internal_index)
      2. gegen die passende Registry-Sektion (technosphere ODER biosphere)
    Gibt (resolution_dict_oder_None, kind) zurueck, kind in {'internal', 'external', None}.
    Das resolution_dict enthaelt, wenn bekannt, auch 'unit' der Zielaktivitaet/
    des Zielflows fuer den Einheiten-Abgleich in build_exchanges().
    """
    key = normalize(name)

    if key in internal_index:
        entry = internal_index[key]
        return {"database": entry["database"], "code": entry["code"], "unit": entry.get("unit")}, "internal"

    if key in registry_section:
        return registry_section[key], "external"

    return None, None


def suggest(name: str, candidates: list, n=3, cutoff=0.6):
    """Fuzzy-Vorschlaege fuer einen nicht aufgeloesten Flow-Namen (z.B. bei Tippfehlern)."""
    return difflib.get_close_matches(normalize(name), candidates, n=n, cutoff=cutoff)


def _try_ecoinvent_match(flow_name, exch_type, value, registry_section, new_entries, ecoinvent_project, ecoinvent_prompt):
    """
    Letzter Auflösungsversuch, bevor ein Flow als unresolved gilt: Abgleich
    gegen eine lokale ecoinvent-/biosphere-Datenbank in `ecoinvent_project`
    (siehe ecoinvent_matcher.py; `ecoinvent_project` ist None, wenn kein
    Projekt gewaehlt wurde/verfuegbar ist - dann wird der Versuch
    stillschweigend uebersprungen). `ecoinvent_prompt` erlaubt bei
    mehrdeutiger 'database'-Zuordnung eine interaktive Rueckfrage (siehe
    cli.py); ohne Callback bleibt der Flow in dem Fall unresolved. Bei
    Treffer wird der Eintrag in `registry_section` (Cache fuer diesen Lauf)
    und in `new_entries` (zum Anhaengen an flow_registry.yaml) abgelegt.
    Gibt den registry-kompatiblen Eintrag oder None zurueck.
    """
    if not ecoinvent_project:
        return None

    try:
        from . import ecoinvent_matcher
    except ImportError:
        return None

    try:
        if exch_type == "technosphere":
            entry = ecoinvent_matcher.match_technosphere(
                ecoinvent_project, flow_name, location_hint=value.get("location"),
                database_hint=value.get("database"), prompt=ecoinvent_prompt,
            )
        else:
            entry = ecoinvent_matcher.match_biosphere(
                ecoinvent_project, flow_name,
                database_hint=value.get("database"), prompt=ecoinvent_prompt,
            )
    except Exception as exc:
        print(f"WARNUNG: ecoinvent-Abgleich fuer '{flow_name}' fehlgeschlagen: {exc}")
        return None

    if entry is None:
        return None

    resolution = {"display_name": flow_name, **entry}
    registry_section[normalize(flow_name)] = resolution
    if new_entries is not None:
        new_entries.setdefault(exch_type, {})[flow_name] = entry
    return resolution


def build_exchanges(
    activity, internal_index, registry, candidate_pool, unresolved, new_entries=None,
    ecoinvent_project=None, ecoinvent_prompt=None,
):
    """
    Baut die Exchange-Liste einer Aktivitaet (bw2io-artige Struktur) und
    sammelt nicht aufloesbare Flows in `unresolved`. Jeder aufgeloeste
    technosphere-/biosphere-Exchange referenziert seine Zielaktivitaet
    einheitlich ueber 'input_database'/'input_code' - egal ob intern,
    ueber die Registry oder per ecoinvent-Abgleich gefunden -, damit
    Stufe 2 (Laden in Brightway) ohne erneuten Namens-Abgleich auskommt.
    `ecoinvent_project`/`ecoinvent_prompt` siehe _try_ecoinvent_match();
    `ecoinvent_project=None` deaktiviert den ecoinvent-Abgleich.
    """
    exchanges = []

    prod = parse_amount_unit(activity["output"])
    exchanges.append({
        "type": "production",
        "name": activity["name"],
        "amount": prod["amount"],
        "unit": prod["unit"],
    })

    for section, exch_type in (("in", "technosphere"), ("emit", "biosphere")):
        for flow_name, raw_value in (activity.get(section) or {}).items():
            value = parse_amount_unit(raw_value)
            registry_section = registry["technosphere"] if exch_type == "technosphere" else registry["biosphere"]

            resolution, kind = resolve_flow(flow_name, internal_index, registry_section)

            if resolution is None and ecoinvent_project:
                resolution = _try_ecoinvent_match(
                    flow_name, exch_type, value, registry_section, new_entries,
                    ecoinvent_project, ecoinvent_prompt,
                )
                if resolution is not None:
                    kind = "external"

            if resolution is None:
                unresolved.append({
                    "activity_id": activity["id"],
                    "database": activity["_database"],
                    "source_file": activity["_source_file"],
                    "flow_name": flow_name,
                    "exchange_type": exch_type,
                    "amount": value["amount"],
                    "unit": value["unit"],
                    "reason": "not_found",
                    "suggestions": suggest(flow_name, candidate_pool),
                })
                continue

            expected_unit = resolution.get("unit")
            if expected_unit and normalize_unit(expected_unit) != normalize_unit(value["unit"]):
                unresolved.append({
                    "activity_id": activity["id"],
                    "database": activity["_database"],
                    "source_file": activity["_source_file"],
                    "flow_name": flow_name,
                    "exchange_type": exch_type,
                    "amount": value["amount"],
                    "unit": value["unit"],
                    "reason": "unit_mismatch",
                    "expected_unit": expected_unit,
                    "suggestions": [],
                })
                continue

            if kind == "internal":
                input_database = resolution["database"]
            else:
                # external (registry oder ecoinvent-Abgleich): die Excel-Spalte
                # 'database' ist nur ein menschenlesbares Label (z.B. 'ecoinvent 3.12'),
                # nicht zwingend der echte Brightway-Datenbankname (z.B.
                # 'ecoinvent-3.12-cutoff') - der aufgeloeste Eintrag ist massgeblich,
                # die Excel-Spalte nur Fallback, falls der Eintrag keinen hat.
                input_database = resolution.get("database") or value.get("database")
            input_code = resolution.get("code")

            if kind == "external" and not input_code:
                unresolved.append({
                    "activity_id": activity["id"],
                    "database": activity["_database"],
                    "source_file": activity["_source_file"],
                    "flow_name": flow_name,
                    "exchange_type": exch_type,
                    "amount": value["amount"],
                    "unit": value["unit"],
                    "reason": "missing_code",
                    "suggestions": [],
                })
                continue

            exchange = {
                "type": exch_type,
                "name": flow_name,
                "amount": value["amount"],
                "unit": value["unit"],
                "input_database": input_database,
                "input_code": input_code,
            }

            if value.get("comment"):
                exchange["comment"] = value["comment"]
            if value.get("formula"):
                exchange["formula"] = value["formula"]

            exchanges.append(exchange)

    return exchanges


def _append_to_section(lines: list, section_key: str, block_lines: list) -> list:
    """
    Fuegt `block_lines` am Ende einer Top-Level-Sektion (z.B.
    'technosphere:') ein. Eine leere Sektion kann im Text als Flow-Stil
    vorliegen (z.B. 'biosphere: {}' - so schreibt yaml.dump() ein leeres
    dict, und so sieht auch die REGISTRY_TEMPLATE in config.py aus, bzw.
    so kann tidy_flow_registry.py eine Sektion nach dem Bereinigen
    hinterlassen). Eingerueckte Zeilen einfach darunter zu haengen waere
    dann kein gueltiges YAML mehr (Flow-Wert + Block-Kinder auf demselben
    Schluessel) - der Header wird daher zuerst auf Block-Stil normalisiert.
    """
    start = None
    for i, line in enumerate(lines):
        if line.startswith(f"{section_key}:"):
            start = i
            break
    if start is None:
        return lines + ["", f"{section_key}:"] + block_lines

    if lines[start].strip() != f"{section_key}:":
        lines = lines[:start] + [f"{section_key}:"] + lines[start + 1:]

    end = len(lines)
    for i in range(start + 1, len(lines)):
        if lines[i] and not lines[i][0].isspace():
            end = i
            break

    return lines[:end] + block_lines + lines[end:]


def _persist_new_registry_entries(registry_path: Path, new_entries: dict):
    """
    Haengt per ecoinvent-Abgleich automatisch gefundene Flows an
    flow_registry.yaml an (mit Kommentar + Datum). Bestehender Inhalt,
    inklusive Kommentare, bleibt dabei erhalten (reiner Text-Append,
    kein Neuschreiben der ganzen Datei ueber yaml.dump).
    """
    from datetime import date

    lines = registry_path.read_text(encoding="utf-8").splitlines()
    today = date.today().isoformat()

    for section in ("technosphere", "biosphere"):
        entries = new_entries.get(section)
        if not entries:
            continue

        block_lines = []
        for name, entry in entries.items():
            block_lines.append(f"  # auto: ecoinvent-Abgleich, {today}")
            snippet = yaml.dump({name: entry}, allow_unicode=True, sort_keys=False)
            block_lines.extend("  " + line for line in snippet.splitlines())

        lines = _append_to_section(lines, section, block_lines)

    registry_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


_UNSAFE_FILENAME_RE = re.compile(r"[^A-Za-z0-9._-]+")


def _safe_filename(name: str) -> str:
    """Macht einen Datenbanknamen dateisystemtauglich (fuer resolved/<name>.yaml)."""
    return _UNSAFE_FILENAME_RE.sub("_", name).strip("_") or "database"


def _write_resolved_files(
    parsed_activities: dict, unresolved: list, resolved_dir: Path,
    project_parameters: list = None, database_parameters: list = None,
) -> tuple[list, list]:
    """
    Schreibt fuer jede Ziel-Datenbank ohne unresolved-Flows eine
    resolved/<database>.yaml - vollstaendig aufgeloeste Aktivitaeten,
    bereit fuer Stufe 2 (siehe brightway_writer.py). Datenbanken mit
    mindestens einem unresolved-Flow werden uebersprungen, damit dort nie
    unvollstaendige Daten geladen werden koennen.
    `project_parameters`/`database_parameters` (aus excel_loader.load_clusters,
    ueber alle Cluster-Dateien dieses Laufs zusammengefuehrt) werden - falls
    vorhanden - jeder geschriebenen resolved-Datei vorangestellt; Stufe 2
    schreibt sie als echte, im Activity Browser editierbare bw2data-Parameter.
    Gibt (geschriebene_datenbanken, uebersprungene_datenbanken) zurueck.
    """
    unresolved_databases = {entry["database"] for entry in unresolved}

    by_database = {}
    for (database, code), act in parsed_activities.items():
        by_database.setdefault(database, []).append({"code": code, **act})

    written, skipped = [], []
    for database, acts in sorted(by_database.items()):
        if database in unresolved_databases:
            skipped.append(database)
            continue

        resolved_dir.mkdir(parents=True, exist_ok=True)
        path = resolved_dir / f"{_safe_filename(database)}.yaml"
        content = {"database": database}
        if project_parameters:
            content["project_parameters"] = project_parameters
        if database_parameters:
            content["database_parameters"] = database_parameters
        content["activities"] = acts
        with open(path, "w", encoding="utf-8") as f:
            yaml.dump(content, f, allow_unicode=True, sort_keys=False)
        written.append(database)

    return written, skipped


def run(
    clusters: Path, registry: Path, out: Path,
    ecoinvent_fallback: bool = True, resolved_dir: Path = None,
    warnings_out: Path = None, ecoinvent_project: str = None, ecoinvent_prompt=None,
):
    """
    Fuehrt den kompletten Test-/Abgleich-Durchlauf aus (ohne Kommandozeile).
    `clusters` ist das Verzeichnis mit den LCI-Cluster-Excel-Dateien
    (*.xlsx im Brightway2-Excel-Format), die ueber bw2io eingelesen werden.
    `ecoinvent_fallback`: wenn True, werden Flows, die weder intern noch
    in flow_registry.yaml gefunden werden, zusaetzlich automatisch gegen
    eine lokale ecoinvent-/biosphere-Datenbank abgeglichen (siehe
    ecoinvent_matcher.py, read-only). Treffer werden in flow_registry.yaml
    ergaenzt.
    `ecoinvent_project`: Brightway-Projekt, in dem diese Datenbanken liegen
    - kommt von aussen (siehe cli.py, das den Nutzer einmal fragt), nichts
    davon ist hier hardcodiert. Ist `ecoinvent_fallback` True, aber
    `ecoinvent_project` None (z.B. bw2data nicht installiert, kein Projekt
    gewaehlt, oder direkter Bibliotheksaufruf ohne Projekt), wird der
    Abgleich stillschweigend uebersprungen.
    `ecoinvent_prompt`: optionaler Callback `prompt(label, candidates) ->
    db_name_oder_None` fuer den Fall, dass die 'database'-Spalte einer
    Exchange-Zeile nicht eindeutig einer Datenbank im Projekt zugeordnet
    werden kann (siehe ecoinvent_matcher._resolve_database). Ohne Callback
    bleibt der betroffene Flow dann unresolved statt geraten zu werden.
    `resolved_dir`: Verzeichnis fuer resolved/<database>.yaml (Stufe-2-
    Eingabe). Wird `None` uebergeben, wird kein resolved-Output geschrieben.
    `warnings_out`: Zieldatei fuer die beim Einlesen gesammelten Warnungen
    (siehe excel_loader.load_clusters) - fehlende Pflichtfelder, unbekannter
    Exchange-Typ, fehlende amount/reference product/name usw. Wird `None`
    uebergeben (Default), wird keine Datei geschrieben; die einzelnen
    Meldungen bleiben dann nur als Rueckgabewert erreichbar.
    Gibt (parsed_activities, unresolved) zurueck.
      parsed_activities: {(database, id): bw2io-artiges Aktivitaets-Dict}
      unresolved: Liste nicht aufgeloester Flows (leer = alles sauber)
    """
    registry_data = load_registry(registry)
    exclude = {registry.resolve(), out.resolve()}

    from . import excel_loader  # lazy: bw2io nur noetig, wenn der Abgleich tatsaechlich laeuft
    activities, internal_index, project_parameters, database_parameters, load_warnings = excel_loader.load_clusters(
        clusters, exclude=exclude,
    )

    # Pool aller bekannten Namen fuer Fuzzy-Vorschlaege (einmal vorab bauen)
    candidate_pool = (
        list(internal_index.keys())
        + list(registry_data["technosphere"].keys())
        + list(registry_data["biosphere"].keys())
    )

    unresolved = []
    parsed_activities = {}
    new_registry_entries = {}

    for act in activities:
        exchanges = build_exchanges(
            act, internal_index, registry_data, candidate_pool, unresolved,
            new_entries=new_registry_entries,
            ecoinvent_project=ecoinvent_project if ecoinvent_fallback else None,
            ecoinvent_prompt=ecoinvent_prompt,
        )
        parsed = {
            "name": act["name"],
            "unit": act["unit"],
            "location": act.get("location"),
            "reference product": act["name"],
            "exchanges": exchanges,
        }
        if act.get("comment"):
            parsed["comment"] = act["comment"]
        if act.get("parameters"):
            parsed["parameters"] = act["parameters"]
        parsed_activities[(act["_database"], act["id"])] = parsed

    print(f"{len(activities)} Aktivitaeten geladen, {len(unresolved)} Flow(s) nicht aufgeloest.")

    if warnings_out is not None:
        with open(warnings_out, "w", encoding="utf-8") as f:
            yaml.dump(load_warnings, f, allow_unicode=True, sort_keys=False)
        if load_warnings:
            print(f"-> {len(load_warnings)} Warnung(en) beim Einlesen, Details in {warnings_out}")
        else:
            print(f"-> keine Warnungen beim Einlesen; {warnings_out} geleert.")

    matched_count = sum(len(v) for v in new_registry_entries.values())
    if matched_count:
        _persist_new_registry_entries(registry, new_registry_entries)
        print(f"{matched_count} Flow(s) automatisch gegen ecoinvent abgeglichen und in {registry} ergaenzt.")

    if unresolved:
        with open(out, "w", encoding="utf-8") as f:
            yaml.dump(unresolved, f, allow_unicode=True, sort_keys=False)
        print(f"-> Details in {out}")
        print("Naechste Schritte pro Eintrag:")
        print("  a) Flow-Name in der Cluster-Excel korrigieren (siehe 'suggestions'), oder")
        print("  b) Flow in flow_registry.yaml ergaenzen,")
        print("dann das Skript erneut laufen lassen.")
    else:
        # Sauberer Lauf: alten Report nicht stehen lassen (waere irrefuehrend).
        with open(out, "w", encoding="utf-8") as f:
            yaml.dump([], f)
        print(f"-> keine unresolved Flows; {out} geleert.")

    if resolved_dir is not None:
        written, skipped = _write_resolved_files(
            parsed_activities, unresolved, resolved_dir,
            project_parameters=project_parameters, database_parameters=database_parameters,
        )
        for database in written:
            print(f"-> {resolved_dir / (_safe_filename(database) + '.yaml')} geschrieben (bereit zum Laden).")
        for database in skipped:
            print(f"-> Datenbank '{database}' hat noch unresolved Flows, kein resolved-File geschrieben.")

    return parsed_activities, unresolved


def main():
    """Niedrigschwelliger CLI-Direktaufruf (ohne Workspace-Konzept, siehe Moduldoc oben)."""
    parser = argparse.ArgumentParser(description="LCI-Cluster-Excel (bw2io) -> Testen & Abgleichen (Stufe 1)")
    parser.add_argument(
        "--clusters", type=Path, default=Path("."),
        help="Verzeichnis mit LCI-Cluster-Excel-Dateien (*.xlsx im Brightway2-Excel-Format)",
    )
    parser.add_argument("--registry", type=Path, required=True, help="Pfad zu flow_registry.yaml")
    parser.add_argument("--out", type=Path, default=Path("unresolved.yaml"), help="Ausgabedatei fuer unresolved Flows")
    parser.add_argument(
        "--warnings-out", type=Path, default=Path("load_warnings.yaml"),
        help="Ausgabedatei fuer Warnungen beim Einlesen (fehlende Pflichtfelder, unbekannter Exchange-Typ, ...)",
    )
    parser.add_argument(
        "--resolved-dir", type=Path, default=Path("resolved"),
        help="Ausgabeverzeichnis fuer resolved/<database>.yaml (Stufe-2-Eingabe); leer lassen mit --no-resolved-dir",
    )
    parser.add_argument(
        "--no-resolved-dir", dest="resolved_dir", action="store_const", const=None,
        help="Keinen resolved-Output schreiben",
    )
    parser.add_argument(
        "--no-ecoinvent", dest="ecoinvent_fallback", action="store_false",
        help="Keinen automatischen Abgleich gegen ecoinvent fuer unresolved Flows versuchen",
    )
    parser.add_argument(
        "--ecoinvent-project", default=None,
        help=(
            "Brightway-Projekt fuer den ecoinvent-Abgleich (enthaelt die ecoinvent-/"
            "biosphere-Datenbank). Ohne Angabe wird der Abgleich stillschweigend "
            "uebersprungen - dieser Direktaufruf fragt nicht interaktiv nach (siehe "
            "cli.py/ex2bw-check fuer die interaktive Variante)."
        ),
    )
    args = parser.parse_args()

    run(
        args.clusters, args.registry, args.out,
        ecoinvent_fallback=args.ecoinvent_fallback, resolved_dir=args.resolved_dir,
        warnings_out=args.warnings_out, ecoinvent_project=args.ecoinvent_project,
    )


if __name__ == "__main__":
    main()
