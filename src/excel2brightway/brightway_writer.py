"""
brightway_writer.py
====================
Stufe 2 von 2: Laden in Brightway. Liest ausschliesslich resolved/*.yaml
(Output von matcher.py) und schreibt sie per bw2data in die jeweilige
Ziel-Datenbank. Enthaelt bewusst KEINE Namens-Aufloesung und keinen
ecoinvent-Abgleich mehr - jede Aktivitaet/jeder Exchange in einer
resolved-Datei referenziert bereits konkret (database, code).

Parameter (project_parameters/database_parameters auf oberster Ebene,
'parameters' je Aktivitaet, 'formula' je Exchange - siehe excel_loader.py/
matcher.py) werden ueber bw2io.importers.base_lci.LCIImporter geschrieben:
das ist dieselbe, bereits getestete Maschinerie, die bw2io beim direkten
Excel-Import benutzt (ProjectParameter/DatabaseParameter/ActivityParameter/
ParameterizedExchange, inkl. Formel-Neuberechnung) - im Activity Browser
editierbar/neu berechenbar, nicht nur einmalig statisch ausgewertet.

WICHTIG: bw2data MUSS exakt in der Version installiert sein, die auch das
Zielprojekt zuletzt geschrieben hat (siehe requirements-lock.txt).

Fuer den normalen Gebrauch siehe den Konsolenbefehl ex2bw-load (cli.py).
Niedrigschwelliger Direktaufruf ohne Workspace-Konzept:

    python -m excel2brightway.brightway_writer resolved/methanol-cluster.yaml [weitere ...]
"""

import argparse
import re
from pathlib import Path

import yaml

_VERSIONED_DB_RE = re.compile(r"(?i)^(?P<family>[a-z]+)[-_ ]?(?P<version>\d+(?:[.\-_]\d+)*)")


def _check_dependencies(bd, activities: list, database: str) -> list:
    """
    Prueft VOR dem Schreiben, ob alle von 'activities' referenzierten
    externen Datenbanken (z.B. 'ecoinvent-3.12-cutoff', 'biosphere3') im
    aktuell aktiven Brightway-Projekt tatsaechlich vorhanden sind - ohne
    das wuerde bw2data erst mitten im Schreibvorgang mit einer rohen
    UnknownObject-Exception abbrechen (und eine halb geschriebene
    Datenbank hinterlassen).
    Gibt eine Liste Warnmeldungen zurueck (leer = alle Abhaengigkeiten da).
    Bei einer fehlenden versionierten Datenbank (z.B. 'ecoinvent-3.12-
    cutoff') wird zusaetzlich im Projekt nach anderen Versionen derselben
    Familie gesucht (z.B. 'ecoinvent-3.10-cutoff'), um einen Versions-
    Mismatch explizit zu benennen statt nur "fehlt" zu melden.
    """
    referenced = {
        exch["input"][0]
        for act in activities
        for exch in act.get("exchanges", [])
        if exch["input"][0] != database
    }
    if not referenced:
        return []

    existing = set(bd.databases)
    warnings = []
    for dep in sorted(referenced):
        if dep in existing:
            continue

        match = _VERSIONED_DB_RE.match(dep)
        family = match.group("family").lower() if match else None
        same_family = sorted(
            name for name in existing
            if name != dep
            and (m2 := _VERSIONED_DB_RE.match(name))
            and family is not None
            and m2.group("family").lower() == family
        )
        if same_family:
            warnings.append(
                f"Datenbank '{dep}' fehlt im Projekt '{bd.projects.current}' - "
                f"stattdessen vorhanden: {', '.join(same_family)} (Versions-Mismatch)."
            )
        else:
            warnings.append(f"Datenbank '{dep}' fehlt im Projekt '{bd.projects.current}'.")

    return warnings


def _resolve_project(project_name=None):
    """
    Aktiviert das Ziel-Brightway-Projekt fuer den Schreibvorgang.
    project_name=None -> aktuell aktiviertes Projekt (bd.projects.current)
    wird unveraendert verwendet. Sonst wird geprueft, dass project_name
    unter den vorhandenen Projekten existiert (kein versehentliches
    Neuanlegen bei Tippfehlern - bw2data.set_current() wuerde das sonst
    stillschweigend tun) und dorthin gewechselt.
    Gibt (bw2data-Modul, tatsaechlich aktiver Projektname) zurueck.
    """
    import bw2data as bd

    if project_name is None:
        return bd, bd.projects.current

    existing = {p.name for p in bd.projects}
    if project_name not in existing:
        raise ValueError(
            f"Brightway-Projekt '{project_name}' existiert nicht. "
            f"Vorhandene Projekte: {', '.join(sorted(existing))}"
        )
    if bd.projects.current != project_name:
        bd.projects.set_current(project_name)
    return bd, project_name


def load_resolved(path: Path):
    """
    Liest eine resolved/<database>.yaml und baut das Format, das
    bw2io.importers.base_lci.LCIImporter.write_database() erwartet: eine
    Liste von Aktivitaets-Dicts (je mit 'database'/'code'), Exchanges mit
    'input' als (db, code)-Tupel. Production-Exchanges ohne expliziten
    input_database/input_code bekommen automatisch 'input' = die Aktivitaet
    selbst.
    Gibt (database_name, activities, project_parameters, database_parameters)
    zurueck; die beiden Parameter-Listen sind None, wenn die resolved-Datei
    keine entsprechenden Abschnitte hat.
    """
    with open(path, encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}

    database = raw.get("database")
    if not database:
        raise ValueError(f"{path}: kein 'database'-Feld, keine gueltige resolved-Datei.")

    activities = []
    for act in raw.get("activities", []):
        code = act["code"]

        exchanges = []
        for exch in act.get("exchanges", []):
            exch_type = exch["type"]
            if exch_type == "production":
                input_database = exch.get("input_database", database)
                input_code = exch.get("input_code", code)
            else:
                input_database = exch["input_database"]
                input_code = exch["input_code"]

            exchange = {
                "type": exch_type,
                "amount": exch["amount"],
                "unit": exch.get("unit"),
                "input": (input_database, input_code),
            }
            if exch.get("comment"):
                exchange["comment"] = exch["comment"]
            if exch.get("formula"):
                exchange["formula"] = exch["formula"]
            exchanges.append(exchange)

        activity = {
            "database": database,
            "code": code,
            "name": act["name"],
            "unit": act["unit"],
            "location": act.get("location"),
            "reference product": act.get("reference product", act["name"]),
            "exchanges": exchanges,
        }
        if act.get("comment"):
            activity["comment"] = act["comment"]
        if act.get("parameters"):
            activity["parameters"] = act["parameters"]
        activities.append(activity)

    return database, activities, raw.get("project_parameters"), raw.get("database_parameters")


def write_to_brightway(paths, project_name=None) -> list:
    """
    Schreibt eine Liste resolved/<database>.yaml-Dateien nach Brightway.
    project_name=None -> aktuell aktiviertes Brightway-Projekt (Default).
    Gibt eine Liste (database, Anzahl_Aktivitaeten) der geschriebenen
    Datenbanken zurueck.
    """
    from bw2io.importers.base_lci import LCIImporter
    from bw2data.parameters import ActivityParameter

    bd, active_project = _resolve_project(project_name)

    written = []
    touched_groups = set()
    for path in paths:
        path = Path(path)
        database, activities, project_parameters, database_parameters = load_resolved(path)

        dependency_warnings = _check_dependencies(bd, activities, database)
        if dependency_warnings:
            print(f"\nWARNUNG - {path.name} in Projekt '{active_project}':")
            for msg in dependency_warnings:
                print(f"  - {msg}")
            print(f"  '{database}' wird NICHT geschrieben (fehlende Abhaengigkeiten wuerden den Schreibvorgang abbrechen).")
            continue

        # Gruppennamen VOR write_database() einsammeln (siehe bw2io.importers.
        # base_lci._prepare_activity_parameters: explizite 'group' aus dem
        # Excel-Parameters-Block, sonst Default '<database>:<code>') - danach
        # sind die 'parameters'-Keys nicht mehr da: write_database() poppt sie
        # aus genau diesen (per Referenz identischen) Aktivitaets-Dicts heraus.
        for act in activities:
            params = act.get("parameters")
            if params:
                touched_groups.add(params[0].get("group") or f"{act['database']}:{act['code']}")

        imp = LCIImporter(database)
        imp.project_parameters = project_parameters
        imp.database_parameters = database_parameters
        if project_parameters:
            imp.write_project_parameters()
        imp.write_database(data=activities, activate_parameters=True, delete_existing=True)

        print(f"{database}: {len(activities)} Aktivitaet(en) nach Brightway geschrieben.")
        written.append((database, len(activities)))

    # LCIImporter.write_database() registriert Formel-Exchanges (ParameterizedExchange),
    # wertet sie an dieser Stelle aber noch nicht aus - die betroffene Gruppe bleibt
    # 'expired' und der Exchange behaelt seinen Platzhalter-amount, bis jemand
    # recalculate() fuer diese Gruppe aufruft (sonst z.B. erst beim naechsten Oeffnen
    # im Activity Browser). Das erledigen wir hier sofort, damit resolved/*.yaml und
    # der tatsaechliche Brightway-Stand nicht auseinanderlaufen - GEZIELT nur fuer die
    # Gruppen, die wir selbst gerade geschrieben haben (nicht der globale
    # bw2data.parameters.recalculate()-Sweep ueber das ganze Projekt: der bricht ab,
    # sobald irgendwo im geteilten Projekt eine unabhaengige, verwaiste Parameter-Gruppe
    # liegt - reales Beispiel: eine leere, 'expired' Gruppe ohne jeden Parameter/
    # ParameterizedExchange, vermutlich Altlast von ausserhalb dieses Tools).
    for group in touched_groups:
        ActivityParameter.recalculate(group)

    if written:
        print(f"\n-> Geschrieben in Brightway-Projekt: '{active_project}'")
    else:
        print(f"\n-> Nichts geschrieben in Brightway-Projekt '{active_project}' (siehe Warnungen oben).")

    return written


def main():
    parser = argparse.ArgumentParser(description="resolved/*.yaml -> Brightway (Stufe 2)")
    parser.add_argument("paths", type=Path, nargs="+", help="Eine oder mehrere resolved/<database>.yaml")
    parser.add_argument(
        "--project", default=None,
        help="Ziel-Brightway-Projekt (Default: aktuell aktiviertes Projekt)",
    )
    args = parser.parse_args()

    write_to_brightway(args.paths, project_name=args.project)


if __name__ == "__main__":
    main()
