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
from pathlib import Path

import yaml

PROJECT_NAME = "ecoinvent-3.12-cutoff"


def _ensure_project():
    import bw2data as bd

    if bd.projects.current != PROJECT_NAME:
        bd.projects.set_current(PROJECT_NAME)
    return bd


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


def write_to_brightway(paths) -> list:
    """
    Schreibt eine Liste resolved/<database>.yaml-Dateien nach Brightway.
    Gibt eine Liste (database, Anzahl_Aktivitaeten) der geschriebenen
    Datenbanken zurueck.
    """
    from bw2io.importers.base_lci import LCIImporter
    from bw2data.parameters import ActivityParameter

    _ensure_project()

    written = []
    touched_groups = set()
    for path in paths:
        database, activities, project_parameters, database_parameters = load_resolved(Path(path))

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

    return written


def main():
    parser = argparse.ArgumentParser(description="resolved/*.yaml -> Brightway (Stufe 2)")
    parser.add_argument("paths", type=Path, nargs="+", help="Eine oder mehrere resolved/<database>.yaml")
    args = parser.parse_args()

    write_to_brightway(args.paths)


if __name__ == "__main__":
    main()
