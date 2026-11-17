"""
brightway_writer.py
====================
Schreibt fertig aufgeloeste Aktivitaeten per bw2io/bw2data in eine
Brightway-Datenbank. Enthaelt bewusst KEINE Namensaufloesung - jede
exchange verweist bereits konkret auf (database, code). Aufgerufen wird es
ausschliesslich von core/write.py (nach erneuter Pruefung der Excel-Dateien).

Parameter (Project-/Database-Parameter, 'parameters' je Aktivitaet,
'formula' je exchange) werden ueber bw2io.importers.base_lci.LCIImporter
geschrieben: dieselbe Maschinerie, die bw2io beim direkten Excel-Import
benutzt (ProjectParameter/DatabaseParameter/ActivityParameter/
ParameterizedExchange) - im Activity Browser editier- und neu berechenbar.

Die Ziel-Datenbank wird vollstaendig ersetzt (delete_existing=True). Was
dabei verloren ginge, prueft vorher der Verweis-Test (core/references.py).

WICHTIG: bw2data MUSS exakt in der Major-Version installiert sein, die auch
das Zielprojekt zuletzt geschrieben hat (siehe uv.lock/pyproject.toml).
"""

import re

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


class MissingDependencyError(RuntimeError):
    """Eine referenzierte Datenbank fehlt im Zielprojekt."""


def write_database(project, database, activities, project_parameters=None, database_parameters=None) -> int:
    """
    Schreibt eine Datenbank (ersetzt sie vollstaendig) und berechnet die
    Parameter der geschriebenen Gruppen neu. Gibt die Anzahl Aktivitaeten
    zurueck. Wirft MissingDependencyError, wenn eine referenzierte
    Datenbank im Projekt fehlt (dann wird nichts geschrieben).
    """
    from bw2io.importers.base_lci import LCIImporter
    from bw2data.parameters import ActivityParameter

    bd, active_project = _resolve_project(project)

    dependency_warnings = _check_dependencies(bd, activities, database)
    if dependency_warnings:
        raise MissingDependencyError(
            f"'{database}' wird NICHT geschrieben (Projekt '{active_project}'): " + " ".join(dependency_warnings)
        )

    # Gruppennamen VOR write_database() einsammeln (siehe bw2io.importers.
    # base_lci._prepare_activity_parameters: explizite 'group', sonst Default
    # '<database>:<code>') - danach sind die 'parameters'-Keys nicht mehr da.
    touched_groups = set()
    for act in activities:
        params = act.get("parameters")
        if params:
            touched_groups.add(params[0].get("group") or f"{act['database']}:{act['code']}")

    imp = LCIImporter(database)
    imp.data = activities  # bw2io >=0.9: write_database() liest self.data auch bei explizitem data=
    imp.project_parameters = project_parameters
    imp.database_parameters = database_parameters
    if project_parameters:
        imp.write_project_parameters()
    imp.write_database(data=activities, activate_parameters=True, delete_existing=True)

    # Formel-Exchanges sofort auswerten - gezielt nur fuer die eben geschriebenen
    # Gruppen (der globale recalculate()-Sweep bricht bei verwaisten Gruppen
    # anderer Herkunft im geteilten Projekt ab).
    for group in touched_groups:
        ActivityParameter.recalculate(group)
    return len(activities)
