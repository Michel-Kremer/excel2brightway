"""
references.py
=============
Verweis-Test: Was passiert beim Schreiben mit der Ziel-Datenbank in Brightway?

Die Excel-Dateien sind der einzige Stand: Eine Datenbank wird beim Schreiben
aus ALLEN Dateien des Ordners gebildet, die sie definieren (auch wenn nicht
alle ausgewaehlt sind), und ersetzt dann die Datenbank in Brightway.

  * HINWEIS 'weitere_datei'   - eine nicht ausgewaehlte Datei definiert
                                dieselbe Datenbank und wird mitgeschrieben.
  * WARNUNG 'wird_geloescht'  - Aktivitaeten stehen in Brightway, aber in
                                keiner Datei; sie wuerden geloescht (Schreiben
                                nur nach ausdruecklicher Bestaetigung).
  * FEHLER  'verweis_bricht'  - andere Datenbanken zeigen auf genau diese
                                Aktivitaeten; ihre Verweise wuerden ins Leere
                                laufen (Schreiben nur nach Bestaetigung).
"""

from dataclasses import dataclass, field

from .model import FEHLER, HINWEIS, TEST_VERWEIS, WARNUNG, Issue


@dataclass
class DatabaseImpact:
    database: str
    files: list  # alle Dateien, die die Datenbank definieren
    unselected_files: list
    planned_codes: set
    exists_in_project: bool = False
    to_delete: list = field(default_factory=list)  # Record
    broken_references: list = field(default_factory=list)  # InboundReference


def database_impacts(selected_models, all_models, backend=None) -> dict:
    """Je Datenbank der ausgewaehlten Dateien: welche Dateien, was wuerde geloescht, welche Verweise brechen."""
    selected_paths = {m.path for m in selected_models}
    impacts = {}
    for database in sorted({m.database for m in selected_models if m.database}):
        files = [m for m in all_models if m.database == database and m.readable]
        planned = {str(a.code) for m in files for a in m.activities if a.name}
        impact = DatabaseImpact(
            database=database,
            files=[m.path for m in files],
            unselected_files=[m.path for m in files if m.path not in selected_paths],
            planned_codes=planned,
        )
        if backend is not None and database in backend.databases():
            impact.exists_in_project = True
            impact.to_delete = [r for r in backend.records(database) if str(r.code) not in planned]
            delete_codes = {str(r.code) for r in impact.to_delete}
            if delete_codes:
                impact.broken_references = [
                    ref for ref in backend.inbound_references(database) if str(ref.target_code) in delete_codes
                ]
        impacts[database] = impact
    return impacts


def _first_model(selected_models, database):
    for model in selected_models:
        if model.database == database:
            return model
    return None


def check_references(selected_models, all_models, backend=None) -> list:
    issues = []
    if backend is None:
        issues.append(Issue(
            test=TEST_VERWEIS, severity=WARNUNG, kind="kein_projekt",
            message="Verweis-Test braucht ein Brightway-Projekt - es wurde nur innerhalb der Dateien geprueft.",
        ))

    for database, impact in database_impacts(selected_models, all_models, backend).items():
        model = _first_model(selected_models, database)
        sheet, row = model.database_location if model and model.database_location else (None, None)
        base = dict(test=TEST_VERWEIS, file=model.path if model else None, sheet=sheet, row=row, column=2)

        for path in impact.unselected_files:
            issues.append(Issue(
                severity=HINWEIS, kind="weitere_datei",
                message=f"Datenbank '{database}' wird auch in '{path.name}' definiert; "
                        f"die Datei wird beim Schreiben automatisch mitgeschrieben.",
                **base,
            ))

        if impact.to_delete:
            names = ", ".join(sorted({r.name or r.code for r in impact.to_delete})[:5])
            more = " ..." if len(impact.to_delete) > 5 else ""
            issues.append(Issue(
                severity=WARNUNG, kind="wird_geloescht",
                message=f"{len(impact.to_delete)} Aktivitaet(en) stehen in Brightway in '{database}', aber in "
                        f"keiner Datei im Ordner ({names}{more}). Beim Schreiben wuerden sie geloescht - "
                        f"nur nach ausdruecklicher Bestaetigung.",
                **base,
            ))

        if impact.broken_references:
            by_target = {r.code: r for r in impact.to_delete}
            examples = []
            for ref in impact.broken_references[:5]:
                target = by_target.get(ref.target_code)
                examples.append(
                    f"{ref.source_database}/{ref.source_name} -> {target.name if target else ref.target_code}"
                )
            more = " ..." if len(impact.broken_references) > 5 else ""
            issues.append(Issue(
                severity=FEHLER, kind="verweis_bricht",
                message=f"{len(impact.broken_references)} Verweis(e) aus anderen Datenbanken zeigen auf Aktivitaeten, "
                        f"die beim Schreiben geloescht wuerden: {'; '.join(examples)}{more}",
                **base,
            ))
    return issues
