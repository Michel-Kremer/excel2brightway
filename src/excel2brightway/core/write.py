"""
write.py
========
Nach Brightway schreiben - ohne Zwischendateien. Die Excel-Dateien sind der
einzige Stand; unmittelbar vor dem Schreiben wird alles erneut geprueft.

Ablauf (plan_write -> execute_write):
  1. Je Datenbank der Auswahl werden ALLE Dateien des Ordners genommen, die
     sie definieren (nicht ausgewaehlte werden automatisch ergaenzt). Die
     Datenbank in Brightway wird aus genau diesen Dateien ersetzt.
  2. Alle Tests laufen erneut (Format ueberbrueckbar). Bei Fehlern wird nicht
     geschrieben; das Pruefblatt zeigt, was fehlt.
  3. Aktivitaeten, die in Brightway stehen, aber in keiner Datei, werden nur
     mit allow_delete=True geloescht; zeigen andere Datenbanken darauf, ist
     zusaetzlich allow_broken_references=True noetig.
  4. execute_write() schreibt ueber bw2io/bw2data (brightway_writer.py).
"""

from dataclasses import dataclass, field, replace
from pathlib import Path

from .checks import CheckOptions, run_checks
from .model import ALL_TESTS
from .references import database_impacts
from .text import split_categories, to_number

UNCERTAINTY_FIELDS = ("uncertainty type", "loc", "scale", "shape", "minimum", "maximum")
_ACTIVITY_FIELDS_HANDLED = {"code", "unit", "location", "production amount", "reference product", "comment", "name"}


class WriteBlockedError(RuntimeError):
    """Schreiben ist (noch) nicht erlaubt - Grund steht in der Meldung."""


@dataclass
class WritePlan:
    database: str
    files: list
    auto_added: list
    activities: list
    project_parameters: list
    database_parameters: list
    to_delete: list = field(default_factory=list)
    broken_references: list = field(default_factory=list)
    skipped: list = field(default_factory=list)  # Texte zu ausgelassenen Zeilen (nur mit Ueberbrueckung)
    exists_in_project: bool = False


def _clean_parameter(param):
    result = {k: v for k, v in param.items() if not k.startswith("_") and v is not None}
    amount = to_number(result.get("amount"))
    if amount is not None:
        result["amount"] = amount
    elif "amount" in result:
        del result["amount"]
    for key in UNCERTAINTY_FIELDS:
        if key in result:
            number = to_number(result[key])
            if number is None:
                del result[key]
            else:
                result[key] = int(number) if key == "uncertainty type" else number
    return result


def _exchange_fields(exchange, activity_unit=None):
    """amount/formula/unit/comment/Unsicherheit einer exchange-Zeile; None, wenn amount fehlt."""
    amount = to_number(exchange.values.get("amount"))
    formula = exchange.get("formula")
    if amount is None:
        if not formula:
            return None
        amount = 0.0  # Wert kommt beim Schreiben aus der Formel
    data = {"amount": amount, "unit": exchange.get("unit") or activity_unit}
    if formula:
        data["formula"] = str(formula)
    if exchange.get("comment"):
        data["comment"] = exchange.get("comment")
    for key in UNCERTAINTY_FIELDS:
        number = to_number(exchange.values.get(key))
        if number is not None:
            data[key] = int(number) if key == "uncertainty type" else number
    return data


def build_activities(models, resolutions, database):
    """
    Baut die Aktivitaeten einer Datenbank im Format von
    bw2io.importers.base_lci.LCIImporter (input = (database, code)).
    Gibt (activities, skipped) zurueck.
    """
    activities, skipped = [], []
    for model in models:
        if model.database != database:
            continue
        for activity in model.activities:
            if not activity.name or not activity.unit:
                skipped.append(f"{model.path.name}/{activity.sheet} Zeile {activity.row}: Aktivitaet ohne name/unit")
                continue
            code = str(activity.code)
            exchanges = []
            production_rows = activity.production_rows
            if production_rows:
                prod = _exchange_fields(production_rows[0], activity.unit)
                if prod is None:
                    prod = {"amount": to_number(activity.fields.get("production amount")) or 1.0, "unit": activity.unit}
            else:
                prod = {"amount": to_number(activity.fields.get("production amount")) or 1.0, "unit": activity.unit}
            prod.update({"type": "production", "input": (database, code)})
            exchanges.append(prod)

            for exchange in activity.exchanges:
                if exchange.type not in ("technosphere", "biosphere"):
                    continue
                where = f"{model.path.name}/{exchange.sheet} Zeile {exchange.row}"
                resolution = resolutions.get((model.path, exchange.sheet, exchange.row))
                if resolution is None:
                    skipped.append(f"{where}: Fluss nicht aufgeloest")
                    continue
                data = _exchange_fields(exchange)
                if data is None:
                    skipped.append(f"{where}: ohne gueltige amount/formula")
                    continue
                if not data.get("unit"):
                    skipped.append(f"{where}: ohne unit")
                    continue
                data.update({"type": exchange.type, "input": resolution.record.key})
                exchanges.append(data)

            node = {
                "database": database,
                "code": code,
                "name": activity.name,
                "unit": activity.unit,
                "location": activity.location,
                "reference product": activity.reference_product,
                "exchanges": exchanges,
            }
            if activity.fields.get("comment"):
                node["comment"] = activity.fields["comment"]
            for key, value in activity.fields.items():
                if key in _ACTIVITY_FIELDS_HANDLED or value is None:
                    continue
                node[key] = split_categories(value) if key == "categories" else value
            if activity.parameters:
                node["parameters"] = [_clean_parameter(p) for p in activity.parameters if p.get("name")]
            activities.append(node)
    return activities, skipped


def _merged_parameters(models, attribute):
    merged = {}
    for model in models:
        for param in getattr(model, attribute):
            if param.get("name"):
                merged[param["name"]] = _clean_parameter(param)
    return list(merged.values())


def plan_write(paths, backend, options: CheckOptions = None, allow_broken_references=False):
    """
    Prueft erneut und plant das Schreiben. Gibt (report, plans) zurueck;
    `plans` ist leer, wenn noch blockierende Meldungen offen sind.
    """
    options = replace(options or CheckOptions(), tests=ALL_TESTS, write_review=True)
    first = run_checks(paths, backend=backend, options=options)
    selected_models = [first.models[p] for p in first.selected if p in first.models]
    databases = sorted({m.database for m in selected_models if m.database})

    # Alle Dateien derselben Datenbank in die Pruefung nehmen
    extra = [
        m.path for m in first.models.values()
        if m.database in databases and m.path not in first.selected
    ]
    report = run_checks(list(first.selected) + extra, backend=backend, options=options) if extra else first

    blocking = [
        i for i in report.blocking_issues()
        if not (i.kind == "verweis_bricht" and allow_broken_references)
    ]
    if blocking or report.errors:
        return report, []

    all_models = list(report.models.values())
    selected_models = [report.models[p] for p in report.selected if p in report.models]
    impacts = database_impacts(selected_models, all_models, backend)
    project_parameters = _merged_parameters(selected_models, "project_parameters")
    plans = []
    for database in databases:
        models = [m for m in selected_models if m.database == database]
        activities, skipped = build_activities(models, report.resolutions, database)
        impact = impacts[database]
        plans.append(WritePlan(
            database=database,
            files=[m.path for m in models],
            auto_added=[p for p in extra if report.models[p].database == database],
            activities=activities,
            project_parameters=project_parameters,
            database_parameters=_merged_parameters(models, "database_parameters"),
            to_delete=impact.to_delete,
            broken_references=impact.broken_references,
            skipped=skipped,
            exists_in_project=impact.exists_in_project,
        ))
    try:
        plans = write_order(plans)  # Anzeige in Schreib-Reihenfolge
    except WriteBlockedError:
        pass  # Kreisverweis: execute_write meldet ihn, bevor etwas geschrieben wird
    return report, plans


def assert_allowed(plans, allow_delete=False, allow_broken_references=False):
    """Wirft WriteBlockedError, wenn Loeschungen oder gebrochene Verweise nicht freigegeben sind."""
    for plan in plans:
        if plan.broken_references and not allow_broken_references:
            raise WriteBlockedError(
                f"'{plan.database}': {len(plan.broken_references)} Verweis(e) aus anderen Datenbanken wuerden "
                f"ins Leere laufen. Schreiben nur mit ausdruecklicher Freigabe."
            )
        if plan.to_delete and not allow_delete:
            raise WriteBlockedError(
                f"'{plan.database}': {len(plan.to_delete)} Aktivitaet(en) in Brightway stehen in keiner Datei "
                f"und wuerden geloescht. Schreiben nur mit ausdruecklicher Freigabe."
            )


def write_order(plans):
    """
    Reihenfolge zum Schreiben: Datenbanken, auf die andere geplante Datenbanken
    verweisen, zuerst (sonst fehlt die Abhaengigkeit noch im Projekt). Bei
    gleichem Rang alphabetisch. Verweisen Datenbanken im Kreis aufeinander,
    wird nichts geschrieben (WriteBlockedError).
    """
    by_name = {plan.database: plan for plan in plans}
    needs = {
        plan.database: {
            exchange["input"][0]
            for activity in plan.activities
            for exchange in activity.get("exchanges", [])
            if exchange.get("input") and exchange["input"][0] in by_name and exchange["input"][0] != plan.database
        }
        for plan in plans
    }
    ordered, done = [], set()
    while len(ordered) < len(plans):
        ready = sorted(name for name in by_name if name not in done and needs[name] <= done)
        if not ready:
            cycle = sorted(name for name in by_name if name not in done)
            details = "; ".join(f"{name} -> {', '.join(sorted(needs[name] - done))}" for name in cycle)
            raise WriteBlockedError(
                f"Die Datenbanken {', '.join(cycle)} verweisen im Kreis aufeinander ({details}). "
                "Das kann noch nicht in einem Schritt geschrieben werden - nichts wurde geschrieben."
            )
        for name in ready:
            ordered.append(by_name[name])
            done.add(name)
    return ordered


def execute_write(plans, project, allow_delete=False, allow_broken_references=False):
    """
    Schreibt die geplanten Datenbanken in Abhaengigkeits-Reihenfolge (write_order).
    Gibt [(database, anzahl_aktivitaeten), ...] zurueck.
    """
    assert_allowed(plans, allow_delete=allow_delete, allow_broken_references=allow_broken_references)
    plans = write_order(plans)
    from .. import brightway_writer

    written = []
    for plan in plans:
        count = brightway_writer.write_database(
            project=project,
            database=plan.database,
            activities=plan.activities,
            project_parameters=plan.project_parameters or None,
            database_parameters=plan.database_parameters or None,
        )
        written.append((plan.database, count))
    return written


def describe_plan(plan) -> str:
    lines = [f"Datenbank '{plan.database}': {len(plan.activities)} Aktivitaet(en) aus {len(plan.files)} Datei(en)"]
    for path in plan.auto_added:
        lines.append(f"  + automatisch mitgenommen: {Path(path).name}")
    if plan.to_delete:
        lines.append(f"  ! {len(plan.to_delete)} Aktivitaet(en) in Brightway wuerden geloescht")
    if plan.broken_references:
        lines.append(f"  ! {len(plan.broken_references)} Verweis(e) aus anderen Datenbanken wuerden brechen")
    for text in plan.skipped:
        lines.append(f"  - ausgelassen: {text}")
    return "\n".join(lines)


__all__ = [
    "plan_write", "execute_write", "assert_allowed", "describe_plan", "build_activities",
    "WritePlan", "WriteBlockedError",
]
