"""
format_check.py
===============
Format-Test: prueft, ob die Excel-Dateien dem Brightway-Excel-Format
entsprechen, unabhaengig davon, ob die Fluesse aufloesbar sind.

Was Pflicht und was optional ist, steht in FORMAT_SCHEMA. Das Schema ist
bewusst eine einfache Datenstruktur: Spaeter soll es aus einer Standarddatei
kommen, die zusaetzlich Metadaten beschreibt und Pflicht/optional markiert
(siehe README, "Geplant").

Der Format-Test ist ueberbrueckbar (siehe checks.CheckOptions.bypass_format):
dann blockieren seine FEHLER auf Zeilenebene das Schreiben nicht; solche
Zeilen werden beim Schreiben ausgelassen und gemeldet. Strukturfehler
(NOT_BYPASSABLE: Datei unlesbar, keine/zwei Database-Zeilen, Aktivitaet ohne
Name oder unit, doppelte codes, mehrere production-Zeilen) blockieren immer.
"""

from .model import FEHLER, HINWEIS, TEST_FORMAT, WARNUNG, Issue
from .text import normalize, to_number

# Strukturfehler, die auch bei ueberbruecktem Format-Test blockieren: ohne sie
# laesst sich die Datenbank gar nicht oder nur mit gebrochenen Verweisen bauen.
NOT_BYPASSABLE = frozenset({
    "datei_gesperrt", "datei_unlesbar", "database_fehlt", "database_mehrfach", "database_name_fehlt",
    "aktivitaet_name_fehlt", "pflichtfeld_fehlt", "code_doppelt", "produktion_mehrfach",
})

FORMAT_SCHEMA = {
    # Metadaten einer Aktivitaet (Zeilen unter 'Activity'); 'name' steht in der Activity-Zeile.
    "activity": {
        "required": ["unit"],
        "optional": ["code", "location", "production amount", "reference product", "comment", "categories", "type"],
    },
    # Spalten der Exchanges-Tabelle.
    "exchange_columns": {
        "required": ["name", "amount", "unit", "type"],
        "optional": [
            "database", "location", "categories", "reference product", "comment", "formula",
            "uncertainty type", "loc", "scale", "shape", "minimum", "maximum",
        ],
    },
    # Gueltige Werte der Spalte 'type'.
    "exchange_types": ["production", "technosphere", "biosphere"],
    # Spalten von Parameter-Tabellen (Activity-, Database- und Project-Parameter).
    "parameter_columns": {
        "required": ["name"],
        "optional": ["amount", "formula", "unit", "comment", "uncertainty type", "loc", "scale", "shape", "minimum", "maximum"],
    },
}


def _issue(severity, kind, message, model=None, sheet=None, row=None, column=None, activity=None, exchange=None):
    return Issue(
        test=TEST_FORMAT, severity=severity, kind=kind, message=message,
        file=model.path if model is not None else (exchange.file if exchange is not None else None),
        sheet=sheet, row=row, column=column, activity=activity, exchange=exchange,
    )


def _check_parameters(model, params, label, activity_name=None):
    issues = []
    for param in params:
        name = param.get("name")
        sheet, row = param.get("_sheet"), param.get("_row")
        if not name:
            issues.append(_issue(
                FEHLER, "parameter_name_fehlt", f"{label}: Parameter ohne Namen.",
                model=model, sheet=sheet, row=row, activity=activity_name,
            ))
            continue
        if to_number(param.get("amount")) is None and not param.get("formula"):
            issues.append(_issue(
                FEHLER, "parameter_wert_fehlt",
                f"{label}: Parameter '{name}' hat weder eine Zahl in 'amount' noch eine 'formula'.",
                model=model, sheet=sheet, row=row, activity=activity_name,
            ))
    return issues


def check_exchange_row(exchange, schema=FORMAT_SCHEMA):
    """Prueft eine einzelne exchange-Zeile. Gibt eine Liste von Issues zurueck."""
    issues = []
    activity_label = None
    exch_type = exchange.type
    valid_types = schema["exchange_types"]

    if exch_type is None:
        issues.append(_issue(
            FEHLER, "typ_fehlt", "exchange-Zeile ohne 'type'.",
            sheet=exchange.sheet, row=exchange.row, column=exchange.columns.get("type"), exchange=exchange,
        ))
    elif exch_type not in valid_types:
        issues.append(_issue(
            FEHLER, "typ_unbekannt",
            f"Unbekannter type '{exchange.values.get('type')}' (erlaubt: {', '.join(valid_types)}).",
            sheet=exchange.sheet, row=exchange.row, column=exchange.columns.get("type"), exchange=exchange,
        ))

    if not exchange.get("name") and not exchange.get("reference product"):
        issues.append(_issue(
            FEHLER, "name_fehlt", "exchange-Zeile ohne 'name'.",
            sheet=exchange.sheet, row=exchange.row, column=exchange.columns.get("name"), exchange=exchange,
        ))

    if to_number(exchange.values.get("amount")) is None and not exchange.get("formula"):
        raw = exchange.values.get("amount")
        detail = "ist leer" if raw is None else f"ist keine Zahl ({raw!r})"
        issues.append(_issue(
            FEHLER, "menge_ungueltig",
            f"'amount' {detail} und es gibt keine 'formula'.",
            sheet=exchange.sheet, row=exchange.row, column=exchange.columns.get("amount"), exchange=exchange,
        ))

    if not exchange.get("unit") and exch_type in ("technosphere", "biosphere"):
        issues.append(_issue(
            FEHLER, "einheit_fehlt", "exchange-Zeile ohne 'unit'.",
            sheet=exchange.sheet, row=exchange.row, column=exchange.columns.get("unit"), exchange=exchange,
        ))

    for issue in issues:
        issue.activity = activity_label
    return issues


def check_format(models, schema=FORMAT_SCHEMA, scope=None) -> list:
    """
    Format-Test ueber eingelesene Dateien (WorkbookModel). `scope` ist eine
    Menge von Dateipfaden, fuer die Meldungen erzeugt werden (None = alle);
    die uebrigen Dateien zaehlen nur fuer dateiuebergreifende Pruefungen
    (doppelte codes).
    Enthaelt auch die Meldungen, die schon beim Einlesen entstanden sind.
    """
    issues = []
    in_scope = (lambda m: True) if scope is None else (lambda m: m.path in scope)

    # Doppelte (database, code) ueber alle Dateien
    seen_codes = {}
    seen_names = {}
    for model in models:
        for activity in model.activities:
            if not activity.name or not model.database:
                continue
            key = (model.database, activity.code)
            seen_codes.setdefault(key, []).append(activity)
            name_key = (model.database, normalize(activity.name), normalize(activity.location))
            seen_names.setdefault(name_key, []).append(activity)

    for model in models:
        if not in_scope(model):
            continue
        issues.extend(model.issues)
        if not model.readable:
            continue

        _check = _check_parameters
        issues.extend(_check(model, model.project_parameters, "Project parameters"))
        issues.extend(_check(model, model.database_parameters, "Database parameters"))

        for activity in model.activities:
            label = activity.name
            for field_name in schema["activity"]["required"]:
                if activity.fields.get(field_name) in (None, ""):
                    issues.append(_issue(
                        FEHLER, "pflichtfeld_fehlt",
                        f"Aktivitaet '{label}' hat kein '{field_name}'.",
                        model=model, sheet=activity.sheet, row=activity.row, activity=label,
                    ))

            production_amount = activity.fields.get("production amount")
            if production_amount is not None and to_number(production_amount) is None:
                issues.append(_issue(
                    FEHLER, "produktionsmenge_ungueltig",
                    f"Aktivitaet '{label}': 'production amount' ist keine Zahl ({production_amount!r}).",
                    model=model, sheet=activity.sheet, row=activity.field_rows.get("production amount"), activity=label,
                ))

            if model.database:
                duplicates = seen_codes.get((model.database, activity.code), [])
                if len(duplicates) > 1:
                    others = [
                        f"{d.file.name}/{d.sheet} Zeile {d.row}" for d in duplicates if d is not activity
                    ]
                    issues.append(_issue(
                        FEHLER, "code_doppelt",
                        f"Aktivitaet '{label}': code '{activity.code}' ist in Datenbank '{model.database}' "
                        f"mehrfach vergeben (auch: {', '.join(others)}).",
                        model=model, sheet=activity.sheet, row=activity.field_rows.get("code", activity.row),
                        activity=label,
                    ))
                same_name = seen_names.get((model.database, normalize(activity.name), normalize(activity.location)), [])
                if len(same_name) > 1 and len(duplicates) <= 1:
                    issues.append(_issue(
                        WARNUNG, "name_doppelt",
                        f"Aktivitaetsname '{label}' ({activity.location or 'ohne location'}) kommt in Datenbank "
                        f"'{model.database}' mehrfach vor; Verweise ueber den Namen sind dann mehrdeutig.",
                        model=model, sheet=activity.sheet, row=activity.row, activity=label,
                    ))

            issues.extend(_check(model, activity.parameters, f"Parameters von '{label}'", label))

            if activity.exchange_header_row is None:
                issues.append(_issue(
                    HINWEIS, "keine_exchanges",
                    f"Aktivitaet '{label}' hat keinen 'Exchanges'-Abschnitt; es wird nur der Referenzfluss geschrieben.",
                    model=model, sheet=activity.sheet, row=activity.row, activity=label,
                ))
                continue

            missing_columns = [
                c for c in schema["exchange_columns"]["required"] if c not in activity.exchange_columns
            ]
            # 'name' darf durch 'reference product' ersetzt sein
            if "name" in missing_columns and "reference product" in activity.exchange_columns:
                missing_columns.remove("name")
            if missing_columns:
                issues.append(_issue(
                    FEHLER, "spalte_fehlt",
                    f"Exchanges von '{label}': Pflichtspalte(n) fehlen: {', '.join(missing_columns)}.",
                    model=model, sheet=activity.sheet, row=activity.exchange_header_row, activity=label,
                ))

            production = activity.production_rows
            if len(production) > 1:
                issues.append(_issue(
                    FEHLER, "produktion_mehrfach",
                    f"Aktivitaet '{label}' hat {len(production)} production-Zeilen; unterstuetzt wird genau eine.",
                    model=model, sheet=activity.sheet, row=production[1].row, activity=label,
                ))
            elif not production:
                issues.append(_issue(
                    HINWEIS, "produktion_fehlt",
                    f"Aktivitaet '{label}' hat keine production-Zeile; der Referenzfluss wird aus "
                    f"'production amount' (Standard 1) und 'unit' gebildet.",
                    model=model, sheet=activity.sheet, row=activity.exchange_header_row, activity=label,
                ))

            for exchange in activity.exchanges:
                for issue in check_exchange_row(exchange, schema):
                    issue.activity = label
                    issues.append(issue)

    return issues
