"""
excel_loader.py
===============
Laedt LCI-Cluster-Excel-Dateien (*.xlsx) im nativen Brightway2-Excel-
Format ueber bw2io (brightway2-io) ein und bildet sie auf dieselbe
Struktur ab wie frueher der handgeschriebene Parser - damit
matcher.build_exchanges() unveraendert weiterlaeuft (gleiche Registry-
Aufloesung, gleicher interner Namensindex, gleiche unresolved-Erkennung,
gleicher ecoinvent-Abgleich).

Das Parsen uebernimmt vollstaendig bw2io.ExcelImporter:

  * Blaetter mit 'skip' in Zelle A1 werden ignoriert (Doku, Rohdaten,
    Aenderungshistorie usw.).
  * 'cutoff' / <n> in Zelle A1 schneidet alle Spalten ab Index <n> ab
    (dort stehen z.B. Fundstellen, die nicht importiert werden sollen).
    Wird dadurch eine von Brightway erkannte Spalte (z.B. 'formula')
    verdeckt, wird das nur als WARNUNG gemeldet - wir tasten weder die
    Excel-Datei noch den Cutoff-Wert an, die Behebung liegt beim Nutzer.
  * Die Abschnitte 'Database', 'Activity' und 'Exchanges' werden erkannt;
    Zahlen werden aus den Zellen uebernommen.
  * Ein 'Parameters'-Block je Aktivitaet sowie 'Database parameters'/
    'Project parameters' werden eingelesen und bis nach Brightway
    durchgereicht (siehe matcher.py/brightway_writer.py) - dort als
    echte, im Activity Browser editierbare bw2data-Parameter geschrieben.

Format einer Cluster-Excel-Datei (siehe bw2io ExcelImporter bzw. die
Vorlagen in tests/fixtures/excel_formats/):

    Database        <Datenbankname>
    (Leerzeile)
    Activity        <Aktivitaetsname>
    code             <id>
    unit             <Einheit>
    location         <Ort>
    production amount  <Menge>
    comment          <Kommentar>              (optional)
    Parameters                                (optional)
    name  amount  formula
    <Zeile pro Parameter>
    Exchanges
    name  amount  unit  database  categories  location  type  comment  formula  ...
    <Zeile pro Exchange>
    (Leerzeile)
    Activity        <naechste Aktivitaet>
    ...

Pro Exchange-Zeile bestimmt die Spalte 'type':
  - 'production'   -> Referenzfluss der Aktivitaet, wird uebersprungen
                       (matcher.build_exchanges baut ihn aus 'output' neu)
  - 'technosphere'  -> Input, Schluessel ist 'reference product'
                       (Fallback: 'name'), wird gegen den internen
                       Namensindex bzw. flow_registry.yaml bzw. ecoinvent
                       aufgeloest
  - 'biosphere'     -> Emission/Aufnahme, Schluessel ist 'name', wird
                       gegen flow_registry.yaml bzw. biosphere3 aufgeloest

'database'/'location'/'comment' aus der Exchange-Zeile werden - sofern
gesetzt - als Override in den Exchange uebernommen. Hat eine Exchange-
Zeile eine 'formula', darf 'amount' bewusst leer sein (der Wert kommt
dann erst beim Schreiben nach Brightway aus der Formel) - das ist kein
Fehlerfall.

Probleme bei einzelnen Aktivitaeten/Exchanges (fehlende Pflichtfelder,
unbekannter Exchange-Typ, fehlende amount/reference product/name, ...)
werden nicht mehr einzeln auf der Konsole ausgegeben, sondern in
load_clusters() gesammelt zurueckgegeben (siehe dort) und pro Datei nur als
Anzahl angezeigt - der Aufrufer (matcher.run) schreibt die Details nach
load_warnings.yaml.
"""

from pathlib import Path

from bw2io.extractors.excel import ExcelExtractor
from bw2io.importers.excel import ExcelImporter

from .matcher import build_internal_index

UNKNOWN_MARKERS = (None, "", "(Unknown)")

# Von bw2io/bw2data erkannte Spaltennamen je Tabellenabschnitt - dient nur
# dazu, in _warn_cutoff_hidden_columns zu erkennen, ob ein cutoff eine fuer
# uns/Brightway relevante Spalte verdeckt (siehe Docstring oben).
_KNOWN_EXCHANGE_FIELDS = {
    "name", "amount", "unit", "database", "categories", "location", "type",
    "uncertainty type", "loc", "scale", "shape", "minimum", "maximum",
    "comment", "formula", "reference product",
}
_KNOWN_PARAMETER_FIELDS = {
    "name", "amount", "formula", "uncertainty type", "loc", "scale", "shape",
    "minimum", "maximum", "comment",
}
_SECTION_FIELDS = {
    "exchanges": _KNOWN_EXCHANGE_FIELDS,
    "parameters": _KNOWN_PARAMETER_FIELDS,
    "database parameters": _KNOWN_PARAMETER_FIELDS,
    "project parameters": _KNOWN_PARAMETER_FIELDS,
}


def _clean(value):
    return None if value in UNKNOWN_MARKERS else value


def _warn_cutoff_hidden_columns(path: Path, warnings: list):
    """
    Rein lesende Pruefung (keine Datei- oder Datenaenderung): meldet pro
    Blatt mit 'cutoff' in Zelle A1, ob eine von bw2io/bw2data erkannte
    Spalte (z.B. 'formula') durch den deklarierten Cutoff-Wert verdeckt
    wird. Die Behebung (Cutoff im Excel erhoehen oder Spalte davor
    verschieben) liegt beim Nutzer - wir greifen nicht ein. Selten (hoechstens
    einmal pro Blatt) und blattuebergreifend relevant, daher weiterhin sofort
    auf der Konsole gemeldet (zusaetzlich zum Eintrag in `warnings`).
    """
    try:
        raw_sheets = ExcelExtractor.extract(str(path))
    except Exception:
        return  # Formatfehler werden an anderer Stelle (ExcelImporter) gemeldet

    for sheet_name, rows in raw_sheets:
        if not rows or not rows[0]:
            continue
        first_cell = rows[0][0]
        if first_cell is None or str(first_cell).strip().lower() != "cutoff":
            continue
        try:
            cutoff = int(rows[0][1])
        except (IndexError, TypeError, ValueError):
            continue

        hidden = {}
        for i, row in enumerate(rows[:-1]):
            marker_cell = row[0] if row else None
            marker = str(marker_cell).strip().lower() if marker_cell is not None else ""
            known_fields = _SECTION_FIELDS.get(marker)
            if known_fields is None:
                continue

            header = rows[i + 1]
            for col_index, cell in enumerate(header):
                if col_index < cutoff or cell is None:
                    continue
                header_name = str(cell).strip().lower()
                if header_name in known_fields:
                    hidden.setdefault(header_name, col_index)

        if hidden:
            cols = ", ".join(
                f"'{name}' (Spalte {index + 1})"
                for name, index in sorted(hidden.items(), key=lambda kv: kv[1])
            )
            message = (
                f"{path.name}: Blatt '{sheet_name}' hat cutoff={cutoff}, dadurch "
                f"werden folgende von Brightway erkannte Spalten NICHT eingelesen: {cols}. "
                f"Falls das nicht gewollt ist: Cutoff im Excel erhoehen oder die Spalte(n) "
                f"davor verschieben."
            )
            print(f"WARNUNG: {message}")
            warnings.append({
                "source_file": path.name, "activity": None, "level": "WARNUNG",
                "kind": "cutoff_hidden_column", "message": message,
            })


def _activity_from_bw2io(raw: dict, source_file: str, warnings: list):
    """
    Baut aus einem bw2io-Aktivitaetsdict (imp.data) das gemeinsame
    Aktivitaets-Dict, das matcher.build_exchanges() erwartet. Gibt None
    zurueck, wenn Pflichtfelder fehlen (Grund landet in `warnings`, siehe
    load_clusters() fuer die pro-Datei-Zusammenfassung auf der Konsole -
    einzelne Exchange-Probleme werden hier NICHT mehr sofort gedruckt,
    das haette bei vielen Fehlern die Konsole zugemuellt).
    """
    name = raw.get("name")
    database = raw.get("database")
    act_id = raw.get("code") or name
    unit = raw.get("unit")
    production_amount = raw.get("production amount", 1)

    missing = [
        field
        for field, value in (("code/name (id)", act_id), ("unit", unit), ("database", database))
        if not value
    ]
    if missing:
        warnings.append({
            "source_file": source_file, "activity": name, "level": "WARNUNG",
            "kind": "missing_required_field",
            "message": f"Aktivitaet '{name}' hat fehlende Pflichtfelder {missing}, wird uebersprungen.",
        })
        return None

    act = {
        "id": act_id,
        "name": name,
        "unit": unit,
        "location": _clean(raw.get("location")),
        "output": {"amount": float(production_amount), "unit": unit},
        "in": {},
        "emit": {},
        "_database": database,
        "_source_file": source_file,
    }

    comment = _clean(raw.get("comment"))
    if comment:
        act["comment"] = comment

    activity_parameters = raw.get("parameters")
    if activity_parameters:
        # bw2io liefert dies als {name: {amount, formula, ...}} - matcher/
        # brightway_writer erwarten eine Liste (bw2data-Parameterformat).
        act["parameters"] = [
            {"name": pname, **pvalue} for pname, pvalue in activity_parameters.items()
        ]

    for row in raw.get("exchanges", []):
        exch_type = str(row.get("type") or "").strip().lower()

        if exch_type == "production":
            continue

        if exch_type not in ("technosphere", "biosphere"):
            warnings.append({
                "source_file": source_file, "activity": name, "level": "WARNUNG",
                "kind": "unknown_exchange_type",
                "message": (
                    f"Exchange-Zeile mit unbekanntem type '{row.get('type')}' "
                    f"in Aktivitaet '{name}' wird uebersprungen."
                ),
            })
            continue

        formula = _clean(row.get("formula"))

        try:
            amount = float(row["amount"])
        except (KeyError, TypeError, ValueError):
            if formula:
                # Amount bewusst leer gelassen: der Wert kommt aus der Formel,
                # sobald sie beim Schreiben nach Brightway ausgewertet wird.
                # Platzhalter wie im bw2io-Beispiel (sample_activities_with_variables.xlsx).
                amount = 0.0
            else:
                warnings.append({
                    "source_file": source_file, "activity": name, "level": "WARNUNG",
                    "kind": "missing_amount",
                    "message": (
                        f"Exchange-Zeile ohne gueltige 'amount' in Aktivitaet "
                        f"'{name}' wird uebersprungen ({row})."
                    ),
                })
                continue

        value = {"amount": amount, "unit": row.get("unit")}
        database_override = _clean(row.get("database"))
        location = _clean(row.get("location"))
        comment = _clean(row.get("comment"))
        if database_override:
            value["database"] = database_override
        if location:
            value["location"] = location
        if comment:
            value["comment"] = comment
        if formula:
            value["formula"] = formula

        if exch_type == "technosphere":
            flow_key = _clean(row.get("reference product")) or _clean(row.get("name"))
            if not flow_key:
                warnings.append({
                    "source_file": source_file, "activity": name, "level": "WARNUNG",
                    "kind": "missing_technosphere_key",
                    "message": (
                        f"technosphere-Exchange ohne 'reference product'/'name' "
                        f"in Aktivitaet '{name}' wird uebersprungen."
                    ),
                })
                continue
            act["in"][flow_key] = value
        else:
            flow_key = _clean(row.get("name"))
            if not flow_key:
                warnings.append({
                    "source_file": source_file, "activity": name, "level": "WARNUNG",
                    "kind": "missing_biosphere_name",
                    "message": (
                        f"biosphere-Exchange ohne 'name' in Aktivitaet '{name}' "
                        f"wird uebersprungen."
                    ),
                })
                continue
            act["emit"][flow_key] = value

    return act


def _parse_workbook(path: Path, warnings: list):
    """
    Liest eine Cluster-Excel-Datei ueber bw2io. Gibt (activities,
    project_parameters, database_parameters) zurueck; die beiden
    Parameter-Listen sind None, wenn die Datei keine entsprechenden
    Bloecke enthaelt (siehe bw2io.ExcelImporter.project_parameters/
    .database_parameters).
    """
    _warn_cutoff_hidden_columns(path, warnings)

    importer = ExcelImporter(str(path))
    activities = []
    for raw in importer.data or []:
        finalized = _activity_from_bw2io(raw, path.name, warnings)
        if finalized is not None:
            activities.append(finalized)
    return activities, importer.project_parameters, importer.database_parameters


def _merge_parameters(target: dict, new_entries, kind_label: str, source_file: str, warnings: list):
    """Fuegt project_parameters/database_parameters aus einer Datei zusammen (name-Konflikt -> Warnung, letzter Wert gewinnt)."""
    if not new_entries:
        return
    for entry in new_entries:
        entry_name = entry.get("name")
        existing = target.get(entry_name)
        if existing is not None and existing != entry:
            warnings.append({
                "source_file": source_file, "activity": None, "level": "WARNUNG",
                "kind": "parameter_conflict",
                "message": (
                    f"{kind_label}-Parameter '{entry_name}' war bereits mit anderem Wert "
                    f"definiert, wird durch diese Datei ueberschrieben."
                ),
            })
        target[entry_name] = entry


def load_clusters(directory: Path, exclude: set = frozenset()):
    """
    Laedt alle Cluster-Excel-Dateien (*.xlsx) aus einem Verzeichnis
    (nur oberste Ebene, nicht rekursiv).
    `exclude` sind aufgeloeste Pfade, die NICHT als Cluster-Datei
    behandelt werden sollen, auch wenn sie im selben Ordner liegen.
    Gibt (activities, internal_index, project_parameters,
    database_parameters, warnings) zurueck. project_parameters/
    database_parameters sind Listen von {name, amount, formula, ...}-Dicts
    (None, wenn keine Datei entsprechende Bloecke hat), ueber alle
    geladenen Dateien zusammengefuehrt.

    `warnings` ist eine flache Liste von {source_file, activity, level,
    kind, message}-Dicts - alle beim Einlesen einzelner Aktivitaeten/
    Exchanges aufgetretenen Probleme (fehlende Pflichtfelder, unbekannter
    Exchange-Typ, fehlende amount/reference product/name, Parameter-
    Konflikte). Diese werden NICHT mehr einzeln auf der Konsole ausgegeben
    (bei vielen Fehlern waere das unuebersichtlich) - stattdessen druckt
    diese Funktion pro Datei nur eine Zusammenfassungszeile mit Anzahl, und
    der Aufrufer (siehe matcher.run) schreibt die volle Liste z.B. nach
    load_warnings.yaml. Datei-weite Probleme (gesperrte/kaputte Datei,
    keine gueltigen Aktivitaeten) sind selten und werden weiterhin sofort
    gedruckt.
    """
    activities = []
    project_parameters = {}
    database_parameters = {}
    warnings = []

    for path in sorted(directory.glob("*.xlsx")):
        if path.name.startswith("~$"):
            # Excel-Sperrdatei einer geoeffneten Arbeitsmappe, keine echte Cluster-Datei
            continue
        if path.resolve() in exclude:
            continue

        warnings_before = len(warnings)
        try:
            file_activities, file_project_params, file_database_params = _parse_workbook(path, warnings)
        except PermissionError:
            message = f"{path.name} ist gerade gesperrt (z.B. in Excel geoeffnet), wird uebersprungen."
            print(f"WARNUNG: {message}")
            warnings.append({
                "source_file": path.name, "activity": None, "level": "WARNUNG",
                "kind": "file_locked", "message": message,
            })
            continue
        except Exception as exc:
            # Formatfehler (bw2io.ExcelImporter, z.B. fehlende/doppelte 'Database'-Zeile,
            # fehlende Spaltenkoepfe in Exchanges, kaputtes 'cutoff') sollen NICHT den
            # gesamten Lauf abbrechen - nur diese eine Datei wird uebersprungen, die
            # uebrigen Excel-Dateien werden trotzdem geprueft.
            message = f"{path.name} hat kein gueltiges Brightway2-Excel-Format ({type(exc).__name__}: {exc}), wird uebersprungen."
            print(f"FEHLER: {message}")
            warnings.append({
                "source_file": path.name, "activity": None, "level": "FEHLER",
                "kind": "invalid_format", "message": message,
            })
            continue

        if not file_activities:
            message = f"{path.name} enthaelt keine (gueltigen) Aktivitaeten, wird uebersprungen."
            print(f"WARNUNG: {message}")
            warnings.append({
                "source_file": path.name, "activity": None, "level": "WARNUNG",
                "kind": "no_activities", "message": message,
            })
            continue

        activities.extend(file_activities)
        _merge_parameters(project_parameters, file_project_params, "Projekt", path.name, warnings)
        _merge_parameters(database_parameters, file_database_params, "Datenbank", path.name, warnings)

        file_warning_count = len(warnings) - warnings_before
        summary = f"{path.name}: {len(file_activities)} Aktivitaet(en) eingelesen"
        if file_warning_count:
            summary += f", {file_warning_count} Warnung(en)"
        print(summary)

    internal_index = build_internal_index(activities)
    return (
        activities,
        internal_index,
        list(project_parameters.values()) or None,
        list(database_parameters.values()) or None,
        warnings,
    )
