"""
workbook.py
===========
Liest Excel-Dateien im nativen Brightway-Excel-Format (wie bw2io.ExcelImporter)
direkt mit openpyxl ein - mit Zellpositionen fuer jede Aktivitaet und jede
exchange-Zeile. Dieselbe Struktur wird fuer alle Tests UND fuer das Schreiben
nach Brightway verwendet: geschrieben wird damit genau das, was geprueft wurde.

Unterstuetzte Konventionen (siehe tests/fixtures/excel_formats/):

  * 'skip' in Zelle A1: Blatt wird ignoriert. Blaetter, deren Name mit
    '_ex2bw_' beginnt (Pruefblatt, Kandidaten, Aenderungsprotokoll), werden
    immer ignoriert.
  * 'cutoff' in A1 und eine Zahl n in B1: nur die Spalten 1..n werden gelesen.
  * Abschnitte (Markierung in Spalte A, Gross-/Kleinschreibung egal):
      Database <name>            danach Metadaten (key/value) bis Leerzeile
      Project parameters         Kopfzeile, dann Parameter bis Leerzeile
      Database parameters        Kopfzeile, dann Parameter bis Leerzeile
      Activity <name>            danach Metadaten (key/value)
      Parameters [<gruppe>]      (in einer Aktivitaet) Kopfzeile, dann Parameter
      Exchanges                  (in einer Aktivitaet) Kopfzeile, dann Zeilen
    Eine Tabelle endet an der ersten Leerzeile. Eine Aktivitaet endet erst an
    der naechsten Markierung 'Activity'/'Database'/'... parameters' oder am
    Blattende (Leerzeilen zwischen Metadaten, Parameters und Exchanges sind
    erlaubt). Freitext ausserhalb von Abschnitten wird ignoriert.
  * Eine Datei hat genau eine 'Database'-Zeile (blattuebergreifend).
"""

import hashlib
from pathlib import Path

import openpyxl

from .model import FEHLER, TEST_FORMAT, WARNUNG, ActivityBlock, ExchangeRow, Issue, WorkbookModel
from .text import clean

EX2BW_SHEET_PREFIX = "_ex2bw_"

_MARKER_DATABASE = "database"
_MARKER_ACTIVITY = "activity"
_MARKER_EXCHANGES = "exchanges"
_MARKER_PARAMETERS = "parameters"
_MARKER_DB_PARAMETERS = "database parameters"
_MARKER_PROJECT_PARAMETERS = "project parameters"
_MARKERS = {
    _MARKER_DATABASE, _MARKER_ACTIVITY, _MARKER_EXCHANGES, _MARKER_PARAMETERS,
    _MARKER_DB_PARAMETERS, _MARKER_PROJECT_PARAMETERS,
}

# Spalten, die Brightway in den Tabellen auswertet (fuer die cutoff-Warnung).
KNOWN_EXCHANGE_COLUMNS = {
    "name", "amount", "unit", "database", "categories", "location", "type",
    "uncertainty type", "loc", "scale", "shape", "minimum", "maximum",
    "comment", "formula", "reference product",
}
KNOWN_PARAMETER_COLUMNS = {
    "name", "amount", "formula", "unit", "uncertainty type", "loc", "scale", "shape",
    "minimum", "maximum", "comment",
}


def is_ex2bw_sheet(title: str) -> bool:
    return str(title).startswith(EX2BW_SHEET_PREFIX)


def _marker(value):
    if not isinstance(value, str):
        return None
    key = " ".join(value.strip().lower().split())
    return key if key in _MARKERS else None


def _is_blank(row) -> bool:
    return all(clean(v) is None for v in row)


def _format_issue(model, severity, kind, message, sheet=None, row=None, column=None, activity=None):
    model.issues.append(Issue(
        test=TEST_FORMAT, severity=severity, kind=kind, message=message,
        file=model.path, sheet=sheet, row=row, column=column, activity=activity,
    ))


def _read_cutoff(model, sheet_title, rows):
    """Gibt (cutoff_oder_None, erste_datenzeile_index) zurueck."""
    if not rows or not rows[0]:
        return None, 0
    first = rows[0][0]
    if not (isinstance(first, str) and first.strip().lower() == "cutoff"):
        return None, 0
    raw = rows[0][1] if len(rows[0]) > 1 else None
    try:
        cutoff = int(raw)
    except (TypeError, ValueError):
        cutoff = None
    if cutoff is None or cutoff <= 0:
        _format_issue(
            model, WARNUNG, "cutoff_ungueltig",
            f"Blatt '{sheet_title}': 'cutoff' in A1 braucht in B1 eine Zahl > 0 (gefunden: {raw!r}); "
            f"es werden alle Spalten gelesen.",
            sheet=sheet_title, row=1,
        )
        return None, 1
    return cutoff, 1


def _warn_hidden_columns(model, sheet_title, rows, cutoff):
    """Meldet, wenn der cutoff eine von Brightway ausgewertete Spalte verdeckt."""
    for index, row in enumerate(rows[:-1]):
        marker = _marker(row[0]) if row else None
        if marker == _MARKER_EXCHANGES:
            known = KNOWN_EXCHANGE_COLUMNS
        elif marker in (_MARKER_PARAMETERS, _MARKER_DB_PARAMETERS, _MARKER_PROJECT_PARAMETERS):
            known = KNOWN_PARAMETER_COLUMNS
        else:
            continue
        header = rows[index + 1]
        hidden = []
        for col_index, cell in enumerate(header):
            if col_index < cutoff or cell is None:
                continue
            name = str(cell).strip().lower()
            if name in known:
                hidden.append(f"'{name}' (Spalte {col_index + 1})")
        if hidden:
            _format_issue(
                model, WARNUNG, "cutoff_verdeckt_spalte",
                f"Blatt '{sheet_title}' hat cutoff={cutoff}, dadurch werden folgende von Brightway "
                f"erkannte Spalten NICHT eingelesen: {', '.join(hidden)}. Falls ungewollt: cutoff erhoehen "
                f"oder die Spalte(n) davor verschieben.",
                sheet=sheet_title, row=index + 2,
            )


def _header_columns(row) -> dict:
    columns = {}
    for col_index, cell in enumerate(row, start=1):
        cell = clean(cell)
        if cell is None:
            continue
        name = " ".join(str(cell).strip().lower().split())
        columns.setdefault(name, col_index)
    return columns


def _row_values(row, columns: dict) -> dict:
    values = {}
    for name, col_index in columns.items():
        values[name] = clean(row[col_index - 1]) if col_index - 1 < len(row) else None
    return values


def _parameter_from_values(values: dict) -> dict:
    param = {k: v for k, v in values.items() if v is not None}
    return param


class _SheetParser:
    """Zustandsautomat fuer ein Blatt. Aktivitaeten/Parameter landen direkt im Modell."""

    def __init__(self, model: WorkbookModel, sheet_title: str):
        self.model = model
        self.sheet = sheet_title
        self.activity = None  # aktuelle ActivityBlock
        self.state = None  # None | 'dbmeta' | 'actmeta' | 'table'
        self.table = None  # 'exchanges' | 'parameters' | 'db_parameters' | 'project_parameters'
        self.table_columns = None  # None = Kopfzeile erwartet
        self.table_header_row = None

    def _end_table(self):
        if self.state == "table" and self.table_columns is None:
            # Markierung ohne Kopfzeile (direkt danach Leerzeile/Blattende)
            _format_issue(
                self.model, FEHLER, "kopfzeile_fehlt",
                f"Nach '{self._table_label()}' fehlt die Kopfzeile.",
                sheet=self.sheet, row=self.table_header_row,
                activity=self.activity.name if self.activity else None,
            )
        self.state = "actmeta_done" if self.activity is not None else None
        self.table = None
        self.table_columns = None

    def _table_label(self):
        return {
            "exchanges": "Exchanges", "parameters": "Parameters",
            "db_parameters": "Database parameters", "project_parameters": "Project parameters",
        }.get(self.table, self.table)

    def _start_table(self, table, row_number):
        self.state = "table"
        self.table = table
        self.table_columns = None
        self.table_header_row = row_number + 1

    def feed(self, row_number: int, row: tuple):
        if _is_blank(row):
            if self.state == "table":
                self._end_table()
            elif self.state == "dbmeta":
                self.state = None
            elif self.state == "actmeta":
                self.state = "actmeta_done"
            return

        marker = _marker(row[0]) if row else None

        if marker == _MARKER_DATABASE:
            if self.state == "table":
                self._end_table()
            self.activity = None
            name = clean(row[1]) if len(row) > 1 else None
            if self.model.database_location is not None:
                first_sheet, first_row = self.model.database_location
                _format_issue(
                    self.model, FEHLER, "database_mehrfach",
                    f"Zweite 'Database'-Zeile gefunden (erste: Blatt '{first_sheet}', Zeile {first_row}). "
                    f"Eine Datei darf genau eine Datenbank definieren.",
                    sheet=self.sheet, row=row_number,
                )
            else:
                self.model.database = str(name) if name is not None else None
                self.model.database_location = (self.sheet, row_number)
                if name is None:
                    _format_issue(
                        self.model, FEHLER, "database_name_fehlt",
                        "Die 'Database'-Zeile hat keinen Namen in Spalte B.",
                        sheet=self.sheet, row=row_number, column=2,
                    )
            self.state = "dbmeta"
            return

        if marker in (_MARKER_DB_PARAMETERS, _MARKER_PROJECT_PARAMETERS):
            if self.state == "table":
                self._end_table()
            self.activity = None
            self._start_table("db_parameters" if marker == _MARKER_DB_PARAMETERS else "project_parameters", row_number)
            return

        if marker == _MARKER_ACTIVITY:
            if self.state == "table":
                self._end_table()
            name = clean(row[1]) if len(row) > 1 else None
            self.activity = ActivityBlock(
                file=self.model.path, sheet=self.sheet, row=row_number,
                name=str(name) if name is not None else None, database=None,
            )
            self.model.activities.append(self.activity)
            if name is None:
                _format_issue(
                    self.model, FEHLER, "aktivitaet_name_fehlt",
                    "'Activity'-Zeile ohne Namen in Spalte B.",
                    sheet=self.sheet, row=row_number, column=2,
                )
            self.state = "actmeta"
            return

        if marker in (_MARKER_EXCHANGES, _MARKER_PARAMETERS):
            if self.state == "table":
                self._end_table()
            if self.activity is None:
                _format_issue(
                    self.model, FEHLER, "abschnitt_ohne_aktivitaet",
                    f"'{row[0]}' steht ausserhalb einer Aktivitaet und wird ignoriert.",
                    sheet=self.sheet, row=row_number,
                )
                self.state = None
                return
            if marker == _MARKER_EXCHANGES:
                if self.activity.exchange_header_row is not None:
                    _format_issue(
                        self.model, FEHLER, "exchanges_mehrfach",
                        f"Aktivitaet '{self.activity.name}' hat mehr als einen 'Exchanges'-Abschnitt.",
                        sheet=self.sheet, row=row_number, activity=self.activity.name,
                    )
                self._start_table("exchanges", row_number)
            else:
                group = clean(row[1]) if len(row) > 1 else None
                if group is not None:
                    self.activity.parameter_group = str(group)
                self._start_table("parameters", row_number)
            return

        if self.state == "table":
            self._feed_table(row_number, row)
            return

        if self.state == "dbmeta":
            key = clean(row[0])
            if key is not None:
                self.model.database_metadata[str(key).strip()] = clean(row[1]) if len(row) > 1 else None
            return

        if self.state in ("actmeta", "actmeta_done") and self.activity is not None:
            key = clean(row[0])
            if key is None:
                return
            if self.state == "actmeta_done":
                # Metadaten nach einer Leerzeile: zaehlen weiterhin zur Aktivitaet
                self.state = "actmeta"
            key = " ".join(str(key).strip().lower().split())
            self.activity.fields[key] = clean(row[1]) if len(row) > 1 else None
            self.activity.field_rows[key] = row_number
            return

        # Freitext ausserhalb von Abschnitten (z.B. Erlaeuterungen) wird ignoriert.

    def _feed_table(self, row_number, row):
        if self.table_columns is None:
            self.table_columns = _header_columns(row)
            self.table_header_row = row_number
            if self.table == "exchanges":
                self.activity.exchange_header_row = row_number
                self.activity.exchange_columns = dict(self.table_columns)
            return

        values = _row_values(row, self.table_columns)
        if self.table == "exchanges":
            self.activity.exchanges.append(ExchangeRow(
                file=self.model.path, sheet=self.sheet, row=row_number,
                columns=dict(self.table_columns), values=values,
                activity_index=len(self.model.activities) - 1,
            ))
        elif self.table == "parameters":
            param = _parameter_from_values(values)
            param["_row"] = row_number
            param["_sheet"] = self.sheet
            self.activity.parameters.append(param)
        else:
            param = _parameter_from_values(values)
            param["_row"] = row_number
            param["_sheet"] = self.sheet
            target = self.model.database_parameters if self.table == "db_parameters" else self.model.project_parameters
            target.append(param)

    def finish(self):
        if self.state == "table":
            self._end_table()


def read_workbook(path) -> WorkbookModel:
    """
    Liest eine Excel-Datei. Fehler beim Oeffnen (gesperrt, kaputt) werden nicht
    geworfen, sondern als FEHLER-Meldung im Modell abgelegt (readable=False).
    """
    path = Path(path)
    model = WorkbookModel(path=path)
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    except PermissionError:
        model.readable = False
        _format_issue(
            model, FEHLER, "datei_gesperrt",
            f"{path.name} ist gesperrt (z.B. in Excel geoeffnet) und kann nicht gelesen werden.",
        )
        return model
    except Exception as exc:  # kaputte oder keine xlsx-Datei
        model.readable = False
        _format_issue(
            model, FEHLER, "datei_unlesbar",
            f"{path.name} kann nicht als Excel-Datei gelesen werden ({type(exc).__name__}: {exc}).",
        )
        return model

    try:
        for ws in wb.worksheets:
            if is_ex2bw_sheet(ws.title):
                continue
            rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
            if not rows:
                continue
            first = rows[0][0] if rows[0] else None
            if isinstance(first, str) and first.strip().lower() == "skip":
                continue

            cutoff, start = _read_cutoff(model, ws.title, rows)
            if cutoff is not None:
                model.cutoffs[ws.title] = cutoff
                _warn_hidden_columns(model, ws.title, rows, cutoff)
                rows = [row[:cutoff] for row in rows]

            parser = _SheetParser(model, ws.title)
            for index in range(start, len(rows)):
                parser.feed(index + 1, rows[index])
            parser.finish()
    finally:
        wb.close()

    for activity in model.activities:
        activity.database = model.database

    if model.database_location is None and model.activities:
        _format_issue(
            model, FEHLER, "database_fehlt",
            f"{path.name} hat Aktivitaeten, aber keine 'Database'-Zeile.",
        )
    if not model.activities:
        _format_issue(
            model, WARNUNG, "keine_aktivitaeten",
            f"{path.name} enthaelt keine Aktivitaeten.",
        )
    return model


def data_hash(path_or_workbook) -> str:
    """
    Pruefsumme ueber alle Daten-Blaetter (ohne '_ex2bw_'-Blaetter), berechnet
    aus den gespeicherten Zellinhalten. Dient beim Uebernehmen dazu, zu
    erkennen, ob die Daten seit dem Pruefen veraendert wurden (das Ausfuellen
    des Pruefblatts selbst zaehlt nicht). Nimmt einen Pfad oder eine bereits
    geoeffnete openpyxl-Arbeitsmappe.
    """
    own = not hasattr(path_or_workbook, "worksheets")
    wb = openpyxl.load_workbook(Path(path_or_workbook)) if own else path_or_workbook
    digest = hashlib.sha256()
    try:
        for ws in wb.worksheets:
            if is_ex2bw_sheet(ws.title):
                continue
            digest.update(f"#{ws.title}\n".encode("utf-8"))
            for row in ws.iter_rows(values_only=True):
                if all(v is None for v in row):
                    continue
                digest.update(repr(tuple(v for v in row)).encode("utf-8"))
                digest.update(b"\n")
    finally:
        if own:
            wb.close()
    return digest.hexdigest()


LCI_MARKER = "cutoff"


def is_lci_file(path) -> bool:
    """
    True, wenn mindestens ein Blatt in A1 'cutoff' hat - so beginnen die
    LCI-Blaetter mit den Aktivitaeten. Andere Excel-Dateien im Ordner
    (Recherche, Rechentools, ...) werden damit nicht eingelesen.
    Nicht lesbare Dateien gelten als LCI-Datei, damit der Lesefehler gemeldet
    und die Datei nicht still uebergangen wird.
    """
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
    except Exception:
        return True
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows(min_row=1, max_row=1, max_col=1, values_only=True):
                value = row[0] if row else None
                if isinstance(value, str) and value.strip().lower() == LCI_MARKER:
                    return True
        return False
    finally:
        wb.close()


def list_excel_files(folder, skipped=None) -> list:
    """
    LCI-Dateien (*.xlsx mit 'cutoff' in A1 eines Blatts, siehe is_lci_file)
    eines Ordners, ohne Excel-Sperrdateien (~$...). Uebergangene Dateien
    werden an `skipped` (Liste) angehaengt, falls angegeben.
    """
    folder = Path(folder)
    result = []
    for path in sorted(p for p in folder.glob("*.xlsx") if not p.name.startswith("~$") and p.is_file()):
        if is_lci_file(path):
            result.append(path)
        elif skipped is not None and path.resolve() not in skipped:
            skipped.append(path.resolve())
    return result
