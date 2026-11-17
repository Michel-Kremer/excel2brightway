"""
apply.py
========
Uebernehmen: schreibt die im Pruefblatt gewaehlten Kandidaten in die
exchange-Zeilen der Eingabe-Excel.

  * Vorher wird geprueft, ob sich die Daten seit dem Pruefen veraendert haben
    (Pruefsumme). Wenn ja, wird abgebrochen - bitte erneut pruefen.
  * Vor dem Speichern wird eine Sicherung in den Unterordner
    'ex2bw_sicherung/' gelegt (nicht in den Eingabeordner selbst, sonst wuerde
    sie beim naechsten Lauf als Eingabe gelesen).
  * Geschrieben werden name, reference product (technosphere), categories
    (biosphere), location und database. Fehlt eine dieser Spalten in der
    Exchanges-Tabelle, wird sie hinten angefuegt (sofern ein cutoff das
    zulaesst). Die Einheit wird NIE geaendert: passt sie nicht zum
    Kandidaten, wird die Zeile nicht uebernommen (Menge muss von Hand
    umgerechnet werden).
  * Aktion 'Proxy' traegt zusaetzlich einen Vermerk in die Spalte 'comment' ein.
  * Jede Aenderung landet im Blatt '_ex2bw_Aenderungen' (alt -> neu).
  * Danach werden Pruefblatt und Markierungen entfernt; das naechste Pruefen
    erzeugt sie fuer verbleibende Meldungen neu.
"""

import shutil
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

from .review_sheet import (
    clear_marks, read_review, remove_review_sheets, save_workbook,
)
from .text import clean, normalize_unit
from .workbook import data_hash, read_workbook

CHANGELOG_SHEET = "_ex2bw_Aenderungen"
BACKUP_DIRNAME = "ex2bw_sicherung"
CHANGELOG_COLUMNS = [
    "Zeitpunkt", "Blatt", "Zeile", "Aktivitaet", "Aktion", "Feld", "alt", "neu", "Kommentar",
]


class DataChangedError(RuntimeError):
    """Die Daten wurden seit dem Pruefen veraendert."""


@dataclass
class ApplyResult:
    path: Path
    applied: list = field(default_factory=list)  # (sheet, row, beschreibung)
    skipped: list = field(default_factory=list)  # (sheet, row, grund)
    backup: Path = None

    @property
    def changed(self) -> bool:
        return bool(self.applied)


def backup_file(path) -> Path:
    path = Path(path)
    folder = path.parent / BACKUP_DIRNAME
    folder.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    target = folder / f"{path.stem}_{stamp}{path.suffix}"
    shutil.copy2(path, target)
    return target


def _ensure_column(ws, activity, column_name, cutoff):
    """Spaltennummer einer Exchanges-Spalte; legt sie bei Bedarf hinten an. None, wenn cutoff es verhindert."""
    if column_name in activity.exchange_columns:
        return activity.exchange_columns[column_name]
    header_row = activity.exchange_header_row
    last = max(activity.exchange_columns.values()) if activity.exchange_columns else 0
    # erste wirklich freie Spalte in der Kopfzeile suchen
    col = last + 1
    while ws.cell(row=header_row, column=col).value not in (None, ""):
        col += 1
    if cutoff is not None and col > cutoff:
        return None
    ws.cell(row=header_row, column=col, value=column_name)
    activity.exchange_columns[column_name] = col
    return col


def _changelog_sheet(wb):
    if CHANGELOG_SHEET in wb.sheetnames:
        return wb[CHANGELOG_SHEET]
    ws = wb.create_sheet(CHANGELOG_SHEET)
    ws.cell(row=1, column=1, value="skip")
    ws.cell(row=1, column=2, value="ex2bw-Aenderungsprotokoll").font = Font(bold=True)
    for col, name in enumerate(CHANGELOG_COLUMNS, start=1):
        ws.cell(row=2, column=col, value=name).font = Font(bold=True)
    return ws


def _target_values(record, activity_columns, exchange_type):
    values = {"name": record.name, "location": record.location, "database": record.database}
    if exchange_type == "biosphere":
        values.pop("location")
        values["categories"] = "::".join(record.categories) if record.categories else None
    else:
        product = record.product or record.name
        if "reference product" in activity_columns or product != record.name:
            values["reference product"] = product
    return values


def apply_review(path, force=False) -> ApplyResult:
    """
    Uebernimmt die Auswahl aus dem Pruefblatt. `force=True` ueberspringt die
    Pruefsummen-Kontrolle (nur fuer Ausnahmefaelle).
    """
    path = Path(path)
    result = ApplyResult(path=path)
    state = read_review(path)
    if state is None:
        return result

    if not force and state.data_hash and data_hash(path) != state.data_hash:
        raise DataChangedError(
            f"{path.name}: Die Daten wurden seit dem Pruefen veraendert. Bitte zuerst erneut pruefen."
        )

    selections = [r for r in state.rows if r.choice and r.choice in state.candidates]
    for row in state.rows:
        if row.choice and row.choice not in state.candidates and not row.choice.startswith("("):
            result.skipped.append((row.sheet, row.row, f"Auswahl '{row.choice}' ist kein gueltiger Kandidat"))
    if not selections:
        return result

    model = read_workbook(path)
    exchanges = {(e.sheet, e.row): e for e in model.exchange_rows()}

    wb = openpyxl.load_workbook(path)
    log = None
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    applied_positions = set()

    for selection in selections:
        record = state.candidates[selection.choice]
        exchange = exchanges.get((selection.sheet, selection.row))
        if exchange is None:
            result.skipped.append((selection.sheet, selection.row, "exchange-Zeile nicht mehr vorhanden"))
            continue
        activity = model.activities[exchange.activity_index]
        unit = exchange.get("unit")
        if unit and record.unit and normalize_unit(unit) != normalize_unit(record.unit):
            result.skipped.append((
                selection.sheet, selection.row,
                f"Einheit '{unit}' passt nicht zu '{record.unit}' - Menge und Einheit bitte von Hand anpassen",
            ))
            continue

        action = (selection.action or "ersetzen").strip()
        ws = wb[exchange.sheet]
        cutoff = model.cutoffs.get(exchange.sheet)
        if log is None:
            log = _changelog_sheet(wb)

        changes = []
        for column_name, new_value in _target_values(record, activity.exchange_columns, exchange.type).items():
            old_value = clean(exchange.values.get(column_name))
            if (old_value or None) == (new_value or None):
                continue
            col = _ensure_column(ws, activity, column_name, cutoff)
            if col is None:
                result.skipped.append((
                    selection.sheet, selection.row,
                    f"Spalte '{column_name}' fehlt und kann wegen cutoff nicht angelegt werden",
                ))
                continue
            ws.cell(row=exchange.row, column=col, value=new_value)
            changes.append((column_name, old_value, new_value))

        if action.lower() == "proxy":
            col = _ensure_column(ws, activity, "comment", cutoff)
            if col is not None:
                old_comment = clean(ws.cell(row=exchange.row, column=col).value)
                note = f"Proxy (ex2bw): {selection.comment}" if selection.comment else "Proxy (ex2bw)"
                new_comment = f"{old_comment}; {note}" if old_comment else note
                ws.cell(row=exchange.row, column=col, value=new_comment)
                changes.append(("comment", old_comment, new_comment))

        for column_name, old_value, new_value in changes:
            log.append([
                stamp, exchange.sheet, exchange.row, activity.name, action, column_name,
                old_value, new_value, selection.comment,
            ])
        if changes:
            applied_positions.add((exchange.sheet, exchange.row))
            result.applied.append((
                exchange.sheet, exchange.row,
                f"{activity.name}: -> {record.name} ({record.location or '-'}, {record.database})",
            ))

    if not result.applied:
        wb.close()
        return result

    result.backup = backup_file(path)
    clear_marks(wb)
    remove_review_sheets(wb)
    save_workbook(wb, path)
    wb.close()
    return result


def summarize(result: ApplyResult) -> str:
    lines = [f"{result.path.name}: {len(result.applied)} Zeile(n) uebernommen, {len(result.skipped)} uebersprungen."]
    for sheet, row, text in result.applied:
        lines.append(f"  + {sheet} Zeile {row}: {text}")
    for sheet, row, reason in result.skipped:
        lines.append(f"  - {sheet} Zeile {row}: {reason}")
    if result.backup:
        lines.append(f"  Sicherung: {result.backup}")
    return "\n".join(lines)


__all__ = ["apply_review", "ApplyResult", "DataChangedError", "summarize"]
