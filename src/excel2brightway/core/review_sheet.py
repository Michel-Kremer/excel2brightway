"""
review_sheet.py
===============
Pruefblatt in der Eingabe-Excel:

  _ex2bw_Pruefung    eine Zeile je Meldung (Format, Fluss, Verweis). Fluss-
                     Meldungen haben eine Auswahlliste mit Kandidaten, eine
                     Spalte "Eigene Suche" und eine Aktion (ersetzen / Proxy).
  _ex2bw_Kandidaten  ausgeblendet: die Kandidaten hinter den Auswahllisten
                     plus Pruefsumme der Daten beim Pruefen.

Beide Blaetter haben 'skip' in A1, damit Brightway/bw2io sie nie als Daten
liest. In den Datenblaettern wird die betroffene Zelle jeder Meldung rot
markiert und bekommt einen Kommentar (Autor 'ex2bw'); beim naechsten Pruefen
bzw. Uebernehmen werden diese Markierungen wieder entfernt.
"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink

from .model import FEHLER, HINWEIS, TEST_FLUSS, TEST_LABELS, WARNUNG, Record
from .resolve import STATUS_LABELS
from .workbook import data_hash as compute_data_hash
from .workbook import is_ex2bw_sheet

REVIEW_SHEET = "_ex2bw_Pruefung"
CANDIDATE_SHEET = "_ex2bw_Kandidaten"
COMMENT_AUTHOR = "ex2bw"
MARK_COLOR = "FFFFC7CE"
ACTIONS = ("ersetzen", "Proxy")
FORMAT_VERSION = 1

COLUMNS = [
    ("Nr", 6), ("Test", 9), ("Schwere", 10), ("Stelle", 30), ("Aktivitaet", 30),
    ("Fluss (Original)", 34), ("Einheit", 10), ("Datenbank (Excel)", 20), ("Problem", 50),
    ("Auswahl", 70), ("Eigene Suche", 24), ("Aktion", 11), ("Kommentar", 30),
    ("Blatt", 14), ("Zeile", 7), ("Art", 22),
]
COL = {name: index for index, (name, _) in enumerate(COLUMNS, start=1)}
HEADER_ROW = 3
FIRST_ROW = 4

CANDIDATE_COLUMNS = [
    "Nr", "Kandidat", "Auswahltext", "database", "code", "name", "reference product",
    "location", "unit", "categories", "kind", "origin", "score",
]
CCOL = {name: index for index, name in enumerate(CANDIDATE_COLUMNS, start=1)}
CANDIDATE_FIRST_ROW = 3

_SEVERITY_FILL = {
    FEHLER: PatternFill("solid", fgColor="FFFFC7CE"),
    WARNUNG: PatternFill("solid", fgColor="FFFFEB9C"),
    HINWEIS: PatternFill("solid", fgColor="FFDDEBF7"),
}
_TEST_ORDER = {"format": 0, "fluss": 1, "verweis": 2}


class WorkbookLockedError(RuntimeError):
    """Die Datei ist (z.B. in Excel) geoeffnet und kann nicht gespeichert werden."""


def save_workbook(wb, path):
    try:
        wb.save(path)
    except PermissionError as exc:
        raise WorkbookLockedError(
            f"{Path(path).name} kann nicht gespeichert werden - bitte die Datei in Excel schliessen."
        ) from exc


def candidate_label(number, record: Record, score=None) -> str:
    origin = f"[Datei] {record.database}" if record.origin == "datei" else record.database
    if record.kind == "biosphere":
        parts = [str(number), origin, record.name, "::".join(record.categories) or "-", record.unit or "-"]
    else:
        name = record.name
        if record.product and record.product != record.name:
            name = f"{record.name} | {record.product}"
        parts = [str(number), origin, name, record.location or "-", record.unit or "-"]
    if score is not None:
        parts.append(f"{min(score, 1.0):.2f}")
    return " · ".join(parts)


@dataclass
class ReviewRow:
    nr: int
    test: str
    severity: str
    sheet: Optional[str]
    row: Optional[int]
    kind: str
    flow: Optional[str]
    choice: Optional[str] = None
    search: Optional[str] = None
    action: Optional[str] = None
    comment: Optional[str] = None


@dataclass
class ReviewState:
    data_hash: Optional[str]
    rows: list = field(default_factory=list)
    candidates: dict = field(default_factory=dict)  # Auswahltext -> Record

    def by_position(self) -> dict:
        return {(r.sheet, r.row): r for r in self.rows if r.sheet and r.row}


# --------------------------------------------------------------- Markierungen
def clear_marks(wb):
    """Entfernt alle von ex2bw gesetzten Kommentare und Markierungsfarben in den Datenblaettern."""
    for ws in wb.worksheets:
        if is_ex2bw_sheet(ws.title):
            continue
        for row in ws.iter_rows():
            for cell in row:
                comment = cell.comment
                if comment is not None and comment.author == COMMENT_AUTHOR:
                    cell.comment = None
                    fill = cell.fill
                    if fill is not None and fill.fgColor is not None and fill.fgColor.rgb == MARK_COLOR:
                        cell.fill = PatternFill(fill_type=None)


def remove_review_sheets(wb):
    for title in (REVIEW_SHEET, CANDIDATE_SHEET):
        if title in wb.sheetnames:
            del wb[title]


def _mark(wb, issue):
    if not issue.sheet or not issue.row or issue.sheet not in wb.sheetnames:
        return
    ws = wb[issue.sheet]
    cell = ws.cell(row=issue.row, column=issue.column or 1)
    text = f"[{TEST_LABELS.get(issue.test, issue.test)}] {issue.message}"
    if cell.comment is not None and cell.comment.author == COMMENT_AUTHOR:
        text = cell.comment.text + "\n" + text
    cell.comment = Comment(text, COMMENT_AUTHOR, width=360, height=140)
    if issue.severity == FEHLER:
        cell.fill = PatternFill("solid", fgColor=MARK_COLOR)


# ----------------------------------------------------------------- Lesen
def read_review(path_or_wb) -> Optional[ReviewState]:
    """Liest Pruefblatt und Kandidaten. None, wenn die Datei kein Pruefblatt hat."""
    own = not hasattr(path_or_wb, "sheetnames")
    wb = openpyxl.load_workbook(path_or_wb) if own else path_or_wb
    try:
        if REVIEW_SHEET not in wb.sheetnames:
            return None
        ws = wb[REVIEW_SHEET]
        data_hash = None
        candidates = {}
        if CANDIDATE_SHEET in wb.sheetnames:
            cs = wb[CANDIDATE_SHEET]
            data_hash = cs.cell(row=1, column=3).value
            for r in range(CANDIDATE_FIRST_ROW, cs.max_row + 1):
                label = cs.cell(row=r, column=CCOL["Auswahltext"]).value
                if not label:
                    continue
                categories = cs.cell(row=r, column=CCOL["categories"]).value
                candidates[label] = Record(
                    database=cs.cell(row=r, column=CCOL["database"]).value,
                    code=str(cs.cell(row=r, column=CCOL["code"]).value),
                    name=cs.cell(row=r, column=CCOL["name"]).value,
                    product=cs.cell(row=r, column=CCOL["reference product"]).value,
                    location=cs.cell(row=r, column=CCOL["location"]).value,
                    unit=cs.cell(row=r, column=CCOL["unit"]).value,
                    categories=tuple(p for p in str(categories).split("::") if p) if categories else (),
                    kind=cs.cell(row=r, column=CCOL["kind"]).value or "technosphere",
                    origin=cs.cell(row=r, column=CCOL["origin"]).value or "brightway",
                )
        rows = []
        for r in range(FIRST_ROW, ws.max_row + 1):
            nr = ws.cell(row=r, column=COL["Nr"]).value
            if nr is None:
                continue

            def value(name, _r=r):
                v = ws.cell(row=_r, column=COL[name]).value
                return v.strip() if isinstance(v, str) else v

            rows.append(ReviewRow(
                nr=int(nr), test=value("Test"), severity=value("Schwere"),
                sheet=value("Blatt"), row=int(value("Zeile")) if value("Zeile") else None,
                kind=value("Art"), flow=value("Fluss (Original)"),
                choice=value("Auswahl") or None, search=value("Eigene Suche") or None,
                action=value("Aktion") or None, comment=value("Kommentar") or None,
            ))
        return ReviewState(data_hash=data_hash, rows=rows, candidates=candidates)
    finally:
        if own:
            wb.close()


# ----------------------------------------------------------------- Schreiben
def _sorted(issues):
    return sorted(
        issues,
        key=lambda i: (_TEST_ORDER.get(i.test, 9), i.sheet or "", i.row or 0, i.kind),
    )


def write_review(path, issues, previous: Optional[ReviewState] = None):
    """
    Schreibt das Pruefblatt fuer eine Datei (ersetzt ein vorhandenes) und setzt
    die Markierungen. Ohne Meldungen werden Pruefblatt und Markierungen nur
    entfernt. Eintraege des Nutzers aus `previous` (Eigene Suche, Kommentar,
    Aktion, noch gueltige Auswahl) bleiben je Zelle erhalten. Die Pruefsumme
    der Daten wird aus genau dem gespeicherten Stand berechnet.
    """
    path = Path(path)
    wb = openpyxl.load_workbook(path)
    data_hash = compute_data_hash(wb)
    clear_marks(wb)
    remove_review_sheets(wb)

    if not issues:
        save_workbook(wb, path)
        wb.close()
        return

    kept = previous.by_position() if previous is not None else {}

    ws = wb.create_sheet(REVIEW_SHEET, 0)
    cs = wb.create_sheet(CANDIDATE_SHEET)
    cs.sheet_state = "hidden"

    ws.cell(row=1, column=1, value="skip")
    ws.cell(row=1, column=2, value="ex2bw-Pruefblatt").font = Font(bold=True, size=13)
    ws.cell(row=2, column=2, value=(
        "Fluss-Meldungen: Kandidat unter 'Auswahl' waehlen (oder Suchbegriff unter 'Eigene Suche' eintragen "
        "und erneut pruefen), Datei speichern und schliessen, dann 'Uebernehmen'. "
        "Aktion leer = ersetzen; 'Proxy' kennzeichnet die Zuordnung zusaetzlich im Kommentar."
    ))
    for col, (name, width) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=HEADER_ROW, column=col, value=name)
        cell.font = Font(bold=True)
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width
    ws.freeze_panes = ws.cell(row=FIRST_ROW, column=1)

    cs.cell(row=1, column=1, value="skip")
    cs.cell(row=1, column=2, value="hash")
    cs.cell(row=1, column=3, value=data_hash)
    cs.cell(row=1, column=4, value="version")
    cs.cell(row=1, column=5, value=FORMAT_VERSION)
    for col, name in enumerate(CANDIDATE_COLUMNS, start=1):
        cs.cell(row=2, column=col, value=name).font = Font(bold=True)

    action_validation = DataValidation(type="list", formula1='"' + ",".join(ACTIONS) + '"', allow_blank=True)
    ws.add_data_validation(action_validation)

    candidate_row = CANDIDATE_FIRST_ROW
    for nr, issue in enumerate(_sorted(issues), start=1):
        r = FIRST_ROW + nr - 1
        exchange = issue.exchange
        values = {
            "Nr": nr,
            "Test": TEST_LABELS.get(issue.test, issue.test),
            "Schwere": issue.severity,
            "Stelle": issue.location_label() or "(Datei)",
            "Aktivitaet": issue.activity,
            "Fluss (Original)": (exchange.get("reference product") or exchange.get("name")) if exchange else None,
            "Einheit": exchange.get("unit") if exchange else None,
            "Datenbank (Excel)": exchange.get("database") if exchange else None,
            "Problem": f"{STATUS_LABELS.get(issue.kind, issue.kind)}: {issue.message}",
            "Blatt": issue.sheet,
            "Zeile": issue.row,
            "Art": issue.kind,
        }
        for name, value in values.items():
            ws.cell(row=r, column=COL[name], value=value)
        ws.cell(row=r, column=COL["Schwere"]).fill = _SEVERITY_FILL.get(issue.severity, PatternFill())
        ws.cell(row=r, column=COL["Problem"]).alignment = Alignment(wrap_text=True, vertical="top")

        if issue.sheet and issue.row and issue.sheet in wb.sheetnames:
            link_cell = ws.cell(row=r, column=COL["Stelle"])
            link_cell.hyperlink = Hyperlink(
                ref=link_cell.coordinate, location=f"'{issue.sheet}'!A{issue.row}",
                display=str(link_cell.value),
            )
            link_cell.font = Font(color="FF0563C1", underline="single")

        old = kept.get((issue.sheet, issue.row))
        if old is not None:
            if old.search:
                ws.cell(row=r, column=COL["Eigene Suche"], value=old.search)
            if old.comment:
                ws.cell(row=r, column=COL["Kommentar"], value=old.comment)
            if old.action:
                ws.cell(row=r, column=COL["Aktion"], value=old.action)

        if issue.test == TEST_FLUSS and issue.candidates:
            start = candidate_row
            labels = []
            for number, candidate in enumerate(issue.candidates, start=1):
                record = candidate.record
                label = candidate_label(number, record, candidate.score)
                # Auswahltext eindeutig je Meldung machen (Excel-Liste zeigt nur den Text)
                label = f"{nr}.{label}"
                labels.append(label)
                row_values = {
                    "Nr": nr, "Kandidat": number, "Auswahltext": label, "database": record.database,
                    "code": record.code, "name": record.name, "reference product": record.product,
                    "location": record.location, "unit": record.unit,
                    "categories": "::".join(record.categories) if record.categories else None,
                    "kind": record.kind, "origin": record.origin, "score": round(candidate.score, 3),
                }
                for name, value in row_values.items():
                    cs.cell(row=candidate_row, column=CCOL[name], value=value)
                candidate_row += 1
            end = candidate_row - 1
            validation = DataValidation(
                type="list",
                formula1=f"'{CANDIDATE_SHEET}'!${openpyxl.utils.get_column_letter(CCOL['Auswahltext'])}${start}:"
                         f"${openpyxl.utils.get_column_letter(CCOL['Auswahltext'])}${end}",
                allow_blank=True,
            )
            ws.add_data_validation(validation)
            validation.add(ws.cell(row=r, column=COL["Auswahl"]))
            action_validation.add(ws.cell(row=r, column=COL["Aktion"]))
            if old is not None and old.choice in labels:
                ws.cell(row=r, column=COL["Auswahl"], value=old.choice)
        elif issue.test == TEST_FLUSS:
            ws.cell(row=r, column=COL["Auswahl"], value="(keine Kandidaten - 'Eigene Suche' nutzen)")

        _mark(wb, issue)

    save_workbook(wb, path)
    wb.close()


def strip_review(path):
    """Entfernt Pruefblatt, Kandidaten und Markierungen aus einer Datei."""
    wb = openpyxl.load_workbook(Path(path))
    changed = REVIEW_SHEET in wb.sheetnames or CANDIDATE_SHEET in wb.sheetnames
    clear_marks(wb)
    remove_review_sheets(wb)
    save_workbook(wb, path)
    wb.close()
    return changed
