"""Tests fuer Pruefblatt (core/review_sheet.py), Uebernehmen (core/apply.py) und Fehler-Log."""

import openpyxl
import pytest

from conftest import EI, activity_rows, make_backend, write_workbook
from excel2brightway.core.apply import BACKUP_DIRNAME, CHANGELOG_SHEET, DataChangedError, apply_review
from excel2brightway.core.checks import LOG_FILENAME, CheckOptions, run_checks
from excel2brightway.core.review_sheet import (
    CANDIDATE_SHEET, COL, COMMENT_AUTHOR, FIRST_ROW, REVIEW_SHEET, read_review,
)
from excel2brightway.core.workbook import read_workbook


def _file(tmp_path, exchanges, name="soec.xlsx"):
    rows = activity_rows("SOEC operation", "op", exchanges, unit="kWh", database_row="SOEC")
    return write_workbook(tmp_path / name, {"Data": rows})


def _choose(path, row_index, choice_index=0, action=None, comment=None, search=None):
    """Simuliert die Auswahl im Pruefblatt (wie in Excel)."""
    state = read_review(path)
    wb = openpyxl.load_workbook(path)
    ws = wb[REVIEW_SHEET]
    review_row = state.rows[row_index]
    r = FIRST_ROW + review_row.nr - 1
    if choice_index is not None:
        labels = [label for label in state.candidates if label.startswith(f"{review_row.nr}.")]
        ws.cell(row=r, column=COL["Auswahl"], value=labels[choice_index])
    if action:
        ws.cell(row=r, column=COL["Aktion"], value=action)
    if comment:
        ws.cell(row=r, column=COL["Kommentar"], value=comment)
    if search:
        ws.cell(row=r, column=COL["Eigene Suche"], value=search)
    wb.save(path)


def test_review_sheet_structure(tmp_path):
    path = _file(tmp_path, [["electricty medium voltage", 2, "kWh", EI, "DE", "technosphere"]])
    report = run_checks([path], backend=make_backend())
    assert not report.ok

    wb = openpyxl.load_workbook(path)
    assert wb.sheetnames[0] == REVIEW_SHEET
    assert wb[REVIEW_SHEET]["A1"].value == "skip"  # bw2io liest das Blatt nie
    assert wb[CANDIDATE_SHEET].sheet_state == "hidden"
    ws = wb[REVIEW_SHEET]
    validations = [dv for dv in ws.data_validations.dataValidation if CANDIDATE_SHEET in (dv.formula1 or "")]
    assert len(validations) == 1
    link = ws.cell(row=FIRST_ROW, column=COL["Stelle"]).hyperlink
    assert link is not None and "'Data'!A" in link.location

    data = wb["Data"]
    marked = [c for row in data.iter_rows() for c in row if c.comment is not None]
    assert len(marked) == 1 and marked[0].comment.author == COMMENT_AUTHOR

    # Excel-Datei bleibt fuer den Leser gueltig
    model = read_workbook(path)
    assert [a.name for a in model.activities] == ["SOEC operation"]


def test_clean_file_has_no_review_sheet_and_old_one_is_removed(tmp_path):
    path = _file(tmp_path, [["electricty medium voltage", 2, "kWh", EI, "DE", "technosphere"]])
    run_checks([path], backend=make_backend())
    _choose(path, 0)
    apply_review(path)
    report = run_checks([path], backend=make_backend())
    assert report.ok
    wb = openpyxl.load_workbook(path)
    assert REVIEW_SHEET not in wb.sheetnames and CANDIDATE_SHEET not in wb.sheetnames
    assert not [c for row in wb["Data"].iter_rows() for c in row if c.comment is not None]


def test_apply_writes_selection_into_exchange_row(tmp_path):
    path = _file(tmp_path, [["electricty medium voltage", 2, "kWh", "ecoinvent", "DE", "technosphere"]])
    run_checks([path], backend=make_backend())
    _choose(path, 0, choice_index=0)

    result = apply_review(path)
    assert len(result.applied) == 1 and not result.skipped
    assert result.backup.parent.name == BACKUP_DIRNAME and result.backup.exists()

    model = read_workbook(path)
    exchange = model.activities[0].exchanges[1]
    assert exchange.get("name") == "market for electricity, medium voltage"
    assert exchange.get("reference product") == "electricity, medium voltage"
    assert exchange.get("database") == EI
    assert exchange.get("location") == "DE"
    assert exchange.get("amount") == 2  # Menge unveraendert

    wb = openpyxl.load_workbook(path)
    log = wb[CHANGELOG_SHEET]
    assert log["A1"].value == "skip"
    changed_fields = {log.cell(row=r, column=6).value for r in range(3, log.max_row + 1)}
    assert {"name", "database"} <= changed_fields


def test_missing_columns_are_added(tmp_path):
    header = ["name", "amount", "unit", "type"]
    rows = activity_rows("SOEC operation", "op", [["electricty medium voltage", 2, "kWh", "technosphere"]],
                         unit="kWh", database_row="SOEC", header=header)
    rows[rows.index(["SOEC operation", 1, "kWh", "SOEC", "DE", "production"])] = ["SOEC operation", 1, "kWh", "production"]
    path = write_workbook(tmp_path / "s.xlsx", {"Data": rows})
    run_checks([path], backend=make_backend())
    _choose(path, 0)
    apply_review(path)
    exchange = read_workbook(path).activities[0].exchanges[1]
    assert exchange.get("database") == EI and exchange.get("location") == "DE"
    assert run_checks([path], backend=make_backend()).ok


def test_proxy_action_is_documented_in_comment(tmp_path):
    path = _file(tmp_path, [["electricty medium voltage", 2, "kWh", EI, "DE", "technosphere"]])
    run_checks([path], backend=make_backend())
    _choose(path, 0, action="Proxy", comment="kein SOEC-spezifischer Datensatz")
    apply_review(path)
    exchange = read_workbook(path).activities[0].exchanges[1]
    assert exchange.get("comment") == "Proxy (ex2bw): kein SOEC-spezifischer Datensatz"


def test_unit_mismatch_is_not_applied(tmp_path):
    path = _file(tmp_path, [["nickel, class 1", 1, "kWh", EI, None, "technosphere"]])
    run_checks([path], backend=make_backend())
    _choose(path, 0)
    result = apply_review(path)
    assert not result.applied and "Einheit" in result.skipped[0][2]
    assert read_workbook(path).activities[0].exchanges[1].get("unit") == "kWh"


def test_apply_refuses_when_data_changed_after_check(tmp_path):
    path = _file(tmp_path, [["electricty medium voltage", 2, "kWh", EI, "DE", "technosphere"]])
    run_checks([path], backend=make_backend())
    _choose(path, 0)
    wb = openpyxl.load_workbook(path)
    wb["Data"]["B10"] = 99
    wb.save(path)
    with pytest.raises(DataChangedError):
        apply_review(path)


def test_own_search_is_kept_and_used_on_next_check(tmp_path):
    path = _file(tmp_path, [["Strom Mittelspannung", 2, "kWh", EI, "DE", "technosphere"]])
    run_checks([path], backend=make_backend())
    _choose(path, 0, choice_index=None, search="electricity medium voltage", comment="Notiz")
    run_checks([path], backend=make_backend())
    state = read_review(path)
    assert state.rows[0].search == "electricity medium voltage"
    assert state.rows[0].comment == "Notiz"
    first = next(label for label in state.candidates if label.startswith("1.1"))
    assert "market for electricity, medium voltage" in first


def test_log_file_and_other_errors(tmp_path):
    path = _file(tmp_path, [["electricty medium voltage", 2, "kWh", EI, "DE", "technosphere"]])
    (tmp_path / "kaputt.xlsx").write_text("keine Excel-Datei")
    report = run_checks([path, tmp_path / "kaputt.xlsx"], backend=make_backend())
    log = (tmp_path / LOG_FILENAME).read_text(encoding="utf-8")
    assert report.log_path == tmp_path / LOG_FILENAME
    assert "datei_unlesbar" in log and "nicht_gefunden" in log


def test_tests_can_run_separately(tmp_path):
    path = _file(tmp_path, [["electricty medium voltage", "viel", "kWh", EI, "DE", "technosphere"]])
    only_format = run_checks([path], backend=make_backend(), options=CheckOptions(tests=("format",)))
    assert {i.test for i in only_format.issues} == {"format"}
    only_flow = run_checks([path], backend=make_backend(), options=CheckOptions(tests=("fluss",)))
    assert {i.test for i in only_flow.issues} == {"fluss"}
    both = run_checks([path], backend=make_backend(), options=CheckOptions(tests=("format", "fluss")))
    assert {i.test for i in both.issues} == {"format", "fluss"}
