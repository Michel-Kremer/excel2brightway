"""Tests fuer den Format-Test (core/format_check.py) und seine Ueberbrueckung."""

from conftest import BLANK, write_workbook
from excel2brightway.core.checks import CheckOptions, run_checks
from excel2brightway.core.format_check import FORMAT_SCHEMA, check_format
from excel2brightway.core.model import FEHLER, HINWEIS
from excel2brightway.core.workbook import read_workbook


def _rows(*exchanges, unit="kg", header=None, production=True):
    header = header or ["name", "amount", "unit", "database", "location", "type"]
    rows = [
        ["Database", "DB1"], BLANK,
        ["Activity", "Act A"], ["code", "a"], ["location", "GLO"],
    ]
    if unit:
        rows.append(["unit", unit])
    rows += [["Exchanges"], header]
    if production:
        rows.append(["Act A", 1, "kg", "DB1", "GLO", "production"])
    rows += list(exchanges)
    return rows


def _kinds(tmp_path, rows, name="a.xlsx"):
    model = read_workbook(write_workbook(tmp_path / name, {"Data": rows}))
    return {i.kind: i for i in check_format([model])}


def test_schema_marks_required_and_optional():
    assert "unit" in FORMAT_SCHEMA["activity"]["required"]
    assert set(FORMAT_SCHEMA["exchange_columns"]["required"]) == {"name", "amount", "unit", "type"}
    assert "location" in FORMAT_SCHEMA["exchange_columns"]["optional"]


def test_clean_file_has_no_errors(tmp_path):
    issues = _kinds(tmp_path, _rows(["Flow", 1, "kg", "DB1", "GLO", "technosphere"]))
    assert not [i for i in issues.values() if i.severity == FEHLER]


def test_missing_activity_unit(tmp_path):
    issue = _kinds(tmp_path, _rows(unit=None))["pflichtfeld_fehlt"]
    assert issue.severity == FEHLER and "unit" in issue.message


def test_missing_required_exchange_column(tmp_path):
    rows = _rows(header=["name", "amount", "database", "location", "type"], production=False)
    rows.append(["Act A", 1, "DB1", "GLO", "production"])
    issue = _kinds(tmp_path, rows)["spalte_fehlt"]
    assert "unit" in issue.message


def test_unknown_type_and_invalid_amount(tmp_path):
    issues = _kinds(tmp_path, _rows(
        ["Flow", 1, "kg", "DB1", "GLO", "input"],
        ["Flow 2", "viel", "kg", "DB1", "GLO", "technosphere"],
    ))
    assert issues["typ_unbekannt"].row == 10
    assert issues["menge_ungueltig"].row == 11
    assert issues["menge_ungueltig"].column == 2


def test_formula_without_amount_is_fine(tmp_path):
    rows = _rows(header=["name", "amount", "unit", "database", "location", "type", "formula"])
    rows.append(["Flow", None, "kg", "DB1", "GLO", "technosphere", "p * 2"])
    assert "menge_ungueltig" not in _kinds(tmp_path, rows)


def test_missing_production_is_only_a_hint(tmp_path):
    issue = _kinds(tmp_path, _rows(["Flow", 1, "kg", "DB1", "GLO", "technosphere"], production=False))["produktion_fehlt"]
    assert issue.severity == HINWEIS


def test_two_production_rows(tmp_path):
    issues = _kinds(tmp_path, _rows(["Act A", 1, "kg", "DB1", "GLO", "production"]))
    assert issues["produktion_mehrfach"].severity == FEHLER


def test_duplicate_code_across_files(tmp_path):
    a = read_workbook(write_workbook(tmp_path / "a.xlsx", {"Data": _rows()}))
    b = read_workbook(write_workbook(tmp_path / "b.xlsx", {"Data": _rows()}))
    issues = check_format([a, b], scope={a.path})
    duplicate = [i for i in issues if i.kind == "code_doppelt"]
    assert len(duplicate) == 1 and "b.xlsx" in duplicate[0].message
    assert all(i.file == a.path for i in issues)  # nur Meldungen fuer die Auswahl


def test_format_errors_block_unless_bypassed(tmp_path):
    path = write_workbook(tmp_path / "a.xlsx", {"Data": _rows(["Flow", "viel", "kg", "DB1", "GLO", "technosphere"])})
    report = run_checks([path], options=CheckOptions(tests=("format",), write_review=False))
    assert not report.ok
    report = run_checks([path], options=CheckOptions(tests=("format",), bypass_format=True, write_review=False))
    assert report.ok
    assert report.counts()[FEHLER] == 1  # Meldung bleibt sichtbar, blockiert aber nicht


def test_format_test_alone_does_not_touch_flows(tmp_path):
    path = write_workbook(tmp_path / "a.xlsx", {"Data": _rows(["voellig unbekannt", 1, "kg", None, None, "technosphere"])})
    report = run_checks([path], options=CheckOptions(tests=("format",), write_review=False))
    assert report.ok
    assert not [i for i in report.issues if i.test == "fluss"]


def test_structural_errors_block_even_when_bypassed(tmp_path):
    path = write_workbook(tmp_path / "a.xlsx", {"Data": _rows(unit=None)})
    report = run_checks([path], options=CheckOptions(tests=("format",), bypass_format=True, write_review=False))
    assert not report.ok
    assert [i.kind for i in report.blocking_issues()] == ["pflichtfeld_fehlt"]
