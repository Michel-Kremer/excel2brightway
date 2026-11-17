"""Tests fuer die Konsolenbefehle (ohne Brightway-Projekt)."""

from conftest import activity_rows, write_workbook
from excel2brightway import cli


def test_check_format_only_returns_zero_for_clean_file(tmp_path):
    rows = activity_rows("A", "a", [["B", 1, "kg", "DB", None, "technosphere"]], database_row="DB")
    rows += activity_rows("B", "b", [])
    path = write_workbook(tmp_path / "a.xlsx", {"Data": rows})
    assert cli.check([str(path), "--tests", "format,fluss"]) == 0


def test_check_returns_one_with_issues_and_writes_log(tmp_path, capsys):
    rows = activity_rows("A", "a", [["fehlt", 1, "kg", None, None, "technosphere"]], database_row="DB")
    path = write_workbook(tmp_path / "a.xlsx", {"Data": rows})
    assert cli.check([str(tmp_path)]) == 1
    out = capsys.readouterr().out
    assert "nicht bereit" in out and (tmp_path / "ex2bw.log").exists()


def test_unknown_test_name_is_rejected(tmp_path):
    import pytest

    with pytest.raises(SystemExit):
        cli.check([str(tmp_path), "--tests", "format,irgendwas"])


def test_write_needs_project(tmp_path, capsys):
    assert cli.write([str(tmp_path)]) == 2
