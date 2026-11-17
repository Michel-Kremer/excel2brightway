"""Tests fuer den Excel-Leser (core/workbook.py): Format-Konventionen und Zellpositionen."""

import shutil

import openpyxl

from conftest import BLANK, write_workbook
from excel2brightway.core.workbook import data_hash, list_excel_files, read_workbook


def _minimal(cutoff=None, extra_header=(), extra_cells=()):
    rows = [
        ["Database", "DB1"],
        BLANK,
        ["Activity", "Act A"],
        ["code", "a"],
        ["unit", "kg"],
        ["location", "GLO"],
        ["production amount", 1],
        ["Exchanges"],
        ["name", "amount", "unit", "database", "location", "type", *extra_header],
        ["Act A", 1, "kg", "DB1", "GLO", "production", *extra_cells],
    ]
    if cutoff is not None:
        rows.insert(0, ["cutoff", cutoff])
    return rows


def test_skip_and_ex2bw_sheets_are_ignored(tmp_path):
    path = write_workbook(tmp_path / "a.xlsx", {
        "SkipMe": [["skip", "wird ignoriert"], ["Activity", "nicht lesen"]],
        "_ex2bw_Pruefung": [["Activity", "auch nicht"]],
        "Data": _minimal(),
    })
    model = read_workbook(path)
    assert [a.name for a in model.activities] == ["Act A"]
    assert model.activities[0].database == "DB1"


def test_positions_of_activity_and_exchanges(tmp_path):
    rows = _minimal() + [["Flow X", 2, "kg", "DB1", "GLO", "technosphere"]]
    model = read_workbook(write_workbook(tmp_path / "a.xlsx", {"Data": rows}))
    activity = model.activities[0]
    assert (activity.sheet, activity.row) == ("Data", 3)
    assert activity.exchange_header_row == 9
    flow = activity.exchanges[1]
    assert (flow.sheet, flow.row, flow.get("name")) == ("Data", 11, "Flow X")
    assert flow.columns["amount"] == 2


def test_duplicate_flow_names_in_one_activity_are_kept(tmp_path):
    # Frueher ueberschrieb die zweite Zeile die erste (dict nach Flussname).
    rows = _minimal() + [
        ["electricity", 1, "kWh", "ei", "DE", "technosphere"],
        ["electricity", 2, "kWh", "ei", "FR", "technosphere"],
    ]
    model = read_workbook(write_workbook(tmp_path / "a.xlsx", {"Data": rows}))
    techno = [e for e in model.activities[0].exchanges if e.type == "technosphere"]
    assert [e.get("location") for e in techno] == ["DE", "FR"]


def test_cutoff_hiding_known_column_warns(tmp_path):
    rows = _minimal(cutoff=6, extra_header=["formula"], extra_cells=[None])
    model = read_workbook(write_workbook(tmp_path / "a.xlsx", {"Data": rows}))
    kinds = [i.kind for i in model.issues]
    assert "cutoff_verdeckt_spalte" in kinds
    assert "formula" not in model.activities[0].exchange_columns


def test_cutoff_hiding_unknown_column_is_silent(tmp_path):
    rows = _minimal(cutoff=6, extra_header=["notes"], extra_cells=["egal"])
    model = read_workbook(write_workbook(tmp_path / "a.xlsx", {"Data": rows}))
    assert not [i for i in model.issues if i.kind.startswith("cutoff")]


def test_activity_continues_across_blank_lines(fixtures_dir):
    model = read_workbook(fixtures_dir / "sample_activities_with_variables.xlsx")
    mpcb = next(a for a in model.activities if a.code == "mpcb")
    assert mpcb.fields["comment"] == "something important here maybe?"
    assert [p["name"] for p in mpcb.parameters] == ["PCB_mass_total"]
    assert mpcb.parameter_group == "alpha group!"
    techno = next(e for e in mpcb.exchanges if e.type == "technosphere")
    assert techno.get("formula") == "PCB_area * 2"
    assert {p["name"] for p in model.project_parameters} == {"PCB_area"}
    assert {p["name"] for p in model.database_parameters} == {
        "PCB_cap_mass_film", "PCB_cap_mass_SMD", "PCB_cap_mass_Tantalum",
    }


def test_bw2io_examples_with_blank_lines_and_reference_product(fixtures_dir):
    model = read_workbook(fixtures_dir / "blank_lines.xlsx")
    assert model.database == "BW2 Excel example"
    assert len(model.activities) == 3

    model = read_workbook(fixtures_dir / "basic_example_with_reference_product.xlsx")
    acetic = model.activities[0]
    assert acetic.name == "Acetic acid production"
    assert acetic.reference_product == "Acetic acid, at plant"


def test_unknown_marker_values_become_none(fixtures_dir):
    model = read_workbook(fixtures_dir / "basic_example.xlsx")
    ammonia = model.activities[0].exchanges[0]
    assert ammonia.get("location") is None  # '(Unknown)'


def test_second_database_row_is_an_error(tmp_path):
    rows = _minimal() + [BLANK, ["Database", "DB2"]]
    model = read_workbook(write_workbook(tmp_path / "a.xlsx", {"Data": rows}))
    assert "database_mehrfach" in [i.kind for i in model.issues]
    assert model.database == "DB1"


def test_missing_database_row(tmp_path):
    rows = _minimal()[2:]
    model = read_workbook(write_workbook(tmp_path / "a.xlsx", {"Data": rows}))
    assert "database_fehlt" in [i.kind for i in model.issues]


def test_unreadable_file_is_reported_not_raised(tmp_path):
    path = tmp_path / "kaputt.xlsx"
    path.write_text("keine Excel-Datei")
    model = read_workbook(path)
    assert not model.readable
    assert model.issues[0].kind == "datei_unlesbar"


def test_data_hash_ignores_ex2bw_sheets(tmp_path):
    path = write_workbook(tmp_path / "a.xlsx", {"Data": _minimal()})
    before = data_hash(path)
    wb = openpyxl.load_workbook(path)
    wb.create_sheet("_ex2bw_Pruefung").append(["skip", "Auswahl"])
    wb.save(path)
    assert data_hash(path) == before
    wb = openpyxl.load_workbook(path)
    wb["Data"]["B4"] = "geaendert"
    wb.save(path)
    assert data_hash(path) != before


def test_list_excel_files_skips_lock_files(tmp_path, fixtures_dir):
    shutil.copy(fixtures_dir / "basic_example.xlsx", tmp_path / "a.xlsx")
    (tmp_path / "~$a.xlsx").write_text("lock")
    (tmp_path / "notizen.txt").write_text("x")
    assert [p.name for p in list_excel_files(tmp_path)] == ["a.xlsx"]
