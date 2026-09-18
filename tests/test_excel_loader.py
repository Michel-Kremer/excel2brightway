"""
Tests fuer excel2brightway.excel_loader: skip-Blaetter, cutoff-Warnung,
Parameter/Formel-Uebernahme. Nutzt echte, minimale .xlsx-Dateien (via
workbook_factory) sowie die bw2io-Beispieldatei in fixtures/excel_formats/,
damit der komplette Pfad ueber bw2io.ExcelImporter mitgetestet wird.
"""

from pathlib import Path

from excel2brightway import excel_loader

BLANK = [None]


def _minimal_activity_rows(name="Act A", code="a", extra_header=(), extra_exchange_cells=(), cutoff=None):
    header = ["name", "amount", "unit", "database", "location", "type", *extra_header]
    rows = [
        ["Database", "DB1"],
        BLANK,
        ["Activity", name],
        ["code", code],
        ["unit", "kg"],
        ["location", "GLO"],
        ["production amount", 1],
        ["Exchanges"],
        header,
        [name, 1, "kg", "DB1", "GLO", "production", *extra_exchange_cells],
    ]
    if cutoff is not None:
        rows.insert(0, ["cutoff", cutoff])
    return rows


def test_skip_sheet_is_ignored(workbook_factory):
    path = workbook_factory(
        "skip_test.xlsx",
        {
            "SkipMe": [["skip", "wird ignoriert"], ["irgendwas"]],
            "Data": _minimal_activity_rows(),
        },
    )

    activities, internal_index, project_params, database_params, load_warnings = excel_loader.load_clusters(path.parent)

    assert len(activities) == 1
    assert activities[0]["name"] == "Act A"
    assert activities[0]["_database"] == "DB1"


def test_cutoff_hiding_known_field_warns(workbook_factory, capsys):
    # 'formula' ist Spalte Index 6 (7. Spalte); cutoff=6 behaelt nur Index 0-5.
    rows = _minimal_activity_rows(extra_header=["formula"], extra_exchange_cells=[None], cutoff=6)
    path = workbook_factory("cutoff_formula.xlsx", {"Data": rows})

    excel_loader.load_clusters(path.parent)

    out = capsys.readouterr().out
    assert "WARNUNG" in out
    assert "formula" in out
    assert "cutoff=6" in out


def test_cutoff_hiding_unknown_field_is_silent(workbook_factory, capsys):
    # 'notes' ist kein von excel_loader erkanntes Feld -> keine Warnung, obwohl abgeschnitten.
    rows = _minimal_activity_rows(extra_header=["notes"], extra_exchange_cells=["irrelevant"], cutoff=6)
    path = workbook_factory("cutoff_notes.xlsx", {"Data": rows})

    excel_loader.load_clusters(path.parent)

    out = capsys.readouterr().out
    assert "WARNUNG" not in out


def test_formula_exchange_without_amount_is_not_an_error(workbook_factory, capsys):
    rows = _minimal_activity_rows()
    rows[-2] = ["name", "amount", "unit", "database", "location", "type", "formula"]
    rows.append(["Some Flow", None, "kg", "DB1", "GLO", "technosphere", "my_param * 2"])
    path = workbook_factory("formula_blank_amount.xlsx", {"Data": rows})

    activities, *_ = excel_loader.load_clusters(path.parent)

    exch = activities[0]["in"]["Some Flow"]
    assert exch["amount"] == 0.0
    assert exch["formula"] == "my_param * 2"

    out = capsys.readouterr().out
    assert "Some Flow" not in out
    assert "WARNUNG" not in out


def test_exchange_warnings_are_collected_not_printed(workbook_factory, capsys):
    # Exchange ohne 'amount' und ohne 'formula' -> Warnung. Frueher wurde das
    # sofort einzeln gedruckt; jetzt landet es in load_warnings und die
    # Konsole zeigt nur eine Anzahl (siehe excel_loader.load_clusters).
    rows = _minimal_activity_rows()
    rows.append(["Some Flow", None, "kg", "DB1", "GLO", "technosphere"])
    path = workbook_factory("bad_exchange.xlsx", {"Data": rows})

    activities, internal_index, project_params, database_params, load_warnings = excel_loader.load_clusters(path.parent)

    assert len(load_warnings) == 1
    assert load_warnings[0]["kind"] == "missing_amount"
    assert load_warnings[0]["source_file"] == "bad_exchange.xlsx"
    assert load_warnings[0]["activity"] == "Act A"

    out = capsys.readouterr().out
    assert "WARNUNG" not in out
    assert "1 Warnung(en)" in out


def test_activity_comment_is_captured(workbook_factory):
    rows = _minimal_activity_rows()
    rows.insert(4, ["comment", "eine Beschreibung"])
    path = workbook_factory("comment_test.xlsx", {"Data": rows})

    activities, *_ = excel_loader.load_clusters(path.parent)

    assert activities[0]["comment"] == "eine Beschreibung"


def test_sample_activities_with_variables_fixture(fixtures_dir: Path, tmp_path):
    # In ein isoliertes Verzeichnis kopieren, damit die anderen bw2io-
    # Beispieldateien (Doppel-Namen, absichtlich kaputte Formate) diesen
    # Test nicht verrauschen - load_clusters() liest ein ganzes Verzeichnis.
    import shutil
    shutil.copy(fixtures_dir / "sample_activities_with_variables.xlsx", tmp_path)

    activities, internal_index, project_params, database_params, load_warnings = excel_loader.load_clusters(tmp_path)

    mpcb = next(a for a in activities if a["id"] == "mpcb")
    assert mpcb["comment"] == "something important here maybe?"
    assert {p["name"] for p in mpcb["parameters"]} == {"PCB_mass_total"}

    exch = mpcb["in"]["unmounted printed circuit board"]
    assert exch["formula"] == "PCB_area * 2"
    assert exch["amount"] == 0.0

    assert {p["name"]: p["amount"] for p in project_params} == {"PCB_area": 0.25}
    db_param_names = {p["name"] for p in database_params}
    assert db_param_names == {"PCB_cap_mass_film", "PCB_cap_mass_SMD", "PCB_cap_mass_Tantalum"}
