"""Tests fuer Verweis-Test (core/references.py) und Schreiben (core/write.py) - ohne Brightway."""

import pytest

from conftest import BLANK, EI, InboundReference, Record, activity_rows, make_backend, write_workbook
from excel2brightway.core import write as write_module
from excel2brightway.core.checks import CheckOptions, run_checks
from excel2brightway.core.write import WriteBlockedError, assert_allowed, build_activities, execute_write, plan_write


def _soec_files(tmp_path):
    op = activity_rows(
        "SOEC operation", "op",
        [["stack production, SOEC", 0.001, "unit", "SOEC", None, "technosphere"],
         ["electricity, medium voltage", 2, "kWh", EI, "DE", "technosphere"],
         ["Carbon dioxide, fossil", 0.1, "kg", "biosphere3", None, "biosphere", None, "air"]],
        unit="kWh", database_row="SOEC",
    )
    stack = activity_rows("stack production, SOEC", "stack", [["nickel, class 1", 0.5, "kg", EI, None, "technosphere"]],
                          unit="unit", database_row="SOEC")
    a = write_workbook(tmp_path / "betrieb.xlsx", {"Data": op})
    b = write_workbook(tmp_path / "stack.xlsx", {"Data": stack})
    return a, b


def test_other_file_of_same_database_is_noted_and_written_along(tmp_path):
    a, b = _soec_files(tmp_path)
    report = run_checks([a], backend=make_backend())
    assert report.ok
    hints = [i for i in report.issues if i.kind == "weitere_datei"]
    assert len(hints) == 1 and "stack.xlsx" in hints[0].message

    report, plans = plan_write([a], make_backend())
    assert [p.database for p in plans] == ["SOEC"]
    plan = plans[0]
    assert [p.name for p in plan.auto_added] == ["stack.xlsx"]
    assert {act["code"] for act in plan.activities} == {"op", "stack"}


def test_activities_only_in_brightway_need_confirmation(tmp_path):
    a, _ = _soec_files(tmp_path)
    old = {"SOEC": [Record("SOEC", "op", "SOEC operation", None, "DE", "kWh"),
                    Record("SOEC", "old", "alte Aktivitaet", None, "DE", "kg")]}
    backend = make_backend(extra=old)
    report = run_checks([a], backend=backend)
    warning = next(i for i in report.issues if i.kind == "wird_geloescht")
    assert "alte Aktivitaet" in warning.message
    assert report.ok  # Warnung, kein Fehler

    _, plans = plan_write([a], backend)
    assert [r.code for r in plans[0].to_delete] == ["old"]
    with pytest.raises(WriteBlockedError):
        assert_allowed(plans)
    assert_allowed(plans, allow_delete=True)


def test_broken_references_block_until_allowed(tmp_path):
    a, _ = _soec_files(tmp_path)
    old = {"SOEC": [Record("SOEC", "old", "alte Aktivitaet", None, "DE", "kg")]}
    refs = [InboundReference("Szenario", "x1", "Szenario A", "SOEC", "old")]
    backend = make_backend(extra=old, references=refs)

    report = run_checks([a], backend=backend)
    broken = next(i for i in report.issues if i.kind == "verweis_bricht")
    assert "Szenario/Szenario A -> alte Aktivitaet" in broken.message
    assert not report.ok

    _, plans = plan_write([a], backend)
    assert plans == []  # ohne Freigabe nicht einmal geplant
    _, plans = plan_write([a], backend, allow_broken_references=True)
    with pytest.raises(WriteBlockedError):
        assert_allowed(plans, allow_delete=True)
    assert_allowed(plans, allow_delete=True, allow_broken_references=True)


def test_nothing_is_planned_while_flows_are_open(tmp_path):
    a, _ = _soec_files(tmp_path)
    report, plans = plan_write([a], make_backend(), CheckOptions())
    assert plans  # Kontrolle: so ist alles gruen
    rows = activity_rows("X", "x", [["gibt es nicht", 1, "kg", EI, None, "technosphere"]], database_row="X")
    bad = write_workbook(tmp_path / "x.xlsx", {"Data": rows})
    report, plans = plan_write([bad], make_backend())
    assert plans == [] and not report.ok


def test_build_activities_structure(tmp_path):
    a, b = _soec_files(tmp_path)
    report = run_checks([a, b], backend=make_backend())
    activities, skipped = build_activities([report.models[a], report.models[b]], report.resolutions, "SOEC")
    assert not skipped
    op = next(act for act in activities if act["code"] == "op")
    by_type = {}
    for exchange in op["exchanges"]:
        by_type.setdefault(exchange["type"], []).append(exchange)
    assert by_type["production"][0]["input"] == ("SOEC", "op")
    inputs = {e["input"] for e in by_type["technosphere"]}
    assert inputs == {("SOEC", "stack"), (EI, "e-de-market")}
    assert by_type["biosphere"][0]["input"] == ("ecoinvent-3.10-biosphere", "co2-air")
    assert op["reference product"] == "SOEC operation"


def test_build_keeps_parameters_formulas_and_uncertainty(tmp_path, fixtures_dir):
    import shutil

    shutil.copy(fixtures_dir / "sample_activities_with_variables.xlsx", tmp_path / "pcb.xlsx")
    report = run_checks([tmp_path / "pcb.xlsx"], options=CheckOptions(tests=("format", "fluss")))
    assert report.ok, report.issues
    activities, _ = build_activities(list(report.models.values()), report.resolutions, "PCB")
    mpcb = next(act for act in activities if act["code"] == "mpcb")
    assert mpcb["parameters"] == [{
        "name": "PCB_mass_total", "amount": 0.6,
        "formula": "PCB_cap_mass_film + PCB_cap_mass_SMD + PCB_cap_mass_Tantalum",
    }]
    production = next(e for e in mpcb["exchanges"] if e["type"] == "production")
    assert production["formula"] == "PCB_mass_total"
    techno = next(e for e in mpcb["exchanges"] if e["type"] == "technosphere")
    assert techno["formula"] == "PCB_area * 2" and techno["input"] == ("PCB", "unmounted printed circuit board")
    unmounted = next(act for act in activities if act["code"] == "unmounted printed circuit board")
    assert unmounted["categories"] == ("electronics", "board")
    assert next(e for e in unmounted["exchanges"] if e["type"] == "production")["uncertainty type"] == 0


def test_execute_write_uses_writer(tmp_path, monkeypatch):
    a, _ = _soec_files(tmp_path)
    calls = []

    def fake_write_database(project, database, activities, project_parameters=None, database_parameters=None):
        calls.append((project, database, len(activities)))
        return len(activities)

    from excel2brightway import brightway_writer
    monkeypatch.setattr(brightway_writer, "write_database", fake_write_database)
    _, plans = plan_write([a], make_backend())
    written = execute_write(plans, "proj")
    assert written == [("SOEC", 2)]
    assert calls == [("proj", "SOEC", 2)]


def test_bypassed_format_errors_skip_rows_when_writing(tmp_path):
    rows = activity_rows(
        "A", "a",
        [["nickel, class 1", 0.5, "kg", EI, None, "technosphere"],
         ["nickel, class 1", None, "kg", EI, None, "technosphere"]],
        database_row="DBX",
    )
    path = write_workbook(tmp_path / "a.xlsx", {"Data": rows + [BLANK]})
    _, plans = plan_write([path], make_backend())
    assert plans == []
    _, plans = plan_write([path], make_backend(), CheckOptions(bypass_format=True))
    assert len(plans[0].activities[0]["exchanges"]) == 2  # production + gueltige Zeile
    assert "ohne gueltige amount" in plans[0].skipped[0]
    assert write_module.describe_plan(plans[0]).count("ausgelassen") == 1
