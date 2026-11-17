"""Tests fuer Datenbank/Version umstellen (core/migrate.py, Grundform)."""

import openpyxl
import pytest

from conftest import BIO, BIOSPHERE, EI, ECOINVENT, Record, activity_rows, write_workbook
from excel2brightway.core.backend import MemoryBackend
from excel2brightway.core.migrate import MIGRATION_SHEET, MigrationBlockedError, migrate_files
from excel2brightway.core.review_sheet import read_review
from excel2brightway.core.workbook import read_workbook

EI12 = "ecoinvent-3.12-cutoff"
BIO12 = "ecoinvent-3.12-biosphere"


def _target_backend():
    records = []
    for r in ECOINVENT:
        if r.code == "n-glo":
            continue  # in 3.12 umbenannt (siehe unten)
        if r.code == "h-rer":
            continue  # in 3.12 entfallen
        records.append(Record(EI12, r.code + "-12", r.name, r.product, r.location, r.unit))
    records.append(Record(EI12, "n-12", "market for nickel, class 1, refined", "nickel, class 1", "GLO", "kilogram"))
    biosphere = [Record(BIO12, r.code + "-12", r.name, None, None, r.unit, r.categories, "biosphere") for r in BIOSPHERE]
    return MemoryBackend({EI12: records, BIO12: biosphere}, project="neu")


def _source_file(tmp_path):
    rows = activity_rows(
        "SOEC operation", "op",
        [["electricity, medium voltage", 2, "kWh", EI, "DE", "technosphere"],
         ["nickel, class 1", 0.1, "kg", EI, None, "technosphere"],
         ["heat, district or industrial, natural gas", 1, "MJ", EI, "RER", "technosphere"],
         ["Carbon dioxide, fossil", 0.1, "kg", BIO, None, "biosphere", None, "air"]],
        unit="kWh", database_row="SOEC",
    )
    return write_workbook(tmp_path / "soec.xlsx", {"Data": rows})


def _source_backend():
    return MemoryBackend({EI: ECOINVENT, BIO: BIOSPHERE}, project="alt")


def test_migration_maps_renames_and_leaves_open(tmp_path):
    source = _source_file(tmp_path)
    result = migrate_files(
        [source], _source_backend(), EI, _target_backend(), EI12,
        source_biosphere=BIO, target_biosphere=BIO12,
    )
    assert result.unchanged == 2  # Strom + CO2
    assert result.renamed == 1  # Nickel
    assert result.open == 1  # Waerme RER gibt es im Ziel nicht

    target = result.files[0]
    assert target.parent.name == EI12 and target.name == "soec.xlsx"
    exchanges = read_workbook(target).activities[0].exchanges
    assert {e.get("database") for e in exchanges[1:]} == {EI12, BIO12}
    nickel = exchanges[2]
    assert nickel.get("name") == "market for nickel, class 1, refined"

    # Original unveraendert
    assert read_workbook(source).activities[0].exchanges[1].get("database") == EI

    wb = openpyxl.load_workbook(target)
    assert MIGRATION_SHEET in wb.sheetnames and wb[MIGRATION_SHEET]["A1"].value == "skip"

    # offene Fluesse stehen im Pruefblatt der neuen Datei, mit Kandidaten aus dem Ziel
    assert not result.report.ok
    state = read_review(target)
    assert [r.kind for r in state.rows if r.test == "Fluss"] == ["standort_nicht_gefunden"]
    assert all(record.database == EI12 for record in state.candidates.values())


def test_migration_requires_clean_source(tmp_path):
    rows = activity_rows("A", "a", [["gibt es nicht", 1, "kg", EI, None, "technosphere"]], database_row="X")
    path = write_workbook(tmp_path / "a.xlsx", {"Data": rows})
    with pytest.raises(MigrationBlockedError):
        migrate_files([path], _source_backend(), EI, _target_backend(), EI12)


def test_migration_requires_existing_target(tmp_path):
    with pytest.raises(MigrationBlockedError):
        migrate_files([_source_file(tmp_path)], _source_backend(), EI, _target_backend(), "gibt-es-nicht")
