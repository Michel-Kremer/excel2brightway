"""Gemeinsame Fixtures und Helfer fuer die excel2brightway-Tests (ohne Brightway)."""

from pathlib import Path

import openpyxl
import pytest

from excel2brightway.core.backend import InboundReference, MemoryBackend
from excel2brightway.core.model import Record

FIXTURES_DIR = Path(__file__).resolve().parent / "fixtures" / "excel_formats"
BLANK = [None]
EI = "ecoinvent-3.10-cutoff"
BIO = "ecoinvent-3.10-biosphere"


@pytest.fixture
def fixtures_dir() -> Path:
    return FIXTURES_DIR


def write_workbook(path: Path, sheets: dict) -> Path:
    """
    Schreibt eine .xlsx mit den gegebenen Blaettern.
    `sheets`: {sheet_name: [[zelle, zelle, ...], ...]} - eine Liste von Zeilen.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sheet_name, rows in sheets.items():
        ws = wb.create_sheet(sheet_name)
        for row in rows:
            ws.append(row)
    wb.save(path)
    return path


@pytest.fixture
def workbook_factory(tmp_path):
    """Gibt eine Funktion zurueck, die (name, sheets) zu einem Pfad unter tmp_path schreibt."""

    def _factory(name: str, sheets: dict) -> Path:
        return write_workbook(tmp_path / name, sheets)

    return _factory


def activity_rows(name, code, exchanges, unit="kg", location="DE", header=None, database_row=None):
    """Zeilen fuer eine Aktivitaet; `exchanges` ohne production-Zeile (wird ergaenzt)."""
    header = header or ["name", "amount", "unit", "database", "location", "type", "reference product", "categories"]
    rows = []
    if database_row:
        rows += [["Database", database_row], BLANK]
    rows += [
        ["Activity", name],
        ["code", code],
        ["unit", unit],
        ["location", location],
        ["Exchanges"],
        header,
        [name, 1, unit, database_row or None, location, "production"],
    ]
    rows += exchanges
    rows.append(BLANK)
    return rows


def ei(code, name, product, location, unit="kilogram"):
    return Record(EI, code, name, product, location, unit)


def bio(code, name, categories, unit="kilogram"):
    return Record(BIO, code, name, None, None, unit, tuple(categories), "biosphere")


ECOINVENT = [
    ei("e-de-market", "market for electricity, medium voltage", "electricity, medium voltage", "DE", "kilowatt hour"),
    ei("e-de-transf", "electricity voltage transformation from high to medium voltage", "electricity, medium voltage", "DE", "kilowatt hour"),
    ei("e-fr-market", "market for electricity, medium voltage", "electricity, medium voltage", "FR", "kilowatt hour"),
    ei("h-rer", "market for heat, district or industrial, natural gas", "heat, district or industrial, natural gas", "RER", "megajoule"),
    ei("h-row", "market for heat, district or industrial, natural gas", "heat, district or industrial, natural gas", "RoW", "megajoule"),
    ei("w-at", "market for tap water", "tap water", "AT", "kilogram"),
    ei("w-ch", "market for tap water", "tap water", "CH", "kilogram"),
    ei("w-row", "market for tap water", "tap water", "RoW", "kilogram"),
    ei("s-1", "steel production, converter, low-alloyed", "steel, low-alloyed", "RER", "kilogram"),
    ei("s-2", "steel production, electric, low-alloyed", "steel, low-alloyed", "RER", "kilogram"),
    ei("n-glo", "market for nickel, class 1", "nickel, class 1", "GLO", "kilogram"),
]

BIOSPHERE = [
    bio("co2-air", "Carbon dioxide, fossil", ["air"]),
    bio("co2-urban", "Carbon dioxide, fossil", ["air", "urban air close to ground"]),
    bio("co2-low", "Carbon dioxide, fossil", ["air", "lower stratosphere + upper troposphere"]),
    bio("water-river", "Water", ["water", "surface water"], "cubic meter"),
    bio("water-ground", "Water", ["water", "ground-"], "cubic meter"),
]


@pytest.fixture
def backend():
    return MemoryBackend({EI: ECOINVENT, BIO: BIOSPHERE}, project="proj")


def make_backend(extra=None, references=None, project="proj"):
    databases = {EI: ECOINVENT, BIO: BIOSPHERE}
    databases.update(extra or {})
    return MemoryBackend(databases, references=references, project=project)


__all__ = [
    "BLANK", "EI", "BIO", "ECOINVENT", "BIOSPHERE", "activity_rows", "write_workbook", "make_backend",
    "InboundReference", "Record",
]
