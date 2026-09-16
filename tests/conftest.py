"""Gemeinsame Fixtures fuer die excel2brightway-Tests."""

from pathlib import Path

import openpyxl
import pytest

FIXTURES_DIR = Path(__file__).resolve().parent / "fixtures" / "excel_formats"


@pytest.fixture
def fixtures_dir() -> Path:
    return FIXTURES_DIR


def write_workbook(path: Path, sheets: dict) -> None:
    """
    Schreibt eine .xlsx mit den gegebenen Blaettern.
    `sheets`: {sheet_name: [[zelle, zelle, ...], ...]} - eine Liste von
    Zeilen (jede Zeile eine Liste von Zellwerten, None fuer leere Zellen).
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sheet_name, rows in sheets.items():
        ws = wb.create_sheet(sheet_name)
        for row in rows:
            ws.append(row)
    wb.save(path)


@pytest.fixture
def workbook_factory(tmp_path):
    """Gibt eine Funktion zurueck, die (name, sheets) zu einem Pfad unter tmp_path schreibt."""

    def _factory(name: str, sheets: dict) -> Path:
        path = tmp_path / name
        write_workbook(path, sheets)
        return path

    return _factory
