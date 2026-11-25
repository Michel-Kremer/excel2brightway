"""Tests: beim Einlesen eines Ordners nur LCI-Dateien ('cutoff' in A1 eines Blatts)."""

from conftest import EI, activity_rows, make_backend, write_workbook
from excel2brightway.core.checks import expand_selection, run_checks
from excel2brightway.core.workbook import is_lci_file


def _lci(folder, name="lci.xlsx"):
    rows = activity_rows("act a", "a", [["electricity, medium voltage", 1, "kWh", EI, "DE", "technosphere"]],
                         unit="kWh", database_row="DB")
    return write_workbook(folder / name, {"Data": rows}, cutoff=True)


def _other(folder, name="recherche.xlsx"):
    return write_workbook(folder / name, {"Quellen": [["Autor", "Jahr"], ["Muster", 2024]]})


def test_is_lci_file_erkennt_cutoff_in_a1_irgendeines_blatts(tmp_path):
    path = write_workbook(tmp_path / "x.xlsx", {
        "Metadata": [["skip"], ["Titel"]],
        "Inventory": [["cutoff", 19], ["Database", "DB"]],
    })
    assert is_lci_file(path)
    assert not is_lci_file(_other(tmp_path))


def test_is_lci_file_mit_beispieldateien(fixtures_dir):
    assert is_lci_file(fixtures_dir / "basic_example.xlsx")
    # bw2io-Format ohne 'cutoff' (Database in A1) zaehlt nicht als LCI-Datei im Ordner
    assert not is_lci_file(fixtures_dir / "with_products.xlsx")


def test_unlesbare_datei_gilt_als_lci_damit_der_fehler_gemeldet_wird(tmp_path):
    broken = tmp_path / "kaputt.xlsx"
    broken.write_bytes(b"keine excel-datei")
    assert is_lci_file(broken)


def test_ordner_liest_nur_lci_dateien(tmp_path):
    lci = _lci(tmp_path)
    other = _other(tmp_path)
    skipped = []
    assert expand_selection([tmp_path], skipped) == [lci.resolve()]
    assert skipped == [other.resolve()]

    report = run_checks([tmp_path], backend=make_backend())
    assert list(report.models) == [lci.resolve()]
    assert report.skipped == [other.resolve()]
    assert "recherche.xlsx" in report.summary()
    assert report.ok


def test_nachbardateien_ohne_lci_blatt_werden_nicht_als_kontext_gelesen(tmp_path):
    lci = _lci(tmp_path)
    _other(tmp_path)
    report = run_checks([lci], backend=make_backend())
    assert list(report.models) == [lci.resolve()]
    assert [p.name for p in report.skipped] == ["recherche.xlsx"]


def test_einzeln_gewaehlte_datei_wird_immer_gelesen(tmp_path):
    other = _other(tmp_path)
    report = run_checks([other], backend=make_backend())
    assert other.resolve() in report.models
    assert report.skipped == []
