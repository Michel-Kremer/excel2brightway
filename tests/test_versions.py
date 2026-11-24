"""Tests fuer core/versions.py: mehrere Fassungen derselben Datei im Ordner."""

import pytest

from conftest import EI, activity_rows, make_backend, write_workbook
from excel2brightway.core.checks import run_checks
from excel2brightway.core.versions import split_version
from excel2brightway.core.write import plan_write


@pytest.mark.parametrize("stem, expected", [
    ("LCI_X_v0.3", ("lci_x", (0, 3))),
    ("LCI_X_v0.10", ("lci_x", (0, 10))),
    ("LCI_X-V2", ("lci_x", (2,))),
    ("LCI_X v1.0", ("lci_x", (1,))),
    ("LCI_X_final", ("lci_x_final", None)),
    ("LCI_X_v0.3 - Kopie", ("lci_x_v0.3 - kopie", None)),
    ("Inventar", ("inventar", None)),
])
def test_split_version(stem, expected):
    assert split_version(stem) == expected


def _file(folder, name, code, database="DB"):
    rows = activity_rows(f"act {code}", code, [["electricity, medium voltage", 1, "kWh", EI, "DE", "technosphere"]],
                         unit="kWh", database_row=database)
    return write_workbook(folder / name, {"Data": rows})


def _kinds(report):
    return {(i.kind, i.severity) for i in report.issues if i.kind in
            ("andere_version", "variante_ohne_version", "version_unklar", "weitere_datei")}


def test_einzelauswahl_zaehlt_die_ausgewaehlte_version(tmp_path):
    old = _file(tmp_path, "LCI_X_v0.1.xlsx", "a")
    new = _file(tmp_path, "LCI_X_v0.3.xlsx", "a")  # gleicher code: wuerde sonst code_doppelt geben
    report = run_checks([new], backend=make_backend())
    assert report.ok
    assert _kinds(report) == {("andere_version", "HINWEIS")}
    assert report.ignored == {old.resolve(): new.resolve()}
    assert old.resolve() not in report.models

    report = run_checks([old], backend=make_backend())
    hint = next(i for i in report.issues if i.kind == "andere_version")
    assert "neuere Version" in hint.message and report.ok


def test_ordnerauswahl_nimmt_die_neueste(tmp_path):
    _file(tmp_path, "LCI_X_v0.1.xlsx", "a")
    _file(tmp_path, "LCI_X_v0.10.xlsx", "a")
    _file(tmp_path, "LCI_X_v0.9.xlsx", "a")
    report = run_checks([tmp_path], backend=make_backend())
    assert report.ok
    assert [p.name for p in report.selected] == ["LCI_X_v0.10.xlsx"]

    _, plans = plan_write([tmp_path], make_backend())
    assert [p.name for p in plans[0].files] == ["LCI_X_v0.10.xlsx"] and not plans[0].auto_added


def test_variante_ohne_version_bei_einzelauswahl_nur_warnung(tmp_path):
    new = _file(tmp_path, "LCI_X_v0.3.xlsx", "a")
    _file(tmp_path, "LCI_X_final.xlsx", "a")
    report = run_checks([new], backend=make_backend())
    assert report.ok
    assert _kinds(report) == {("variante_ohne_version", "WARNUNG")}

    _, plans = plan_write([new], make_backend())
    assert [p.name for p in plans[0].files] == ["LCI_X_v0.3.xlsx"]


def test_variante_ohne_version_bei_ordnerauswahl_blockiert(tmp_path):
    _file(tmp_path, "LCI_X_v0.3.xlsx", "a")
    _file(tmp_path, "LCI_X_final.xlsx", "b")
    report = run_checks([tmp_path], backend=make_backend())
    assert ("version_unklar", "FEHLER") in _kinds(report)
    assert not report.ok
    _, plans = plan_write([tmp_path], make_backend())
    assert plans == []


def test_mehrfachauswahl_zweier_fassungen_nimmt_die_neueste(tmp_path):
    a = _file(tmp_path, "LCI_X_v1.xlsx", "a")
    b = _file(tmp_path, "LCI_X_v2.xlsx", "a")
    report = run_checks([a, b], backend=make_backend())
    assert report.ok and [p.name for p in report.selected] == ["LCI_X_v2.xlsx"]


def test_gleiche_versionsnummer_bei_mehrfachauswahl_blockiert(tmp_path):
    a = _file(tmp_path, "LCI_X_v1.xlsx", "a")
    b = _file(tmp_path, "LCI_X v1.0.xlsx", "b")
    report = run_checks([a, b], backend=make_backend())
    assert ("version_unklar", "FEHLER") in _kinds(report) and not report.ok


def test_echte_teildateien_bleiben_weitere_dateien(tmp_path):
    a = _file(tmp_path, "betrieb.xlsx", "a")
    _file(tmp_path, "stack.xlsx", "b")
    _file(tmp_path, "anderes_v1.xlsx", "c", database="ANDERE")
    _file(tmp_path, "anderes_v2.xlsx", "c", database="ANDERE")
    report = run_checks([a], backend=make_backend())
    assert _kinds(report) == {("weitere_datei", "HINWEIS")}
    assert report.ok


def test_gleicher_stamm_aber_andere_datenbank_ist_keine_fassung(tmp_path):
    a = _file(tmp_path, "LCI_X_v1.xlsx", "a", database="DB1")
    _file(tmp_path, "LCI_X_v2.xlsx", "b", database="DB2")
    report = run_checks([a], backend=make_backend())
    assert not _kinds(report) and report.ok
