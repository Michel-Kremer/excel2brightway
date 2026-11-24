"""Tests fuer core/importer.py - nur der Teil ohne Brightway (Namen, Entpacken, Ordner finden)."""

import zipfile

import pytest

from excel2brightway.core.importer import ImportProblem, extract_archive, find_ecospold_dirs, suggest_names


@pytest.mark.parametrize("name, expected", [
    ("ecoinvent 3.11_cutoff_ecoSpold02.7z", ("ecoinvent-3.11-cutoff", "ecoinvent-3.11-biosphere", "3.11")),
    ("ecoinvent 3.10.1_cutoff_ecoSpold02", ("ecoinvent-3.10.1-cutoff", "ecoinvent-3.10.1-biosphere", "3.10.1")),
    ("ecoinvent 3.5_apos_ecoSpold02.zip", ("ecoinvent-3.5-apos", "ecoinvent-3.5-biosphere", "3.5")),
    ("meine daten.zip", ("meine-daten", "meine-daten-biosphere", None)),
])
def test_suggest_names(name, expected):
    assert suggest_names(name) == expected


def _make_ecospold_tree(root):
    (root / "datasets").mkdir(parents=True)
    (root / "datasets" / "a.spold").write_text("<x/>", encoding="utf-8")
    (root / "MasterData").mkdir()
    (root / "MasterData" / "ElementaryExchanges.xml").write_text("<x/>", encoding="utf-8")


def test_zip_wird_neben_das_archiv_entpackt_und_gefunden(tmp_path):
    src = tmp_path / "src"
    _make_ecospold_tree(src)
    archive = tmp_path / "ecoinvent 3.11_cutoff_ecoSpold02.zip"
    with zipfile.ZipFile(archive, "w") as zf:
        for f in src.rglob("*"):
            zf.write(f, f.relative_to(src))

    folder = extract_archive(archive, log=lambda *_: None)
    assert folder == tmp_path / "ecoinvent 3.11_cutoff_ecoSpold02"
    datasets, masterdata = find_ecospold_dirs(folder)
    assert datasets == folder / "datasets"
    assert masterdata == folder / "MasterData" / "ElementaryExchanges.xml"


def test_bereits_entpackter_ordner_wird_wiederverwendet(tmp_path):
    archive = tmp_path / "x.zip"
    archive.write_bytes(b"kein gueltiges zip")
    _make_ecospold_tree(tmp_path / "x")
    messages = []
    assert extract_archive(archive, log=messages.append) == tmp_path / "x"
    assert "Bereits entpackt" in messages[0]


def test_ordner_als_quelle_und_verschachtelte_datasets(tmp_path):
    _make_ecospold_tree(tmp_path / "inner")
    assert extract_archive(tmp_path) == tmp_path
    datasets, masterdata = find_ecospold_dirs(tmp_path)
    assert datasets == tmp_path / "inner" / "datasets"
    assert masterdata is not None


def test_ohne_spold_dateien_gibt_es_eine_klare_meldung(tmp_path):
    with pytest.raises(ImportProblem, match="spold"):
        find_ecospold_dirs(tmp_path)


def test_unbekanntes_format(tmp_path):
    f = tmp_path / "x.rar"
    f.write_bytes(b"")
    with pytest.raises(ImportProblem, match=".7z"):
        extract_archive(f)
