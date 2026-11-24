"""Tests fuer den Schutz gegen eine falsche bw2data-Version (core/backend.py)."""

from importlib.metadata import PackageNotFoundError

import pytest

from excel2brightway.core import backend
from excel2brightway.core.backend import WrongBw2dataVersion, list_projects, require_bw2data


def _fake_version(value):
    def fake(name):
        if isinstance(value, Exception):
            raise value
        return value
    return fake


def test_passende_major_version_ist_erlaubt(monkeypatch):
    monkeypatch.setattr(backend, "version", _fake_version("4.7"))
    require_bw2data()


@pytest.mark.parametrize("found", ["3.6.6", "5.0"])
def test_andere_major_version_wird_abgelehnt(monkeypatch, found):
    monkeypatch.setattr(backend, "version", _fake_version(found))
    with pytest.raises(WrongBw2dataVersion, match=f"bw2data {found}"):
        require_bw2data()


def test_falsche_version_blockiert_auch_projektliste(monkeypatch):
    monkeypatch.setattr(backend, "version", _fake_version("3.6.6"))
    with pytest.raises(WrongBw2dataVersion):
        list_projects()


def test_ohne_bw2data_ist_projektliste_leer(monkeypatch):
    monkeypatch.setattr(backend, "version", _fake_version(PackageNotFoundError("bw2data")))
    assert list_projects() == []
