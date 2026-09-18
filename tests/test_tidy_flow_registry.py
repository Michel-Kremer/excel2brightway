"""
Tests fuer excel2brightway.tidy_flow_registry: Bereinigen veralteter
Eintraege (_prune_stale_entries) und Backfill (_backfill) - beide nehmen
das Brightway-Projekt jetzt als Parameter statt einer hardcodierten
Konstante. Nutzt Monkeypatch von _try_activate_project, damit KEIN echter
bw2data-Projektzugriff stattfindet (das wuerde bei einem unbekannten
Projektnamen sonst ein echtes, leeres Projekt anlegen).
"""

from excel2brightway import tidy_flow_registry as tfr


class _FakeBD:
    def __init__(self, databases, activities=None):
        self.databases = list(databases)
        self._activities = activities or {}

    def get_activity(self, key):
        return self._activities[key]


def _patch_project(monkeypatch, bd):
    monkeypatch.setattr(tfr, "_try_activate_project", lambda project_name: bd)


def test_prune_removes_entries_with_missing_database(monkeypatch):
    _patch_project(monkeypatch, _FakeBD(["ecoinvent-3.10.1-cutoff", "ecoinvent-3.10-biosphere"]))
    registry = {
        "technosphere": {
            "steel production": {"database": "ecoinvent-3.10.1-cutoff", "code": "abc"},
            "stale flow": {"database": "ecoinvent-3.12-cutoff", "code": "old"},
        },
        "biosphere": {
            "Water": {"database": "ecoinvent-3.10-biosphere", "code": "w1"},
            "Carbon dioxide": {"database": "biosphere3", "code": "co2"},
        },
    }

    removed = tfr._prune_stale_entries(registry, "phoenix_ecoinvent-3.10.1-cutoff")

    assert "stale flow" not in registry["technosphere"]
    assert "steel production" in registry["technosphere"]
    assert "Carbon dioxide" not in registry["biosphere"]
    assert "Water" in registry["biosphere"]
    assert len(removed) == 2


def test_prune_without_project_is_skipped(monkeypatch):
    monkeypatch.setattr(tfr, "_try_activate_project", lambda project_name: None)
    registry = {"technosphere": {"x": {"database": "whatever", "code": "1"}}, "biosphere": {}}

    removed = tfr._prune_stale_entries(registry, None)

    assert removed == []
    assert "x" in registry["technosphere"]


def test_backfill_fills_missing_unit_and_categories(monkeypatch):
    bd = _FakeBD(
        ["ecoinvent-3.10.1-cutoff", "ecoinvent-3.10-biosphere"],
        activities={
            ("ecoinvent-3.10.1-cutoff", "abc"): {"unit": "kilogram"},
            ("ecoinvent-3.10-biosphere", "co2"): {"categories": ["air"]},
        },
    )
    _patch_project(monkeypatch, bd)
    registry = {
        "technosphere": {
            "steel production": {"database": "ecoinvent-3.10.1-cutoff", "code": "abc"},
        },
        "biosphere": {
            "Carbon dioxide": {"database": "ecoinvent-3.10-biosphere", "code": "co2", "unit": "kilogram"},
        },
    }

    changed = tfr._backfill(registry, "phoenix_ecoinvent-3.10.1-cutoff")

    assert registry["technosphere"]["steel production"]["unit"] == "kilogram"
    assert registry["biosphere"]["Carbon dioxide"]["categories"] == ["air"]
    assert changed == 2
