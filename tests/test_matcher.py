"""
Tests fuer excel2brightway.matcher: interne Aufloesung, Einheiten-Mismatch,
resolved-Gating pro Datenbank, Formel-/Parameter-Durchreichung bis in
resolved/<db>.yaml.
"""

import yaml

from excel2brightway import matcher

BLANK = [None]
EMPTY_REGISTRY = {"technosphere": {}, "biosphere": {}}


def _write_registry(tmp_path, data=None):
    path = tmp_path / "flow_registry.yaml"
    path.write_text(yaml.dump(data or EMPTY_REGISTRY), encoding="utf-8")
    return path


def test_persist_new_registry_entries_after_empty_flow_style_section(tmp_path):
    # yaml.dump() (bzw. die REGISTRY_TEMPLATE in config.py, bzw. tidy_flow_registry.py
    # nach dem Bereinigen einer leergewordenen Sektion) schreibt ein leeres dict als
    # Flow-Stil ('biosphere: {}'). _persist_new_registry_entries() haengt danach per
    # reinem Text-Append neue Eintraege an - das darf die Datei nicht kaputt machen
    # (Regressionstest fuer genau diesen Fall).
    registry = _write_registry(tmp_path)  # EMPTY_REGISTRY -> "technosphere: {}\nbiosphere: {}\n"

    matcher._persist_new_registry_entries(registry, {
        "biosphere": {
            "Water": {
                "database": "ecoinvent-3.10-biosphere", "code": "w1",
                "categories": ["water"], "unit": "cubic meter",
            },
        },
    })

    reloaded = yaml.safe_load(registry.read_text(encoding="utf-8"))
    assert reloaded["biosphere"]["Water"]["code"] == "w1"
    assert reloaded["technosphere"] == {}


def test_internal_resolution_without_registry(workbook_factory, tmp_path):
    workbook_factory(
        "cluster.xlsx",
        {
            "Data": [
                ["Database", "DB1"],
                BLANK,
                ["Activity", "Act A"],
                ["code", "a"],
                ["unit", "kg"],
                ["location", "GLO"],
                ["production amount", 1],
                ["Exchanges"],
                ["name", "amount", "unit", "database", "location", "type"],
                ["Act A", 1, "kg", "DB1", "GLO", "production"],
                BLANK,
                ["Activity", "Act B"],
                ["code", "b"],
                ["unit", "kg"],
                ["location", "GLO"],
                ["production amount", 1],
                ["Exchanges"],
                ["name", "amount", "unit", "database", "location", "type"],
                ["Act B", 1, "kg", "DB1", "GLO", "production"],
                ["Act A", 2, "kg", "DB1", "GLO", "technosphere"],
            ],
        },
    )
    registry = _write_registry(tmp_path)
    out = tmp_path / "unresolved.yaml"
    resolved_dir = tmp_path / "resolved"

    parsed_activities, unresolved = matcher.run(
        clusters=tmp_path, registry=registry, out=out,
        ecoinvent_fallback=False, resolved_dir=resolved_dir,
    )

    assert unresolved == []
    act_b = parsed_activities[("DB1", "b")]
    tech = next(e for e in act_b["exchanges"] if e["type"] == "technosphere")
    assert tech["input_database"] == "DB1"
    assert tech["input_code"] == "a"
    assert (resolved_dir / "DB1.yaml").exists()


def test_unit_mismatch_is_reported(workbook_factory, tmp_path):
    workbook_factory(
        "cluster.xlsx",
        {
            "Data": [
                ["Database", "DB1"],
                BLANK,
                ["Activity", "Act A"],
                ["code", "a"],
                ["unit", "kg"],
                ["location", "GLO"],
                ["production amount", 1],
                ["Exchanges"],
                ["name", "amount", "unit", "database", "location", "type"],
                ["Act A", 1, "kg", "DB1", "GLO", "production"],
                ["external flow", 5, "litre", "ext-db", "GLO", "technosphere"],
            ],
        },
    )
    registry = _write_registry(tmp_path, {
        "technosphere": {"external flow": {"database": "ext-db", "code": "x1", "unit": "kilogram"}},
        "biosphere": {},
    })
    out = tmp_path / "unresolved.yaml"

    _, unresolved = matcher.run(
        clusters=tmp_path, registry=registry, out=out,
        ecoinvent_fallback=False, resolved_dir=None,
    )

    assert len(unresolved) == 1
    assert unresolved[0]["reason"] == "unit_mismatch"
    assert unresolved[0]["expected_unit"] == "kilogram"


def test_resolved_file_only_written_when_fully_resolved(workbook_factory, tmp_path):
    workbook_factory(
        "cluster.xlsx",
        {
            "Data": [
                ["Database", "DB1"],
                BLANK,
                ["Activity", "Act A"],
                ["code", "a"],
                ["unit", "kg"],
                ["location", "GLO"],
                ["production amount", 1],
                ["Exchanges"],
                ["name", "amount", "unit", "database", "location", "type"],
                ["Act A", 1, "kg", "DB1", "GLO", "production"],
                ["unknown flow", 1, "kg", "ext-db", "GLO", "technosphere"],
            ],
        },
    )
    registry = _write_registry(tmp_path)
    out = tmp_path / "unresolved.yaml"
    resolved_dir = tmp_path / "resolved"

    _, unresolved = matcher.run(
        clusters=tmp_path, registry=registry, out=out,
        ecoinvent_fallback=False, resolved_dir=resolved_dir,
    )

    assert len(unresolved) == 1
    assert not (resolved_dir / "DB1.yaml").exists()


def test_warnings_out_collects_exchange_warnings(workbook_factory, tmp_path):
    workbook_factory(
        "cluster.xlsx",
        {
            "Data": [
                ["Database", "DB1"],
                BLANK,
                ["Activity", "Act A"],
                ["code", "a"],
                ["unit", "kg"],
                ["location", "GLO"],
                ["production amount", 1],
                ["Exchanges"],
                ["name", "amount", "unit", "database", "location", "type"],
                ["Act A", 1, "kg", "DB1", "GLO", "production"],
                ["Broken Flow", None, "kg", "DB1", "GLO", "technosphere"],
            ],
        },
    )
    registry = _write_registry(tmp_path)
    out = tmp_path / "unresolved.yaml"
    warnings_out = tmp_path / "load_warnings.yaml"

    matcher.run(
        clusters=tmp_path, registry=registry, out=out,
        ecoinvent_fallback=False, resolved_dir=None, warnings_out=warnings_out,
    )

    assert warnings_out.exists()
    load_warnings = yaml.safe_load(warnings_out.read_text(encoding="utf-8"))
    assert len(load_warnings) == 1
    assert load_warnings[0]["kind"] == "missing_amount"
    assert load_warnings[0]["source_file"] == "cluster.xlsx"


def test_formula_and_parameters_reach_resolved_yaml(workbook_factory, tmp_path):
    workbook_factory(
        "cluster.xlsx",
        {
            "Data": [
                ["Database", "DB1"],
                BLANK,
                ["Activity", "Act A"],
                ["code", "a"],
                ["unit", "kg"],
                ["location", "GLO"],
                ["production amount", 1],
                ["Parameters"],
                ["name", "amount", "formula"],
                ["my_param", 2, None],
                ["Exchanges"],
                ["name", "amount", "unit", "database", "location", "type", "formula"],
                ["Act A", 1, "kg", "DB1", "GLO", "production", None],
                ["Act A helper", None, "kg", "DB1", "GLO", "technosphere", "my_param * 2"],
                BLANK,
                ["Activity", "Act A helper"],
                ["code", "helper"],
                ["unit", "kg"],
                ["location", "GLO"],
                ["production amount", 1],
                ["Exchanges"],
                ["name", "amount", "unit", "database", "location", "type"],
                ["Act A helper", 1, "kg", "DB1", "GLO", "production"],
            ],
        },
    )
    registry = _write_registry(tmp_path)
    out = tmp_path / "unresolved.yaml"
    resolved_dir = tmp_path / "resolved"

    _, unresolved = matcher.run(
        clusters=tmp_path, registry=registry, out=out,
        ecoinvent_fallback=False, resolved_dir=resolved_dir,
    )

    assert unresolved == []
    written = yaml.safe_load((resolved_dir / "DB1.yaml").read_text(encoding="utf-8"))
    act_a = next(a for a in written["activities"] if a["code"] == "a")
    assert act_a["parameters"] == [{"name": "my_param", "amount": 2}]
    tech = next(e for e in act_a["exchanges"] if e["type"] == "technosphere")
    assert tech["formula"] == "my_param * 2"
    assert tech["amount"] == 0.0
