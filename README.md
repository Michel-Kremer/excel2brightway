# excel2brightway

Prueft LCI-/ecoinvent-Excel-Daten im nativen Brightway2-Excel-Format
(`bw2io.ExcelImporter`) auf Korrektheit - Flow-Aufloesung (intern /
`flow_registry.yaml` / ecoinvent, read-only) und Excel-Format - und schreibt
sie bei vollstaendiger Aufloesung nach Brightway, inklusive Parametern und
Formeln als echte, im Activity Browser editierbare `bw2data`-Parameter.

## Installation

```
pip install -e .
```

Fuer exakt reproduzierbare, verifiziert funktionierende Paketversionen statt
nur kompatibler Bereiche:

```
pip install -r requirements-lock.txt
```

`bw2data` MUSS in derselben Major-Version installiert sein, die auch das
Ziel-Brightway-Projekt zuletzt geschrieben hat (aktuell 3.6.x) - eine neuere
Major-Version loest beim ersten Projektzugriff eine automatische,
irreversible Migration am geteilten Brightway-Projektordner aus.

## Workspace

`excel2brightway` operiert auf einem **Workspace**-Ordner, der getrennt vom
Package-Code liegt:

```
<workspace>/
  *.xlsx                     Eingabe: LCI-Cluster-Excel-Dateien (von Hand gepflegt)
  output_excel2brightway/    Ausgabe: alles, was excel2brightway selbst schreibt
    flow_registry.yaml
    unresolved.yaml
    resolved/
```

Workspace-Aufloesung, in dieser Reihenfolge: `--workspace PATH` > Umgebungs-
variable `EX2BW_WORKSPACE` > aktuelles Arbeitsverzeichnis.

## Nutzung

```
cd <workspace>
ex2bw-check              # Stufe 1: Testen & Abgleichen -> output_excel2brightway/resolved/*.yaml
ex2bw-load                # Stufe 2: resolved/*.yaml -> Brightway (fragt vor dem Schreiben nach)
ex2bw-run                 # interaktives Menue fuer beide Stufen
ex2bw-tidy-registry        # flow_registry.yaml aufraeumen (Dedupe + unit-Backfill)
```

Jeder Befehl akzeptiert `--workspace PATH`, um nicht aus dem Workspace heraus
aufgerufen zu werden.

## Entwicklung

```
pip install -e ".[test]"
pytest
```

`tests/fixtures/excel_formats/` enthaelt bw2io-Beispiel-Excel-Dateien fuer
die Tests (Skip/Cutoff/Parameter/Formel-Handling). Kein Test schreibt nach
Brightway - das beruehrt ein echtes, ggf. geteiltes Brightway-Projekt und
bleibt manuelle Verifikation.
