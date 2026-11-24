"""
importer.py
===========
Neue ecoinvent-Datenbank (ecoSpold02) in ein Brightway-Projekt einlesen.

Quelle ist ein Archiv (.7z oder .zip, wie von ecoinvent geliefert) oder ein
bereits entpackter Ordner mit `datasets/` (und `MasterData/`). Archive werden
neben die Archivdatei in einen gleichnamigen Ordner entpackt; ist der schon
vorhanden, wird er wiederverwendet.

Die Biosphere wird entweder aus dem Projekt genommen (wenn es eine Datenbank
mit dem Namen schon gibt) oder aus MasterData/ElementaryExchanges.xml neu
angelegt. Gibt es nach dem Verknuepfen ungelinkte Exchanges, wird nichts
geschrieben und eine eben angelegte Biosphere wieder entfernt.
"""

import re
import zipfile
from dataclasses import dataclass
from pathlib import Path

ARCHIVE_SUFFIXES = (".7z", ".zip")
_VERSION_RE = re.compile(r"(\d+\.\d+(?:\.\d+)?)")
_SYSTEM_MODELS = ("cutoff", "cut-off", "apos", "consequential", "consequential_longterm", "eol")


class ImportProblem(Exception):
    """Import nicht moeglich oder abgebrochen (Meldung fuer den Benutzer)."""


@dataclass
class ImportResult:
    project: str
    database: str
    activities: int
    biosphere: str
    biosphere_created: bool

    def summary(self):
        bio = f"'{self.biosphere}' neu angelegt" if self.biosphere_created else f"vorhandene '{self.biosphere}'"
        return (f"Datenbank '{self.database}' mit {self.activities} Aktivitaeten in Projekt "
                f"'{self.project}' geschrieben (Biosphere: {bio}).")


def suggest_names(source):
    """
    Namensvorschlag aus dem Datei-/Ordnernamen, z. B.
    'ecoinvent 3.11_cutoff_ecoSpold02.7z' -> ('ecoinvent-3.11-cutoff', 'ecoinvent-3.11-biosphere', '3.11').
    """
    stem = Path(source).name
    for suffix in ARCHIVE_SUFFIXES:
        if stem.lower().endswith(suffix):
            stem = stem[: -len(suffix)]
    version_match = _VERSION_RE.search(stem)
    lowered = stem.lower()
    model = next((m for m in _SYSTEM_MODELS if m in lowered), None)
    if "ecoinvent" in lowered and version_match:
        version = version_match.group(1)
        model = (model or "cutoff").replace("cut-off", "cutoff")
        return f"ecoinvent-{version}-{model}", f"ecoinvent-{version}-biosphere", version
    clean = re.sub(r"[^A-Za-z0-9.]+", "-", stem).strip("-") or "neue-datenbank"
    return clean, f"{clean}-biosphere", version_match.group(1) if version_match else None


def extract_archive(source, log=print):
    """Archiv entpacken (oder Ordner direkt nehmen). Gibt den Ordner zurueck."""
    source = Path(source)
    if source.is_dir():
        return source
    if not source.is_file():
        raise ImportProblem(f"{source} gibt es nicht.")
    suffix = source.suffix.lower()
    if suffix not in ARCHIVE_SUFFIXES:
        raise ImportProblem(f"{source.name}: nur .7z, .zip oder ein Ordner werden unterstuetzt.")
    target = source.with_suffix("")
    if target.is_dir() and any(target.iterdir()):
        log(f"Bereits entpackt, verwende {target}")
        return target
    log(f"Entpacke {source.name} nach {target} ...")
    target.mkdir(parents=True, exist_ok=True)
    if suffix == ".zip":
        with zipfile.ZipFile(source) as archive:
            archive.extractall(target)
    else:
        try:
            import py7zr
        except ImportError as exc:
            raise ImportProblem("Fuer .7z-Archive wird das Paket 'py7zr' benoetigt.") from exc
        with py7zr.SevenZipFile(source, mode="r") as archive:
            archive.extractall(path=target)
    return target


def find_ecospold_dirs(folder):
    """(datasets-Ordner, ElementaryExchanges.xml oder None) im entpackten Ordner suchen."""
    folder = Path(folder)
    candidates = [folder / "datasets"] + sorted(p for p in folder.rglob("datasets") if p.is_dir())
    datasets = next((d for d in candidates if d.is_dir() and any(d.glob("*.spold"))), None)
    if datasets is None:
        if any(folder.glob("*.spold")):
            datasets = folder
        else:
            raise ImportProblem(f"In {folder} wurden keine ecoSpold02-Dateien (*.spold) gefunden.")
    masterdata = datasets.parent / "MasterData" / "ElementaryExchanges.xml"
    if not masterdata.is_file():
        masterdata = next(iter(sorted(folder.rglob("ElementaryExchanges.xml"))), None)
    return datasets, masterdata


def import_ecospold(source, project, database, biosphere, version=None, log=print):
    """Archiv/Ordner `source` als `database` in `project` schreiben (Biosphere `biosphere`)."""
    if not project or not database or not biosphere:
        raise ImportProblem("Projekt, Datenbankname und Biosphere muessen angegeben sein.")
    folder = extract_archive(source, log=log)
    datasets, masterdata = find_ecospold_dirs(folder)
    log(f"Datensaetze: {datasets}")

    import bw2data as bd
    import bw2io as bi
    from bw2io.importers.ecospold2_biosphere import Ecospold2BiosphereImporter

    bd.projects.set_current(project)
    if database in bd.databases:
        raise ImportProblem(f"Datenbank '{database}' gibt es in Projekt '{project}' schon - nichts geaendert.")

    biosphere_created = False
    if biosphere in bd.databases:
        log(f"Verwende vorhandene Biosphere '{biosphere}'.")
    else:
        if masterdata is None:
            raise ImportProblem(
                f"Biosphere '{biosphere}' gibt es nicht im Projekt und im Archiv fehlt "
                "MasterData/ElementaryExchanges.xml, um sie anzulegen."
            )
        log(f"Lege Biosphere '{biosphere}' aus {masterdata.name} an ...")
        bio = Ecospold2BiosphereImporter(name=biosphere, version=version or "3.9", filepath=masterdata)
        bio.apply_strategies()
        bio.write_database()
        biosphere_created = True
        log(f"'{biosphere}': {len(bd.Database(biosphere))} Fluesse.")

    try:
        log("Lese und verknuepfe Datensaetze (kann 20-60 min dauern) ...")
        # use_mp=False: Multiprocessing startet unter Windows aus dem Fenster heraus nicht zuverlaessig
        importer = bi.SingleOutputEcospold2Importer(
            str(datasets), database, biosphere_database_name=biosphere, use_mp=False,
        )
        importer.apply_strategies()
        unlinked = list(importer.unlinked)
        if unlinked:
            examples = "; ".join(f"{e.get('name')} ({e.get('type')})" for e in unlinked[:5])
            raise ImportProblem(
                f"{len(unlinked)} ungelinkte Exchanges - Datenbank wird nicht geschrieben. Beispiele: {examples}"
            )
        log("Schreibe Datenbank ...")
        importer.write_database()
    except BaseException:
        if biosphere_created and biosphere in bd.databases:
            del bd.databases[biosphere]
            log(f"Eben angelegte Biosphere '{biosphere}' wieder entfernt.")
        raise
    return ImportResult(project, database, len(bd.Database(database)), biosphere, biosphere_created)
