"""
Normalisierung von Namen, Einheiten und Datenbank-Labels sowie Wort-Zerlegung
fuer die Kandidatensuche. Reine Funktionen ohne Abhaengigkeiten.
"""

import re

UNKNOWN_MARKERS = (None, "", "(Unknown)", "(unknown)")

_TOKEN_RE = re.compile(r"[a-z0-9]+")
_DB_NAME_NORMALIZE_RE = re.compile(r"[^a-z0-9]")

# Woerter, die in LCA-Namen sehr haeufig sind und beim Wortvergleich wenig
# unterscheiden. "market" bleibt bewusst drin (eigener Bonus beim Ranking).
STOPWORDS = frozenset({
    "for", "of", "at", "the", "to", "and", "in", "from", "with", "by", "a", "an",
    "on", "per", "as",
})

UNIT_ALIASES = {
    "kg": "kilogram", "kilogram": "kilogram", "kilograms": "kilogram",
    "g": "gram", "gram": "gram", "grams": "gram",
    "t": "metric ton", "ton": "metric ton", "tonne": "metric ton", "metric ton": "metric ton",
    "kwh": "kilowatt hour", "kilowatt hour": "kilowatt hour", "kilowatt-hour": "kilowatt hour",
    "mj": "megajoule", "megajoule": "megajoule", "mega joule": "megajoule",
    "kj": "kilojoule", "kilojoule": "kilojoule", "kilo joule": "kilojoule",
    "m3": "cubic meter", "m³": "cubic meter", "cubic meter": "cubic meter", "cubic metre": "cubic meter",
    "km": "kilometer", "kilometer": "kilometer", "kilometre": "kilometer",
    "l": "litre", "liter": "litre", "litre": "litre",
    "unit": "unit", "piece": "unit", "item": "unit", "p": "unit",
    "tkm": "ton kilometer", "ton kilometer": "ton kilometer", "metric ton*km": "ton kilometer",
    "hour": "hour", "h": "hour",
    "m2": "square meter", "m²": "square meter", "square meter": "square meter", "square metre": "square meter",
    "m2a": "square meter-year", "m2*a": "square meter-year", "square meter-year": "square meter-year",
}


def clean(value):
    """Leere Zellen und bw2io-Platzhalter ('(Unknown)') werden zu None, Strings getrimmt."""
    if isinstance(value, str):
        value = value.strip()
    return None if value in UNKNOWN_MARKERS else value


def normalize(name) -> str:
    """Normalisiert Flow-/Aktivitaetsnamen fuer den exakten Abgleich (lowercase, Leerraum vereinheitlicht)."""
    if name is None:
        return ""
    return " ".join(str(name).strip().lower().split())


def normalize_unit(unit):
    """Normalisiert gaengige Einheiten-Schreibweisen (kg == kilogram, kWh == kilowatt hour, ...)."""
    if not unit:
        return None
    key = " ".join(str(unit).strip().lower().split())
    return UNIT_ALIASES.get(key, key)


def units_compatible(a, b) -> bool:
    """True, wenn beide Einheiten nach Normalisierung gleich sind oder eine davon fehlt."""
    na, nb = normalize_unit(a), normalize_unit(b)
    if na is None or nb is None:
        return True
    return na == nb


def normalize_db_name(value) -> str:
    """Datenbanknamen/-Labels ohne Satzzeichen vergleichen ('ecoinvent 3.12' ~ 'ecoinvent-3.12-cutoff')."""
    return _DB_NAME_NORMALIZE_RE.sub("", str(value or "").lower())


def tokens(text) -> list:
    """Zerlegt einen Namen in Woerter (lowercase, ohne Satzzeichen und Fuellwoerter)."""
    if not text:
        return []
    return [t for t in _TOKEN_RE.findall(str(text).lower()) if t not in STOPWORDS]


def split_categories(value) -> tuple:
    """'air::urban air close to ground' -> ('air', 'urban air close to ground'); Tupel/Listen bleiben erhalten."""
    value = clean(value)
    if value is None:
        return ()
    if isinstance(value, (list, tuple)):
        return tuple(str(v).strip() for v in value if clean(v) is not None)
    return tuple(part.strip() for part in str(value).split("::") if part.strip())


def to_number(value):
    """Zahl aus einer Zelle (float/int oder numerischer String, auch mit Dezimalkomma); sonst None."""
    value = clean(value)
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(" ", "")
    if "," in text and "." not in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None
