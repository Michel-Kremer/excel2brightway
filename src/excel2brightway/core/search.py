"""
search.py
=========
Kandidatensuche fuer nicht eindeutig aufgeloeste Fluesse.

Statt der Volltextsuche von bw2data wird ein eigener Index im Speicher
aufgebaut: ex2bw laedt die Datensaetze der Zieldatenbanken fuer den exakten
Abgleich ohnehin, die Suche braucht so keinen Brightway-Suchindex, aendert
nichts am Projekt und laesst sich ohne Brightway testen.

Ablauf:
  1. Kandidaten sammeln: Datensaetze, die mindestens ein Wort mit der Anfrage
     teilen (Woerter mit Tippfehlern werden ueber aehnliche Woerter des
     Vokabulars ergaenzt).
  2. Re-Ranking per Code (score_record): Wortaehnlichkeit gegen reference
     product und name, Einheit, Standort, Bonus fuer Maerkte.
"""

import difflib
from collections import Counter

from .model import Candidate
from .text import normalize, normalize_unit, tokens

MARKET_PREFIXES = ("market for ", "market group for ")

# Gewichte des Re-Rankings. Bewusst an einer Stelle, damit sie sich am
# Testsatz kalibrieren lassen.
WEIGHTS = {
    "unit_match": 0.15,
    "unit_mismatch": -0.35,
    "location_match": 0.10,
    "location_priority": (0.05, 0.04, 0.03),
    "market": 0.05,
    "category_match": 0.10,
}

_MAX_PREFILTER = 400
_SCORE_ALL_BELOW = 2000


def _text_similarity(query_tokens, text) -> float:
    if not text:
        return 0.0
    text_tokens = tokens(text)
    if not query_tokens or not text_tokens:
        return difflib.SequenceMatcher(None, normalize(" ".join(query_tokens)), normalize(text)).ratio() * 0.5
    q, t = set(query_tokens), set(text_tokens)
    dice = 2 * len(q & t) / (len(q) + len(t))
    ratio = difflib.SequenceMatcher(None, " ".join(sorted(q)), " ".join(sorted(t))).ratio()
    return 0.5 * dice + 0.5 * ratio


def score_record(record, query_tokens, unit=None, location=None, categories=(), location_priority=("GLO", "RER", "RoW")):
    """Punktzahl eines Datensatzes fuer eine Anfrage. Gibt (score, reason) zurueck."""
    text_score = max(_text_similarity(query_tokens, record.product), _text_similarity(query_tokens, record.name))
    score = text_score
    reasons = [f"Text {text_score:.2f}"]

    if unit and record.unit:
        if normalize_unit(unit) == normalize_unit(record.unit):
            score += WEIGHTS["unit_match"]
            reasons.append("Einheit passt")
        else:
            score += WEIGHTS["unit_mismatch"]
            reasons.append(f"Einheit {record.unit} statt {unit}")

    if record.kind == "technosphere":
        if location and record.location:
            if str(location).lower() == str(record.location).lower():
                score += WEIGHTS["location_match"]
                reasons.append("Standort passt")
        elif not location and record.location in location_priority:
            bonus = WEIGHTS["location_priority"]
            index = list(location_priority).index(record.location)
            score += bonus[min(index, len(bonus) - 1)]
        if normalize(record.name).startswith(MARKET_PREFIXES):
            score += WEIGHTS["market"]
            reasons.append("Markt")
    elif categories and record.categories:
        wanted = tuple(c.lower() for c in categories)
        have = tuple(c.lower() for c in record.categories)
        if wanted == have:
            score += WEIGHTS["category_match"]
            reasons.append("Kategorie passt")
        elif wanted[0] == have[0]:
            score += WEIGHTS["category_match"] / 2

    return score, ", ".join(reasons)


class SearchIndex:
    """Wort-Index ueber eine Liste von Record-Objekten."""

    def __init__(self, records):
        self.records = list(records)
        self._postings = {}
        for index, record in enumerate(self.records):
            words = set(tokens(record.name)) | set(tokens(record.product))
            for word in words:
                self._postings.setdefault(word, []).append(index)
        self._vocabulary = list(self._postings)

    def _expand(self, word):
        if word in self._postings:
            return [word]
        return difflib.get_close_matches(word, self._vocabulary, n=3, cutoff=0.8)

    def _prefilter(self, query_tokens):
        if len(self.records) <= _SCORE_ALL_BELOW:
            return range(len(self.records))
        counts = Counter()
        for word in query_tokens:
            for expanded in self._expand(word):
                for index in self._postings.get(expanded, ()):
                    counts[index] += 1
        return [index for index, _ in counts.most_common(_MAX_PREFILTER)]

    def search(self, query, unit=None, location=None, categories=(), kind=None, limit=5,
               location_priority=("GLO", "RER", "RoW")):
        """Die besten `limit` Kandidaten (Candidate) fuer einen Suchtext."""
        query_tokens = tokens(query)
        if not query_tokens and not query:
            return []
        scored = []
        for index in self._prefilter(query_tokens):
            record = self.records[index]
            if kind and record.kind != kind:
                continue
            score, reason = score_record(
                record, query_tokens, unit=unit, location=location, categories=categories,
                location_priority=location_priority,
            )
            scored.append(Candidate(record=record, score=score, reason=reason))
        scored.sort(key=lambda c: (-c.score, c.record.name or "", c.record.location or "", c.record.code))
        return scored[:limit]


def merge_candidates(lists, limit=5, min_score=None):
    """Fuehrt Kandidatenlisten mehrerer Indizes zusammen (ohne Duplikate, beste zuerst)."""
    best = {}
    for candidates in lists:
        for candidate in candidates:
            if min_score is not None and candidate.score < min_score:
                continue
            key = candidate.record.key
            if key not in best or candidate.score > best[key].score:
                best[key] = candidate
    merged = sorted(best.values(), key=lambda c: (-c.score, c.record.name or "", c.record.code))
    return merged[:limit]
