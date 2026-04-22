"""Regex-Patterns und Marker für Content-/PII-Regeln (Phase 2 A.1).

Zentrale Sammlung wiederverwendbarer Patterns — damit einzelne Rule-Dateien
nicht jeweils eigene Regex duplizieren. Patterns sind bewusst konservativ
gewählt, um False-Positives auf Produktcodes und Mess-Identifier zu
begrenzen. Strengere Validierung (Prüfziffer, Kontext-Heuristik) passiert
je Rule.
"""

from __future__ import annotations

import re
from typing import Iterable


# ── PII-Patterns ────────────────────────────────────────────
EMAIL_PATTERN = re.compile(
    r"[\w\.\-+]+@[\w\-]+(?:\.[\w\-]+)+",
)

# IBAN: 2 Letters + 2 Digits + 11-30 alphanum. Lockere Match, strenge
# Prüfung über iban_checksum_ok().
IBAN_PATTERN = re.compile(
    r"\b[A-Z]{2}\d{2}[A-Z0-9]{11,30}\b",
)

# Österreichische SV-Nummer: 4 Stellen (Laufnummer + Prüfziffer) + 6 Stellen
# (Geburtsdatum TTMMJJ). Optional mit Leerzeichen.
SVNR_AT_PATTERN = re.compile(
    r"\b(\d{3})(\d)\s?(\d{6})\b",
)

# Telefon DE/AT: sehr konservativ — verlangt Plus-Präfix oder 0-Präfix und
# Mindestlänge. Reduziert Kollision mit Produktcodes.
PHONE_PATTERN = re.compile(
    r"(?<!\d)(?:\+(?:49|43)|0)[\s\-/]?\d{2,4}[\s\-/]?\d{3}[\s\-/]?\d{3,}",
)


def iban_checksum_ok(iban: str) -> bool:
    """Validiert die mod-97-Prüfsumme einer IBAN (ISO 13616).

    Akzeptiert IBANs mit Leerzeichen. Gibt False zurück, wenn die Struktur
    nicht passt oder die Prüfsumme nicht stimmt.
    """
    cleaned = "".join(iban.split()).upper()
    if len(cleaned) < 15 or len(cleaned) > 34:
        return False
    if not cleaned[:2].isalpha() or not cleaned[2:4].isdigit():
        return False
    rearranged = cleaned[4:] + cleaned[:4]
    # Convert letters to numbers (A=10 … Z=35)
    digits = []
    for ch in rearranged:
        if ch.isdigit():
            digits.append(ch)
        elif "A" <= ch <= "Z":
            digits.append(str(ord(ch) - 55))
        else:
            return False
    try:
        return int("".join(digits)) % 97 == 1
    except ValueError:
        return False


def svnr_at_checksum_ok(match: re.Match) -> bool:
    """Prüfziffer der österreichischen Sozialversicherungsnummer.

    Gewichte (für die ersten 3 Ziffern der Laufnummer + 6 Ziffern
    Geburtsdatum): 3, 7, 9, 5, 8, 4, 2, 1, 6.
    """
    serial = match.group(1)
    check_digit = match.group(2)
    birthday = match.group(3)
    numbers = serial + birthday
    if len(numbers) != 9 or not numbers.isdigit() or not check_digit.isdigit():
        return False
    weights = (3, 7, 9, 5, 8, 4, 2, 1, 6)
    total = sum(int(numbers[i]) * weights[i] for i in range(9))
    return (total % 11) == int(check_digit)


# ── Encoding / Whitespace / Number-as-Text ─────────────────
ZERO_WIDTH_CHARS: frozenset[str] = frozenset({
    "​",  # Zero-Width Space
    "‌",  # Zero-Width Non-Joiner
    "‍",  # Zero-Width Joiner
    "﻿",  # BOM / Zero-Width No-Break Space
})

# Typische Mojibake-Muster (UTF-8 als Latin-1 gelesen → UTF-8 re-encoded)
ENCODING_MOJIBAKE_MARKERS: tuple[str, ...] = (
    "Ã¤", "Ã¶", "Ã¼", "Ã„", "Ã–", "Ãœ", "ÃŸ",
    "â€™", "â€œ", "â€\x9d", "â€“", "â€”",
    "Â ", "Â©", "Â®",
)

NUMBER_AS_TEXT_PATTERN = re.compile(
    r"^\s*-?\d{1,}(?:[.,]\d+)?\s*$",
)

# Typische Datumsformate — wenn mehrere parallel in einer Spalte auftreten,
# ist das ein Indiz für Inkonsistenz.
DATE_FORMAT_PATTERNS: tuple[tuple[str, re.Pattern], ...] = (
    ("DD.MM.YYYY",   re.compile(r"^\d{1,2}\.\d{1,2}\.\d{2,4}$")),
    ("YYYY-MM-DD",   re.compile(r"^\d{4}-\d{1,2}-\d{1,2}$")),
    ("MM/DD/YYYY",   re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4}$")),
    ("DD-MM-YYYY",   re.compile(r"^\d{1,2}-\d{1,2}-\d{2,4}$")),
    ("DD Mon YYYY",  re.compile(r"^\d{1,2}\s+[A-Za-zÄÖÜäöü]{3,}\s+\d{2,4}$")),
)


def contains_zero_width(s: str) -> bool:
    return any(c in ZERO_WIDTH_CHARS for c in s)


def contains_mojibake(s: str) -> bool:
    return any(m in s for m in ENCODING_MOJIBAKE_MARKERS)


def has_leading_or_trailing_whitespace(s: str) -> bool:
    return bool(s) and (s != s.strip())


def looks_like_number_text(s: str) -> bool:
    """True, wenn der String einem ganzen/dezimalen Zahlenwert entspricht."""
    return bool(NUMBER_AS_TEXT_PATTERN.match(s)) if isinstance(s, str) else False


def date_format_key(s: str) -> str | None:
    """Liefert den Namen des erkannten Datumsformats oder None."""
    if not isinstance(s, str):
        return None
    s = s.strip()
    for name, pat in DATE_FORMAT_PATTERNS:
        if pat.match(s):
            return name
    return None


# ── Free-Text-in-Enum-Helper ───────────────────────────────
def normalize_enum_candidate(s: str) -> str:
    """Normalisierung für Fuzzy-Match auf Enum-Werte."""
    if not isinstance(s, str):
        return ""
    return " ".join(s.strip().lower().split())


def group_by_normalized(values: Iterable[str]) -> dict[str, list[str]]:
    """Gruppiert Rohwerte nach ihrer Normalform. Gruppen mit mehr als einer
    eindeutigen Variante sind Kandidaten für Free-Text-in-Enum-Findings.
    """
    groups: dict[str, list[str]] = {}
    for v in values:
        key = normalize_enum_candidate(v)
        if not key:
            continue
        groups.setdefault(key, []).append(v)
    return {k: vs for k, vs in groups.items() if len({x.strip() for x in vs}) > 1}
