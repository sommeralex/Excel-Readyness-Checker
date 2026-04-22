"""Internationalisierung (i18n) für den Excel-Reifecheck.

Stellt eine zentrale _()-Funktion bereit, die Übersetzungsschlüssel
in die aktuelle Sprache auflöst. Standardsprache ist Deutsch.

Verwendung:
    from excel_checker.i18n import _, set_language

    set_language("en")        # auf Englisch umschalten
    print(_("cli.score", score=85))  # -> "Score: 85/100"
"""

from __future__ import annotations

import threading

_lock = threading.Lock()
_current_lang: str = "de"
_translations: dict[str, dict[str, str]] = {}


def set_language(lang: str) -> None:
    """Setzt die aktive Sprache (z.B. 'de', 'en')."""
    global _current_lang
    with _lock:
        _current_lang = lang


def get_language() -> str:
    """Gibt den aktiven Sprachcode zurück."""
    return _current_lang


def register(lang: str, strings: dict[str, str]) -> None:
    """Registriert ein Übersetzungs-Dict für eine Sprache."""
    with _lock:
        if lang not in _translations:
            _translations[lang] = {}
        _translations[lang].update(strings)


def _(key: str, **kwargs) -> str:
    """Übersetzt einen Schlüssel in die aktive Sprache.

    Optionale Keyword-Argumente werden per .format() eingesetzt.
    Fallback: Deutsch → Schlüssel selbst.
    """
    lang_dict = _translations.get(_current_lang, {})
    text = lang_dict.get(key)
    if text is None:
        # Fallback auf Deutsch
        text = _translations.get("de", {}).get(key)
    if text is None:
        # Schlüssel nirgends gefunden
        return key
    if kwargs:
        try:
            text = text.format(**kwargs)
        except (KeyError, IndexError):
            pass
    return text


def _load_all() -> None:
    """Lädt alle verfügbaren Locale-Module."""
    from excel_checker.locale import de, en
    register("de", de.STRINGS)
    register("en", en.STRINGS)


# Beim Import automatisch laden
_load_all()
