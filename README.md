# Excel-Reifecheck

Automatisierter Datenreife-Check für Excel-Dateien. Findet Anti-Patterns (verbundene
Zellen, Magic Numbers, implizite Farblogik, …), die saubere Weiterverarbeitung und
KI-Integration verhindern. Gibt einen HTML-Report mit Health-Score und konkreten
Empfehlungen aus.

- **CLI**: Batch-Analyse einer `.xlsx`
- **Flask-Webapp**: Upload-UI mit SSE-Progress und optionalem LLM-Kommentar

## Setup (pro PC einmalig)

Python-venvs sind **nicht portierbar** — auf jedem PC neu bauen.

### Voraussetzungen
- Python ≥ 3.10 (getestet mit 3.12.10 von python.org). Inkscape's mitgeliefertes
  MinGW-Python funktioniert **nicht** — dort fehlen die Standard-Windows-Wheels für
  C-Erweiterungen.
- Git Bash oder ein anderes Bash unter Windows (die Commands unten sind POSIX-Syntax)

### Install
```bash
cd /c/Users/<you>/ExcelChecker
python -m venv .venv
source .venv/Scripts/activate
pip install -e .
```

`.env` mit `CLAUDE_API_KEY`, `AZURE_ENDPOINT`, `CLAUDE_MODEL`, `APP_VERSION` anlegen
(wird von [`excel_checker/__init__.py`](excel_checker/__init__.py) via `python-dotenv`
geladen). Ohne funktioniert die Kern-Analyse trotzdem — nur die LLM-Kommentierung
ist deaktiviert.

### Starten

```bash
# Webapp (http://127.0.0.1:5000)
PYTHONIOENCODING=utf-8 python -m excel_checker.webapp

# CLI
excel-reifecheck datei.xlsx --html report.html --open --lang de
```

`PYTHONIOENCODING=utf-8` ist auf Windows nötig, damit das `⚡`-Emoji im Startup-Banner
nicht an der cp1252-Konsole scheitert.

## Projektstruktur

- [`excel_checker/engine.py`](excel_checker/engine.py) — Analysekern, `analyze_with_progress`-Generator
- [`excel_checker/rules/`](excel_checker/rules/) — Regel-Klassen, kategorisiert nach
  Struktur / Formeln / Volumen / Implizites Wissen
- [`excel_checker/webapp.py`](excel_checker/webapp.py) — Flask-Routes, SSE-Stream, Template `UPLOAD_PAGE`
- [`excel_checker/_upload_page.html`](excel_checker/_upload_page.html) — Sidecar-Template (wird beim Modul-Import gelesen)
- [`excel_checker/_webapp_bytecode.pyc`](excel_checker/_webapp_bytecode.pyc) — Safety-Net-Backup, siehe unten
- [`excel_checker/report.py`](excel_checker/report.py) — HTML-Report-Generator
- [`excel_checker/llm_analysis.py`](excel_checker/llm_analysis.py) — optionale Claude-API-Anbindung für narrative Deutung
- [`excel_checker/i18n.py`](excel_checker/i18n.py) + [`excel_checker/locale/`](excel_checker/locale/) — DE/EN-Übersetzungen

## Disaster-Recovery

`_webapp_bytecode.pyc` liegt als kompilierte Kopie des intakten `webapp.py` neben der
Quelldatei und ist **versioniert** (Ausnahme-Regel in [`.gitignore`](.gitignore)).
Falls die Quelle erneut korrumpiert wird, lässt sich der Inhalt über
`marshal.load` aus dem Bytecode rekonstruieren:

```python
import marshal, dis
with open('excel_checker/_webapp_bytecode.pyc','rb') as f:
    f.read(16)           # pyc-Header überspringen
    code = marshal.load(f)
# code.co_consts enthält alle String-Literale (HTML-Templates etc.),
# code.co_consts[i] mit i entsprechend co_name='<funcname>' die Code-Objekte.
# dis.dis(code) zeigt den vollen Bytecode für manuelle Rekonstruktion.
```

Die Rekonstruktions-Session vom April 2026 (siehe Commits `173a251`, `8973880`) ist
der Referenzfall.

## Entwickeln

Stand der Dependency-Liste ist [`pyproject.toml`](pyproject.toml). Neue Dependencies
dort hinzufügen, dann `pip install -e .` erneut laufen lassen.

Unit-Tests liegen im Root als `test_*.py` und lassen sich mit `pytest` ausführen.
