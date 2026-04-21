"""Web-Oberfläche für den Excel-Reifecheck."""

from __future__ import annotations

import json
import traceback
import os
import re
import tempfile
import uuid
from urllib.parse import urlparse, unquote

import requests as http_requests
from flask import Flask, render_template_string, request, redirect, url_for, Response, session

from excel_checker.engine import analyze, analyze_with_progress
from excel_checker.models import WorkbookReport
from excel_checker.report import generate_html
from excel_checker import __version__, BUILD_TIMESTAMP
from excel_checker.llm_analysis import (
    get_default_api_key, get_default_endpoint, get_default_model, mask_key,
    test_api_key, analyze_with_llm,
)


def _get_version_line() -> str:
    return f"Version {__version__} (Build {BUILD_TIMESTAMP})"

app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # Max 100 MB

# Temporäre Speicher für Upload→SSE Übergabe
import threading
import queue as queue_mod

_pending_lock = threading.Lock()
_reports: dict[str, str] = {}
_report_data: dict[str, WorkbookReport] = {}
_sessions: dict[str, queue_mod.Queue] = {}


def _recover_from_backup():
    """Restoration shim — the original source for UPLOAD_PAGE and _start_analysis
    was lost on 2026-04-21 when webapp.py was accidentally truncated. The intact
    compiled bytecode was preserved in _webapp_bytecode.pyc next to this file.
    UPLOAD_PAGE is read from the _upload_page.html sidecar; _start_analysis is
    reconstituted from the backup bytecode's code object (too complex a nested
    closure to decompile cleanly by hand)."""
    import marshal, types
    from pathlib import Path
    pkg_dir = Path(__file__).parent
    with open(pkg_dir / '_upload_page.html', 'r', encoding='utf-8') as f:
        upload = f.read()
    with open(pkg_dir / '_webapp_bytecode.pyc', 'rb') as f:
        f.read(16)  # pyc header
        module_code = marshal.load(f)
    start_code = next(c for c in module_code.co_consts
                      if isinstance(c, types.CodeType) and c.co_name == '_start_analysis')
    func = types.FunctionType(start_code, globals(), name='_start_analysis')
    return upload, func


UPLOAD_PAGE, _start_analysis = _recover_from_backup()


ERROR_PAGE = """<!DOCTYPE html>
<html lang="de"><head><meta charset="UTF-8">
<title>Fehler</title>
<link rel="stylesheet" href="/static/theme.css">
<style>
  body { display: flex; align-items: center; justify-content: center;
         min-height: 100vh; }
  .box { background: var(--card); border: 1px solid #fecaca; border-radius: 12px;
         padding: 2rem; max-width: 500px; text-align: center; }
  .box h1 { color: var(--red); }
  a { color: var(--accent); text-decoration: none; }
</style>
<script>
  (function() {
    if (localStorage.getItem('theme') === 'dark') {
      document.documentElement.setAttribute('data-theme', 'dark');
    }
  })();
</script>
</head>
<body><div class="box">
  <h1>⚠️ {{ title }}</h1>
  <p style="margin:1rem 0;">{{ message }}</p>
  <a href="/">← Zurück zum Upload</a>
</div></body></html>"""


PRUEFREGELN_PAGE = """<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Prüfregeln – Excel-Reifecheck</title>
<link rel="stylesheet" href="/static/theme.css">
<style>
  /* Vars + Reset + body-Defaults jetzt in theme.css.
     Hier nur seitenspezifisches Padding (Original-Wert). */
  body { padding: 2rem; }
  .container { max-width: 960px; margin: 0 auto; }
  h1 { font-size: 1.8rem; margin-bottom: 0.5rem; }
  .subtitle { color: var(--muted); margin-bottom: 2rem; font-size: 1rem; }
  .back { display: inline-block; margin-bottom: 1.5rem; color: var(--accent);
           text-decoration: none; font-weight: 600; font-size: 0.9rem; }
  .back:hover { text-decoration: underline; }

  .category { margin-bottom: 2.5rem; }
  .category-header {
    display: flex; align-items: center; gap: 0.75rem;
    margin-bottom: 1rem; padding-bottom: 0.5rem;
    border-bottom: 2px solid var(--border);
  }
  .category-header .emoji { font-size: 1.5rem; }
  .category-header h2 { font-size: 1.2rem; font-weight: 700; }
  .category-header .count {
    font-size: 0.75rem; background: var(--accent); color: white;
    padding: 0.15rem 0.5rem; border-radius: 999px; font-weight: 600;
  }

  .rule-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 0.75rem; }
  .rule-card {
    background: var(--card); border: 1px solid var(--border);
    border-radius: 10px; padding: 1rem 1.1rem;
    transition: box-shadow 0.2s;
  }
  .rule-card:hover { box-shadow: 0 4px 16px rgba(0,0,0,0.06); }
  .rule-id { font-size: 0.7rem; font-weight: 700; color: var(--muted);
             text-transform: uppercase; letter-spacing: 0.05em; }
  .rule-name { font-size: 0.92rem; font-weight: 600; margin: 0.25rem 0; }
  .rule-desc { font-size: 0.8rem; color: var(--muted); line-height: 1.5; }
  .severity {
    display: inline-block; font-size: 0.65rem; font-weight: 700;
    padding: 0.1rem 0.4rem; border-radius: 4px; margin-top: 0.5rem;
    text-transform: uppercase; letter-spacing: 0.04em;
  }
  .severity--critical { background: #fef2f2; color: #dc2626; }
  .severity--error { background: #fff7ed; color: #ea580c; }
  .severity--warning { background: #fefce8; color: #ca8a04; }
  .severity--info { background: #eff6ff; color: #2563eb; }

  .why-section {
    background: var(--card); border: 1px solid var(--border);
    border-radius: 14px; padding: 1.5rem 1.75rem; margin-bottom: 2.5rem;
  }
  .why-section h2 { font-size: 1.1rem; margin-bottom: 1rem; }
  .why-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 1.5rem; }
  .why-col h3 { font-size: 0.95rem; margin-bottom: 0.5rem; display: flex; align-items: center; gap: 0.4rem; }
  .why-col ul { padding-left: 1.2rem; }
  .why-col li { font-size: 0.85rem; color: var(--muted); margin-bottom: 0.35rem; }
  @media (max-width: 640px) { .why-grid { grid-template-columns: 1fr; } }
  @media (max-width: 768px) {
    body { padding: 1rem; }
    .rule-grid { grid-template-columns: 1fr; }
  }
</style>
<script>
  // Dark-Mode Init (synchron, vor Body-Render, um FOUC zu vermeiden).
  (function() {
    if (localStorage.getItem('theme') === 'dark') {
      document.documentElement.setAttribute('data-theme', 'dark');
    }
  })();
</script>
</head>
<body>
<div class="container">
  <a href="/" class="back">← Zurück zur Startseite</a>
  <h1>🔍 Was wird geprüft – und warum?</h1>
  <p class="subtitle">26 automatisierte Regeln in 4 Kategorien decken die häufigsten Excel-Anti-Patterns auf.</p>

  <div class="why-section">
    <h2>Warum ist das wichtig?</h2>
    <div class="why-grid">
      <div class="why-col">
        <h3>👥 Für Teams &amp; Zusammenarbeit</h3>
        <ul>
          <li>Verborgenes Wissen (Farben, versteckte Blätter) geht verloren, wenn Kolleg:innen wechseln</li>
          <li>Zusammengeführte Zellen &amp; leere Trennzeilen machen Copy-Paste und Filtern unzuverlässig</li>
          <li>Inkonsistente IDs und Datentypen führen zu falschen SVERWEIS-Ergebnissen</li>
          <li>Fehlende Primärschlüssel machen Daten-Verknüpfungen unmöglich</li>
          <li>Volatile Funktionen &amp; Zirkelbezüge erzeugen schwer nachvollziehbare Rechenfehler</li>
        </ul>
      </div>
      <div class="why-col">
        <h3>🤖 Für KI &amp; Automatisierung</h3>
        <ul>
          <li>KI-Modelle können nur saubere, tabellarische Daten zuverlässig lesen</li>
          <li>Merged Cells, Farb-Logik &amp; Magic Numbers sind für KI unsichtbar</li>
          <li>Implizites Wissen in Kommentaren oder bedingter Formatierung ist nicht maschinenlesbar</li>
          <li>Zu große Dateien überschreiten KI-Kontext-Limits oder scheitern beim Import</li>
          <li>Saubere Kopfzeilen &amp; konsistente Typen sind Voraussetzung für jeden ETL-Prozess</li>
        </ul>
      </div>
    </div>
  </div>

  <div class="category">
    <div class="category-header">
      <span class="emoji">🏗️</span>
      <h2>Struktur &amp; Normalform</h2>
      <span class="count">7 Regeln</span>
    </div>
    <div class="rule-grid">
      <div class="rule-card">
        <div class="rule-id">STR-001</div>
        <div class="rule-name">Zusammengeführte Zellen</div>
        <div class="rule-desc">Erkennt verbundene Zellen, die Filtern, Sortieren und maschinelles Lesen verhindern.</div>
        <span class="severity severity--error">Error</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">STR-002</div>
        <div class="rule-name">Datentyp-Homogenität</div>
        <div class="rule-desc">Prüft, ob Spalten einheitliche Datentypen enthalten (1.&nbsp;Normalform).</div>
        <span class="severity severity--warning">Warning</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">STR-003</div>
        <div class="rule-name">Kopfzeilen-Erkennung</div>
        <div class="rule-desc">Validiert, ob eine saubere, einzelne Kopfzeile existiert – Grundlage jeder Datenverarbeitung.</div>
        <span class="severity severity--error">Error</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">STR-004</div>
        <div class="rule-name">Leere Trennzeilen/-spalten</div>
        <div class="rule-desc">Erkennt leere Zeilen/Spalten als Datentrenner – verhindert korrektes Erkennen zusammenhängender Bereiche.</div>
        <span class="severity severity--warning">Warning</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">STR-005</div>
        <div class="rule-name">ID-Konsistenz</div>
        <div class="rule-desc">Prüft, ob Identifikator-Spalten einheitlich formatiert sind (z.B. DEM-001 vs. DEM-2).</div>
        <span class="severity severity--warning">Warning</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">STR-006</div>
        <div class="rule-name">Fehlender Primärschlüssel</div>
        <div class="rule-desc">Erkennt Datentabellen ohne eindeutige Schlüsselspalte – macht Verknüpfungen und Deduplizierung unmöglich.</div>
        <span class="severity severity--critical">Critical</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">STR-007</div>
        <div class="rule-name">Freitext-IDs</div>
        <div class="rule-desc">Erkennt ID-Spalten mit Freitext-Werten, die weder sortierbar noch eindeutig sind.</div>
        <span class="severity severity--warning">Warning</span>
      </div>
    </div>
  </div>

  <div class="category">
    <div class="category-header">
      <span class="emoji">📐</span>
      <h2>Formeln &amp; Bezüge</h2>
      <span class="count">5 Regeln</span>
    </div>
    <div class="rule-grid">
      <div class="rule-card">
        <div class="rule-id">FRM-001</div>
        <div class="rule-name">Absolute vs. relative Bezüge</div>
        <div class="rule-desc">Prüft das Verhältnis von absoluten ($) zu relativen Zellbezügen – ein Indikator für Copy-Paste-Fehler.</div>
        <span class="severity severity--warning">Warning</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">FRM-002</div>
        <div class="rule-name">Volatile Funktionen</div>
        <div class="rule-desc">Erkennt Funktionen wie INDIRECT, OFFSET, NOW, die bei jedem Öffnen neu berechnet werden.</div>
        <span class="severity severity--warning">Warning</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">FRM-003</div>
        <div class="rule-name">SVERWEIS-Ketten</div>
        <div class="rule-desc">Erkennt übermäßigen SVERWEIS-Einsatz – ein Zeichen, dass Excel als Datenbank missbraucht wird.</div>
        <span class="severity severity--warning">Warning</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">FRM-004</div>
        <div class="rule-name">Zirkelbezüge</div>
        <div class="rule-desc">Sucht nach selbstreferenzierenden Formeln, die zu unvorhersehbaren Ergebnissen führen.</div>
        <span class="severity severity--error">Error</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">FRM-005</div>
        <div class="rule-name">Blattübergreifende Bezüge</div>
        <div class="rule-desc">Analysiert Abhängigkeiten zwischen Blättern und externen Dateien – Risiko bei Datei-Umzügen.</div>
        <span class="severity severity--info">Info</span>
      </div>
    </div>
  </div>

  <div class="category">
    <div class="category-header">
      <span class="emoji">📊</span>
      <h2>Volumen &amp; Performance</h2>
      <span class="count">4 Regeln</span>
    </div>
    <div class="rule-grid">
      <div class="rule-card">
        <div class="rule-id">VOL-001</div>
        <div class="rule-name">Zeilenanzahl</div>
        <div class="rule-desc">Prüft, ob das Datenvolumen die praktiablen Excel-Grenzen überschreitet.</div>
        <span class="severity severity--warning">Warning</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">VOL-002</div>
        <div class="rule-name">Formel-Dichte</div>
        <div class="rule-desc">Prüft den Anteil an Formeln – zu viele destabilisieren Excel und verlangsamen das Öffnen.</div>
        <span class="severity severity--warning">Warning</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">VOL-003</div>
        <div class="rule-name">Blattanzahl</div>
        <div class="rule-desc">Erkennt übermäßig viele Arbeitsblätter – ein Zeichen für fehlende Datenmodellierung.</div>
        <span class="severity severity--warning">Warning</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">VOL-004</div>
        <div class="rule-name">Dateigröße</div>
        <div class="rule-desc">Prüft die Dateigröße – große Dateien sind langsam, fehleranfällig und schwer per E-Mail zu teilen.</div>
        <span class="severity severity--critical">Critical</span>
      </div>
    </div>
  </div>

  <div class="category">
    <div class="category-header">
      <span class="emoji">🕵️</span>
      <h2>Implizites Wissen</h2>
      <span class="count">10 Regeln</span>
    </div>
    <div class="rule-grid">
      <div class="rule-card">
        <div class="rule-id">IMP-001</div>
        <div class="rule-name">Undokumentierte Farbcodes</div>
        <div class="rule-desc">Erkennt farbliche Markierungen ohne Legende – Wissen das nur im Kopf der Ersteller:in existiert.</div>
        <span class="severity severity--warning">Warning</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">IMP-002</div>
        <div class="rule-name">Versteckte Blätter</div>
        <div class="rule-desc">Erkennt ausgeblendete oder „sehr versteckte" Arbeitsblätter mit potenziell kritischen Daten.</div>
        <span class="severity severity--critical">Critical</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">IMP-003</div>
        <div class="rule-name">Versteckte Zeilen/Spalten</div>
        <div class="rule-desc">Erkennt ausgeblendete Zeilen und Spalten, die bei Analysen übersehen werden.</div>
        <span class="severity severity--warning">Warning</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">IMP-004</div>
        <div class="rule-name">Bedingte Formatierung</div>
        <div class="rule-desc">Erkennt bedingte Formatierung als versteckte Geschäftslogik – für KI unsichtbar.</div>
        <span class="severity severity--warning">Warning</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">IMP-005</div>
        <div class="rule-name">Magic Numbers</div>
        <div class="rule-desc">Erkennt hartcodierte Zahlen in Formeln ohne benannte Bereiche – schwer wartbar und fehleranfällig.</div>
        <span class="severity severity--warning">Warning</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">IMP-006</div>
        <div class="rule-name">Datenvalidierung</div>
        <div class="rule-desc">Erkennt Dropdown-Listen mit hartcodierten Werten statt dynamischer Quellen.</div>
        <span class="severity severity--warning">Warning</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">IMP-007</div>
        <div class="rule-name">Kryptische Blattnamen</div>
        <div class="rule-desc">Erkennt generische Namen wie „Tabelle1" oder „Sheet2" – erschweren Navigation und Verständnis.</div>
        <span class="severity severity--info">Info</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">IMP-008</div>
        <div class="rule-name">Kommentare mit Logik</div>
        <div class="rule-desc">Erkennt Zellkommentare die Geschäftsregeln enthalten – nicht automatisierbar.</div>
        <span class="severity severity--info">Info</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">IMP-009</div>
        <div class="rule-name">Irreführende Zahlenformate</div>
        <div class="rule-desc">Erkennt Formate, bei denen der angezeigte Wert vom tatsächlichen abweicht (z.B. gerundete Anzeige).</div>
        <span class="severity severity--warning">Warning</span>
      </div>
      <div class="rule-card">
        <div class="rule-id">IMP-010</div>
        <div class="rule-name">Geschützte Bereiche</div>
        <div class="rule-desc">Erkennt Blattschutz als Hinweis auf undokumentierte Regeln und versteckte Abhängigkeiten.</div>
        <span class="severity severity--info">Info</span>
      </div>
    </div>
  </div>

  <div style="text-align:center; margin-top:2rem; padding-top:1.5rem; border-top:1px solid var(--border);">
    <a href="/" style="color:var(--accent); text-decoration:none; font-weight:600;">← Zurück zur Startseite</a>
  </div>
</div>
</body>
</html>"""


@app.route("/")
def index():
    return render_template_string(UPLOAD_PAGE)


@app.route("/learn")
def learn():
    """Interaktive 'Warum Datenbank?' Lernseite."""
    from excel_checker.learn_page import generate_learn_page
    return generate_learn_page()


@app.route("/pruefregeln")
def pruefregeln():
    """Übersichtsseite: Was wird geprüft und warum."""
    return render_template_string(PRUEFREGELN_PAGE)


@app.route("/logo.png")
def logo():
    """Liefert das Logo-Bild aus."""
    logo_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "Excel Checker.png")
    if os.path.exists(logo_path):
        from flask import send_file
        return send_file(logo_path, mimetype="image/png")
    return "", 404



@app.route("/report/<report_id>")
def show_report(report_id):
  """Zeigt einen gespeicherten Report an."""
  with _pending_lock:
    html = _reports.get(report_id)
  if not html:
    return render_template_string(
      ERROR_PAGE,
      title="Report nicht gefunden",
      message="Der Report ist abgelaufen oder wurde bereits angezeigt. Bitte eine neue Analyse starten.",
    ), 404
  return html

# Download-Endpoint für Report als HTML
from flask import make_response
@app.route("/download-report/<report_id>")
def download_report(report_id):
  """Ermöglicht das Herunterladen des Reports als HTML-Datei."""
  with _pending_lock:
    html = _reports.get(report_id)
  if not html:
    return "Report nicht gefunden", 404
  response = make_response(html)
  response.headers['Content-Type'] = 'text/html; charset=utf-8'
  response.headers['Content-Disposition'] = f'attachment; filename=ExcelChecker_Report_{report_id}.html'
  return response


@app.route("/upload", methods=["POST"])
def upload_file():
    """Nimmt den Upload entgegen, speichert temporär, gibt session_id zurück."""
    if "file" not in request.files:
        return {"error": "Keine Datei ausgewählt."}, 400

    file = request.files["file"]
    if not file.filename:
        return {"error": "Keine Datei ausgewählt."}, 400

    original_name = file.filename
    ext = os.path.splitext(original_name)[1].lower()
    if ext not in ('.xlsx', '.xlsm'):
        return {"error": "Nur .xlsx und .xlsm Dateien werden unterstützt."}, 400

    tmp_dir = tempfile.mkdtemp()
    safe_name = f"{uuid.uuid4().hex}{ext}"
    tmp_path = os.path.join(tmp_dir, safe_name)
    file.save(tmp_path)

    session_id = uuid.uuid4().hex
    _sessions[session_id] = queue_mod.Queue()
    _start_analysis(session_id, tmp_path, original_name)

    return {"session_id": session_id}


@app.route("/upload-url", methods=["POST"])
def upload_url():
    """Nimmt eine URL entgegen, lädt die Datei, gibt session_id zurück."""
    data = request.get_json(silent=True) or {}
    url = data.get("url", "").strip()
    if not url:
        return {"error": "Bitte eine URL eingeben."}, 400

    parsed = urlparse(url)
    if parsed.scheme not in ('http', 'https'):
        return {"error": "Bitte eine gültige https:// URL eingeben."}, 400

    tmp_path, display_name, error = _download_file(url)
    if error:
        return {"error": error}, 400

    session_id = uuid.uuid4().hex
    _sessions[session_id] = queue_mod.Queue()
    _start_analysis(session_id, tmp_path, display_name)

    return {"session_id": session_id}


@app.route("/progress/<session_id>")
def progress_stream(session_id):
    """SSE-Endpoint: Streamt Analyse-Fortschritt live zum Browser mit Heartbeat."""
    q = _sessions.get(session_id)

    if q is None:
        def error_stream():
            yield f'event: error_msg\ndata: {json.dumps({"message": "Session nicht gefunden."})}\n\n'
        return Response(error_stream(), mimetype='text/event-stream')

    def generate():
        try:
            while True:
                try:
                    msg = q.get(timeout=5)
                except queue_mod.Empty:
                    # Heartbeat um die Verbindung offen zu halten
                    yield ": heartbeat\n\n"
                    continue

                kind, data = msg
                if kind == "progress":
                    yield f'event: progress\ndata: {json.dumps(data)}\n\n'
                elif kind == "report":
                    yield f'event: report\ndata: {json.dumps(data)}\n\n'
                    break
                elif kind == "error":
                    yield f'event: error_msg\ndata: {json.dumps({"message": data})}\n\n'
                    break
                elif kind == "done":
                    break
        finally:
            _sessions.pop(session_id, None)

    return Response(generate(), mimetype='text/event-stream',
                    headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})


# ── LLM API routes ──────────────────────────────────────────────

@app.route("/api/llm-config")
def llm_config():
    """Gibt den vorgeladenen API-Key (maskiert) und Endpoint zurück."""
    key = get_default_api_key()
    endpoint = get_default_endpoint()
    model = get_default_model()
    return {
        "has_key": bool(key),
        "masked_key": mask_key(key) if key else "",
        "has_endpoint": bool(endpoint),
        "endpoint": endpoint,
        "model": model,
    }


@app.route("/api/llm-test", methods=["POST"])
def llm_test():
    """Testet den API-Key mit einem minimalen API-Call."""
    data = request.get_json(silent=True) or {}
    api_key = data.get("api_key", "").strip()
    if not api_key:
        api_key = get_default_api_key()
    endpoint = data.get("endpoint", "").strip() or get_default_endpoint()

    ok, msg = test_api_key(api_key, endpoint)
    return {"success": ok, "message": msg}


@app.route("/api/llm-analyze/<report_id>", methods=["POST"])
def llm_analyze(report_id):
    """Führt die LLM-Analyse auf einem gespeicherten Report aus."""
    data = request.get_json(silent=True) or {}

    # Report aus Server-State ODER aus Client-Context
    with _pending_lock:
        report = _report_data.get(report_id)

    client_context = data.get("context") if not report else None
    if not report and not client_context:
        return {"error": "Report nicht gefunden. Bitte zuerst eine klassische Analyse durchführen."}, 404

    api_key = data.get("api_key", "").strip()
    if not api_key:
        api_key = get_default_api_key()
    endpoint = data.get("endpoint", "").strip() or get_default_endpoint()

    if not api_key:
        return {"error": "Kein API-Key angegeben. Bitte unter ⚙️ KI-Einstellungen konfigurieren."}, 400

    if not endpoint:
        return {"error": "Kein Endpoint konfiguriert. Bitte unter ⚙️ KI-Einstellungen die Azure AI Endpoint-URL eingeben."}, 400

    model = data.get("model", "").strip() or get_default_model()

    try:
        from excel_checker.report import generate_llm_section_html
        report_or_ctx = report if report else client_context
        analysis = analyze_with_llm(report_or_ctx, api_key, endpoint, model=model)
        html = generate_llm_section_html(analysis)
        return {"success": True, "html": html, "summary": analysis.summary}
    except Exception as e:
        return {"error": f"LLM-Analyse fehlgeschlagen: {e}"}, 500


@app.route("/api/report-cleanup/<report_id>", methods=["POST"])
def report_cleanup(report_id):
    """Räumt gespeicherte Report-Daten auf (bei neuer Analyse)."""
    with _pending_lock:
        _reports.pop(report_id, None)
        _report_data.pop(report_id, None)
    return {"success": True}


@app.route("/analyze", methods=["POST"])
def analyze_file():
    """Fallback für nicht-JS Clients."""
    # Prüfe ob URL im Form statt Datei
    url = request.form.get("url", "").strip()
    if url:
        return _analyze_from_url(url)

    if "file" not in request.files:
        return render_template_string(
            ERROR_PAGE, title="Keine Datei", message="Bitte eine Excel-Datei auswählen."
        ), 400

    file = request.files["file"]
    if not file.filename:
        return render_template_string(
            ERROR_PAGE, title="Keine Datei", message="Bitte eine Excel-Datei auswählen."
        ), 400

    original_name = file.filename
    ext = os.path.splitext(original_name)[1].lower()
    if ext not in ('.xlsx', '.xlsm'):
        return render_template_string(
            ERROR_PAGE,
            title="Falsches Format",
            message="Nur .xlsx und .xlsm Dateien werden unterstützt.",
        ), 400

    # Sichere temporäre Datei erstellen
    tmp_dir = tempfile.mkdtemp()
    safe_name = f"{uuid.uuid4().hex}{ext}"
    tmp_path = os.path.join(tmp_dir, safe_name)

    try:
        file.save(tmp_path)
        report = analyze(tmp_path)
        report.file_path = original_name  # Originalnamen im Report anzeigen
        html = generate_html(report)
        return html
    except Exception as e:
        return render_template_string(
            ERROR_PAGE,
            title="Analyse fehlgeschlagen",
            message=f"Die Datei konnte nicht analysiert werden: {e}",
        ), 500
    finally:
        # Datei sofort löschen
        try:
            os.unlink(tmp_path)
            os.rmdir(tmp_dir)
        except OSError:
            pass


if __name__ == "__main__":
    print("\n" + "=" * 50)
    print("⚡ Excel-Reifecheck")
    print("=" * 50)
    print("\n🌐 Öffne im Browser: http://localhost:5000")
    print("   Strg+C zum Beenden\n")
    app.run(debug=False, host="127.0.0.1", port=5000)


def _convert_sharepoint_url(url: str) -> tuple[str, str]:
    """Konvertiert SharePoint/OneDrive URLs in Download-URLs.

    Returns (download_url, display_name).
    """
    parsed = urlparse(url)
    host = parsed.hostname or ""
    path = unquote(parsed.path)

    # Dateinamen extrahieren
    name_match = re.search(r'([^/]+\.xlsx?)(?:\?|$)', url, re.IGNORECASE)
    if not name_match:
        name_match = re.search(r'([^/]+\.xlsm?)(?:\?|$)', url, re.IGNORECASE)
    display_name = name_match.group(1) if name_match else "download.xlsx"

    # SharePoint: /:x:/r/ Share-Links → download=1 anhängen
    if 'sharepoint.com' in host:
        # Format: https://tenant.sharepoint.com/:x:/s/site/ENCODED?e=TOKEN
        if '/:x:/' in url or '/:w:/' in url:
            separator = '&' if '?' in url else '?'
            return url + separator + 'download=1', display_name

        # Direkte Datei-URL: /sites/SiteName/Documents/file.xlsx
        if path.lower().endswith(('.xlsx', '.xlsm')):
            separator = '&' if '?' in url else '?'
            return url + separator + 'download=1', display_name

        # _layouts/download.aspx URL
        if '_layouts' in path and 'download' in path.lower():
            return url, display_name

    # OneDrive for Business
    if 'my.sharepoint.com' in host or '1drv.ms' in host:
        if path.lower().endswith(('.xlsx', '.xlsm')):
            separator = '&' if '?' in url else '?'
            return url + separator + 'download=1', display_name

    # Direkte URL zu .xlsx Datei (generisch)
    if path.lower().endswith(('.xlsx', '.xlsm')):
        return url, display_name

    # Fallback: URL unverändert versuchen
    return url, display_name


def _download_file(url: str) -> tuple[str, str, str]:
    """Lädt eine Excel-Datei von einer URL herunter.

    Returns (tmp_path, display_name, error_msg).
    error_msg ist leer wenn erfolgreich.
    """
    download_url, display_name = _convert_sharepoint_url(url)

    tmp_dir = tempfile.mkdtemp()
    ext = os.path.splitext(display_name)[1].lower()
    if ext not in ('.xlsx', '.xlsm'):
        ext = '.xlsx'
    safe_name = f"{uuid.uuid4().hex}{ext}"
    tmp_path = os.path.join(tmp_dir, safe_name)

    try:
        resp = http_requests.get(
            download_url,
            timeout=60,
            stream=True,
            allow_redirects=True,
            headers={
                'User-Agent': 'ExcelHealthChecker/1.0',
            },
        )
        resp.raise_for_status()

        # Content-Type prüfen
        ct = resp.headers.get('content-type', '').lower()
        is_excel = any(x in ct for x in [
            'spreadsheet', 'excel', 'octet-stream', 'zip',
            'officedocument', 'ms-excel',
        ])
        # Bei HTML-Antwort: wahrscheinlich Login-Seite
        if 'text/html' in ct:
            return "", display_name, (
                "Der Link erfordert eine Anmeldung. Bitte den Link über "
                "'Freigeben → Link kopieren' mit 'Jeder mit dem Link' "
                "erstellen, oder die Datei direkt hochladen."
            )

        # Dateigröße prüfen
        content_length = int(resp.headers.get('content-length', 0))
        if content_length > 100 * 1024 * 1024:
            return "", display_name, "Die Datei ist größer als 100 MB."

        # Datei speichern
        with open(tmp_path, 'wb') as f:
            for chunk in resp.iter_content(chunk_size=8192):
                f.write(chunk)

        # Rudimentäre Prüfung ob es wirklich ein ZIP/XLSX ist
        with open(tmp_path, 'rb') as f:
            magic = f.read(4)
        if magic[:2] != b'PK':  # ZIP magic bytes
            os.unlink(tmp_path)
            os.rmdir(tmp_dir)
            return "", display_name, (
                "Die heruntergeladene Datei ist keine gültige Excel-Datei. "
                "Möglicherweise verweist der Link auf eine Login-Seite. "
                "Bitte die Datei stattdessen direkt hochladen."
            )

        return tmp_path, display_name, ""

    except http_requests.exceptions.Timeout:
        return "", display_name, "Timeout: Der Server hat nicht rechtzeitig geantwortet."
    except http_requests.exceptions.ConnectionError:
        return "", display_name, "Verbindungsfehler: Der Server ist nicht erreichbar."
    except http_requests.exceptions.HTTPError as e:
        status = e.response.status_code if e.response is not None else "?"
        if status == 401 or status == 403:
            return "", display_name, (
                "Zugriff verweigert. Bitte einen Freigabe-Link mit "
                "'Jeder mit dem Link' erstellen, oder die Datei direkt hochladen."
            )
        return "", display_name, f"HTTP-Fehler {status} beim Herunterladen."
    except Exception as e:
        return "", display_name, f"Fehler beim Herunterladen: {e}"


def _analyze_from_url(url: str):
    """Analysiert eine Excel-Datei von einer URL."""
    # URL-Validierung
    parsed = urlparse(url)
    if parsed.scheme not in ('http', 'https'):
        return render_template_string(
            ERROR_PAGE,
            title="Ungültige URL",
            message="Bitte eine gültige https:// URL eingeben.",
        ), 400

    tmp_path, display_name, error = _download_file(url)

    if error:
        return render_template_string(
            ERROR_PAGE,
            title="Download fehlgeschlagen",
            message=error,
        ), 400

    try:
        report = analyze(tmp_path)
        report.file_path = display_name
        html = generate_html(report)
        return html
    except Exception as e:
        return render_template_string(
            ERROR_PAGE,
            title="Analyse fehlgeschlagen",
            message=f"Die Datei konnte nicht analysiert werden: {e}",
        ), 500
    finally:
        try:
            os.unlink(tmp_path)
            os.rmdir(os.path.dirname(tmp_path))
        except OSError:
            pass


@app.route("/analyze-url", methods=["POST"])
def analyze_url():
    """Endpunkt für URL-basierte Analyse."""
    url = request.form.get("url", "").strip()
    if not url:
        return render_template_string(
            ERROR_PAGE, title="Keine URL", message="Bitte eine URL eingeben."
        ), 400
    return _analyze_from_url(url)
