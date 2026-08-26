/* Web Worker: führt den Analysekern per Pyodide im Browser aus.
 *
 * Warum ein Worker und nicht der Haupt-Thread: Die Analyse ist minutenlang
 * synchron (openpyxl streamt, Python blockiert). Im Haupt-Thread würde die
 * Seite in dieser Zeit einfrieren — keine Fortschrittsanzeige, kein Scrollen.
 *
 * Warum ein FRISCHER Worker je Analyse: Der WASM-Heap schrumpft nie. Nach
 * einer großen Datei bleibt er auf seinem Höchststand, auch nach
 * gc.collect(). Über mehrere Analysen hinweg würde sich das bis zum Absturz
 * summieren. Deshalb beendet browser_analysis.js den Worker nach jedem Lauf
 * — siehe docs/deployment/PLAN.md, Abschnitt 1.3.
 *
 * Protokoll (main -> worker):
 *   {type:'init',    base}                      Pyodide + Wheels laden
 *   {type:'analyze', buffer, filename}          Analyse starten
 * Protokoll (worker -> main):
 *   {type:'boot',     step}                     Ladefortschritt vor der Analyse
 *   {type:'ready'}                              bereit für 'analyze'
 *   {type:'progress', event}                    exakt das ProgressEvent des Generators
 *   {type:'report',   html, score, findings, sheets}
 *   {type:'error',    message, detail}
 */

let pyodide = null;

function boot(step) {
  self.postMessage({ type: 'boot', step: step });
}

async function init(base) {
  boot('Analyse-Umgebung wird geladen…');

  // Zuerst das Manifest: Fehlt es, ist das Vendor-Verzeichnis nicht
  // eingerichtet. Ohne diese Prüfung scheitert weiter unten der Import mit
  // einer Meldung, aus der niemand die Ursache ablesen kann.
  const manifestResp = await fetch(base + '/MANIFEST.json');
  if (!manifestResp.ok) {
    throw new Error(
      'Die Analyse-Umgebung fehlt (' + base + '/MANIFEST.json nicht gefunden). ' +
      'Einmalig einrichten mit:  python tools/fetch_vendor.py'
    );
  }
  const manifest = await manifestResp.json();

  // Module-Worker: Pyodide ab Version 0.28 laeuft nicht mehr als klassischer
  // Worker ("Classic web workers are not supported"). Deshalb dynamisches
  // import() statt importScripts — der Pfad steht erst zur Laufzeit fest.
  const mod = await import(base + '/pyodide/pyodide.mjs');
  pyodide = await mod.loadPyodide({ indexURL: base + '/pyodide/' });

  boot('Prüfregeln werden geladen…');
  const names = [];
  for (const wheel of manifest.wheels) {
    const bytes = new Uint8Array(await (await fetch(base + '/wheels/' + wheel)).arrayBuffer());
    pyodide.FS.writeFile('/tmp/' + wheel, bytes);
    names.push(wheel);
  }

  // Wheels direkt entpacken statt über micropip: Die sind bereits lokal,
  // micropip würde nur einen Index im Netz suchen, den es hier nicht gibt.
  pyodide.globals.set('_wheel_names', names.join('\n'));
  pyodide.runPython(`
import os, sys, zipfile
os.makedirs("/site", exist_ok=True)
for _name in _wheel_names.split("\\n"):
    with zipfile.ZipFile("/tmp/" + _name) as _z:
        _z.extractall("/site")
    os.unlink("/tmp/" + _name)
if "/site" not in sys.path:
    sys.path.insert(0, "/site")
`);
  pyodide.globals.delete('_wheel_names');

  // Einmal importieren, damit der erste Klick nicht auch noch den
  // Modul-Import bezahlt.
  pyodide.runPython('import excel_checker.engine, excel_checker.report');
  self.postMessage({ type: 'ready' });
}

function analyze(buffer, filename) {
  pyodide.FS.writeFile('/tmp/input.xlsx', new Uint8Array(buffer));

  // Der Generator meldet Fortschritt; jedes Event geht sofort an den
  // Haupt-Thread. postMessage blockiert nicht, der Python-Lauf darf also
  // synchron bleiben.
  pyodide.globals.set('_emit', (payload) => {
    self.postMessage({ type: 'progress', event: JSON.parse(payload) });
  });

  const summary = pyodide.runPython(`
import io, json
from excel_checker.engine import analyze_with_progress
from excel_checker.models import WorkbookReport
from excel_checker.report import generate_html

_buf = io.BytesIO(open("/tmp/input.xlsx", "rb").read())
_report = None
for _item in analyze_with_progress(_buf, filename=_filename):
    if isinstance(_item, WorkbookReport):
        _report = _item
    else:
        _emit(json.dumps(_item, default=str))

json.dumps({
    "html": generate_html(_report),
    "score": _report.health_score,
    "findings": len(_report.findings),
    "sheets": _report.sheet_count,
})
`);

  pyodide.globals.delete('_emit');
  pyodide.globals.delete('_filename');
  try { pyodide.FS.unlink('/tmp/input.xlsx'); } catch (e) { /* schon weg */ }

  const parsed = JSON.parse(summary);
  self.postMessage({
    type: 'report',
    html: parsed.html,
    score: parsed.score,
    findings: parsed.findings,
    sheets: parsed.sheets,
  });
}

self.onmessage = async (e) => {
  const msg = e.data;
  try {
    if (msg.type === 'init') {
      await init(msg.base);
    } else if (msg.type === 'analyze') {
      pyodide.globals.set('_filename', msg.filename);
      analyze(msg.buffer, msg.filename);
    }
  } catch (err) {
    self.postMessage({
      type: 'error',
      message: String(err && err.message ? err.message : err),
      detail: String(err && err.stack ? err.stack : ''),
    });
  }
};
