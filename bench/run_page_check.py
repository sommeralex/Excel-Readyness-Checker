"""Abnahme fuer Etappe 3: Datei ablegen, Fortschritt, Report — ohne Upload.

Faehrt die echte Upload-Seite in Chromium, legt eine Datei per File-Input
hinein und prueft drei Dinge:

1. Der Fortschritt kommt an (dieselben ProgressEvents wie frueher per SSE).
2. Der Report entsteht im Browser und ist inhaltsgleich zum CPython-Report.
3. **Nichts geht zum Server**, ausser statischen Dateien der eigenen Domain.
   Kein /upload, kein /progress, kein /report — die Datei bleibt im Browser.

Aufruf:

    .venv/bin/python -m flask --app excel_checker.webapp run --port 5000 &
    .venv/bin/python bench/run_page_check.py http://127.0.0.1:5000 test_messy.xlsx
"""

from __future__ import annotations

import pathlib
import re
import sys
import time

from playwright.sync_api import sync_playwright

BASE = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:5000"
SAMPLES = sys.argv[2:] or ["test_messy.xlsx", "test_pii.xlsx"]

# Endpunkte, die es in der Browser-Variante nicht mehr geben darf.
FORBIDDEN = re.compile(r"/(upload|upload-url|progress|report|download-report|analyze)\b")

_TIMESTAMP = re.compile(r"\d{2}\.\d{2}\.\d{4},? \d{2}:\d{2}(:\d{2})?")


def _launch(p):
    import glob
    import os

    try:
        return p.chromium.launch(args=["--no-sandbox"])
    except Exception as exc:
        print(f"Standard-Launch fehlgeschlagen: {str(exc).splitlines()[0][:110]}")
    pattern = os.path.expanduser(
        "~/Library/Caches/ms-playwright/chromium_headless_shell-*/"
        "chrome-headless-shell-*/chrome-headless-shell"
    )
    for path in sorted(glob.glob(pattern), reverse=True):
        try:
            return p.chromium.launch(executable_path=path, args=["--no-sandbox"])
        except Exception:
            continue
    return p.chromium.launch(channel="chrome", args=["--no-sandbox"])


def cpython_report(path: pathlib.Path) -> str:
    import io

    from excel_checker.engine import analyze
    from excel_checker.report import generate_html

    return generate_html(
        analyze(io.BytesIO(path.read_bytes()), filename=path.name)
    )


def run_one(page, sample: pathlib.Path, progress_steps: list) -> str:
    """Legt eine Datei ein, wartet auf den Report und gibt sein HTML zurueck."""
    progress_steps.clear()
    page.set_input_files("#fileInput", str(sample.resolve()))
    page.click("#submitBtn")

    deadline = time.time() + 600
    while time.time() < deadline:
        if page.is_visible("#showReportBtn"):
            break
        if page.is_visible("#errorMsg.visible"):
            raise RuntimeError(page.inner_text("#errorMsg"))
        time.sleep(1)
    else:
        raise TimeoutError("kein Report erschienen")

    status = page.inner_text("#progressText")
    page.click("#showReportBtn")
    page.wait_for_selector("#reportFrame")
    html = page.evaluate("document.getElementById('reportFrame').srcdoc")
    # Overlay wieder schliessen, damit der naechste Lauf die Seite bedienen kann.
    page.evaluate("""() => {
        const o = document.getElementById('reportOverlay');
        if (o) o.style.display = 'none';
    }""")
    return status, html


def main() -> int:
    samples = [pathlib.Path(s) for s in SAMPLES]
    for sample in samples:
        if not sample.is_file():
            print(f"Testdatei nicht gefunden: {sample.resolve()}")
            return 1

    requests_after_load: list[str] = []
    progress_steps: list[str] = []
    page_errors: list[str] = []
    results: list = []

    with sync_playwright() as p:
        browser = _launch(p)
        page = browser.new_page()
        page.on("pageerror", lambda e: page_errors.append(str(e)))

        page.goto(BASE, wait_until="load")
        # Vorwaermen laeuft beim load-Event los; ab jetzt wird mitgeschrieben.
        page.on("request", lambda r: requests_after_load.append(r.url))

        # Fortschritts-Schritte mitschneiden, so wie die Seite sie anzeigt.
        # expose_function braucht eine echte Funktion, keine gebundene
        # Builtin-Methode.
        page.expose_function("__recordStep", lambda step: progress_steps.append(step))
        # Als Pfeilfunktion, sonst ruft Playwright den Rueckgabewert der
        # Zuweisung (die neue Funktion) gleich selbst auf.
        page.evaluate("""() => {
            const orig = window.applyProgress;
            window.applyProgress = function (data, els, state) {
                window.__recordStep(String((data && data.step) || ''));
                return orig(data, els, state);
            };
        }""")

        # Mehrere Laeufe hintereinander in DERSELBEN Seite: Das ist der
        # eigentliche Test fuer "frischer Worker je Analyse". Wuerde der
        # Worker weiterverwendet oder nicht sauber beendet, scheitert der
        # zweite Lauf oder liefert Reste des ersten.
        for idx, sample in enumerate(samples, 1):
            print(f"\nLauf {idx}/{len(samples)}: {sample.name}")
            try:
                status, html = run_one(page, sample, progress_steps)
            except Exception as exc:
                print(f"  FEHLGESCHLAGEN: {exc}")
                browser.close()
                return 1
            results.append((sample, status, html, len(progress_steps)))
            print(f"  {len(progress_steps)} Fortschritts-Schritte, "
                  f"{len(html)} Zeichen Report")
            print(f"  Statuszeile: {status}")

        browser.close()

    failures = 0

    for sample, status, html, steps in results:
        if steps < 5:
            print(f"\n{sample.name}: ZU WENIG Fortschritt ({steps} Schritte)")
            failures += 1

    print("\nNetzwerk nach dem Seitenaufbau:")
    offending = [u for u in requests_after_load if FORBIDDEN.search(u)]
    static_only = [u for u in requests_after_load if not FORBIDDEN.search(u)]
    print(f"  {len(requests_after_load)} Anfragen, davon {len(static_only)} statisch")
    if offending:
        failures += 1
        print("  SERVER-AUFRUFE, die es nicht geben darf:")
        for url in sorted(set(offending)):
            print(f"    {url}")
    else:
        print("  kein /upload, kein /progress, kein /report — Datei blieb im Browser")

    print("\nReports gegen CPython:")
    for sample, _status, browser_html, _steps in results:
        native_html = cpython_report(sample)
        same = (_TIMESTAMP.sub("<ZEIT>", native_html)
                == _TIMESTAMP.sub("<ZEIT>", browser_html))
        print(f"  {sample.name:20} {'IDENTISCH' if same else 'ABWEICHUNG'} "
              f"({len(browser_html)} Zeichen im Browser, {len(native_html)} lokal)")
        if not same:
            failures += 1

    if page_errors:
        failures += 1
        print("\nJavaScript-Fehler auf der Seite:")
        for err in page_errors[:5]:
            print("   ", err.splitlines()[0])

    print("\n" + ("ABNAHME BESTANDEN" if not failures else f"{failures} PUNKT(E) OFFEN"))
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
