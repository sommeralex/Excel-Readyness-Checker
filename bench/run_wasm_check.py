"""Faehrt bench/wasm_check.html im echten Browser und vergleicht mit CPython.

Abnahme fuer Etappe 2.4: derselbe Report aus demselben Wheel, einmal unter
CPython und einmal unter Pyodide. Aufruf siehe bench/README.md.
"""

import hashlib
import json
import pathlib
import re
import sys
import time
from dataclasses import asdict

from playwright.sync_api import sync_playwright

# Der Report traegt seinen Erstellungs- und Build-Zeitstempel im Fusstext.
# Die koennen zwischen zwei Laeufen gar nicht gleich sein — alles andere
# muss es.
_TIMESTAMP = re.compile(r"\d{2}\.\d{2}\.\d{4},? \d{2}:\d{2}(:\d{2})?")


def html_fingerprint(html: str) -> tuple:
    """Laenge und Hash des Reports, Zeitstempel herausgerechnet."""
    normalised = _TIMESTAMP.sub("<ZEIT>", html)
    return len(normalised), hashlib.sha256(normalised.encode("utf-8")).hexdigest()

BASE = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8765"
FILES = sys.argv[2:] or ["test_messy.xlsx", "test_pii.xlsx"]

# Die Dateinamen sind so gemeint, wie der HTTP-Server sie ausliefert. Lokal
# koennen dieselben Dateien woanders liegen (Testdaten im Repo-Wurzel,
# Benchmark-Workbooks in bench/), deshalb wird hier gesucht statt angenommen.
_LOCAL_DIRS = [pathlib.Path("."), pathlib.Path("bench")]


def local_path(name: str) -> pathlib.Path:
    for directory in _LOCAL_DIRS:
        candidate = directory / name
        if candidate.is_file():
            return candidate
    raise FileNotFoundError(
        f"{name} lokal nicht gefunden – gesucht in: "
        + ", ".join(str(d.resolve()) for d in _LOCAL_DIRS)
    )


def cpython_payload(path: str) -> dict:
    """Dieselbe Nutzlast wie die Browser-Seite, nur lokal gerechnet."""
    import io

    from excel_checker.engine import analyze_with_progress
    from excel_checker.models import WorkbookReport
    from excel_checker.report import generate_html

    resolved = local_path(path)
    buf = io.BytesIO(resolved.read_bytes())
    events, report = [], None
    for item in analyze_with_progress(buf, filename=path):
        if isinstance(item, WorkbookReport):
            report = item
        else:
            events.append(item.get("step", ""))

    html = generate_html(report)
    length, digest = html_fingerprint(html)
    return {
        "score": report.health_score,
        "findings": sorted(
            f"{f.rule_id}|{f.severity.value}|{f.sheet}|{f.message}"
            for f in report.findings
        ),
        "stats": sorted([asdict(s) for s in report.sheet_stats], key=lambda d: d["name"]),
        "profiles": sorted(
            [asdict(c) for c in report.column_profiles],
            key=lambda d: (d["sheet"], d["column_letter"]),
        ),
        "recs": sorted(x.title for x in report.recommendations),
        "events": len(events),
        "html_len": length,
        "html_sha": digest,
    }


def browser_payload(page, file: str) -> dict:
    page.goto(f"{BASE}/wasm_check.html?file={file}")
    deadline = time.time() + 900
    while time.time() < deadline:
        if page.evaluate("window.DONE === true"):
            break
        time.sleep(2)
    else:
        raise TimeoutError(f"Browser-Lauf fuer {file} nicht fertig geworden")
    print(page.inner_text("#log"))
    raw = json.loads(page.evaluate("window.RESULT"))
    if "error" in raw:
        return raw
    payload = raw["payload"]
    payload["html_len"], payload["html_sha"] = html_fingerprint(raw["html"])
    return payload


def _launch(p):
    """Startet Chromium, auch wenn die Playwright-Version einen anderen
    Browser-Build erwartet als lokal installiert ist."""
    import glob
    import os

    try:
        return p.chromium.launch(args=["--no-sandbox"])
    except Exception as exc:
        print(f"Standard-Launch fehlgeschlagen: {str(exc).splitlines()[0][:120]}")

    patterns = [
        os.path.expanduser(
            "~/Library/Caches/ms-playwright/chromium_headless_shell-*/"
            "chrome-headless-shell-*/chrome-headless-shell"
        ),
        "/opt/pw-browsers/chromium-*/chrome-linux/chrome",
    ]
    for pattern in patterns:
        for path in sorted(glob.glob(pattern), reverse=True):
            try:
                return p.chromium.launch(executable_path=path, args=["--no-sandbox"])
            except Exception:
                continue
    # Letzter Ausweg: das installierte Google Chrome.
    return p.chromium.launch(channel="chrome", args=["--no-sandbox"])


def main() -> int:
    failures = 0
    with sync_playwright() as p:
        browser = _launch(p)
        page = browser.new_page()
        page.on("console", lambda m: None)
        for file in FILES:
            print(f"\n===== {file} =====")
            wasm = browser_payload(page, file)
            if "error" in wasm:
                print(f"  FEHLER im Browser: {wasm['error']}")
                failures += 1
                continue
            native = cpython_payload(file)
            # Der Anzeigename steckt nicht in der Nutzlast — verglichen wird
            # nur, was die Analyse selbst produziert.
            if wasm == native:
                print(f"  IDENTISCH zu CPython (score={native['score']}, "
                      f"{len(native['findings'])} findings, html sha "
                      f"{native['html_sha'][:12]}…)")
            else:
                failures += 1
                print("  ABWEICHUNG:")
                for key in sorted(set(wasm) | set(native)):
                    if wasm.get(key) != native.get(key):
                        print(f"    {key}:")
                        print(f"      cpython: {str(native.get(key))[:300]}")
                        print(f"      pyodide: {str(wasm.get(key))[:300]}")
        browser.close()
    print("\n" + ("ALLE LAEUFE IDENTISCH" if not failures else f"{failures} ABWEICHUNG(EN)"))
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
