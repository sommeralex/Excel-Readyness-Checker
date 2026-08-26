"""Holt Pyodide und die benötigten Wheels nach ``excel_checker/static/vendor``.

Die Analyse läuft im Browser, also braucht die Seite den Python-Interpreter
und die Wheels als eigene Dateien. Bewusst **kein Fremd-CDN zur Laufzeit**:
Alles wird einmal heruntergeladen und danach von derselben Domain
ausgeliefert wie die Seite selbst (siehe docs/deployment/PLAN.md, Etappe 4.1).

Aufruf:

    python tools/fetch_vendor.py            # alles holen und Wheel bauen
    python tools/fetch_vendor.py --check    # nur prüfen, ob es vollständig ist

Das Ergebnis ist nicht versioniert (rund 14 MB); ``vendor/MANIFEST.json``
sagt der Worker-Datei, welche Wheels sie entpacken muss.
"""

from __future__ import annotations

import argparse
import io
import json
import shutil
import subprocess
import sys
import tarfile
import urllib.request
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
VENDOR = REPO / "excel_checker" / "static" / "vendor"
PYODIDE_DIR = VENDOR / "pyodide"
WHEEL_DIR = VENDOR / "wheels"
MANIFEST = VENDOR / "MANIFEST.json"

# Version bewusst festgenagelt: Ein stillschweigender Sprung auf eine neue
# Pyodide-Version würde die Analyse im Browser ändern, ohne dass am Code
# etwas passiert ist.
PYODIDE_VERSION = "314.0.6"
NPM_TARBALL = f"https://registry.npmjs.org/pyodide/-/pyodide-{PYODIDE_VERSION}.tgz"

# Was Pyodide zum Starten wirklich braucht. Der Rest des npm-Pakets
# (TypeScript-Definitionen, Source-Maps, Beispielkonsolen) bleibt draußen.
PYODIDE_FILES = [
    "pyodide.js",
    "pyodide.mjs",
    # Je nach Pyodide-Version heisst der Emscripten-Glue .js oder .mjs.
    # Beide anfragen, fehlende werden uebersprungen.
    "pyodide.asm.js",
    "pyodide.asm.mjs",
    "pyodide.asm.wasm",
    "python_stdlib.zip",
    "pyodide-lock.json",
]

RUNTIME_DEPS = ["openpyxl", "et_xmlfile"]


def fetch_pyodide() -> list[str]:
    print(f"Pyodide {PYODIDE_VERSION} laden …")
    with urllib.request.urlopen(NPM_TARBALL, timeout=120) as resp:
        blob = resp.read()
    print(f"  {len(blob) / 1048576:.1f} MB geladen")

    PYODIDE_DIR.mkdir(parents=True, exist_ok=True)
    written = []
    with tarfile.open(fileobj=io.BytesIO(blob), mode="r:gz") as tar:
        available = {Path(m.name).name: m for m in tar.getmembers() if m.isfile()}
        for name in PYODIDE_FILES:
            member = available.get(name)
            if member is None:
                # Nicht jede Pyodide-Version liefert jede Datei (js vs mjs).
                print(f"  übersprungen (nicht im Paket): {name}")
                continue
            src = tar.extractfile(member)
            if src is None:
                continue
            target = PYODIDE_DIR / name
            target.write_bytes(src.read())
            written.append(name)
            print(f"  {name} ({target.stat().st_size / 1024:.0f} KB)")
    if "pyodide.js" not in written:
        raise SystemExit("pyodide.js fehlt im npm-Paket – Version prüfen.")
    return written


def fetch_wheels() -> list[str]:
    print("Laufzeit-Wheels laden …")
    if WHEEL_DIR.exists():
        shutil.rmtree(WHEEL_DIR)
    WHEEL_DIR.mkdir(parents=True)
    subprocess.run(
        [sys.executable, "-m", "pip", "download", *RUNTIME_DEPS,
         "-d", str(WHEEL_DIR), "--no-deps", "-q"],
        check=True,
    )
    return [p.name for p in sorted(WHEEL_DIR.glob("*.whl"))]


def build_own_wheel() -> str:
    print("excel_checker als Wheel bauen …")
    out = VENDOR / "_build"
    if out.exists():
        shutil.rmtree(out)
    subprocess.run(
        [sys.executable, "-m", "build", "--wheel", "--outdir", str(out)],
        check=True, cwd=str(REPO),
    )
    built = sorted(out.glob("*.whl"))
    if not built:
        raise SystemExit("Wheel-Bau hat nichts erzeugt.")
    # Ein altes Wheel derselben Distribution muss weg, sonst entpackt der
    # Worker zwei Versionen übereinander.
    for old in WHEEL_DIR.glob("excel_reifecheck-*.whl"):
        old.unlink()
    target = WHEEL_DIR / built[-1].name
    shutil.copy2(built[-1], target)
    shutil.rmtree(out)
    print(f"  {target.name} ({target.stat().st_size / 1024:.0f} KB)")
    return target.name


def write_manifest(pyodide_files: list[str], wheels: list[str]) -> None:
    # Reihenfolge zählt: Abhängigkeiten vor dem Paket, das sie braucht.
    ordered = sorted(wheels, key=lambda n: n.startswith("excel_reifecheck"))
    MANIFEST.write_text(
        json.dumps(
            {
                "pyodide_version": PYODIDE_VERSION,
                "pyodide_files": pyodide_files,
                "wheels": ordered,
            },
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"MANIFEST.json geschrieben ({len(ordered)} Wheels)")


def check() -> int:
    if not MANIFEST.is_file():
        print("FEHLT: vendor/MANIFEST.json – bitte 'python tools/fetch_vendor.py' laufen lassen.")
        return 1
    data = json.loads(MANIFEST.read_text(encoding="utf-8"))
    missing = [n for n in data["pyodide_files"] if not (PYODIDE_DIR / n).is_file()]
    missing += [n for n in data["wheels"] if not (WHEEL_DIR / n).is_file()]
    if missing:
        print("FEHLT:", ", ".join(missing))
        return 1
    total = sum(f.stat().st_size for f in VENDOR.rglob("*") if f.is_file())
    print(f"vollständig – Pyodide {data['pyodide_version']}, "
          f"{len(data['wheels'])} Wheels, {total / 1048576:.1f} MB")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true",
                        help="nur prüfen, nichts herunterladen")
    args = parser.parse_args()

    if args.check:
        return check()

    VENDOR.mkdir(parents=True, exist_ok=True)
    pyodide_files = fetch_pyodide()
    wheels = fetch_wheels()
    wheels.append(build_own_wheel())
    write_manifest(pyodide_files, sorted(set(wheels)))
    return check()


if __name__ == "__main__":
    raise SystemExit(main())
