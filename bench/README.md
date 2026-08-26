# Benchmark-Harness — Speicher- und Laufzeitgrenzen

Misst, wie viel Arbeitsspeicher und Zeit openpyxl und der Analysekern bei
wachsenden Workbooks brauchen — in CPython und im echten Browser (Pyodide).
Grundlage für die Auslieferungsentscheidung in `docs/deployment/PLAN.md`.

## Ausführen

```bash
python -m venv .venv-bench && .venv-bench/bin/pip install -e . playwright

# Testdateien erzeugen (nicht versioniert — sie sind groß)
.venv-bench/bin/python bench/gen.py  20000 15 bench/wb_S.xlsx
.venv-bench/bin/python bench/gen.py 100000 20 bench/wb_M.xlsx
.venv-bench/bin/python bench/gen.py 300000 25 bench/wb_L.xlsx

# CPython: load / load_ro / analyze
for f in S M L; do for m in load load_ro analyze; do
  .venv-bench/bin/python bench/bench.py $m bench/wb_$f.xlsx
done; done

# Browser: Pyodide von npm holen, lokal servieren, headless messen
npm pack pyodide && mkdir -p srv/pyodide && tar xzf pyodide-*.tgz -C srv/pyodide --strip-components=1
.venv-bench/bin/pip download openpyxl et_xmlfile -d srv/wheels --no-deps
cp bench/bench.html bench/wb_*.xlsx srv/
(cd srv && python -m http.server 8765 &)
.venv-bench/bin/python bench/run_bench.py
```

## Die Seite als Ganzes prüfen

`run_page_check.py` fährt die echte Upload-Seite in Chromium, legt eine
Datei ein und prüft, was dabei passiert — vor allem, was **nicht** passiert:

```bash
python tools/fetch_vendor.py                 # einmalig
PYTHONIOENCODING=utf-8 python -m excel_checker.webapp &
.venv/bin/python bench/run_page_check.py http://127.0.0.1:5000 test_messy.xlsx test_pii.xlsx
```

Der Test schlägt fehl, sobald nach dem Seitenaufbau ein Request an
`/upload`, `/upload-url`, `/progress`, `/report`, `/download-report` oder
`/analyze` geht. Mehrere Dateien in einem Aufruf laufen nacheinander in
derselben Seite — das prüft nebenbei, dass je Analyse ein frischer Worker
gestartet und danach beendet wird.

## Analysekern im Browser prüfen

`wasm_check.html` lädt das gebaute Wheel in Pyodide, analysiert eine Datei
aus einem `BytesIO` und gibt den fertigen Report zurück; `run_wasm_check.py`
rechnet dieselbe Analyse unter CPython und vergleicht beides Feld für Feld.

```bash
# Wheel bauen und mit den Abhängigkeiten bereitstellen
.venv/bin/python -m build --wheel --outdir dist
mkdir -p srv/wheels && cp dist/*.whl srv/wheels/
.venv/bin/pip download openpyxl et_xmlfile -d srv/wheels --no-deps
cp bench/wasm_check.html test_messy.xlsx test_pii.xlsx srv/

(cd srv && python -m http.server 8765 &)
.venv/bin/python bench/run_wasm_check.py http://127.0.0.1:8765 test_messy.xlsx
```

Verglichen werden Score, Findings, Sheet-Statistik, Column-Profiles,
Empfehlungen und der Report selbst. Die beiden Zeitstempel im Fußtext werden
vor dem Vergleich herausgerechnet — alles andere muss übereinstimmen.

`run_bench.py` sucht Chromium unter `/opt/pw-browsers/chromium-*/chrome-linux/chrome`;
Pfad ggf. anpassen.

Zwei Fallstricke beim Nachmessen:

- `ru_maxrss` ist **auf Linux in KB, auf macOS in Bytes**. `bench.py` rechnet
  das seit dem Etappe-1-Lauf plattformabhängig um; wer die Zahlen anders
  erhebt, liegt sonst um Faktor 1024 daneben.
- Die von `gen.py` erzeugten Dateien enthalten **kein `<dimension>`-Element**
  (openpyxl schreibt im write-only-Modus keines). Im read-only-Modus liefert
  `ws.max_row` dafür `None`; der Analysekern rechnet die Maße dann per
  zusätzlichem Durchlauf nach. Reale Excel-Dateien bringen das Element mit
  und sparen sich diesen Lauf — bei wb_L rund 35 s.

## Ergebnisse (26.08.2026, Codestand 3d0b075)

Testdateien: gemischte Typen (String-IDs, Ganzzahlen, Fließkomma, Enum-Text,
Datum), ein Blatt.

### CPython 3.11, openpyxl 3.1.5

| Datei | Größe | Modus | Zeit | Peak-RSS |
|---|---|---|---|---|
| wb_S (20k × 15) | 1,8 MB | `load` | 3,3 s | **146 MB** |
| wb_S | 1,8 MB | `load read_only` | 2,6 s | **26 MB** |
| wb_S | 1,8 MB | `analyze` (voll) | 8,9 s | 157 MB |
| wb_M (100k × 20) | 11,9 MB | `load` | 28,5 s | **849 MB** |
| wb_M | 11,9 MB | `load read_only` | 10,6 s | **33 MB** |
| wb_M | 11,9 MB | `analyze` (voll) | 56,2 s | 869 MB |
| wb_L (300k × 25) | 44,4 MB | `load` | 105,4 s | **3134 MB** |
| wb_L | 44,4 MB | `load read_only` | 22,6 s | **49 MB** |
| wb_L | 44,4 MB | `analyze` (voll) | 134,9 s | 3146 MB |

### Pyodide 314.0.6 in Chromium 1194 (headless)

| Datei | Größe | Modus | Zeit | WASM-Heap |
|---|---|---|---|---|
| Boot Pyodide | — | — | 2,0 s | 30 MB |
| openpyxl-Wheel entpacken | — | — | 0,1 s | +0 MB |
| wb_S | 1,8 MB | `load` | 5,4 s | 30 → 108 MB |
| wb_M | 11,9 MB | `load` | 33,7 s | 108 → 653 MB |
| wb_L | 44,4 MB | `load` | 123,7 s | 653 → **2053 MB** |
| wb_L | 44,4 MB | `load read_only` (frischer Heap) | 42,9 s | 30 → **43 MB** |
| wb_M | 11,9 MB | `load read_only` | 20,7 s | 43 → 43 MB |

## Was daraus folgt

1. **`read_only` ist der entscheidende Hebel.** Ohne ihn wächst der Speicher
   mit etwa dem **70-fachen der Dateigröße**; mit ihm bleibt er ungefähr bei
   der **einfachen Dateigröße**. Bei wb_L: 3134 MB gegen 49 MB.
2. **`read_only` ist seit Etappe 1 aktiv** (ab Tier 2). Gegenmessung mit
   demselben Harness auf wb_L: `analyze` braucht **58 MB** statt **3501 MB**
   Peak-RSS. Details in `docs/deployment/PLAN.md`, Abschnitt 7.
3. **openpyxl läuft unverändert unter Pyodide** — bestätigt, nicht vermutet.
4. **Der Pyodide-Aufschlag ist klein**: Faktor 1,2 bis 1,6 gegenüber CPython,
   nicht 3 bis 10.
5. **Der WASM-Heap schrumpft nie.** Nach wb_L bleibt er bei 2 GB, auch nach
   `gc.collect()`. Jede Analyse braucht deshalb einen **frischen Worker, der
   danach beendet wird.**
6. **Pyodide-Boot ist unkritisch**: 2,0 s, 14 MB entpackt, danach im Cache.

## Etappe 2.4 — Analysekern im Browser (26.08.2026, Codestand c4e4939)

Gemessen mit `run_wasm_check.py`, Analyse jeweils aus einem `BytesIO`:

| Datei | CPython | Pyodide | Speicher CPython | Speicher Pyodide | Report |
|---|---|---|---|---|---|
| test_messy.xlsx (11 KB) | < 1 s | 0,1 s | — | 36 MB Heap | identisch |
| test_pii.xlsx (11 KB) | < 1 s | 0,1 s | — | 36 MB Heap | identisch |
| wb_L.xlsx (44,4 MB) | 108,9 s | 165,6 s | 102 MB RSS | 106 MB Heap | identisch |

„identisch" heißt: Score, Findings, Sheet-Statistik, Column-Profiles,
Empfehlungen und der Report-HTML stimmen überein — Letzterer nach Abzug der
beiden Zeitstempel im Fußtext.

Nebenbefund: Pyodide meldet `threads verfuegbar: False`. Der Timeout-Wächter
für Regeln fällt dort also tatsächlich auf den threadlosen Pfad zurück.
