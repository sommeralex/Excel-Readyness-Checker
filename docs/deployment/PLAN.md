# Ausliefern statt Hosten — Reifecheck als Web-App und als Download

Status: Entwurf v2 (Messungen ergänzt) · Basis: Codestand `bb1121f` · Entscheidungsvorlage

---

## 0. Die Frage

Challenge vorerst parken. Stattdessen: den Excel-Reifecheck als Online-Dienst
verfügbar machen (Vercel) plus Download-Variante für Firmen mit Datenbedenken.
Viralität später.

**Antwort: Descoping ist richtig. Vercel ist richtig. Flask auf Vercel ist es
nicht** — und die Auflösung dieses Widerspruchs löst gleichzeitig das Problem
mit den „polierten Excels".

---

## 1. Gemessene Grenzen

Die Frage „welche Limits existieren da eigentlich" ist am 26.08.2026 gemessen
worden, nicht geschätzt. Harness und Rohdaten liegen in
[`bench/`](../../bench/README.md) und lassen sich jederzeit neu rechnen.

### 1.1 Die harte Formatgrenze

Das XLSX-Format selbst deckelt bei **1.048.576 Zeilen × 16.384 Spalten je
Blatt**. Größer kann eine Excel-Datei gar nicht sein. Alles darunter muss das
Werkzeug schaffen.

### 1.2 Der einzige Wert, auf den es ankommt

| Lademodus | Speicherbedarf | Praktische Obergrenze im Browser (~2 GB) |
|---|---|---|
| `read_only=False` — **heutiger Stand** | ≈ **70 × Dateigröße** | **~25–30 MB Datei** |
| `read_only=True` | ≈ **1 × Dateigröße** | jenseits jeder realen Excel-Datei |

Belege: 44,4-MB-Workbook (300k Zeilen × 25 Spalten) in CPython — **3134 MB**
Peak-RSS normal gegen **49 MB** mit `read_only`. Im Browser dasselbe Bild:
**2053 MB** WASM-Heap normal gegen **43 MB** mit `read_only`.

### 1.3 Drei Befunde, die den Plan bestimmen

1. **`read_only` ist heute nicht aktiv.** `_select_tier` setzt in allen drei
   Tiers `read_only: False` (`engine.py:242`, der Docstring sagt es
   ausdrücklich); der Kommentar in `engine.py:359` behauptet das Gegenteil und
   ist veraltet. **Das ist die höchstwirksame Einzeländerung im ganzen
   Projekt** — sie nützt der Browser-, der Server- und der Desktop-Variante
   gleichermaßen. Voraussetzung: 21 Zugriffe über `ws.cell()` in vier
   Regelmodulen müssen auf `iter_rows()` umgestellt werden;
   `rules/_sampling.py` hat das Muster bereits.
2. **Der Pyodide-Aufschlag ist klein.** Gemessen Faktor **1,2 bis 1,6**
   gegenüber CPython — nicht 3 bis 10, wie in der ersten Fassung dieses
   Dokuments geschätzt. Boot 2,0 s, 14 MB entpackt, danach im Browser-Cache.
   openpyxl läuft dort unverändert.
3. **Der WASM-Heap schrumpft nie.** Nach einer großen Datei bleibt er auf
   seinem Höchststand, auch nach `gc.collect()`. Jede Analyse braucht deshalb
   einen **frischen Worker, der danach beendet wird** — sonst summiert sich der
   Verbrauch über die Sitzung bis zum Absturz.

### 1.4 Laufzeit

| Datei | CPython | Pyodide |
|---|---|---|
| 1,8 MB | 3,3 s | 5,4 s |
| 11,9 MB | 28,5 s | 33,7 s |
| 44,4 MB (normal) | 105,4 s | 123,7 s |
| 44,4 MB (`read_only`) | 22,6 s | 42,9 s |

Minutenlange Analysen sind also real. Im Browser ist das unproblematisch —
kein Timeout, Fortschrittsanzeige vorhanden. In einer Serverless-Funktion ist
es ein Abbruch.

---

## 2. Warum die polierten Excels kein Anreizproblem sind, sondern ein Vertrauensproblem

Der Einwand aus der Challenge-Diskussion war: Leute laden ihre saubersten
Dateien hoch. Der Grund dafür ist nicht Wettbewerbslogik — der Grund ist, dass
die *echte* Datei Kundendaten, Gehälter, Margen und Lieferantenkonditionen
enthält. Niemand zieht die auf die Website eines Anbieters, den er nicht kennt.

Wer die tagtäglich eingesetzten Excels sehen will, muss den Grund beseitigen,
sie zurückzuhalten. Das ist keine Regel- und keine Anreizfrage, sondern eine
Architekturfrage: **Wenn die Datei den Rechner nicht verlässt, gibt es nichts
zurückzuhalten.**

Dieselbe Entscheidung, die das Vercel-Problem löst, löst also auch dieses.

---

## 3. Realitätscheck: Die App läuft so nicht auf Vercel

Vier harte Blocker im aktuellen Code:

| # | Blocker | Fundstelle |
|---|---|---|
| 1 | **4,5 MB Request-Body-Limit** bei Serverless Functions, gegen `MAX_CONTENT_LENGTH = 100 MB`. Genau die überladenen Workbooks — die Zielgruppe — liegen bei 20–80 MB. | `webapp.py:31` |
| 2 | **SSE bricht.** `/progress/<id>` liest aus dem In-Memory-Dict `_sessions`, gefüttert von einem Daemon-Thread. Serverless hat weder geteilten Speicher zwischen Invocations noch Threads, die die Response überleben. Die gesamte Progress-UX entfällt. | `webapp.py:40`, `48–90`, `563` |
| 3 | **`/report/<id>` bricht.** Der Report liegt in `_reports`; der Folge-Request landet auf einer anderen Instanz. | `webapp.py:38`, `485` |
| 4 | **Laufzeit und Speicher.** openpyxl braucht bei großen Workbooks leicht 2–3 GB und Minuten. Serverless-Limits (Hobby-Tier im Bereich einer Minute, Pro einige Minuten) sind dafür knapp bis unzureichend. | `engine.py:360` |

*Die konkreten Vercel-Limits vor der Umsetzung nachprüfen — sie ändern sich.
Blocker 1 bis 3 sind davon unabhängig strukturell.*

### 3.1 Zur Klarstellung: Vercel ja, Flask-auf-Vercel nein

Die Empfehlung lautet **Vercel**. Das „nicht Vercel" in Abschnitt 5 bezieht
sich ausschließlich auf die Brückenvariante, bei der die Flask-App unverändert
bleibt. Zwei verschiedene Dinge:

| Variante | Plattform | Bewertung |
|---|---|---|
| **A — statische App, Analyse im Browser** | **Vercel** | **Empfehlung.** Genau das, wofür Vercel gebaut ist |
| B — Flask unverändert, Server nötig | Container (Fly, Hetzner, Railway) | Brücke, wenn es diese Woche live sein muss |
| C — Flask serverless-tauglich umgebaut | Vercel | Möglich, aber teuer — siehe 3.2 |

### 3.2 Warum die Blocker strukturell sind und nicht konfigurierbar

Vercels Modell ist **zustandslose, kurzlebige Funktionsaufrufe hinter einem
CDN**. Die Flask-App setzt **einen langlebigen Prozess mit gemeinsamem
Speicher** voraus. Die Unvereinbarkeit hat drei Ursachen, und keine davon ist
eine Einstellung:

1. **Geteilter Zustand über Requests hinweg.** `_reports`, `_report_data` und
   `_sessions` sind modulweite Dicts. Der Upload-Request schreibt hinein, ein
   *anderer* Request (`/progress/<id>`, `/report/<id>`) liest heraus. Auf
   Vercel landet der zweite Request potenziell auf einer anderen Instanz — das
   ist kein Timing-Problem, sondern ein direkter 404.
2. **Arbeit, die die Response überlebt.** `_start_analysis` startet einen
   Daemon-Thread und kehrt sofort zurück; die Analyse läuft weiter, während der
   Client den SSE-Stream öffnet. Serverless friert die Instanz nach der
   Response ein oder beendet sie. Der Thread ist tot.
3. **Minutenlang offene Streaming-Response.** SSE hält die Verbindung über die
   gesamte Analyse — das ist ein Funktionsaufruf, der gegen die maximale
   Laufzeit läuft und die ganze Zeit abgerechnet wird.

Dazu das Body-Limit (Blocker 1). Für das gibt es zwar einen Standard-Workaround
— der Client lädt direkt in einen Blob-Store, die Funktion liest von dort —
aber **dieser Workaround bedeutet, dass die Datei gespeichert wird.** Genau das
zerstört die Vertrauensaussage aus Abschnitt 2 und damit den Zugang zu den
tagtäglich eingesetzten Excels.

### 3.3 Variante C wäre machbar — und teurer als die Portierung

Ehrlichkeitshalber: Flask *lässt* sich serverless-tauglich umbauen. Nötig wären

- Blob-Storage für den Upload (Vercel Blob oder S3),
- eine Queue plus separater Worker für die Analyse, weil sie die
  Funktionslaufzeit sprengt,
- Redis oder Postgres für Session- und Report-Zustand,
- Polling statt SSE für den Fortschritt.

Das ist ein verteiltes System mit vier zusätzlichen Diensten. Es kostet mehr
Aufwand als die WASM-Portierung, verursacht laufende Kosten pro Analyse,
skaliert im viralen Fall nur gegen Geld — und bringt die Datei zurück auf
fremde Server. Variante A ist nicht der Kompromiss, sondern die günstigere und
zugleich stärkere Lösung.

Ein Umbau auf Blob-Upload plus Queue plus externem Report-Store wäre die
Serverless-taugliche Variante — und damit ein größeres Projekt als das, was
stattdessen möglich ist.

---

## 4. Die Auflösung: Die Analyse läuft im Browser

Wenn der Analysekern per WebAssembly (Pyodide) im Browser läuft, ist Vercel
nicht nur möglich, sondern die ideale Plattform:

- Statisches Hosting, CDN, **keine Serverkosten pro Analyse**
- **Kein Upload-Limit** — die Datei wird nie hochgeladen
- **Beliebig skalierbar** — ein viraler Peak kostet nichts und fällt nicht um
- **Keine DSGVO-Oberfläche** — keine Verarbeitung, kein AVV, kein TOM-Dokument,
  kein EU-Hosting-Thema
- **Die Download-Frage löst sich auf** — es gibt nichts hochzuladen, wovor eine
  Firma sich schützen müsste

### 4.1 Machbarkeit — geprüft am Code

| Prüfpunkt | Befund |
|---|---|
| openpyxl unter Pyodide | Reines Python (`py3-none-any`-Wheel), einzige Abhängigkeit `et-xmlfile`, ebenfalls rein → per `micropip` installierbar |
| Fremdabhängigkeiten im Analysekern | Keine. `flask`, `requests`, `dotenv` liegen ausschließlich in `webapp.py`, `cli.py`, `__init__.py`, `llm_analysis.py` |
| Report-Erzeugung | `report.py` erzeugt reines HTML ohne Datei-I/O → direkt in den DOM |
| Progress | `analyze_with_progress` ist bereits ein Generator → `postMessage` aus dem Web Worker statt SSE. Die Upload-Seite tauscht praktisch nur die Event-Quelle |

**Zwei Anpassungen sind nötig, beide klein:**

1. `analyze()` / `analyze_with_progress()` müssen ein file-like Objekt
   (`BytesIO`) akzeptieren statt nur einen Pfad — heute `os.path.abspath`,
   `os.path.exists`, `os.path.getsize` (`engine.py:316–320`) und
   `load_workbook(file_path, …)` (`engine.py:360`). openpyxl akzeptiert
   file-likes bereits; die Dateigröße kommt dann aus der Puffergröße.
2. Der Regel-Timeout nutzt `threading.Thread` (`engine.py:606`). Threads sind
   in Pyodide standardmäßig nicht verfügbar → als optionalen Fallback bauen,
   der ohne Threads einfach ohne Timeout durchläuft.

### 4.2 Risiken — nach der Messung

| Risiko | Stand nach Messung |
|---|---|
| Pyodide-Erstladung | **Entschärft.** 14 MB entpackt, Boot 2,0 s, danach im Cache |
| Langsamer als CPython | **Entschärft.** Gemessen Faktor 1,2–1,6, nicht 3–10 |
| Browser-Speichergrenze ~2 GB | **Offen bis Etappe 1.** Ohne `read_only` reißt eine 30-MB-Datei die Grenze; mit `read_only` ist sie kein Thema mehr |
| Heap wächst über die Sitzung | **Neu erkannt.** WASM-Heap schrumpft nie → frischer Worker je Analyse, danach `terminate()` |
| Analysedauer bei großen Dateien | Real: 40 s bis über 2 Minuten. Im Browser tolerierbar (kein Timeout, Fortschrittsanzeige da), erfordert aber ehrliche Zeitangabe in der UI |
| Alte oder gesperrte Firmenbrowser | Hinweis plus Download-Variante (Abschnitt 5) |


---

## 5. Was aus dem Download-Wunsch wird

Drei Zielgruppen, drei Artefakte — in dieser Reihenfolge:

| Zielgruppe | Artefakt | Aufwand |
|---|---|---|
| Normalfall („ich will das nicht hochladen") | **Nichts zu tun** — die Web-App ist bereits lokal | — |
| Offline / abgeschottetes Netz | **PWA** mit gebündelten Wheels: dieselbe Seite, installierbar, offline lauffähig | klein |
| IT-Abteilung, Batch, Automatisierung | **Windows-.exe** (PyInstaller, Doppelklick → Browser auf localhost) plus **Docker-Image**, dazu `pipx install` für Entwickler | mittel |

Die `.exe` ist das, was beim Zielpublikum tatsächlich ankommt: Der README zeigt
ein Windows-Publikum (Git Bash, cp1252-Konsole, `PYTHONIOENCODING=utf-8`) —
dort ist „installier dir erst Python" bereits das Ende des Gesprächs.

---

## 6. Brückenvariante B: Wenn der Server trotzdem sein soll

Falls die WASM-Portierung zu weit weg ist und diese Woche etwas live sein muss:
**für diesen einen Fall nicht Vercel, sondern Container.** Der Grund ist nicht
Vercel, sondern die unveränderte Flask-App (siehe 3.2). Fly.io, Railway, Render oder Hetzner mit
Docker, Region Frankfurt. Der Code bleibt unverändert, SSE funktioniert,
kein Body-Limit, DSGVO-Hosting geklärt. Aufwand ein bis zwei Tage.

Das ist eine Brücke, kein Ziel — mit Server bleiben Kosten pro Analyse,
Skalierungsrisiko im viralen Fall und die Vertrauensfrage aus Abschnitt 1
bestehen.

---

## 7. Der Plan zum Abarbeiten

Sechs Etappen. Jede hat ein überprüfbares Abnahmekriterium — erst wenn das
erfüllt ist, geht es weiter. Etappe 0 ist erledigt.

### Etappe 0 — Messgrundlage ✅ erledigt

Benchmark-Harness in [`bench/`](../../bench/README.md), Ergebnisse in
Abschnitt 1. Der Go/No-Go-Spike ist damit vorweggenommen: **openpyxl läuft
unter Pyodide, der Aufschlag ist klein, die Speichergrenze ist der einzige
echte Engpass — und der ist lösbar.** Entscheidung: Variante A.

### Etappe 1 — `read_only` scharf schalten

Das Fundament. Nützt Browser, Server und Desktop gleichermaßen und ist
unabhängig von allem Weiteren wertvoll.

| Schritt | Inhalt |
|---|---|
| 1.1 | Die 21 `ws.cell()`-Zugriffe in `rules/structure.py` (10), `rules/formulas.py` (5), `rules/implicit_knowledge.py` (5), `rules/volume.py` (1) auf `iter_rows()` umstellen. Muster liegt in `rules/_sampling.py` |
| 1.2 | `_select_tier` auf `read_only: True` ab Tier 2 umstellen; Docstring und den veralteten Kommentar in `engine.py:359` korrigieren |
| 1.3 | Regeln, die zwingend Styles brauchen (`needs_styles`), bleiben auf Tier 1 mit vollem Laden — dort ist die Datei klein genug |
| 1.4 | Regressionstest: Scores und Findings auf `test_messy.xlsx` und `test_pii.xlsx` müssen vor und nach der Umstellung identisch sein |

**Abnahme:** `bench/bench.py analyze wb_L.xlsx` bleibt unter **300 MB**
Peak-RSS (heute 3146 MB) bei unveränderten Findings.

### Etappe 2 — Analysekern browserfähig

| Schritt | Inhalt |
|---|---|
| 2.1 | `analyze()` und `analyze_with_progress()` akzeptieren ein file-like Objekt statt nur eines Pfads (`engine.py:316–320`, `360`); Dateigröße aus der Puffergröße |
| 2.2 | Regel-Timeout: Fallback ohne Threads, wenn `threading.Thread` nicht verfügbar ist (`engine.py:606`) |
| 2.3 | `excel_checker` als reines Wheel bauen — ohne `flask`, `requests`, `dotenv` als harte Abhängigkeit (Extras statt Core-Dependencies in `pyproject.toml`) |
| 2.4 | Wheel in Pyodide laden und `analyze_with_progress` vollständig durchlaufen lassen |

**Abnahme:** Vollständiger HTML-Report aus `test_messy.xlsx`, erzeugt im
Browser, inhaltsgleich zum CPython-Report.

### Etappe 3 — Worker und Oberfläche

| Schritt | Inhalt |
|---|---|
| 3.1 | Web Worker: Pyodide plus Wheels laden, Fortschritts-Events des Generators per `postMessage` weiterreichen — Ereignisformat bleibt wie beim SSE-Stream |
| 3.2 | **Frischer Worker je Analyse, danach `terminate()`** (Heap schrumpft nie, siehe 1.3) |
| 3.3 | `_upload_page.html` von `EventSource` auf Worker umstellen — nur die Ereignisquelle tauschen, die Anzeige bleibt |
| 3.4 | Report-HTML direkt in den DOM statt über `/report/<id>`; „Als Datei speichern" ersetzt den Download-Endpunkt |
| 3.5 | Ehrliche Laufzeitangabe in der UI, abgeleitet aus der Dateigröße |

**Abnahme:** Datei per Drag-and-drop → Fortschritt → Report, und im
Netzwerk-Tab passiert nach dem Seitenaufbau nichts mehr.

### Etappe 4 — Vercel-Deploy

| Schritt | Inhalt |
|---|---|
| 4.1 | Pyodide (14 MB) und Wheels selbst ausliefern, nicht per Fremd-CDN |
| 4.2 | Fonts selbst hosten — `static/theme.css` lädt heute von `rsms.me` und `jsdelivr` |
| 4.3 | `vercel.json`: Cache-Header für die unveränderlichen Pyodide-Dateien, Security-Header |
| 4.4 | Domain, `robots.txt`, OG-Meta-Tags |

**Abnahme:** Öffentlich erreichbar; zweiter Aufruf lädt Pyodide aus dem Cache;
keine Anfrage an einen Drittanbieter-Host.

### Etappe 5 — Distribution

| Schritt | Inhalt |
|---|---|
| 5.1 | PWA: Service Worker cached Pyodide und Wheels → offline lauffähig, installierbar |
| 5.2 | Windows-`.exe` mit PyInstaller: Doppelklick startet den lokalen Server und öffnet den Browser |
| 5.3 | Docker-Image für IT-Abteilungen |
| 5.4 | Release-Pipeline, die alle drei aus einem Tag baut |

**Abnahme:** Alle drei Artefakte aus einem Tag gebaut und auf einem Rechner
ohne vorinstalliertes Python getestet.

### Etappe 6 — Vor der Viralität

Die vier Punkte aus Abschnitt 8. Danach steht die Challenge-Vorlage in
`docs/challenge/PLAN.md` bereit.

### Reihenfolge-Logik

Etappe 1 vor allem anderen, weil sie den einzigen echten Engpass beseitigt und
auch dann Wert liefert, wenn die Browser-Variante scheitern sollte. Etappe 2
und 3 sind die eigentliche Portierung. Etappe 4 ist klein, sobald 3 steht.
Etappe 5 ist unabhängig und kann parallel laufen.


---

## 8. Vier Dinge, die jetzt eingebaut gehören

Auch bei geparkter Challenge sind diese vier später teuer nachzurüsten:

1. **Kein Signup, kein E-Mail-Gate. Nie.** Jede Hürde vor dem Ergebnis
   halbiert die Weitergabe.
2. **Das Ergebnis muss screenshot-würdig sein.** Score-Gauge und Radar-Chart
   sind es bereits — ein „Als Bild speichern"-Button und OG-Meta-Tags kosten
   fast nichts und sind die Einheit, die tatsächlich geteilt wird.
3. **„Deine Datei verlässt deinen Browser nie" gehört über den Upload-Button**,
   nicht ins Impressum. Das ist der Satz, der Leute dazu bringt, die echte
   Datei zu nehmen statt der vorzeigbaren — und damit der Satz, der das
   Problem aus Abschnitt 1 auflöst.
4. **Ein winziger, opt-in Perzentil-Zähler.** Nur der Score, nichts sonst,
   ausdrücklich freiwillig. Reicht für „sauberer als 68 % aller geprüften
   Dateien" — und ist gleichzeitig die kleinste Tür, durch die die Challenge
   später hereinkommen kann, ohne dass jetzt etwas davon gebaut werden muss.

---

## 9. Erster Schritt

**Etappe 1.1**: die 21 `ws.cell()`-Zugriffe auf `iter_rows()` umstellen.

Das ist die höchstwirksame Einzeländerung im Projekt — sie senkt den
Speicherbedarf um etwa den Faktor 60, macht große Dateien im Browser überhaupt
erst möglich, beschleunigt gleichzeitig die CLI und die Server-Variante, und
sie ist unabhängig von jeder Entscheidung über Vercel, WASM oder Container
wertvoll.

Der Go/No-Go-Spike, der hier ursprünglich stand, ist durch die Messungen in
Abschnitt 1 bereits beantwortet: **Variante A ist tragfähig.**
