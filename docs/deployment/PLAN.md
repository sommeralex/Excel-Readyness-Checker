# Ausliefern statt Hosten — Reifecheck als Web-App und als Download

Status: Entwurf v1 · Basis: Codestand `06a12c0` · Entscheidungsvorlage

---

## 0. Die Frage

Challenge vorerst parken. Stattdessen: den Excel-Reifecheck als Online-Dienst
verfügbar machen (Vercel) plus Download-Variante für Firmen mit Datenbedenken.
Viralität später.

**Antwort: Descoping ist richtig. Vercel ist richtig. Flask auf Vercel ist es
nicht** — und die Auflösung dieses Widerspruchs löst gleichzeitig das Problem
mit den „polierten Excels".

---

## 1. Warum die polierten Excels kein Anreizproblem sind, sondern ein Vertrauensproblem

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

## 2. Realitätscheck: Die App läuft so nicht auf Vercel

Vier harte Blocker im aktuellen Code:

| # | Blocker | Fundstelle |
|---|---|---|
| 1 | **4,5 MB Request-Body-Limit** bei Serverless Functions, gegen `MAX_CONTENT_LENGTH = 100 MB`. Genau die überladenen Workbooks — die Zielgruppe — liegen bei 20–80 MB. | `webapp.py:31` |
| 2 | **SSE bricht.** `/progress/<id>` liest aus dem In-Memory-Dict `_sessions`, gefüttert von einem Daemon-Thread. Serverless hat weder geteilten Speicher zwischen Invocations noch Threads, die die Response überleben. Die gesamte Progress-UX entfällt. | `webapp.py:40`, `48–90`, `563` |
| 3 | **`/report/<id>` bricht.** Der Report liegt in `_reports`; der Folge-Request landet auf einer anderen Instanz. | `webapp.py:38`, `485` |
| 4 | **Laufzeit und Speicher.** openpyxl braucht bei großen Workbooks leicht 2–3 GB und Minuten. Serverless-Limits (Hobby-Tier im Bereich einer Minute, Pro einige Minuten) sind dafür knapp bis unzureichend. | `engine.py:360` |

*Die konkreten Vercel-Limits vor der Umsetzung nachprüfen — sie ändern sich.
Blocker 1 bis 3 sind davon unabhängig strukturell.*

### 2.1 Zur Klarstellung: Vercel ja, Flask-auf-Vercel nein

Die Empfehlung lautet **Vercel**. Das „nicht Vercel" in Abschnitt 5 bezieht
sich ausschließlich auf die Brückenvariante, bei der die Flask-App unverändert
bleibt. Zwei verschiedene Dinge:

| Variante | Plattform | Bewertung |
|---|---|---|
| **A — statische App, Analyse im Browser** | **Vercel** | **Empfehlung.** Genau das, wofür Vercel gebaut ist |
| B — Flask unverändert, Server nötig | Container (Fly, Hetzner, Railway) | Brücke, wenn es diese Woche live sein muss |
| C — Flask serverless-tauglich umgebaut | Vercel | Möglich, aber teuer — siehe 2.2 |

### 2.2 Warum die Blocker strukturell sind und nicht konfigurierbar

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
zerstört die Vertrauensaussage aus Abschnitt 1 und damit den Zugang zu den
tagtäglich eingesetzten Excels.

### 2.3 Variante C wäre machbar — und teurer als die Portierung

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

## 3. Die Auflösung: Die Analyse läuft im Browser

Wenn der Analysekern per WebAssembly (Pyodide) im Browser läuft, ist Vercel
nicht nur möglich, sondern die ideale Plattform:

- Statisches Hosting, CDN, **keine Serverkosten pro Analyse**
- **Kein Upload-Limit** — die Datei wird nie hochgeladen
- **Beliebig skalierbar** — ein viraler Peak kostet nichts und fällt nicht um
- **Keine DSGVO-Oberfläche** — keine Verarbeitung, kein AVV, kein TOM-Dokument,
  kein EU-Hosting-Thema
- **Die Download-Frage löst sich auf** — es gibt nichts hochzuladen, wovor eine
  Firma sich schützen müsste

### 3.1 Machbarkeit — geprüft am Code

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

### 3.2 Risiken

| Risiko | Einschätzung |
|---|---|
| Pyodide-Erstladung ~10 MB | Einmalig, danach im Browser-Cache. Vertretbar; Ladebalken zeigt ihn ohnehin |
| 3–10× langsamer als CPython | Kein Timeout, fremde CPU, Progress-UI existiert bereits. Immer noch besser als ein Serverless-Timeout mitten in der Analyse |
| Browser-Speichergrenze (praktisch ~2 GB) | Das vorhandene Tier- und Sampling-System (`_select_tier`, `read_only=True` ab Tier 2) ist genau der richtige Hebel — Schwellen ggf. für WASM absenken |
| Alte Browser / Firmen-IE-Reste | Fallback-Hinweis plus Download-Variante (Abschnitt 4) |

---

## 4. Was aus dem Download-Wunsch wird

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

## 5. Brückenvariante B: Wenn der Server trotzdem sein soll

Falls die WASM-Portierung zu weit weg ist und diese Woche etwas live sein muss:
**für diesen einen Fall nicht Vercel, sondern Container.** Der Grund ist nicht
Vercel, sondern die unveränderte Flask-App (siehe 2.2). Fly.io, Railway, Render oder Hetzner mit
Docker, Region Frankfurt. Der Code bleibt unverändert, SSE funktioniert,
kein Body-Limit, DSGVO-Hosting geklärt. Aufwand ein bis zwei Tage.

Das ist eine Brücke, kein Ziel — mit Server bleiben Kosten pro Analyse,
Skalierungsrisiko im viralen Fall und die Vertrauensfrage aus Abschnitt 1
bestehen.

---

## 6. Reihenfolge

| Sprint | Inhalt | Dauer |
|---|---|---|
| **1 — Spike** | `engine` auf file-like Input umstellen, Threading-Fallback, Pyodide-Worker aufsetzen, eine echte 20-MB-Datei durchlaufen lassen. **Go/No-Go-Entscheidung an echten Zahlen.** | 2–3 Tage |
| **2 — Web-App** | Frontend an den Worker hängen (SSE → `postMessage`), Upload-Seite und Report übernehmen, Vercel-Deploy, Domain, Fonts selbst hosten | ~1 Woche |
| **3 — Distribution** | PWA-Manifest, PyInstaller-.exe, Docker-Image, Release-Pipeline | 3–4 Tage |
| **4 — Viralität** | Erst jetzt. Die Challenge-Vorlage liegt in `docs/challenge/PLAN.md` bereit | — |

---

## 7. Vier Dinge, die jetzt eingebaut gehören

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

## 8. Erster Schritt

Sprint 1, Spike. Konkret: `analyze_with_progress` file-like-fähig machen,
Threading-Timeout optional, minimaler Pyodide-Worker, `test_messy.xlsx` und
eine echte große Datei durchlaufen lassen, Laufzeit und Speicher messen.

Das Ergebnis dieses Spikes entscheidet zwischen „Vercel als statische App"
(Abschnitt 3) und „Container als Brücke" (Abschnitt 5). Alles andere hängt
daran.
