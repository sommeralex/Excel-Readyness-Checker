# Excel-Reifecheck Challenge — Konzept-Challenge & Umsetzungsplan

Status: Entwurf v1 · Basis: Codestand `453cffd` · Sprache: DE

---

## 0. Die Idee in einem Satz

Unternehmen, Abteilungen und Töchter prüfen ihre Excels mit dem Reifecheck, die
Ergebnisse werden anonymisiert aggregiert und als Ranking / Karte pro Abteilung,
Firma, Stadt, Land und global sichtbar gemacht. Excels werden nie gespeichert.

Die Erzählung dahinter — *"Wozu von KI träumen, wenn die Grundlagen nicht
stimmen?"* — ist stark, wahr und der eigentliche Marketing-Asset. Das Produkt
muss diese These belegen können, ohne sie zu überdehnen.

---

## 1. Challenge: Was an der Idee bricht, bevor sie viral geht

Neun Einwände, sortiert nach Sprengkraft. Jeder mit Gegenmaßnahme — keiner ist
ein K.-o., aber #1 bis #3 entscheiden über Erfolg oder Blamage.

### 1.1 Das Ranking misst nicht Datenreife, sondern Upload-Taktik  ⚠️ kritisch

Der Score hängt vollständig davon ab, **welche Datei** hochgeladen wird. Ein Team
lädt seine sauberste Datei hoch und führt das Ranking an. Es gibt keinen Nenner:
Wir wissen nie, ob die eingereichten 3 Dateien aus 5 oder aus 50.000 stammen.
Ein öffentlich sichtbares Ranking auf dieser Basis ist methodisch angreifbar —
und der erste Journalist, der das merkt, macht daraus die Story.

**Gegenmaßnahme**
- Ranking niemals auf Einzeldatei-Ebene. Einheit ist die **Kohorte** mit
  Mindest-n (≥ 10 Dateien, ≥ 5 verschiedene Einreicher pro Entität).
- Gerankt wird der **Median**, nicht der Mittelwert (robust gegen Cherry-Picking
  nach oben *und* nach unten).
- Primäre Kennzahl ist **Beteiligung + Verbesserung**, nicht der absolute Score
  (siehe 1.6). „Wie viel Chaos habt ihr abgebaut" lässt sich schlechter frisieren
  als „wie sauber ist eure beste Datei".
- Der Ordner-Scan-Modus (siehe Plan Phase 1) erzeugt echte Stichproben statt
  Einzeleinreichungen — das ist der eigentliche Fix.

### 1.2 Ein öffentliches Firmen-Ranking ist der Anti-Viralitäts-Mechanismus  ⚠️ kritisch

B2B-Logik: Kein CFO, keine Rechtsabteilung, keine Konzernkommunikation gibt frei,
dass die eigene Firma auf einer öffentlichen Karte „Excel-Chaos in Österreich"
auftaucht. Das Risiko ist asymmetrisch — auf der Karte zu stehen kann nur
schaden, nie nützen. Genau die Firmen, die man gewinnen will, springen ab.

**Gegenmaßnahme — Zwei-Ebenen-Modell**
- **Öffentlich**: nur Geografie und Branche (Stadt / Bundesland / Land / NACE-
  Sektor), k-anonymisiert. Nie ein Firmenname ohne aktives Opt-in.
- **Privat**: der interne Wettbewerb Abteilung gegen Abteilung, Tochter gegen
  Tochter — sichtbar nur innerhalb der Organisation. **Hier sitzt die echte
  Virialität.** Interner Wettbewerb ist sozial anschlussfähig, öffentliches
  Pranger-Ranking ist es nicht.
- **Opt-in-Badge**: Wer gut ist, darf sich zeigen („Top-10 % Datenreife 2026").
  Positivauszeichnung statt Negativliste. Nur Gewinner werden namentlich sichtbar.

### 1.3 „Wir speichern keine Excels" ist ein Architektur-Versprechen, kein Satz  ⚠️ kritisch

Vertrauen ist hier das gesamte Produkt. Der aktuelle Stand schreibt Uploads in
`tempfile.mkdtemp()` (`webapp.py:528`) und hält vollständige Reports unbegrenzt
im Prozessspeicher (`_reports`, `_report_data` — `webapp.py:38-40`, Eviction nur
manuell über `/api/report-cleanup`). Das ist für ein lokales Tool in Ordnung und
für eine öffentliche Challenge nicht haltbar.

**Gegenmaßnahme — Messung und Aggregation trennen**
- Die Analyse läuft **lokal** (CLI / später WASM). Zum Server geht nur ein
  **numerisches, schemafestes Payload**. Damit wird die Datenschutz-Aussage von
  einer Behauptung zu einer nachprüfbaren Eigenschaft — und das ist selbst eine
  virale Story: *„Deine Datei verlässt deinen Laptop nie."*
- Auch Metriken leaken: Sheet-Namen, Spaltenüberschriften, PII-Fundstellen,
  Dateinamen dürfen **nicht** ins Payload. Nur Rule-IDs (festes öffentliches
  Vokabular) und Zahlen.
- Exakte Kennzahlen sind Quasi-Identifikatoren. „247.913 Zeilen, 14,2 MB,
  Villach, Energiewirtschaft" ist re-identifizierbar → **Buckets statt Zahlen**.

### 1.4 Die Karte ist zum Start das falsche Hero-Visual

Eine Karte mit drei Punkten sieht nach gescheitertem Projekt aus. Karten
brauchen Dichte. Am Tag 1 gibt es null Daten — Cold-Start-Problem in Reinform.

**Gegenmaßnahme**
- Start-Visual ist das **Perzentil** („Deine Datei ist sauberer als 68 % aller
  geprüften Dateien"). Funktioniert ab n≈50 und ist persönlich relevant.
- Die Karte wird **mit einer Baseline-Studie vorbefüllt** (siehe Phase 1):
  öffentlich verfügbare XLSX von data.gv.at, EU Open Data Portal, Statistik
  Austria, Stadtportalen analysieren. Das liefert (a) eine gefüllte Karte am Tag
  der Veröffentlichung, (b) den Presse-Aufhänger, (c) Referenzwerte für Perzentile
  — ganz ohne einen einzigen Upload. Öffentliche Daten öffentlicher Stellen zu
  bewerten ist zulässig und journalistisch anschlussfähig.
- Karte erst freischalten, wenn eine Region die k-Schwelle erreicht.

### 1.5 Virialität in B2B kommt nicht vom Leaderboard, sondern vom teilbaren Artefakt

Der Loop, der wirklich funktioniert: Einzelperson bekommt Score → teilt Score-
Karte in Teams/LinkedIn → Kollege probiert es → Abteilungs-Challenge → Manager
will Org-Sicht. Das Leaderboard ist die *Folge* der Verbreitung, nicht ihr Motor.

**Gegenmaßnahme**: Die **Score-Karte** (server-gerendertes Bild, OG-tauglich, ohne
Firmenname) ist das Kernartefakt und wird zuerst gebaut — vor der Karte.

### 1.6 Gaming

Eine saubere 3-Spalten-Datei mit 20 Zeilen erreicht 100/100. Ohne Gegenmaßnahme
gewinnt, wer am wenigsten einreicht.

**Gegenmaßnahme**
- Mindestkomplexität für Wertung (Zeilen/Spalten/Sheets über Schwellwert).
- Median aus n Dateien statt Bestwert.
- Dedupe über gesalzenen Datei-Fingerprint (dieselbe Datei zählt einmal).
- **Verbesserungs-Modus** als Hauptmechanik: dieselbe Datei erneut prüfen →
  „+34 Punkte gutgemacht". Schwer zu faken, hoher Eigennutzen, erzeugt
  Wiederkehr statt Einmalbesuch.

### 1.7 DSGVO, AVV und — unterschätzt — der Betriebsrat

- Der Checker enthält **PII-Regeln** (`rules/pii.py`: E-Mail, IBAN, SVNR,
  Telefon). Ein Befund wie „12 IBANs in Sheet 'Gehälter'" wäre selbst
  sensibel. → Im Payload nur gebucketete Zähler, keine Sheet-Bezüge; PII-Regeln
  aus öffentlichen Aggregaten ausschließen.
- Serverseitige Verarbeitung = Auftragsverarbeitung → AVV, TOM-Dokument,
  EU-Hosting. Der lokale Modus umgeht das weitgehend, weil keine
  personenbezogenen Daten übertragen werden.
- **§ 96 Abs. 1 Z 3 ArbVG (AT) / § 87 BetrVG (DE)**: Systeme, die
  Leistungsdaten von Mitarbeitenden erheben, sind zustimmungspflichtig. Ein
  Ranking auf Abteilungsebene kann als Leistungskontrolle gelesen werden und im
  Konzern am Betriebsrat scheitern. → Niemals Scores auf Personenebene, auch
  nicht intern; Abteilungsansicht nur über Mindest-n; fertige
  Betriebsvereinbarungs-Vorlage mitliefern (das ist ein Verkaufsargument, kein
  Nebenschauplatz).

### 1.8 Der Score wird zum Benchmark — und dann angegriffen

`health_score` ist heute eine Summe hausgemachter Penalties
(`models.py:130`). Sobald daraus ein öffentlicher Landesvergleich wird, wird
die Methodik zum Angriffspunkt. Zusätzlich: Ändert sich eine Regel, ändern sich
alle historischen Vergleiche.

**Gegenmaßnahme**
- Methodik **einfrieren und versionieren**: „Datenreife-Index v1.0", Version im
  Payload, öffentliche Methodikseite (die vorhandene `/pruefregeln`-Seite
  ausbauen), Regeln bleiben offener Code. Transparenz ist hier Verteidigung.
- Rohe Findings-Counts speichern (nicht nur den fertigen Score) → Index v1.1
  kann historische Einreichungen **rückwirkend neu berechnen**. Ohne das ist
  jede Methodikänderung ein Datenverlust.

### 1.9 Betrieb: Der aktuelle Stand überlebt keinen viralen Moment

Faktisch aus dem Code: Single-Process-Flask, unbegrenzt wachsende In-Memory-
Dicts, `app.secret_key = os.urandom(24)` (`webapp.py:30` — Sessions brechen bei
Neustart und über mehrere Worker), 100 MB Upload-Limit ohne Rate-Limit,
LLM-Aufruf über offenen Endpoint (`/api/llm-analyze` — offene Kostenquelle).
Ein Launch-Peak legt das lahm oder wird teuer.

**Gegenmaßnahme**: Härtung als Phase 0, *vor* jeder Öffentlichkeitsarbeit.

### 1.10 Scope-Ehrlichkeit

Der Checker misst **Datei-Hygiene**. Das korreliert mit KI-Reife, ist aber nicht
dasselbe. „AI-Readiness-Score Österreichs" ist überdehnt und leicht angreifbar.
Positionierung: **Datenreife-Index**, mit klar formuliertem Geltungsbereich.

---

## 2. Was daraus folgt (Positionierungs-Entscheidungen)

| Entscheidung | Statt |
|---|---|
| Datenreife-Index v1.0, versioniert & offen dokumentiert | „AI-Readiness-Score" |
| Interner Wettbewerb (Abteilung vs. Abteilung) als Kern | öffentliches Firmen-Ranking |
| Öffentlich nur Geografie + Branche, k-anonymisiert | Firmennamen auf der Karte |
| Lokale Analyse, nur Zahlen-Payload zum Server | Server-Upload mit Löschversprechen |
| Perzentil + Score-Karte als Launch-Visual | Karte als Launch-Visual |
| Beteiligung & Verbesserung als Primärkennzahl | absoluter Score als Primärkennzahl |
| Baseline-Studie öffentlicher Daten als Cold-Start | leere Karte am Tag 1 |

---

## 3. Architektur

Kernentscheidung: **Messung und Aggregation trennen.**

```
┌─ lokal beim Nutzer ─────────────────┐   ┌─ Challenge-Service ──────────────┐
│ excel_checker (Bestand, unverändert)│   │                                  │
│   engine.py → WorkbookReport        │   │  POST /api/challenge/submit      │
│         │                           │   │        │                         │
│   challenge/metrics.py              │   │   Validierung gegen JSON-Schema  │
│     build_payload(report) ──────────┼──►│   Dedupe über fingerprint        │
│     nur Zahlen + Rule-IDs           │   │   Persistenz: submissions        │
│         │                           │   │        │                         │
│   challenge/client.py               │   │   Aggregation (Median, k-anon)   │
│     Vorschau + explizites Opt-in    │   │        │                         │
│                                     │   │   /board, /map, /card.png        │
│   [Excel-Datei bleibt hier]         │   │   [Excel-Datei kommt nie an]     │
└─────────────────────────────────────┘   └──────────────────────────────────┘
```

Der bestehende Web-Upload bleibt als bequemer Pfad erhalten (mit Härtung), ist
aber nicht der beworbene Challenge-Pfad für Unternehmen.

### 3.1 Das Submission-Payload (Entwurf v1)

Die wichtigste Einzelentscheidung des Projekts — sie definiert, was das
Datenschutzversprechen wert ist.

```jsonc
{
  "schema_version": 1,
  "index_version": "1.0",
  "checker_version": "0.2.0",
  "submitted_at": "2026-08-26T10:00:00Z",

  "scores": { "health": 42, "ai_readiness": 30, "compliance": 65, "db_readiness": 38 },

  "dimensions": {                       // die 6 Radar-Dimensionen, gerundet auf 5er
    "volume": 60, "formula": 45, "linkage": 33,
    "implicit": 75, "structure": 80, "filesize": 20
  },

  "findings": { "STR-001": 3, "IMP-004": 12, "FRM-002": 1 },   // Rule-ID → Anzahl
                                                              // PII-* nur als Bucket
  "size_class": {                       // Buckets, keine Exaktwerte
    "rows": "10k-100k", "cols": "20-50", "sheets": "5-10", "mb": "5-20"
  },

  "entity": {
    "org_id":  "sha256(org_secret + org_slug)",     // pseudonym, serverseitig salted
    "unit_id": "sha256(org_secret + unit_slug)",    // Abteilung/Tochter
    "country": "AT", "region": "AT21", "city": "Klagenfurt",   // Stadt nur ab Größe X
    "industry": "D"                                            // NACE Abschnitt
  },

  "fingerprint": "sha256(salt + strukturelle Merkmale)",   // Dedupe + Verbesserungs-Match
  "signature": "hmac(challenge_token, payload)"
}
```

**Ausdrücklich nicht enthalten**: Dateiname, Sheet-Namen, Spaltenüberschriften,
Zellinhalte, Zellreferenzen, Beispielwerte, Freitext aus Findings, exakte
Zeilen-/Spaltenzahlen, LLM-Ausgaben, Nutzername, IP.

### 3.2 k-Anonymität

Eine Entität erscheint erst im öffentlichen Board, wenn **≥ 5 verschiedene
Einreicher** und **≥ 10 Dateien** vorliegen. Darunter rollt sie in die
übergeordnete Ebene auf (Stadt → Bundesland → Land). Die Schwelle steht in der
Config, wird auf jeder Ansicht ausgewiesen und ist Teil der Methodikseite.

### 3.3 Datenmodell (minimal)

```
orgs         (id, slug_hash, country, region, city, industry, privacy_level, created_at)
units        (id, org_id, slug_hash, label_encrypted, created_at)
submissions  (id, org_id, unit_id, index_version, scores_json, dimensions_json,
              findings_json, size_class_json, fingerprint, created_at)
aggregates   (scope_type, scope_key, period, n_files, n_submitters,
              median_health, p25, p75, computed_at)     -- materialisiert, stündlich
```

Keine Rohdateien, keine Nutzertabelle in Phase 1–2 (Magic-Link statt Accounts).

---

## 4. Umsetzungsplan

### Phase 0 — Fundament & Härtung (1–2 Wochen, keine Öffentlichkeit)

| # | Aufgabe | Betrifft |
|---|---|---|
| 0.1 | `INDEX_VERSION = "1.0"` einführen, Scoring einfrieren, Änderungen nur noch über Versionssprung | `models.py`, `report.py` |
| 0.2 | Methodikseite ausbauen: jede Regel mit Penalty, Begründung, Beispiel | `learn_page.py`, `/pruefregeln` |
| 0.3 | `challenge/metrics.py`: `build_payload(report)` + JSON-Schema + Golden-Tests gegen `test_messy.xlsx`, `test_pii.xlsx` | neu |
| 0.4 | Leak-Test: automatisierter Test, der jedes Payload gegen eine Blocklist prüft (keine Strings aus der Datei) | neu, **Pflicht-Gate** |
| 0.5 | `_reports`/`_report_data` mit TTL-Eviction + Größenlimit; `secret_key` aus Env | `webapp.py` |
| 0.6 | Rate-Limiting, LLM-Endpoint hinter Token, gunicorn-Setup, EU-Hosting | `webapp.py`, Deployment |
| 0.7 | Persistenzschicht (SQLite lokal / Postgres prod) + Migrationen | neu |

*Gate: 0.4 muss grün sein, bevor irgendetwas nach außen geht.*

### Phase 1 — Lokaler Modus & Baseline-Studie (2–3 Wochen)

| # | Aufgabe |
|---|---|
| 1.1 | `excel-reifecheck challenge <ordner>` — scannt rekursiv, zeigt das vollständige Payload im Klartext an, sendet **erst nach expliziter Bestätigung**. Das ist der unternehmenstaugliche Pfad und die Vertrauensgeschichte. |
| 1.2 | Ordner-Scan liefert echte Stichproben → löst 1.1 (Cherry-Picking) an der Wurzel |
| 1.3 | Baseline-Harvester: XLSX von data.gv.at, EU Open Data, Statistik Austria, Stadtportalen → analysieren → Referenzverteilung |
| 1.4 | Publikation „Datenreife-Report Österreich 2026": gefüllte Karte + Presse-Aufhänger + Perzentil-Basis am Tag 1 |
| 1.5 | Perzentil-Anzeige im bestehenden Report („sauberer als 68 %") |
| — | *Stretch*: Pyodide/WASM-Variante („läuft im Browser, nichts wird übertragen"). Stärkeres Marketing, deutlich mehr Aufwand — erst nach Phase 2 bewerten. |

### Phase 2 — Das Challenge-Produkt (3–4 Wochen)

| # | Aufgabe |
|---|---|
| 2.1 | Org/Unit-Modell, Einladungslinks pro Abteilung, Magic-Link statt Accounts |
| 2.2 | Privacy-Schalter durch die Org selbst: `privat` / `nur aggregiert` / `öffentliches Badge`. Default: `privat`. |
| 2.3 | Internes Board: Abteilung vs. Abteilung, Median + Verteilung + n |
| 2.4 | **Score-Karte** als server-gerendertes PNG/OG-Bild — das virale Kernartefakt |
| 2.5 | Verbesserungs-Modus über Fingerprint-Match: „+34 Punkte" |
| 2.6 | Rechts-Paket: AVV-Vorlage, TOM-Dokument, Betriebsvereinbarungs-Einseiter |

### Phase 3 — Karte & Skalierung (3–4 Wochen)

| # | Aufgabe |
|---|---|
| 3.1 | Öffentliches Board: Land / Bundesland / Stadt / Branche, k-anonymisiert, n immer sichtbar |
| 3.2 | Karte: MapLibre GL, **selbst gehostete** Vector-Tiles (kein externer Tile-CDN — sonst bricht das Datenschutzversprechen an der eigenen Karte). Primärkodierung = Beteiligung (Blasengröße), Sekundär = Median-Index (Farbe). Choropleth allein über Scores ist bei ungleichem n irreführend. |
| 3.3 | Branchen-Benchmarks (NACE-Abschnitt) |
| 3.4 | Quartals-„Seasons" mit Reset → wiederkehrende Nachrichtenanlässe statt Einmal-PR |

### Phase 4 — Governance (laufend)

Methodik-Changelog, Index-Versionierung mit Rückrechnung historischer
Einreichungen, offener Regelsatz, dokumentierte Aggregationsregeln.

---

## 5. Was bewusst *nicht* gebaut wird

- Keine Scores auf Personenebene — nie, auch nicht intern (Betriebsrat, 1.7).
- Keine Nutzerkonten in Phase 1–2 (Magic-Link reicht, spart DSGVO-Oberfläche).
- Kein LLM im Challenge-Pfad (Kosten, Datenabfluss, Nichtdeterminismus).
- Keine Echtzeit-Karte (stündliche Aggregate genügen, sind billiger und stabiler).
- Keine Firmennamen ohne aktives Opt-in der Firma selbst.
- Keine externen Karten-/Font-CDNs auf Challenge-Seiten (`theme.css` lädt heute
  Fonts von `rsms.me` und `jsdelivr` — für die öffentliche Challenge selbst hosten).

---

## 6. Erfolgskriterien und Abbruchbedingungen

**Woche 6 nach Launch** — Zielkorridor:
- ≥ 500 gewertete Dateien, ≥ 30 Organisationen, ≥ 3 Presse-/Fach-Erwähnungen
- ≥ 20 % der Teilnehmer laden eine zweite Datei hoch (Retention)
- ≥ 5 Organisationen mit mehr als einer Abteilung (interner Wettbewerb greift)

**Abbruch / Umbau**, wenn:
- < 3 Regionen die k-Schwelle erreichen → Karte einmotten, Perzentil-Produkt behalten
- Einreichungen konzentrieren sich auf Kleinstdateien → Gaming gewinnt, auf reinen
  Verbesserungs-Wettbewerb umstellen
- Rechts-/Betriebsrats-Einwände blockieren > 2 Pilotkunden → Org-Ebene auf
  „nur aggregiert" fixieren, Abteilungsansicht streichen

---

## 7. Erster konkreter Schritt

Phase 0.3 + 0.4: `challenge/metrics.py` mit Payload-Builder, JSON-Schema und dem
Leak-Test. Das Payload-Format ist die Entscheidung, an der alles Weitere hängt —
und der Leak-Test ist das, was das Versprechen „Excels werden nicht gespeichert"
von einer Marketingzeile in eine geprüfte Eigenschaft verwandelt.
