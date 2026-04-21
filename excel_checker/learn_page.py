"""Interaktive 'Warum Datenbank?' Lernseite – spielerisch und verständlich."""

from __future__ import annotations


def generate_learn_page() -> str:
    """Generiert die interaktive Lernseite als vollständiges HTML."""
    return LEARN_PAGE_HTML


LEARN_PAGE_HTML = r"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>⚡ Warum eine Datenbank? – Excel-Reifecheck</title>
<style>
  :root {
    --bg: #f8fafc; --card: #ffffff; --border: #e2e8f0;
    --text: #1e293b; --muted: #64748b; --accent: #3b82f6;
    --accent-hover: #2563eb; --green: #16a34a; --red: #dc2626;
    --yellow: #ca8a04;
  }
  [data-theme="dark"] {
    --bg: #0f172a; --card: #1e293b; --border: #334155;
    --text: #e2e8f0; --muted: #94a3b8; --accent: #60a5fa;
    --accent-hover: #93c5fd; --green: #4ade80; --red: #f87171;
    --yellow: #fbbf24;
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
    background: var(--bg); color: var(--text);
    line-height: 1.7;
    transition: background 0.3s, color 0.3s;
  }

  /* Hero */
  .hero {
    background: linear-gradient(135deg, #1e40af 0%, #3b82f6 50%, #06b6d4 100%);
    color: white; text-align: center; padding: 4rem 2rem 3rem;
  }
  .hero h1 { font-size: 2.2rem; margin-bottom: 0.5rem; }
  .hero p { font-size: 1.1rem; opacity: 0.9; max-width: 600px; margin: 0 auto; }
  .hero .back-link {
    position: absolute; top: 1rem; left: 1.5rem;
    color: rgba(255,255,255,0.8); text-decoration: none; font-size: 0.9rem;
  }
  .hero .back-link:hover { color: white; }
  .theme-toggle-learn {
    position: absolute; top: 1rem; right: 1.5rem;
    background: rgba(255,255,255,0.2); border: none;
    border-radius: 8px; padding: 0.35rem 0.6rem;
    cursor: pointer; font-size: 1.1rem; line-height: 1;
    color: white;
  }
  .theme-toggle-learn:hover { background: rgba(255,255,255,0.3); }

  .container { max-width: 900px; margin: 0 auto; padding: 2rem 1.5rem; }

  /* Section cards */
  .section {
    background: var(--card); border: 1px solid var(--border);
    border-radius: 16px; padding: 2rem; margin-bottom: 2rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
  }
  .section h2 {
    font-size: 1.4rem; margin-bottom: 1rem;
    display: flex; align-items: center; gap: 0.5rem;
  }

  /* Comparison cards */
  .compare {
    display: grid; grid-template-columns: 1fr 1fr; gap: 1.5rem;
    margin: 1.5rem 0;
  }
  .compare-card {
    padding: 1.5rem; border-radius: 12px; position: relative;
  }
  .compare-card.bad {
    background: linear-gradient(135deg, #fef2f2, #fff7ed);
    border: 1px solid #fecaca;
  }
  .compare-card.good {
    background: linear-gradient(135deg, #f0fdf4, #eff6ff);
    border: 1px solid #bbf7d0;
  }
  .compare-card h3 { font-size: 1rem; margin-bottom: 0.75rem; }
  .compare-card ul { padding-left: 1.2rem; font-size: 0.9rem; }
  .compare-card li { margin-bottom: 0.35rem; }
  .compare-card .tag {
    position: absolute; top: -10px; right: 16px;
    padding: 0.15rem 0.75rem; border-radius: 20px;
    font-size: 0.75rem; font-weight: 700; color: white;
  }
  .compare-card.bad .tag { background: var(--red); }
  .compare-card.good .tag { background: var(--green); }

  /* Animated growth bar */
  .growth-demo {
    background: #f1f5f9; border-radius: 12px; padding: 1.5rem;
    margin: 1.5rem 0; text-align: center;
  }
  .growth-bars {
    display: flex; align-items: flex-end; justify-content: center;
    gap: 1rem; height: 180px; margin: 1rem 0;
  }
  .growth-bar {
    width: 60px; background: var(--accent); border-radius: 8px 8px 0 0;
    transition: height 1s ease; position: relative;
    display: flex; align-items: flex-end; justify-content: center;
  }
  .growth-bar .bar-label {
    position: absolute; bottom: -22px; font-size: 0.7rem;
    color: var(--muted); white-space: nowrap;
  }
  .growth-bar .bar-val {
    color: white; font-size: 0.75rem; font-weight: 700;
    padding: 0.25rem; text-align: center;
  }
  .growth-bar.danger { background: var(--red); }
  .growth-bar.warning { background: var(--yellow); }

  /* Scenario toggle */
  .scenario-toggle {
    display: flex; gap: 0; border-radius: 10px; overflow: hidden;
    border: 2px solid var(--accent); margin: 1rem 0; width: fit-content;
  }
  .scenario-btn {
    padding: 0.5rem 1.25rem; border: none; background: white;
    cursor: pointer; font-size: 0.9rem; font-weight: 600;
    color: var(--accent); transition: all 0.2s;
  }
  .scenario-btn.active {
    background: var(--accent); color: white;
  }
  .scenario-content {
    display: none; padding: 1.25rem; border-radius: 10px;
    background: #f8fafc; border: 1px solid var(--border);
    margin-top: 0.75rem; font-size: 0.9rem;
    animation: fadeIn 0.3s ease;
  }
  .scenario-content.active { display: block; }
  @keyframes fadeIn { from { opacity: 0; transform: translateY(-5px); } to { opacity: 1; } }

  /* Timeline */
  .timeline { position: relative; padding-left: 2rem; margin: 1.5rem 0; }
  .timeline::before {
    content: ''; position: absolute; left: 8px; top: 0; bottom: 0;
    width: 3px; background: var(--border); border-radius: 3px;
  }
  .timeline-item {
    position: relative; margin-bottom: 1.5rem; padding-left: 1.5rem;
  }
  .timeline-item::before {
    content: ''; position: absolute; left: -1.55rem; top: 6px;
    width: 14px; height: 14px; border-radius: 50%;
    background: var(--accent); border: 3px solid white;
    box-shadow: 0 0 0 2px var(--accent);
  }
  .timeline-item.red::before { background: var(--red); box-shadow: 0 0 0 2px var(--red); }
  .timeline-item h4 { font-size: 0.95rem; margin-bottom: 0.25rem; }
  .timeline-item p { font-size: 0.85rem; color: var(--muted); }

  /* Quiz */
  .quiz {
    background: linear-gradient(135deg, #fef9c3, #fef3c7);
    border: 1px solid #fde68a; border-radius: 12px;
    padding: 1.5rem; margin: 1.5rem 0;
  }
  .quiz h3 { font-size: 1rem; margin-bottom: 1rem; }
  .quiz-option {
    display: block; width: 100%; text-align: left;
    padding: 0.75rem 1rem; margin-bottom: 0.5rem;
    border: 2px solid #e5e7eb; border-radius: 10px;
    background: white; cursor: pointer; font-size: 0.9rem;
    transition: all 0.2s; font-family: inherit;
  }
  .quiz-option:hover { border-color: var(--accent); background: #eff6ff; }
  .quiz-option.correct {
    border-color: var(--green); background: #f0fdf4;
    pointer-events: none;
  }
  .quiz-option.wrong {
    border-color: var(--red); background: #fef2f2;
    pointer-events: none;
  }
  .quiz-result {
    display: none; padding: 0.75rem 1rem; margin-top: 0.75rem;
    border-radius: 8px; font-size: 0.9rem;
  }
  .quiz-result.visible { display: block; }
  .quiz-result.correct { background: #f0fdf4; color: #15803d; }
  .quiz-result.wrong { background: #fef2f2; color: #b91c1c; }

  /* Checklist */
  .checklist { list-style: none; padding: 0; }
  .checklist li {
    padding: 0.6rem 0; font-size: 0.9rem;
    display: flex; align-items: center; gap: 0.5rem;
    border-bottom: 1px solid var(--border);
    cursor: pointer; user-select: none;
  }
  .checklist li:last-child { border-bottom: none; }
  .checklist li .check {
    width: 22px; height: 22px; border: 2px solid var(--border);
    border-radius: 6px; display: flex; align-items: center;
    justify-content: center; flex-shrink: 0; transition: all 0.2s;
  }
  .checklist li.checked .check {
    background: var(--green); border-color: var(--green); color: white;
  }

  /* CTA */
  .cta {
    text-align: center; padding: 2rem;
    background: linear-gradient(135deg, #1e40af, #3b82f6);
    border-radius: 16px; color: white; margin: 2rem 0;
  }
  .cta h2 { font-size: 1.3rem; margin-bottom: 0.75rem; }
  .cta p { opacity: 0.9; margin-bottom: 1.25rem; }
  .cta a {
    display: inline-block; padding: 0.75rem 2rem;
    background: white; color: var(--accent); border-radius: 10px;
    text-decoration: none; font-weight: 700; transition: transform 0.1s;
  }
  .cta a:hover { transform: translateY(-2px); }

  .footer {
    text-align: center; padding: 1.5rem; color: var(--muted);
    font-size: 0.8rem;
  }

  @media (max-width: 700px) {
    .compare { grid-template-columns: 1fr; }
    .hero h1 { font-size: 1.6rem; }
    .growth-bars { gap: 0.5rem; }
    .growth-bar { width: 40px; }
  }
</style>
</head>
<body>

<div class="hero" style="position:relative;">
  <a href="/" class="back-link">← Zurück zum Checker</a>
  <button class="theme-toggle-learn" onclick="toggleTheme()" title="Dark/Light Mode">🌙</button>
  <h1>📊 Wann gehört Excel in eine Datenbank?</h1>
  <p>Eine interaktive Reise – warum dein Excel ab einer gewissen Komplexität leidet und was du dagegen tun kannst.</p>
</div>

<div class="container">

<!-- ── Section 1: Wachstums-Demo ── -->
<div class="section">
  <h2>📈 Das Wachstumsproblem</h2>
  <p>Excel ist großartig für kleine Aufgaben. Aber Daten wachsen – und irgendwann wird's eng:</p>

  <div class="growth-demo">
    <p style="font-weight:600;margin-bottom:0.5rem;">Wie fühlt sich dein Excel über die Jahre an?</p>
    <div class="growth-bars" id="growthBars">
      <div class="growth-bar" style="height:20px;"><span class="bar-val">50</span><span class="bar-label">2020</span></div>
      <div class="growth-bar" style="height:40px;"><span class="bar-val">200</span><span class="bar-label">2021</span></div>
      <div class="growth-bar" style="height:70px;"><span class="bar-val">800</span><span class="bar-label">2022</span></div>
      <div class="growth-bar warning" style="height:110px;"><span class="bar-val">3.000</span><span class="bar-label">2023</span></div>
      <div class="growth-bar danger" style="height:150px;"><span class="bar-val">15.000</span><span class="bar-label">2024</span></div>
      <div class="growth-bar danger" style="height:175px;"><span class="bar-val">50.000</span><span class="bar-label">2025</span></div>
    </div>
    <p style="font-size:0.85rem;color:var(--muted);margin-top:1.5rem;">
      ⬆️ Zeilen in einer typischen Tracking-Datei. Ab ~5.000 Zeilen wird es <strong>spürbar langsam</strong>.
    </p>
  </div>

  <div class="scenario-toggle">
    <button class="scenario-btn active" onclick="showScenario(this, 'sc-small')">📄 Kleine Datei</button>
    <button class="scenario-btn" onclick="showScenario(this, 'sc-medium')">📊 Mittelgroß</button>
    <button class="scenario-btn" onclick="showScenario(this, 'sc-large')">💥 Riesig</button>
  </div>

  <div class="scenario-content active" id="sc-small">
    <strong>✅ Excel ist perfekt dafür!</strong><br>
    Bis ~1.000 Zeilen, wenige Formeln, 1-2 Personen arbeiten damit.
    Hier spielt Excel seine Stärken aus: schnell, flexibel, jeder kennt es.
    <br><br>
    <em style="color:var(--green);">→ Alles gut! Weiter so. 👍</em>
  </div>
  <div class="scenario-content" id="sc-medium">
    <strong>⚡ Erste Warnsignale:</strong><br>
    5.000-50.000 Zeilen, verschachtelte SVERWEISe, mehrere Personen.
    Die Datei wird langsam, SVERWEISe über mehrere Sheets brechen
    gelegentlich, und niemand weiß genau, ob die Zahlen noch stimmen.
    <br><br>
    <em style="color:var(--yellow);">→ Über SharePoint-Liste oder Power BI nachdenken!</em>
  </div>
  <div class="scenario-content" id="sc-large">
    <strong>🚨 Dringender Handlungsbedarf:</strong><br>
    100.000+ Zeilen, 20+ Sheets, externe Verknüpfungen, 5+ Personen.
    Die Datei ist >20 MB, öffnet sich minutenlang, Formeln rechnen
    ewig, und keiner traut sich mehr etwas zu ändern.
    <br><br>
    <em style="color:var(--red);">→ Migration in eine Datenbank + professionelles Interface dringend empfohlen!</em>
  </div>
</div>

<!-- ── Section 2: Excel vs. Datenbank ── -->
<div class="section">
  <h2>⚔️ Excel vs. Datenbank – der ehrliche Vergleich</h2>
  <p>Beides hat seine Berechtigung. Die Frage ist nur: <strong>ab wann</strong>?</p>

  <div class="compare">
    <div class="compare-card bad">
      <span class="tag">Excel bei hoher Komplexität</span>
      <h3>😓 Das Problem</h3>
      <ul>
        <li>Datei wird <strong>immer langsamer</strong></li>
        <li>Formeln brechen bei Änderungen</li>
        <li>Keine Versionskontrolle – wer hat was geändert?</li>
        <li>Gleichzeitiges Arbeiten → <strong>Konflikte</strong></li>
        <li>"Kopie_final_v3_NEU.xlsx" 📂</li>
        <li>Ein falscher Klick löscht wichtige Daten</li>
      </ul>
    </div>
    <div class="compare-card good">
      <span class="tag">Datenbank + Interface</span>
      <h3>😊 Die Lösung</h3>
      <ul>
        <li>Blitzschnell, auch bei <strong>Millionen Zeilen</strong></li>
        <li>Regeln & Validierung automatisch</li>
        <li>Audit-Trail: <strong>wer, wann, was</strong></li>
        <li>Gleichzeitig arbeiten ohne Konflikte</li>
        <li>Eine Datenquelle, eine Wahrheit</li>
        <li>Zugriffsrechte schützen sensible Daten</li>
      </ul>
    </div>
  </div>
</div>

<!-- ── Section 3: Zusammenarbeit ── -->
<div class="section">
  <h2>👥 Zusammenarbeit – was Excel 365 kann und was nicht</h2>
  <p>Gute Nachricht: Mit <strong>Microsoft 365 und SharePoint/OneDrive</strong> können mehrere Personen gleichzeitig
     an derselben Excel-Datei arbeiten. Das ist ein großer Fortschritt! Aber auch hier gibt es Grenzen:</p>

  <div class="scenario-toggle">
    <button class="scenario-btn active" onclick="showScenario(this, 'collab-ok')">✅ Was geht</button>
    <button class="scenario-btn" onclick="showScenario(this, 'collab-limits')">⚠️ Die Grenzen</button>
    <button class="scenario-btn" onclick="showScenario(this, 'collab-db')">🏗️ Was besser wäre</button>
  </div>

  <div class="scenario-content active" id="collab-ok">
    <strong>Excel 365 Co-Authoring – das funktioniert gut:</strong>
    <ul style="padding-left:1.2rem;margin-top:0.5rem;font-size:0.9rem;">
      <li>✅ Mehrere Personen können <strong>gleichzeitig in derselben Datei</strong> arbeiten</li>
      <li>✅ Änderungen werden in <strong>Echtzeit synchronisiert</strong></li>
      <li>✅ <strong>Versionsverlauf</strong> über SharePoint – man kann ältere Stände wiederherstellen</li>
      <li>✅ <strong>@Mentions in Kommentaren</strong> für Team-Kommunikation</li>
      <li>✅ AutoSave verhindert Datenverlust bei Abstürzen</li>
    </ul>
    <p style="margin-top:0.75rem;font-size:0.9rem;color:var(--green);font-weight:600;">
      → Für einfache Dateien mit wenigen Personen ist das top! 👍
    </p>
  </div>
  <div class="scenario-content" id="collab-limits">
    <strong>⚠️ Wo Excel 365 Co-Authoring an seine Grenzen stößt:</strong>
    <ul style="padding-left:1.2rem;margin-top:0.5rem;font-size:0.9rem;">
      <li>🚫 <strong>Kein echter Audit-Trail</strong> – du siehst, DASS sich etwas geändert hat, aber nicht WER genau WELCHE Zelle geändert hat</li>
      <li>🚫 <strong>Konflikte bei Formeln:</strong> Wenn Anna und Max gleichzeitig eine Formel bearbeiten, gewinnt der Letzte – ohne Warnung</li>
      <li>🚫 <strong>Performance-Probleme:</strong> Ab ~5 gleichzeitigen Nutzern + großen Dateien wird Co-Authoring spürbar langsam</li>
      <li>🚫 <strong>Makros & VBA funktionieren nicht</strong> im Co-Authoring-Modus</li>
      <li>🚫 <strong>Keine feldgenauen Berechtigungen:</strong> Entweder jemand darf die ganze Datei bearbeiten – oder gar nicht</li>
      <li>🚫 <strong>Strukturänderungen blockieren:</strong> Sheets verschieben, Spalten einfügen → andere Nutzer werden kurzfristig ausgesperrt</li>
      <li>🚫 <strong>Kein Genehmigungs-Workflow:</strong> Jede:r kann alles ändern, keine Freigabe-Prozesse</li>
    </ul>
    <p style="margin-top:0.75rem;font-size:0.9rem;color:var(--yellow);font-weight:600;">
      → Je komplexer die Datei, desto mehr dieser Probleme treten auf.
    </p>
  </div>
  <div class="scenario-content" id="collab-db">
    <strong>🏗️ Was ein echtes Datenbank-System besser kann:</strong>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin-top:0.75rem;font-size:0.85rem;">
      <div style="padding:0.75rem;background:#fef2f2;border-radius:8px;border:1px solid #fecaca;">
        <strong>Excel 365 Co-Authoring</strong>
        <ul style="padding-left:1rem;margin-top:0.25rem;color:var(--muted);">
          <li>Wer hat was geändert? → unklar</li>
          <li>Feld-Berechtigungen? → nein</li>
          <li>Genehmigungen? → manuell per Mail</li>
          <li>100 Nutzer gleichzeitig? → unmöglich</li>
          <li>Historische Auswertung? → mühsam</li>
        </ul>
      </div>
      <div style="padding:0.75rem;background:#f0fdf4;border-radius:8px;border:1px solid #bbf7d0;">
        <strong>SharePoint-Liste / Power App / DB</strong>
        <ul style="padding-left:1rem;margin-top:0.25rem;color:var(--muted);">
          <li>Audit-Trail auf Feldebene ✅</li>
          <li>Spalten-Berechtigungen ✅</li>
          <li>Power Automate Workflows ✅</li>
          <li>Hunderte Nutzer parallel ✅</li>
          <li>Power BI Reports live ✅</li>
        </ul>
      </div>
    </div>
    <p style="margin-top:0.75rem;font-size:0.9rem;color:var(--green);font-weight:600;">
      → Ab einer gewissen Komplexität lohnt sich der Umstieg – und er ist einfacher als man denkt!
    </p>
  </div>
</div>

<!-- ── Section 4: Quiz ── -->
<div class="section">
  <h2>🧠 Schnell-Check: Brauche ich eine Datenbank?</h2>

  <div class="quiz" id="quiz1">
    <h3>Frage 1: Dein Excel hat 25.000 Zeilen und 12 Sheets mit SVERWEISen. Was tun?</h3>
    <button class="quiz-option" onclick="quizAnswer(this, false)">A) Noch mehr Sheets anlegen 😅</button>
    <button class="quiz-option" onclick="quizAnswer(this, false)">B) Einen stärkeren PC kaufen 💻</button>
    <button class="quiz-option" onclick="quizAnswer(this, true)">C) Prüfen, ob eine SharePoint-Liste oder Datenbank besser passt 🎯</button>
    <button class="quiz-option" onclick="quizAnswer(this, false)">D) Einfach weitermachen und hoffen 🤞</button>
    <div class="quiz-result" id="quiz1-result"></div>
  </div>

  <div class="quiz" id="quiz2">
    <h3>Frage 2: 3 Kolleg:innen melden "Die Datei ist gesperrt". Was ist die Ursache?</h3>
    <button class="quiz-option" onclick="quizAnswer(this, false)">A) IT hat den Zugriff eingeschränkt</button>
    <button class="quiz-option" onclick="quizAnswer(this, true)">B) Excel kann nur einen gleichzeitigen Bearbeiter – das ist ein Architekturproblem 🏗️</button>
    <button class="quiz-option" onclick="quizAnswer(this, false)">C) Die Datei ist kaputt</button>
    <button class="quiz-option" onclick="quizAnswer(this, false)">D) Das ist normal und kein Problem</button>
    <div class="quiz-result" id="quiz2-result"></div>
  </div>

  <div class="quiz" id="quiz3">
    <h3>Frage 3: Woran erkennst du, dass dein Excel eigentlich eine Datenbank sein sollte?</h3>
    <button class="quiz-option" onclick="quizAnswer(this, false)">A) Es hat mehr als 3 Farben 🎨</button>
    <button class="quiz-option" onclick="quizAnswer(this, false)">B) Der Dateiname enthält "final"</button>
    <button class="quiz-option" onclick="quizAnswer(this, true)">C) Mehrere Sheets referenzieren sich gegenseitig mit Lookups, und du fragst dich "wer hat das zuletzt geändert?" 🔍</button>
    <button class="quiz-option" onclick="quizAnswer(this, false)">D) Die Datei hat ein hübsches Logo</button>
    <div class="quiz-result" id="quiz3-result"></div>
  </div>

  <div class="quiz" id="quiz4">
    <h3>Frage 4: Du kopierst regelmäßig Daten aus SAP/einem Fachsystem in ein Excel, um sie dort weiterzuverarbeiten. Was ist das Problem?</h3>
    <button class="quiz-option" onclick="quizAnswer(this, false)">A) Kein Problem – ich mache das schon seit Jahren so 🤷</button>
    <button class="quiz-option" onclick="quizAnswer(this, false)">B) Das ist nur ineffizient, wenn man viele Daten hat</button>
    <button class="quiz-option" onclick="quizAnswer(this, true)">C) Das ist ein Medienbruch – die Daten sollten per Schnittstelle direkt verbunden werden, sonst drohen Fehler, Verzögerungen und doppelte Pflege 🔗</button>
    <button class="quiz-option" onclick="quizAnswer(this, false)">D) Der Aufwand ist zu gering, um darüber nachzudenken</button>
    <div class="quiz-result" id="quiz4-result"></div>
  </div>
</div>

<!-- ── Section 5: Wann ist Excel OK? ── -->
<div class="section">
  <h2>✅ Wann ist Excel die richtige Wahl?</h2>
  <p>Excel ist kein schlechtes Tool – es wird nur manchmal falsch eingesetzt. Hier ein ehrlicher Check:</p>

  <ul class="checklist" id="checklist">
    <li onclick="toggleCheck(this)"><span class="check"></span> Weniger als 5.000 Zeilen</li>
    <li onclick="toggleCheck(this)"><span class="check"></span> 1-2 Personen arbeiten damit</li>
    <li onclick="toggleCheck(this)"><span class="check"></span> Keine verschachtelten SVERWEIS/INDEX-Ketten</li>
    <li onclick="toggleCheck(this)"><span class="check"></span> Keine externen Verknüpfungen zu anderen Dateien</li>
    <li onclick="toggleCheck(this)"><span class="check"></span> Niemand kopiert regelmäßig Daten aus einem anderen System hinein</li>
    <li onclick="toggleCheck(this)"><span class="check"></span> Keiner fragt "Wer hat das geändert?"</li>
    <li onclick="toggleCheck(this)"><span class="check"></span> Die Datei öffnet sich in unter 10 Sekunden</li>
    <li onclick="toggleCheck(this)"><span class="check"></span> Es gibt keine Datei-Kopien mit "_v2", "_final", "_NEU"</li>
  </ul>

  <div id="checklist-result" style="margin-top:1rem;padding:1rem;border-radius:10px;display:none;font-weight:600;text-align:center;"></div>
</div>

<!-- ── Section 6: AI-Readiness Maturity Model ── -->
<div class="section">
  <h2>🤖 Der eigentliche Grund: AI-Readiness</h2>
  <p>KI-Tools wie Copilot können grundsätzlich mit <strong>jeder Art von Daten</strong> arbeiten – auch mit Excel.
     Aber: Je sauberer und strukturierter die Daten sind, desto <strong>zuverlässiger, schneller und skalierbarer</strong>
     werden die Ergebnisse. Eine 30-MB-Excel-Datei mit Farbcodes und versteckten Sheets kann eine KI zwar lesen –
     aber sie muss viel mehr raten und liefert weniger präzise Antworten.</p>
  <p style="margin-top:0.5rem;">Das Ziel: <strong>Daten so aufbereiten, dass KI ihr volles Potenzial entfalten kann</strong> –
     idealerweise in strukturierten Datenquellen, die automatisiert und zuverlässig abfragbar sind.</p>

  <div style="margin-top:1.5rem;">
    <h3 style="font-size:1.1rem;margin-bottom:1rem;text-align:center;">📊 Daten-Reifegrad-Modell – Wo stehst du?</h3>
    <p style="text-align:center;font-size:0.85rem;color:var(--muted);margin-bottom:1.5rem;">Klicke auf eine Stufe, um mehr zu erfahren.</p>
    <div style="display:flex;gap:4px;margin-bottom:1rem;height:12px;border-radius:6px;overflow:hidden;">
      <div style="flex:1;background:#dc2626;" title="Level 1"></div>
      <div style="flex:1;background:#ea580c;" title="Level 2"></div>
      <div style="flex:1;background:#ca8a04;" title="Level 3"></div>
      <div style="flex:1;background:#16a34a;" title="Level 4"></div>
      <div style="flex:1;background:#2563eb;" title="Level 5"></div>
    </div>
    <div style="display:flex;flex-direction:column;gap:0.5rem;" id="maturityLevels">
      <div class="maturity-level" onclick="toggleMaturity(this)" style="cursor:pointer;padding:1rem;border-radius:10px;border-left:4px solid #dc2626;background:var(--card);border-top:1px solid var(--border);border-right:1px solid var(--border);border-bottom:1px solid var(--border);">
        <div style="display:flex;align-items:center;gap:0.5rem;">
          <span style="background:#dc2626;color:white;padding:0.1rem 0.5rem;border-radius:4px;font-size:0.75rem;font-weight:700;">Level 1</span>
          <strong>Excel-Wildwuchs</strong>
          <span style="margin-left:auto;color:var(--muted);font-size:0.9rem;">▸</span>
        </div>
        <div class="maturity-detail" style="display:none;margin-top:0.75rem;font-size:0.85rem;color:var(--muted);">
          <p>📂 Dutzende Excel-Dateien auf persönlichen Laufwerken, per Mail verschickt, verschiedene Versionen.
             Farbcodes als Logik, hartcodierte Werte, keine Dokumentation.</p>
          <p style="margin-top:0.5rem;"><strong>KI-Tauglichkeit:</strong> ❌ Nahe null. Daten sind nicht auffindbar, nicht standardisiert, nicht maschinenlesbar.</p>
          <p style="margin-top:0.25rem;"><strong>Typisch:</strong> <em>"Die Datei hat nur die Sabine, und die ist im Urlaub."</em></p>
        </div>
      </div>
      <div class="maturity-level" onclick="toggleMaturity(this)" style="cursor:pointer;padding:1rem;border-radius:10px;border-left:4px solid #ea580c;background:var(--card);border-top:1px solid var(--border);border-right:1px solid var(--border);border-bottom:1px solid var(--border);">
        <div style="display:flex;align-items:center;gap:0.5rem;">
          <span style="background:#ea580c;color:white;padding:0.1rem 0.5rem;border-radius:4px;font-size:0.75rem;font-weight:700;">Level 2</span>
          <strong>Zentrales Excel</strong>
          <span style="margin-left:auto;color:var(--muted);font-size:0.9rem;">▸</span>
        </div>
        <div class="maturity-detail" style="display:none;margin-top:0.75rem;font-size:0.85rem;color:var(--muted);">
          <p>📁 Excel-Dateien liegen zentral auf SharePoint/Teams. Co-Authoring wird genutzt.
             Aber: Die Dateien werden immer größer, Formeln immer verschachtelter, und niemand traut sich, aufzuräumen.</p>
          <p style="margin-top:0.5rem;"><strong>KI-Tauglichkeit:</strong> ⚠️ Eingeschränkt. Daten sind auffindbar, aber das Format ist nicht standardisiert.</p>
          <p style="margin-top:0.25rem;"><strong>Typisch:</strong> <em>"Die Datei ist 45 MB und niemand weiß, welche Formeln noch stimmen."</em></p>
        </div>
      </div>
      <div class="maturity-level" onclick="toggleMaturity(this)" style="cursor:pointer;padding:1rem;border-radius:10px;border-left:4px solid #ca8a04;background:var(--card);border-top:1px solid var(--border);border-right:1px solid var(--border);border-bottom:1px solid var(--border);">
        <div style="display:flex;align-items:center;gap:0.5rem;">
          <span style="background:#ca8a04;color:white;padding:0.1rem 0.5rem;border-radius:4px;font-size:0.75rem;font-weight:700;">Level 3</span>
          <strong>Strukturierte Listen</strong>
          <span style="margin-left:auto;color:var(--muted);font-size:0.9rem;">▸</span>
        </div>
        <div class="maturity-detail" style="display:none;margin-top:0.75rem;font-size:0.85rem;color:var(--muted);">
          <p>📋 Die wichtigsten Daten leben in <strong>SharePoint-Listen</strong> oder einfachen Datenbanken.
             Excel wird nur noch für Ad-hoc-Analysen genutzt, nicht mehr als Datenhaltung.</p>
          <p style="margin-top:0.5rem;"><strong>KI-Tauglichkeit:</strong> 🟡 Gut für einfache Abfragen. Daten sind strukturiert und über APIs zugänglich.</p>
          <p style="margin-top:0.25rem;"><strong>Typisch:</strong> <em>"Unsere Stammdaten pflegen wir über ein Formular, Excel nutzen wir nur zum Auswerten."</em></p>
        </div>
      </div>
      <div class="maturity-level" onclick="toggleMaturity(this)" style="cursor:pointer;padding:1rem;border-radius:10px;border-left:4px solid #16a34a;background:var(--card);border-top:1px solid var(--border);border-right:1px solid var(--border);border-bottom:1px solid var(--border);">
        <div style="display:flex;align-items:center;gap:0.5rem;">
          <span style="background:#16a34a;color:white;padding:0.1rem 0.5rem;border-radius:4px;font-size:0.75rem;font-weight:700;">Level 4</span>
          <strong>Datenbank + Professionelles UI</strong>
          <span style="margin-left:auto;color:var(--muted);font-size:0.9rem;">▸</span>
        </div>
        <div class="maturity-detail" style="display:none;margin-top:0.75rem;font-size:0.85rem;color:var(--muted);">
          <p>🗄️ Echte Datenbanken (SQL, Dataverse) mit <strong>Power Apps oder Web-Interfaces</strong> als Frontend.
             Audit-Trail, Zugriffsrechte, Validierung, Workflows – alles automatisiert.</p>
          <p style="margin-top:0.5rem;"><strong>KI-Tauglichkeit:</strong> ✅ Sehr gut. Daten sind sauber, standardisiert, und über APIs erreichbar. KI kann direkt zugreifen.</p>
          <p style="margin-top:0.25rem;"><strong>Typisch:</strong> <em>"Power BI zeigt live die Dashboards, Copilot kann unsere Daten abfragen."</em></p>
        </div>
      </div>
      <div class="maturity-level" onclick="toggleMaturity(this)" style="cursor:pointer;padding:1rem;border-radius:10px;border-left:4px solid #2563eb;background:var(--card);border-top:1px solid var(--border);border-right:1px solid var(--border);border-bottom:1px solid var(--border);">
        <div style="display:flex;align-items:center;gap:0.5rem;">
          <span style="background:#2563eb;color:white;padding:0.1rem 0.5rem;border-radius:4px;font-size:0.75rem;font-weight:700;">Level 5</span>
          <strong>🌟 AI-Ready Data Platform</strong>
          <span style="margin-left:auto;color:var(--muted);font-size:0.9rem;">▸</span>
        </div>
        <div class="maturity-detail" style="display:none;margin-top:0.75rem;font-size:0.85rem;color:var(--muted);">
          <p>🚀 Alle Unternehmensdaten in einem <strong>Data Lakehouse / Data Mesh</strong> mit klarer Governance.
             Metadaten katalogisiert, Datenqualität automatisch überwacht, APIs für alles.
             KI-Modelle können direkt auf saubere, dokumentierte Datenquellen zugreifen.</p>
          <p style="margin-top:0.5rem;"><strong>KI-Tauglichkeit:</strong> 🌟 Perfekt. KI kann Muster erkennen, Vorhersagen treffen, und automatisiert Entscheidungen unterstützen.</p>
          <p style="margin-top:0.25rem;"><strong>Typisch:</strong> <em>"Unsere KI hat das Nachfragemuster erkannt und automatisch den Einkauf informiert."</em></p>
        </div>
      </div>
    </div>
  </div>

  <div style="margin-top:1.5rem;padding:1.25rem;background:linear-gradient(135deg,#eff6ff 0%,#faf5ff 100%);border:1px solid #bfdbfe;border-radius:10px;">
    <strong>🎯 Die Kernbotschaft:</strong>
    <p style="margin-top:0.5rem;font-size:0.9rem;">Jedes Excel, das du in eine saubere Datenstruktur überführst, ist ein Schritt in Richtung KI-Readiness.
       Es geht nicht darum, Excel zu verbieten – sondern darum, <strong>die richtigen Daten am richtigen Ort</strong> zu haben.</p>
    <p style="margin-top:0.5rem;font-size:0.9rem;">Dieser Checker hilft dir, den ersten Schritt zu machen: <strong>Erkennen, welche Dateien bereit für den nächsten Level sind.</strong></p>
  </div>
</div>

<!-- ── Section 7: Effizienz & Kosten ── -->
<div class="section" id="effizienz">
  <h2>💰 Was aufgeblähte Excel-Dateien wirklich kosten</h2>
  <p>Excel ist oft der richtige Startpunkt – schnell, verfügbar, keine IT-Freigabe nötig.
     Aber ab einer gewissen Komplexität kippt das Kosten-Nutzen-Verhältnis.
     Und was bei einer einzelnen Person wie „nur ein paar Minuten" wirkt, wird bei <strong>1.000+ Mitarbeitern</strong>
     zu einer massiven versteckten Kostenstelle.</p>

  <!-- Szenarien -->
  <h3 style="font-size:1.1rem;margin-top:2rem;margin-bottom:1rem;">👥 Drei Szenarien aus dem Arbeitsalltag</h3>
  <p style="color:var(--muted);font-size:0.85rem;margin-bottom:1.5rem;">Alle Berechnungen konservativ: Vollkosten €65/h, 48 Arbeitswochen/Jahr, niedrigste realistische Schätzung.</p>

  <div style="display:flex;flex-direction:column;gap:1.25rem;">

    <!-- Szenario 1: Gelegenheitsnutzer -->
    <div style="padding:1.5rem;border-radius:12px;background:linear-gradient(135deg,#eff6ff,#ffffff);border:1px solid #bfdbfe;border-left:4px solid #3b82f6;">
      <div style="display:flex;align-items:center;gap:0.5rem;margin-bottom:0.75rem;">
        <span style="font-size:1.5rem;">📋</span>
        <h4 style="font-size:1rem;margin:0;">Szenario 1: „Ich füll nur kurz was aus"</h4>
        <span style="margin-left:auto;font-size:0.75rem;padding:0.2rem 0.6rem;border-radius:4px;background:#dbeafe;color:#1e40af;font-weight:600;">~600 Mitarbeiter</span>
      </div>
      <p style="font-size:0.85rem;color:var(--muted);margin-bottom:0.75rem;">
        Mitarbeiter, die punktuell Daten in ein Excel eintragen müssen – Zeiterfassung, Bestelllisten,
        Statusmeldungen, Investitionsanträge. Sie sind keine „Excel-Profis" und dürfen oft <strong>nicht mal
        sortieren oder filtern</strong>, weil die Sorge besteht, dass Formeln oder Struktur kaputtgehen.</p>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin-top:0.75rem;">
        <div style="padding:0.75rem;background:white;border-radius:8px;border:1px solid #e2e8f0;">
          <div style="font-size:0.8rem;color:var(--muted);">Ineffizienz pro Person</div>
          <div style="font-size:0.85rem;margin-top:0.25rem;">Excel öffnen, richtige Zelle finden, vorsichtig eintragen, nichts kaputt machen, speichern, hoffen.</div>
          <div style="font-weight:700;color:#dc2626;margin-top:0.5rem;">~15 Min/Woche Mehraufwand</div>
          <div style="font-size:0.75rem;color:var(--muted);">(vs. ein einfaches Formular mit Validierung)</div>
        </div>
        <div style="padding:0.75rem;background:white;border-radius:8px;border:1px solid #e2e8f0;">
          <div style="font-size:0.8rem;color:var(--muted);">Hochrechnung bei 600 Betroffenen</div>
          <div style="display:flex;justify-content:space-between;padding:0.3rem 0;border-bottom:1px solid #f1f5f9;font-size:0.85rem;"><span>15 Min × 600 Personen</span><span style="font-weight:600;">150 h/Woche</span></div>
          <div style="display:flex;justify-content:space-between;padding:0.3rem 0;border-bottom:1px solid #f1f5f9;font-size:0.85rem;"><span>× 48 Wochen × €65</span><span style="font-weight:700;color:#dc2626;">€468.000/Jahr</span></div>
          <div style="font-size:0.75rem;color:var(--muted);margin-top:0.5rem;">Dazu kommen Fehler: falsche Zelle, überschriebene Formel, vergessener Eintrag.</div>
        </div>
      </div>
    </div>

    <!-- Szenario 2: Daten-Kopierer -->
    <div style="padding:1.5rem;border-radius:12px;background:linear-gradient(135deg,#fefce8,#ffffff);border:1px solid #fde68a;border-left:4px solid #ca8a04;">
      <div style="display:flex;align-items:center;gap:0.5rem;margin-bottom:0.75rem;">
        <span style="font-size:1.5rem;">🔄</span>
        <h4 style="font-size:1rem;margin:0;">Szenario 2: „Ich kopier das schnell aus SAP rüber"</h4>
        <span style="margin-left:auto;font-size:0.75rem;padding:0.2rem 0.6rem;border-radius:4px;background:#fef3c7;color:#92400e;font-weight:600;">~150 Mitarbeiter</span>
      </div>
      <p style="font-size:0.85rem;color:var(--muted);margin-bottom:0.75rem;">
        Mitarbeiter, die regelmäßig <strong>Daten aus Fachsystemen (SAP, SCADA, Maximo, …) ins Excel kopieren</strong>,
        um sie dort weiterzuverarbeiten, auszuwerten oder an andere weiterzugeben. Ein klassischer <strong>Medienbruch</strong>:
        Die Daten existieren bereits in einem System – werden aber manuell in ein anderes übertragen,
        wo sie sofort veralten, fehleranfällig werden und nicht mehr rückverfolgbar sind.</p>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin-top:0.75rem;">
        <div style="padding:0.75rem;background:white;border-radius:8px;border:1px solid #e2e8f0;">
          <div style="font-size:0.8rem;color:var(--muted);">Ineffizienz pro Person</div>
          <div style="font-size:0.85rem;margin-top:0.25rem;">System öffnen, Daten exportieren/kopieren, ins Excel einfügen, Format anpassen, Formeln aktualisieren, manuell prüfen.</div>
          <div style="font-weight:700;color:#dc2626;margin-top:0.5rem;">~45 Min/Woche Mehraufwand</div>
          <div style="font-size:0.75rem;color:var(--muted);">(vs. eine automatische Schnittstelle / Power Query)</div>
        </div>
        <div style="padding:0.75rem;background:white;border-radius:8px;border:1px solid #e2e8f0;">
          <div style="font-size:0.8rem;color:var(--muted);">Hochrechnung bei 150 Betroffenen</div>
          <div style="display:flex;justify-content:space-between;padding:0.3rem 0;border-bottom:1px solid #f1f5f9;font-size:0.85rem;"><span>45 Min × 150 Personen</span><span style="font-weight:600;">112 h/Woche</span></div>
          <div style="display:flex;justify-content:space-between;padding:0.3rem 0;border-bottom:1px solid #f1f5f9;font-size:0.85rem;"><span>× 48 Wochen × €65</span><span style="font-weight:700;color:#dc2626;">€349.000/Jahr</span></div>
          <div style="font-size:0.75rem;color:var(--muted);margin-top:0.5rem;">Plus das eigentliche Risiko: veraltete Daten führen zu falschen Entscheidungen.</div>
        </div>
      </div>
    </div>

    <!-- Szenario 3: Heavy User -->
    <div style="padding:1.5rem;border-radius:12px;background:linear-gradient(135deg,#fef2f2,#ffffff);border:1px solid #fecaca;border-left:4px solid #dc2626;">
      <div style="display:flex;align-items:center;gap:0.5rem;margin-bottom:0.75rem;">
        <span style="font-size:1.5rem;">⚙️</span>
        <h4 style="font-size:1rem;margin:0;">Szenario 3: „Ohne mein Excel geht hier gar nichts"</h4>
        <span style="margin-left:auto;font-size:0.75rem;padding:0.2rem 0.6rem;border-radius:4px;background:#fee2e2;color:#b91c1c;font-weight:600;">~50 Mitarbeiter</span>
      </div>
      <p style="font-size:0.85rem;color:var(--muted);margin-bottom:0.75rem;">
        Die „Excel-Helden": Abteilungsexperten, die <strong>geschäftskritische Excel-Dateien</strong> pflegen.
        30+ MB, verschachtelte SVERWEISe, Makros, Farblogik, Sheets die aufeinander aufbauen.
        Sie sind die Einzigen, die „ihre" Datei verstehen – und wenn sie krank sind oder gehen,
        steht die Abteilung still.</p>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin-top:0.75rem;">
        <div style="padding:0.75rem;background:white;border-radius:8px;border:1px solid #e2e8f0;">
          <div style="font-size:0.8rem;color:var(--muted);">Ineffizienz pro Person</div>
          <div style="font-size:0.85rem;margin-top:0.25rem;">Formelwartung, Daten abgleichen, Reports zusammenbauen, Kolleg:innen erklären warum man "da nicht reinsortieren darf", Fehler suchen.</div>
          <div style="font-weight:700;color:#dc2626;margin-top:0.5rem;">~6 Stunden/Woche Mehraufwand</div>
          <div style="font-size:0.75rem;color:var(--muted);">(vs. ein strukturiertes System mit automatisierten Reports)</div>
        </div>
        <div style="padding:0.75rem;background:white;border-radius:8px;border:1px solid #e2e8f0;">
          <div style="font-size:0.8rem;color:var(--muted);">Hochrechnung bei 50 Betroffenen</div>
          <div style="display:flex;justify-content:space-between;padding:0.3rem 0;border-bottom:1px solid #f1f5f9;font-size:0.85rem;"><span>6h × 50 Personen</span><span style="font-weight:600;">300 h/Woche</span></div>
          <div style="display:flex;justify-content:space-between;padding:0.3rem 0;border-bottom:1px solid #f1f5f9;font-size:0.85rem;"><span>× 48 Wochen × €65</span><span style="font-weight:700;color:#dc2626;">€936.000/Jahr</span></div>
          <div style="font-size:0.75rem;color:var(--muted);margin-top:0.5rem;">Nicht eingerechnet: Wissensverlust bei Personalwechsel, Fehlerkosten, Audit-Risiken.</div>
        </div>
      </div>
    </div>
  </div>

  <!-- Gesamtrechnung -->
  <div style="margin-top:2rem;padding:1.5rem;border-radius:12px;background:linear-gradient(135deg,#fef9c3 0%,#fef3c7 100%);border:2px solid #fde68a;">
    <h3 style="font-size:1.1rem;margin-bottom:1rem;color:#92400e;">📊 Gesamtrechnung: Was Excel das Unternehmen wirklich kostet</h3>

    <!-- Realisierungsgrad-Slider -->
    <div style="margin-bottom:1.25rem;padding:1rem;background:rgba(255,255,255,0.7);border-radius:10px;border:1px solid #fde68a;">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:1rem;flex-wrap:wrap;">
        <div style="flex:1;min-width:200px;">
          <label for="effSlider" style="font-size:0.9rem;font-weight:600;color:#92400e;">⚙️ Realisierungsgrad der Einsparung</label>
          <p style="font-size:0.78rem;color:var(--muted);margin-top:0.25rem;line-height:1.5;">
            Nicht jede eingesparte Minute wird zu 100% produktiv genutzt. Stelle hier ein,
            welcher Anteil der Zeitersparnis realistisch in wertschöpfende Arbeit fließt.
            30% ist ein konservativer Wert – selbst wenn nur ein Drittel der gesparten Zeit
            produktiv genutzt wird, sprechen die Zahlen eine klare Sprache.</p>
        </div>
        <div style="display:flex;align-items:center;gap:0.75rem;min-width:220px;">
          <input type="range" id="effSlider" min="10" max="100" value="30" step="5"
            style="flex:1;accent-color:#f59e0b;cursor:pointer;"
            oninput="updateEffCalc(this.value)">
          <span id="effPct" style="font-size:1.2rem;font-weight:700;color:#92400e;min-width:3rem;text-align:right;">30%</span>
        </div>
      </div>
    </div>

    <table style="width:100%;border-collapse:collapse;font-size:0.9rem;">
      <thead>
        <tr style="border-bottom:2px solid #fde68a;">
          <th style="padding:0.6rem 0;text-align:left;">Szenario</th>
          <th style="padding:0.6rem 0;text-align:right;">Betroffene</th>
          <th style="padding:0.6rem 0;text-align:right;">Pro Person/Woche</th>
          <th style="padding:0.6rem 0;text-align:right;">Brutto-Kosten</th>
          <th style="padding:0.6rem 0;text-align:right;">Realer Verlust</th>
        </tr>
      </thead>
      <tbody>
        <tr style="border-bottom:1px solid #fde68a;">
          <td style="padding:0.5rem 0;">📋 Gelegenheitsnutzer</td>
          <td style="padding:0.5rem 0;text-align:right;">600</td>
          <td style="padding:0.5rem 0;text-align:right;">15 Min</td>
          <td style="padding:0.5rem 0;text-align:right;color:var(--muted);">€468.000</td>
          <td id="eff-s1" style="padding:0.5rem 0;text-align:right;font-weight:600;">€140.400</td>
        </tr>
        <tr style="border-bottom:1px solid #fde68a;">
          <td style="padding:0.5rem 0;">🔄 Daten-Kopierer</td>
          <td style="padding:0.5rem 0;text-align:right;">150</td>
          <td style="padding:0.5rem 0;text-align:right;">45 Min</td>
          <td style="padding:0.5rem 0;text-align:right;color:var(--muted);">€349.000</td>
          <td id="eff-s2" style="padding:0.5rem 0;text-align:right;font-weight:600;">€104.700</td>
        </tr>
        <tr style="border-bottom:1px solid #fde68a;">
          <td style="padding:0.5rem 0;">⚙️ Heavy User</td>
          <td style="padding:0.5rem 0;text-align:right;">50</td>
          <td style="padding:0.5rem 0;text-align:right;">6 Stunden</td>
          <td style="padding:0.5rem 0;text-align:right;color:var(--muted);">€936.000</td>
          <td id="eff-s3" style="padding:0.5rem 0;text-align:right;font-weight:600;">€280.800</td>
        </tr>
        <tr style="border-top:2px solid #f59e0b;">
          <td style="padding:0.75rem 0;" colspan="3"><strong style="font-size:1rem;">Summe</strong></td>
          <td style="padding:0.75rem 0;text-align:right;color:var(--muted);font-size:0.85rem;">€1,75 Mio</td>
          <td id="eff-total" style="padding:0.75rem 0;text-align:right;font-weight:700;color:#dc2626;font-size:1.2rem;">~€526.000/Jahr</td>
        </tr>
      </tbody>
    </table>
    <p id="eff-note" style="font-size:0.82rem;color:#92400e;margin-top:1rem;line-height:1.6;padding:0.75rem;background:rgba(255,255,255,0.5);border-radius:8px;">
      📌 <strong>Sichtweise „Realisierungsgrad":</strong> Selbst bei nur <strong><span id="eff-note-pct">30</span>% Realisierung</strong>
      entsteht immer noch ein realer Verlust von <strong><span id="eff-note-sum">~€526.000</span>/Jahr</strong>.
      Die restlichen 70% sind keine „geschenkte Zeit" – sie versickern in Kontextwechseln,
      Wartezeiten und Mikrofrustration, die schwer messbar ist, aber die Arbeitszufriedenheit senkt.</p>
    <p style="font-size:0.8rem;color:#92400e;margin-top:0.75rem;line-height:1.6;">
      💡 Nicht eingerechnet: Fehlerkosten durch falsche Daten in Entscheidungen,
      Audit-Risiken bei nicht nachvollziehbaren Änderungen, Wissensverlust bei Pensionierung der „Excel-Experten",
      und die Opportunitätskosten – was könnten diese 800 Mitarbeiter tun, wenn sie sich auf ihre eigentliche Arbeit
      konzentrieren könnten statt auf Excel-Workarounds?</p>
    <p style="font-size:0.85rem;color:#92400e;margin-top:0.75rem;">
      <strong>Zum Vergleich:</strong> Eine professionelle Migration der 20 kritischsten Excel-Dateien in
      SharePoint-Listen oder eine einfache Datenbank kostet einmalig ca. €100.000–250.000.
      <strong id="eff-amort">Selbst bei 30% Realisierung: Amortisation in unter 6 Monaten.</strong></p>
  </div>

  <!-- Alltagsbeispiele -->
  <div style="margin-top:1.5rem;padding:1.25rem;border-radius:10px;background:var(--card);border:1px solid var(--border);">
    <h3 style="font-size:1rem;margin-bottom:1rem;">🎯 Erkennst du dich wieder?</h3>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:1rem;">
      <div style="padding:0.75rem;border-radius:8px;background:#fef2f2;border:1px solid #fecaca;font-size:0.85rem;">
        <strong style="color:#b91c1c;">❌ „Bitte nicht sortieren!"</strong>
        <p style="color:var(--muted);margin-top:0.25rem;">Du bekommst ein Excel zum Ausfüllen, aber darfst nicht filtern
          oder sortieren, weil sonst Formeln kaputtgehen. Also scrollst du manuell durch 500 Zeilen,
          um deine Zeile zu finden.</p>
      </div>
      <div style="padding:0.75rem;border-radius:8px;background:#fef2f2;border:1px solid #fecaca;font-size:0.85rem;">
        <strong style="color:#b91c1c;">❌ „Ich kopier das schnell rüber"</strong>
        <p style="color:var(--muted);margin-top:0.25rem;">Du exportierst wöchentlich Daten aus SAP,
          kopierst sie ins Excel, passt Spalten an, prüfst manuell. Die Daten sind in dem Moment
          schon veraltet, in dem du sie einfügst.</p>
      </div>
      <div style="padding:0.75rem;border-radius:8px;background:#fef2f2;border:1px solid #fecaca;font-size:0.85rem;">
        <strong style="color:#b91c1c;">❌ „Nur die Kollegin versteht das"</strong>
        <p style="color:var(--muted);margin-top:0.25rem;">Ein geschäftskritisches Excel mit 15 Sheets und 300 Formeln.
          Eine einzige Person kennt die Logik. Wenn sie im Urlaub ist, traut sich niemand
          etwas zu ändern.</p>
      </div>
      <div style="padding:0.75rem;border-radius:8px;background:#f0fdf4;border:1px solid #bbf7d0;font-size:0.85rem;">
        <strong style="color:#15803d;">✅ So geht es besser</strong>
        <p style="color:var(--muted);margin-top:0.25rem;">Eingabe über ein Formular mit Validierung. Daten aus
          Fachsystemen per Schnittstelle automatisch aktualisiert. Reports auf Knopfdruck.
          Jeder darf filtern, sortieren, auswerten – ohne etwas kaputt zu machen.</p>
      </div>
    </div>
  </div>

  <!-- KI-Potenzial -->
  <div style="margin-top:1.5rem;padding:1.25rem;border-radius:10px;background:linear-gradient(135deg,#eff6ff,#f0fdf4);border:1px solid #bfdbfe;">
    <h3 style="font-size:1rem;margin-bottom:0.75rem;">🚀 Ungenutztes KI-Potenzial</h3>
    <p style="font-size:0.9rem;color:var(--muted);margin-bottom:1rem;">
      KI braucht saubere, strukturierte Daten. Solange Wissen in Farbcodes, verbundenen Zellen und
      SVERWEIS-Ketten steckt, bleiben diese Möglichkeiten liegen:</p>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:1rem;">
      <div style="padding:0.75rem;background:var(--card);border-radius:8px;border:1px solid var(--border);">
        <div style="font-size:1.2rem;margin-bottom:0.25rem;">📋</div>
        <strong style="font-size:0.85rem;">Automatische Berichte</strong>
        <p style="font-size:0.8rem;color:var(--muted);margin-top:0.25rem;">
          Power BI / Dashboards auf Knopfdruck statt stundenlanger manueller Zusammenstellung.
          Ersparnis: ~2-5h/Woche pro Empfänger.</p>
      </div>
      <div style="padding:0.75rem;background:var(--card);border-radius:8px;border:1px solid var(--border);">
        <div style="font-size:1.2rem;margin-bottom:0.25rem;">🔍</div>
        <strong style="font-size:0.85rem;">Anomalie-Erkennung</strong>
        <p style="font-size:0.8rem;color:var(--muted);margin-top:0.25rem;">
          KI findet Ausreißer und Inkonsistenzen automatisch – aber nur wenn Spalten saubere Typen haben,
          nicht „Text-und-Zahlen-Salat".</p>
      </div>
      <div style="padding:0.75rem;background:var(--card);border-radius:8px;border:1px solid var(--border);">
        <div style="font-size:1.2rem;margin-bottom:0.25rem;">💬</div>
        <strong style="font-size:0.85rem;">Natürlichsprachliche Abfragen</strong>
        <p style="font-size:0.8rem;color:var(--muted);margin-top:0.25rem;">
          „Zeig mir alle Vorfälle im Q3 mit hoher Priorität" – in einer Datenbank in Sekunden.
          In Excel: manuelle Suche über 8 Sheets.</p>
      </div>
      <div style="padding:0.75rem;background:var(--card);border-radius:8px;border:1px solid var(--border);">
        <div style="font-size:1.2rem;margin-bottom:0.25rem;">📈</div>
        <strong style="font-size:0.85rem;">Prognosen & Trends</strong>
        <p style="font-size:0.8rem;color:var(--muted);margin-top:0.25rem;">
          Vorhersagen auf Basis historischer Daten – nur möglich mit sauberen Zeitreihen,
          nicht mit Freitext-IDs und gemischten Formaten.</p>
      </div>
    </div>
    <p style="font-size:0.8rem;color:#1e40af;margin-top:1rem;">
      ℹ️ In strukturierten Systemen sinkt der Zeitanteil für Datenaufbereitung von ~30% auf unter 10%
      der Arbeitszeit. Die gewonnene Zeit steht für Analyse, Entscheidungen und echte Wertschöpfung zur Verfügung.</p>
  </div>
</div>

<!-- ── CTA ── -->
<div class="cta">
  <h2>🔍 Lass dein Excel jetzt prüfen!</h2>
  <p>Unser kostenloser Checker analysiert deine Datei und sagt dir ehrlich, ob alles OK ist – oder ob es Zeit für den nächsten Schritt ist.</p>
  <a href="/">⚡ Excel-Check starten</a>
</div>

<div class="footer">
  ⚡ Excel-Reifecheck &middot; UCS - Digital Workplace Solutions<br>
  Fragen? Melde dich bei uns!
</div>

</div>

<script>
// ── Dark/Light Mode ──
function toggleTheme() {
  const html = document.documentElement;
  const btn = document.querySelector('.theme-toggle-learn');
  if (html.getAttribute('data-theme') === 'dark') {
    html.removeAttribute('data-theme');
    btn.textContent = '🌙';
    localStorage.setItem('theme', 'light');
  } else {
    html.setAttribute('data-theme', 'dark');
    btn.textContent = '☀️';
    localStorage.setItem('theme', 'dark');
  }
}
(function() {
  if (localStorage.getItem('theme') === 'dark') {
    document.documentElement.setAttribute('data-theme', 'dark');
    const btn = document.querySelector('.theme-toggle-learn');
    if (btn) btn.textContent = '☀️';
  }
})();

// ── Scenario Toggle ──
function showScenario(btn, id) {
  btn.closest('.section').querySelectorAll('.scenario-btn').forEach(b => b.classList.remove('active'));
  btn.closest('.section').querySelectorAll('.scenario-content').forEach(c => c.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById(id).classList.add('active');
}

// ── Maturity Model ──
function toggleMaturity(el) {
  const detail = el.querySelector('.maturity-detail');
  const arrow = el.querySelector('span:last-child');
  const isOpen = detail.style.display !== 'none';
  // Close all
  document.querySelectorAll('.maturity-detail').forEach(d => d.style.display = 'none');
  document.querySelectorAll('.maturity-level span:last-child').forEach(a => a.textContent = '▸');
  // Toggle clicked
  if (!isOpen) {
    detail.style.display = 'block';
    arrow.textContent = '▾';
  }
}

// ── Quiz ──
function quizAnswer(btn, correct) {
  const quiz = btn.closest('.quiz');
  const result = quiz.querySelector('.quiz-result');
  quiz.querySelectorAll('.quiz-option').forEach(o => {
    o.style.pointerEvents = 'none';
    o.style.opacity = '0.6';
  });
  btn.style.opacity = '1';
  if (correct) {
    btn.classList.add('correct');
    result.className = 'quiz-result visible correct';
    result.textContent = '🎉 Richtig! Gut erkannt – genau so denkt man professionell über Daten nach.';
  } else {
    btn.classList.add('wrong');
    result.className = 'quiz-result visible wrong';
    result.textContent = '😅 Nicht ganz – aber gut, dass du dich damit beschäftigst! Die grüne Antwort wäre richtig gewesen.';
    // Highlight correct answer
    quiz.querySelectorAll('.quiz-option').forEach(o => {
      o.addEventListener('click', () => {}, true); // no-op
    });
  }
}

// ── Checklist ──
function toggleCheck(li) {
  li.classList.toggle('checked');
  const check = li.querySelector('.check');
  check.innerHTML = li.classList.contains('checked') ? '✓' : '';
  updateChecklist();
}

function updateChecklist() {
  const total = document.querySelectorAll('#checklist li').length;
  const checked = document.querySelectorAll('#checklist li.checked').length;
  const result = document.getElementById('checklist-result');

  if (checked === 0) { result.style.display = 'none'; return; }
  result.style.display = 'block';

  if (checked === total) {
    result.style.background = '#f0fdf4';
    result.style.color = '#15803d';
    result.textContent = '🎉 Alle Punkte erfüllt! Dein Excel ist vermutlich genau richtig so. Weiter so!';
  } else if (checked >= total - 2) {
    result.style.background = '#fefce8';
    result.style.color = '#854d0e';
    result.textContent = '⚡ Fast alle Punkte erfüllt – ein paar Dinge könnten optimiert werden. Schau dir unseren Check an!';
  } else {
    result.style.background = '#fef2f2';
    result.style.color = '#b91c1c';
    result.textContent = '🚨 Mehrere Warnsignale! Deine Datei könnte von einer professionelleren Lösung profitieren.';
  }
}

// ── Effizienz-Rechner ──
function updateEffCalc(pct) {
  pct = parseInt(pct);
  var f = pct / 100;
  var s1 = Math.round(468000 * f);
  var s2 = Math.round(349000 * f);
  var s3 = Math.round(936000 * f);
  var total = s1 + s2 + s3;
  var fmt = function(n) { return '€' + n.toLocaleString('de-DE'); };
  document.getElementById('effPct').textContent = pct + '%';
  document.getElementById('eff-s1').textContent = fmt(s1);
  document.getElementById('eff-s2').textContent = fmt(s2);
  document.getElementById('eff-s3').textContent = fmt(s3);
  var totalFmt = total >= 1000000
    ? '~€' + (total / 1000000).toFixed(2).replace('.', ',') + ' Mio/Jahr'
    : '~' + fmt(total) + '/Jahr';
  document.getElementById('eff-total').textContent = totalFmt;
  document.getElementById('eff-note-pct').textContent = pct;
  document.getElementById('eff-note-sum').textContent = total >= 1000000
    ? '~€' + (total / 1000000).toFixed(2).replace('.', ',') + ' Mio'
    : '~' + fmt(total);
  var rest = 100 - pct;
  var investMin = 100000;
  var months = total > 0 ? Math.max(1, Math.ceil(investMin / (total / 12))) : 99;
  document.getElementById('eff-amort').textContent =
    (pct === 100 ? 'Bei voller Realisierung' : 'Selbst bei ' + pct + '% Realisierung')
    + ': Amortisation in ' + (months <= 1 ? 'unter 1 Monat' : months <= 2 ? 'unter 2 Monaten'
    : 'unter ' + months + ' Monaten') + '.';
}

// ── Growth Bar Animation (on scroll) ──
const observer = new IntersectionObserver((entries) => {
  entries.forEach(entry => {
    if (entry.isIntersecting) {
      const bars = entry.target.querySelectorAll('.growth-bar');
      bars.forEach((bar, i) => {
        const heights = [25, 50, 85, 120, 155, 175];
        setTimeout(() => {
          bar.style.height = heights[i] + 'px';
        }, i * 150);
      });
      observer.unobserve(entry.target);
    }
  });
}, { threshold: 0.3 });
const growthSection = document.getElementById('growthBars');
if (growthSection) observer.observe(growthSection);
</script>
</body>
</html>"""
