"""HTML-Bericht-Generator für die Excel-Qualitätsprüfung – v2 mit Diagrammen, Anti-Mustern & LLM."""

from __future__ import annotations

import html
import math
import os
from collections import defaultdict
from datetime import datetime
from typing import List

import hashlib as _hashlib
import json as _json
import sys as _sys

from excel_checker.models import (
    Category, ColumnProfile, Finding, Recommendation, RecommendationType,
    Severity, WorkbookReport,
)


SEVERITY_COLORS = {
    Severity.CRITICAL: ("#dc2626", "#fef2f2", "🔴"),
    Severity.ERROR: ("#ea580c", "#fff7ed", "🟠"),
    Severity.WARNING: ("#ca8a04", "#fefce8", "🟡"),
    Severity.INFO: ("#2563eb", "#eff6ff", "🔵"),
}

RECTYPE_ICONS = {
    RecommendationType.DATABASE_MIGRATION: "🗄️",
    RecommendationType.POWER_BI: "📊",
    RecommendationType.DATA_WAREHOUSE: "🏗️",
    RecommendationType.SHAREPOINT_LIST: "📋",
    RecommendationType.NORMALIZATION: "🧹",
    RecommendationType.SPLIT_WORKBOOK: "✂️",
    RecommendationType.CLEANUP: "🔧",
    RecommendationType.OK: "✅",
}

# Schweregrad-Stufen für Anti-Muster
GRADE_MEGA = "Mega-Problem"
GRADE_ECHT = "Echtes Problem"
GRADE_SCHOEN = "Schönheitsfehler"

ANTI_PATTERN_GRADES = {
    GRADE_MEGA:   ("🔥", "#dc2626", "#fef2f2"),   # rot
    GRADE_ECHT:   ("😬", "#ea580c", "#fff7ed"),   # orange
    GRADE_SCHOEN: ("💅", "#6366f1", "#eef2ff"),   # indigo
}

GRADE_ORDER = {GRADE_MEGA: 0, GRADE_ECHT: 1, GRADE_SCHOEN: 2}

# Anti-Muster Definitionen: (Name, Icon, Beschreibung, Schweregrad)
ANTI_PATTERNS = {
    # ── Mega-Probleme 🔥 ── objektiv kritisch, nicht diskutierbar ──
    "STR-005": ("ID-Wildwuchs", "🏷️",
                "Inkonsistente Identifier (Formate, Lücken, Duplikate).",
                GRADE_MEGA),
    "STR-006": ("Phantom-Schlüssel", "👻",
                "Datentabelle ohne eindeutige, sortierbare Schlüsselspalte.",
                GRADE_MEGA),
    "STR-007": ("Freitext-IDs", "🔤",
                "IDs aus Freitext statt sortierbarer Codes – nicht maschinell verarbeitbar.",
                GRADE_MEGA),
    "FRM-004": ("Zirkelbezug", "🔄",
                "Formeln referenzieren sich selbst – unberechenbare Ergebnisse.",
                GRADE_MEGA),
    "VOL-001": ("Excel-Sprenger", "💥",
                "Datenvolumen hat die sinnvolle Excel-Kapazität überschritten.",
                GRADE_MEGA),
    "VOL-002": ("Formel-Overload", "⚡",
                "Zu viele Formeln – Instabilität und Langsamkeit.",
                GRADE_MEGA),
    # ── Echte Probleme 😬 ── reale Wartungs-/Zuverlässigkeitsprobleme ──
    "STR-001": ("Verbundene-Zellen-Chaos", "🔗",
                "Verbundene Zellen zerstören maschinelle Lesbarkeit.",
                GRADE_ECHT),
    "STR-002": ("Typ-Salat", "🥗",
                "Gemischte Datentypen in einer Spalte – 1. Normalform verletzt.",
                GRADE_ECHT),
    "STR-003": ("Kopflos-Tabelle", "👻",
                "Fehlende oder unklare Kopfzeilen.",
                GRADE_ECHT),
    "FRM-002": ("Volatile-Falle", "🌋",
                "Volatile Funktionen erzwingen permanente Neuberechnung.",
                GRADE_ECHT),
    "FRM-003": ("SVERWEIS-Datenbank", "🔍",
                "Lookup-Ketten simulieren SQL-JOINs – viel langsamer.",
                GRADE_ECHT),
    "FRM-005": ("Datei-Spinnennetz", "🕸️",
                "Externe Verknüpfungen – fragile Abhängigkeiten.",
                GRADE_ECHT),
    "IMP-001": ("Geheim-Ampel", "🚦",
                "Farbcodes tragen undokumentierte Bedeutung.",
                GRADE_ECHT),
    "IMP-002": ("Versteckspiel", "🙈",
                "Versteckte Blätter mit undokumentiertem Wissen.",
                GRADE_ECHT),
    "IMP-004": ("Format-Logik", "🎭",
                "Geschäftslogik in bedingten Formatierungen versteckt.",
                GRADE_ECHT),
    "IMP-005": ("Zauberzahlen", "🎩",
                "Hartcodierte Zahlen in Formeln ohne Erklärung.",
                GRADE_ECHT),
    "VOL-003": ("Blatt-Hydra", "🐉",
                "Zu viele Blätter bilden eine relationale DB ab – ohne deren Vorteile.",
                GRADE_ECHT),
    "VOL-004": ("Schwergewicht", "🏋️",
                "Überdimensionierte Dateigröße.",
                GRADE_ECHT),
    # ── Schönheitsfehler 💅 ── Best Practices, kein Strick draus drehbar ──
    "STR-004": ("Leerzeilen-Layout", "🕳️",
                "Leere Zeilen als visuelle Trenner fragmentieren den Datenbereich.",
                GRADE_SCHOEN),
    "FRM-001": ("Dollar-Fixierung", "💲",
                "Zu viele starre $-Bezüge – schwer wartbar.",
                GRADE_SCHOEN),
    "IMP-003": ("Unsichtbare Daten", "👁️‍🗨️",
                "Ausgeblendete Zeilen/Spalten.",
                GRADE_SCHOEN),
    "IMP-006": ("Eingetippte Dropdowns", "📝",
                "Validierungslisten mit hardcodierten Werten.",
                GRADE_SCHOEN),
    "IMP-007": ("Namenlose Blätter", "❓",
                "Generische Blatt-Namen ohne Aussagekraft.",
                GRADE_SCHOEN),
    "IMP-008": ("Kommentar-Wissen", "💬",
                "Geschäftsregeln leben nur in Zell-Kommentaren.",
                GRADE_SCHOEN),
    "IMP-009": ("Zahlen-Maskerade", "🎭",
                "Irreführende Zahlenformate.",
                GRADE_SCHOEN),
    "IMP-010": ("Blattschutz-Rätsel", "🔒",
                "Geschützte Bereiche ohne Dokumentation.",
                GRADE_SCHOEN),
}


def _get_version_info() -> str:
    from excel_checker import __version__, BUILD_TIMESTAMP
    return f"Version {__version__} (Build {BUILD_TIMESTAMP})"


def _score_color(score: int) -> str:
    if score >= 80:
        return "#16a34a"
    elif score >= 60:
        return "#ca8a04"
    elif score >= 40:
        return "#ea580c"
    return "#dc2626"


def _compute_ai_readiness(report: 'WorkbookReport') -> tuple[int, str, list[tuple[str, str, str]]]:
    """Berechnet KI-Readiness Score, Farbe und Blocker-Liste.

    Returns: (score, color, blockers) wobei blockers = [(icon, title, desc), ...]
    """
    total_merged = sum(s.merged_regions for s in report.sheet_stats)

    ai_blockers = []
    if total_merged > 0:
        ai_blockers.append(("🔗", f"{total_merged} verbundene Zellbereiche",
                            "KI kann tabellarische Struktur nicht erkennen"))
    imp_count = sum(1 for f in report.findings if f.rule_id.startswith("IMP-"))
    if imp_count > 0:
        ai_blockers.append(("🧠", f"{imp_count}× implizites Wissen",
                            "Farbcodes, versteckte Blätter, Zauberzahlen – für KI unsichtbar"))
    str_count = sum(1 for f in report.findings
                    if f.rule_id in ("STR-002", "STR-005", "STR-006", "STR-007"))
    if str_count > 0:
        ai_blockers.append(("🏷️", f"{str_count}× Strukturproblem",
                            "Gemischte Typen, fehlende IDs – KI kann Datensätze nicht zuordnen"))
    frm_count = sum(1 for f in report.findings
                    if f.rule_id in ("FRM-003", "FRM-004", "FRM-005"))
    if frm_count > 0:
        ai_blockers.append(("🔄", f"{frm_count}× Formel-Verkettung",
                            "SVERWEIS-Ketten & Zirkelbezüge blockieren automatische Auswertung"))
    if not ai_blockers and report.health_score < 80:
        ai_blockers.append(("⚠️", "Strukturelle Schwächen",
                            "Daten sind nicht in maschinenlesbarer Form"))

    ai_ready = max(0, min(100, report.health_score - len(ai_blockers) * 10))
    ai_ready_color = "#16a34a" if ai_ready >= 70 else "#ca8a04" if ai_ready >= 40 else "#dc2626"
    return ai_ready, ai_ready_color, ai_blockers


def _score_label(score: int) -> str:
    if score >= 80:
        return "Gut"
    elif score >= 60:
        return "Verbesserungsbedarf"
    elif score >= 40:
        return "Handlungsbedarf"
    return "Dringender Handlungsbedarf"


def _compute_compliance_score(report: 'WorkbookReport') -> tuple[int, str, list[tuple[str, str, str]]]:
    """Berechnet Compliance-Score (0-100, höher = weniger Risiko).

    Gewichtung: ein einziger PII-Treffer bedeutet "Risiko vorhanden" und
    zieht 25 Basis-Punkte ab; jeder weitere Treffer zählt nur noch
    degressiv (max. 30 Punkte pro Kategorie). So signalisiert ein Score
    von 50 eine relevante Risiko-Lage, nicht "maximal kaputt".
    """
    pii_count = sum(1 for f in report.findings if f.rule_id.startswith("PII-"))
    critical_count = report.critical_count
    imp_count = sum(1 for f in report.findings if f.rule_id.startswith("IMP-"))

    risks: list[tuple[str, str, str]] = []
    if pii_count > 0:
        risks.append(("🔐", f"{pii_count}× Verdacht auf sensible Daten",
                      "E-Mail, IBAN oder Sozialversicherungsnummer in Zellen erkannt"))
    if critical_count > 0:
        risks.append(("⚠️", f"{critical_count}× kritisches Problem",
                      "Datenintegrität gefährdet — Governance-relevant"))
    if imp_count > 0:
        risks.append(("🧠", f"{imp_count}× implizites Wissen",
                      "Farbcodes, versteckte Blätter, undokumentierte Logik"))

    def _tier_penalty(count: int, base: int, per_extra: int, cap: int) -> int:
        if count <= 0:
            return 0
        return min(cap, base + per_extra * (count - 1))

    penalty = (
        _tier_penalty(pii_count, base=20, per_extra=2, cap=35)
        + _tier_penalty(critical_count, base=15, per_extra=5, cap=30)
        + _tier_penalty(imp_count, base=5, per_extra=1, cap=15)
    )
    score = max(0, 100 - penalty)
    color = "#16a34a" if score >= 70 else "#ca8a04" if score >= 40 else "#dc2626"
    return score, color, risks


def _humanize(rule_id: str, fallback: str) -> str:
    """Ersetzt technischen Jargon durch Mitarbeiter-verständliche Titel.

    Fallback = original message. Rückgabe bleibt HTML-sicher — Caller muss
    bei Bedarf _esc() anwenden.
    """
    HUMANIZE = {
        "STR-001": "Verbundene Zellen stören maschinelle Lesbarkeit",
        "STR-002": "Gemischte Datentypen in einer Spalte",
        "STR-003": "Kopfzeile fehlt oder ist unklar",
        "STR-004": "Leerzeilen unterbrechen den Datenbereich",
        "STR-005": "Uneinheitliche Identifier",
        "STR-006": "Eindeutige Schlüsselspalte fehlt",
        "STR-007": "IDs aus Freitext statt sortierbarer Codes",
        "FRM-001": "Zu viele fixe $-Bezüge",
        "FRM-002": "Volatile Funktionen verlangsamen die Datei",
        "FRM-003": "SVERWEIS-Ketten ersetzen eine Datenbank",
        "FRM-004": "Zirkelbezug — unberechenbares Ergebnis",
        "FRM-005": "Externe Datei-Verknüpfung",
        "FRM-006": "Formel liefert Fehlerwert",
        "VOL-001": "Datenvolumen überschreitet Excel-Grenzen",
        "VOL-002": "Zu viele Formeln — Stabilitätsrisiko",
        "VOL-003": "Zu viele Tabellenblätter",
        "VOL-004": "Dateigröße sehr hoch",
        "IMP-001": "Farbcodes tragen undokumentierte Bedeutung",
        "IMP-002": "Versteckte Tabellenblätter",
        "IMP-003": "Ausgeblendete Zeilen oder Spalten",
        "IMP-004": "Geschäftslogik steckt in bedingter Formatierung",
        "IMP-005": "Hartcodierte Zahlen in Formeln",
        "IMP-006": "Dropdown mit hardcodierten Werten",
        "IMP-007": "Generischer Blatt-Name",
        "IMP-008": "Geschäftsregel nur im Kommentar",
        "IMP-009": "Irreführendes Zahlenformat",
        "IMP-010": "Geschützter Bereich ohne Hinweis",
    }
    return HUMANIZE.get(rule_id, fallback)


def _html_score_gauge(label: str, value: int, color: str,
                      hint: str = "", tooltip: str = "") -> str:
    """Rendert einen animierten Score-Ring.

    Nutzt --gauge-value (CSS @property aus theme.css) für Stroke-Animation.
    Breite/Höhe des SVG: 110×110, Radius 40, Umfang 2π·40 ≈ 251.2.
    """
    safe_label = _esc(label)
    safe_hint = _esc(hint)
    safe_tooltip = _esc(tooltip or f"{label}: {value} von 100")
    return f"""
<div class="score-gauge-card" title="{safe_tooltip}">
  <div class="score-gauge" style="--gauge-value:{value}; width:110px; height:110px; margin:0 auto; position:relative;">
    <svg width="110" height="110" style="transform:rotate(-90deg);">
      <circle cx="55" cy="55" r="40" fill="none" stroke="var(--border)" stroke-width="8"
              class="report-svg-ring-bg"/>
      <circle cx="55" cy="55" r="40" fill="none" stroke="{color}" stroke-width="8"
              stroke-linecap="round" class="gauge-ring-fg"/>
    </svg>
    <div style="position:absolute; inset:0; display:flex; align-items:center; justify-content:center;">
      <div class="score-gauge-value" style="color:{color};">{value}</div>
    </div>
  </div>
  <div class="score-gauge-label">{safe_label}</div>
  <div class="score-gauge-hint" style="font-size:0.8rem;color:var(--muted);margin-top:0.25rem;">{safe_hint}</div>
</div>
"""


def _esc(text: str | None) -> str:
    if text is None:
        return ""
    return html.escape(str(text))


def _compute_complexity_dimensions(report: WorkbookReport) -> dict[str, float]:
    """Berechne 6 Komplexitäts-Dimensionen (0-100) für das Radar-Chart."""
    total_rows = sum(s.row_count for s in report.sheet_stats)
    total_formulas = sum(s.formula_count for s in report.sheet_stats)
    total_merged = sum(s.merged_regions for s in report.sheet_stats)

    def _count_prefix(prefix: str) -> int:
        return sum(1 for f in report.findings if f.rule_id.startswith(prefix))

    dims = {}
    # Datenvolumen
    dims["Daten-\nvolumen"] = min(100, total_rows / 500 * 100)
    # Formel-Komplexität
    dims["Formel-\nKomplexität"] = min(100, total_formulas / 200 * 100)
    # Vernetzung (cross-sheet refs, external refs)
    net_count = _count_prefix("FRM-005") + _count_prefix("FRM-003")
    dims["Vernetzung"] = min(100, net_count * 33)
    # Implizites Wissen
    imp_count = _count_prefix("IMP-")
    dims["Implizites\nWissen"] = min(100, imp_count * 15)
    # Strukturprobleme
    str_count = _count_prefix("STR-")
    dims["Strukturelle\nProbleme"] = min(100, str_count * 20 + total_merged * 5)
    # Dateigröße
    dims["Datei-\ngröße"] = min(100, report.file_size_mb / 10 * 100)
    return dims


def _radar_chart_svg(dimensions: dict[str, float], size: int = 400) -> str:
    """Erzeuge ein SVG-Radar-Chart (Spider-Chart) für die Komplexitätsdimensionen."""
    padding = 100  # Platz für Labels außen
    inner = size - 2 * padding
    cx, cy = size // 2, size // 2
    radius = inner // 2
    labels = list(dimensions.keys())
    values = list(dimensions.values())
    n = len(labels)
    if n == 0:
        return ""
    angle_step = 2 * math.pi / n

    # Hintergrund-Ringe (20,40,60,80,100%) – farbcodiert: innen grün → außen rot
    ring_colors = [
        (0.2, "#dcfce7", "#bbf7d0"),   # grün (unkritisch)
        (0.4, "#fef9c3", "#fde68a"),   # gelb
        (0.6, "#ffedd5", "#fed7aa"),   # orange-hell
        (0.8, "#fee2e2", "#fecaca"),   # orange-rot
        (1.0, "#fecaca", "#fca5a5"),   # rot (kritisch)
    ]
    grid_lines = []
    # Ringe von außen nach innen zeichnen (damit innere über äußeren liegen)
    for pct, fill, stroke in reversed(ring_colors):
        r = radius * pct
        points = []
        for i in range(n):
            a = -math.pi / 2 + i * angle_step
            points.append(f"{cx + r * math.cos(a):.1f},{cy + r * math.sin(a):.1f}")
        grid_lines.append(
            f'<polygon points="{" ".join(points)}" fill="{fill}" '
            f'stroke="{stroke}" stroke-width="1" opacity="0.6"/>'
        )

    # Achsenlinien
    axis_lines = []
    for i in range(n):
        a = -math.pi / 2 + i * angle_step
        x2 = cx + radius * math.cos(a)
        y2 = cy + radius * math.sin(a)
        # class für Dark-Mode-Override (siehe theme.css .report-svg-axis-stroke)
        axis_lines.append(
            f'<line x1="{cx}" y1="{cy}" x2="{x2:.1f}" y2="{y2:.1f}" '
            f'stroke="#cbd5e1" stroke-width="1" class="report-svg-axis-stroke"/>'
        )

    # Datenfläche
    data_points = []
    for i in range(n):
        a = -math.pi / 2 + i * angle_step
        r = radius * (values[i] / 100)
        data_points.append(f"{cx + r * math.cos(a):.1f},{cy + r * math.sin(a):.1f}")
    data_polygon = (
        f'<polygon points="{" ".join(data_points)}" '
        f'fill="rgba(59,130,246,0.2)" stroke="#3b82f6" stroke-width="2"/>'
    )

    # Datenpunkte
    dots = []
    for pt in data_points:
        x, y = pt.split(",")
        dots.append(
            f'<circle cx="{x}" cy="{y}" r="4" fill="#3b82f6"/>'
        )

    # Labels
    label_elems = []
    for i in range(n):
        a = -math.pi / 2 + i * angle_step
        lr = radius + 30
        lx = cx + lr * math.cos(a)
        ly = cy + lr * math.sin(a)
        anchor = "middle"
        if abs(math.cos(a)) > 0.3:
            anchor = "start" if math.cos(a) > 0 else "end"
        # Multi-line labels via tspan
        lines = labels[i].split('\n')
        if len(lines) > 1:
            tspans = ''.join(
                f'<tspan x="{lx:.1f}" dy="{"0" if j == 0 else "1.2em"}">'
                f'{html.escape(line)}</tspan>'
                for j, line in enumerate(lines)
            )
            # Shift up for multi-line labels so they center vertically
            label_elems.append(
                f'<text x="{lx:.1f}" y="{ly - 8:.1f}" text-anchor="{anchor}" '
                f'dominant-baseline="central" font-size="12" fill="#475569" '
                f'font-weight="500" class="report-svg-label-fill">'
                f'{tspans}</text>'
            )
        else:
            label_elems.append(
                f'<text x="{lx:.1f}" y="{ly:.1f}" text-anchor="{anchor}" '
                f'dominant-baseline="central" font-size="12" fill="#475569" '
                f'font-weight="500" class="report-svg-label-fill">'
                f'{html.escape(labels[i])}</text>'
            )

    svg = (
        f'<svg viewBox="0 0 {size} {size}" '
        f'xmlns="http://www.w3.org/2000/svg" '
        f'style="max-width:100%;width:100%;height:auto;overflow:visible;">\n'
        + "\n".join(grid_lines) + "\n"
        + "\n".join(axis_lines) + "\n"
        + data_polygon + "\n"
        + "\n".join(dots) + "\n"
        + "\n".join(label_elems) + "\n"
        + "</svg>"
    )
    return svg


def _bar_chart_svg(data: dict[str, float], label: str, max_val: float | None = None,
                   width: int = 450, bar_height: int = 28) -> str:
    """Erzeuge ein horizontales Balkendiagramm als SVG."""
    if not data:
        return ""
    if max_val is None:
        max_val = max(data.values()) or 1
    bar_area = width - 140
    h = len(data) * (bar_height + 8) + 30
    parts = [
        f'<svg width="{width}" height="{h}" viewBox="0 0 {width} {h}" '
        f'xmlns="http://www.w3.org/2000/svg" style="max-width:100%;height:auto;">'
    ]
    y = 10
    for name, val in data.items():
        bw = max(2, bar_area * (val / max_val)) if max_val > 0 else 2
        parts.append(
            f'<text x="130" y="{y + bar_height // 2 + 4}" text-anchor="end" '
            f'font-size="12" fill="#64748b" class="report-svg-bar-label-fill">'
            f'{html.escape(str(name))}</text>'
        )
        parts.append(
            f'<rect x="140" y="{y}" width="{bw:.0f}" height="{bar_height}" '
            f'rx="4" fill="#3b82f6" opacity="0.8"/>'
        )
        parts.append(
            f'<text x="{140 + bw + 6:.0f}" y="{y + bar_height // 2 + 4}" '
            f'font-size="11" fill="#334155" class="report-svg-bar-value-fill">'
            f'{val:,.0f}</text>'
        )
        y += bar_height + 8
    parts.append("</svg>")
    return "\n".join(parts)


def generate_html(report: WorkbookReport) -> str:
    """Generiert einen vollständigen HTML-Bericht mit Diagrammen, Registerkarten und Anti-Mustern."""
    score = report.health_score
    score_color = _score_color(score)
    score_label = _score_label(score)
    filename = os.path.basename(report.file_path)
    now = datetime.now().strftime("%d.%m.%Y %H:%M")

    # Komplexitätsdimensionen
    dims = _compute_complexity_dimensions(report)

    # KI-Readiness berechnen (brauchen wir für Hero + Effizienz)
    ai_ready, ai_ready_color, ai_blockers = _compute_ai_readiness(report)

    # Compliance-Score (Phase 2 C.2) — Governance-Ring
    compliance_score, compliance_color, compliance_risks = _compute_compliance_score(report)

    # DB-Readiness Durchschnitt (wandert in Migration-Tab, wird aber für Hero-Kurzprofil gebraucht)
    total_rows = sum(s.row_count for s in report.sheet_stats)
    if total_rows > 0:
        avg_db = int(sum(s.db_readiness * max(s.row_count, 1) for s in report.sheet_stats) / max(total_rows, 1))
    else:
        avg_db = int(sum(s.db_readiness for s in report.sheet_stats) / max(len(report.sheet_stats), 1)) if report.sheet_stats else 0

    # Anti-Patterns erkennen
    detected_ap: dict[str, tuple[str, str, str, int, str]] = {}
    for f in report.findings:
        if f.rule_id in ANTI_PATTERNS and f.rule_id not in detected_ap:
            name, icon, desc, grade = ANTI_PATTERNS[f.rule_id]
            count = sum(1 for ff in report.findings if ff.rule_id == f.rule_id)
            detected_ap[f.rule_id] = (name, icon, desc, count, grade)

    # HTML zusammenbauen
    parts = [_html_head(filename, now)]
    parts.append(_html_hero(score, score_color, score_label, report, dims,
                            ai_ready, ai_ready_color, avg_db,
                            compliance_score, compliance_color))
    parts.append(_html_score_breakdown(report))
    parts.append(_html_stats_section(report))

    if detected_ap:
        parts.append(_html_anti_patterns(detected_ap, report.findings))

    parts.append(_html_complexity_section(report, dims))

    parts.append(_html_per_sheet_section(report))

    if report.recommendations:
        parts.append(_html_recommendations_tabs(report.recommendations))

    if report.findings:
        parts.append(_html_audience_tabs(report.findings, compliance_score,
                                         compliance_risks, report, avg_db))

    # Report-Kontext als verstecktes JSON einbetten (für LLM-Analyse ohne Server-State)
    try:
        from excel_checker.llm_analysis import extract_context
        ctx = extract_context(report)
        ctx_json = _json.dumps(ctx, ensure_ascii=False)
        parts.append(f'<script type="application/json" id="report-context">{ctx_json}</script>')
    except Exception as exc:  # noqa: BLE001 - Report darf daran nicht scheitern
        # Nicht stillschweigend übergehen: Ein verschluckter ImportError hat
        # diesen Block im Browser-Build unbemerkt verschwinden lassen.
        print(f"[WARN] Report-Kontext konnte nicht eingebettet werden: {exc}",
              file=_sys.stderr)

    parts.append(_html_footer())
    return "\n".join(parts)


def _finding_anchor(f) -> str:
    """Stabile DOM-ID für ein Finding.

    Hier stand ``id(f)`` — die Speicheradresse des Objekts. Die ist bei
    jedem Lauf eine andere, damit auch die Anker im Report: zwei Analysen
    derselben Datei ergaben unterschiedliche HTML-Dateien, und ein
    kopierter Deep-Link zeigte im nächsten Report ins Leere. Jetzt wird die
    ID aus dem Inhalt abgeleitet und ist damit reproduzierbar.
    """
    raw = "|".join(
        str(x) for x in (f.rule_id, f.sheet, f.cell, f.message, f.detail, f.suggestion)
    )
    digest = _hashlib.sha1(raw.encode("utf-8")).hexdigest()[:10]
    return f"finding-{_esc(f.rule_id)}-{digest}"


def _html_head(filename: str, now: str) -> str:
    return f"""<!DOCTYPE html>
<html lang='de'>
<head>
<meta charset='UTF-8'>
<meta name='viewport' content='width=device-width, initial-scale=1.0'>
<title>Excel-Reifecheck – {_esc(filename)}</title>
<link rel='stylesheet' href='/static/theme.css'>
<script>
  // Dark-Mode Init synchron (FOUC vermeiden). Nutzt das gemeinsame
  // 'theme'-Schlüssel-Pattern wie Upload- und Learn-Seite.
  (function() {{
    if (localStorage.getItem('theme') === 'dark') {{
      document.documentElement.setAttribute('data-theme', 'dark');
    }}
  }})();
</script>
<style>
/* Vars + Reset + body-Defaults stammen aus theme.css.
   Hier nur seitenspezifisches Padding (Original-Wert). */
body {{ padding: 1rem; }}
.container {{ max-width: 1200px; margin: 0 auto; }}
.header {{ text-align: center; margin-bottom: 1.5rem; padding: 1.5rem; background: var(--card); border-radius: 12px; border: 1px solid var(--border); }}
.header h1 {{ font-size: 1.5rem; margin-bottom: 0.25rem; }}
.header .subtitle {{ color: var(--muted); font-size: 0.9rem; }}
.card {{ background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 1.5rem; margin-bottom: 1.5rem; }}
.card h2 {{ font-size: 1.15rem; margin-bottom: 1rem; padding-bottom: 0.5rem; border-bottom: 2px solid var(--border); }}
details.card {{ cursor: default; }}
details.card > summary {{ cursor: pointer; list-style: none; display: flex; align-items: center; justify-content: space-between; user-select: none; }}
details.card > summary::-webkit-details-marker {{ display: none; }}
details.card > summary::after {{ content: '▸'; font-size: 1.2rem; color: var(--muted); transition: transform 0.2s; flex-shrink: 0; margin-left: 0.75rem; }}
details.card[open] > summary::after {{ transform: rotate(90deg); }}
details.card > summary h2 {{ margin-bottom: 0; padding-bottom: 0; border-bottom: none; }}
details.card[open] > summary {{ margin-bottom: 1rem; padding-bottom: 0.5rem; border-bottom: 2px solid var(--border); }}
details.card[open] > summary h2 {{ margin-bottom: 0; padding-bottom: 0; border-bottom: none; }}
.hero-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 1.5rem; align-items: center; }}
.score-ring {{ width: 110px; height: 110px; margin: 0.5rem auto; position: relative; display: flex; align-items: center; justify-content: center; }}
.score-ring svg {{ position: absolute; transform: rotate(-90deg); }}
.score-ring .score-ring-text {{ position: relative; z-index: 1; }}
.score-ring .value {{ font-size: 1.8rem; font-weight: 700; }}
.score-ring .label {{ font-size: 0.7rem; color: var(--muted); }}
.score-text {{ text-align: center; margin-top: 0.25rem; }}
.score-text strong {{ font-size: 0.85rem; }}
.score-text span {{ font-size: 0.75rem !important; }}
.hero-tabs {{ display: flex; border-bottom: 2px solid var(--border); margin-bottom: 1rem; }}
.hero-tab-btn {{ flex: 1; padding: 0.5rem 1rem; text-align: center; font-size: 0.85rem; font-weight: 600; color: var(--muted); background: none; border: none; border-bottom: 2px solid transparent; margin-bottom: -2px; cursor: pointer; transition: all 0.2s; }}
.hero-tab-btn:hover {{ color: var(--text); background: var(--bg); }}
.hero-tab-btn.active {{ color: var(--accent); border-bottom-color: var(--accent); }}
.hero-tab-panel {{ display: none; }}
.hero-tab-panel.active {{ display: block; }}
.stats-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 1rem; }}
.stat-item {{ text-align: center; padding: 1rem; background: var(--bg); border-radius: 8px; }}
.stat-item .num {{ font-size: 1.8rem; font-weight: 700; color: var(--accent); }}
.stat-item .lbl {{ font-size: 0.8rem; color: var(--muted); }}
.ap-grade-section {{ margin-bottom: 1.5rem; }}
.ap-grade-header {{ display: flex; align-items: center; gap: 0.5rem; font-size: 1.1rem; font-weight: 700; margin-bottom: 0.75rem; padding-bottom: 0.5rem; border-bottom: 2px solid var(--border); }}
.ap-grade-header .grade-icon {{ font-size: 1.3rem; }}
.ap-grade-header .grade-count {{ font-size: 0.8rem; font-weight: 400; color: var(--muted); }}
.ap-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 1rem; }}
.ap-flip-container {{ perspective: 800px; cursor: pointer; min-height: 260px; min-width: 0; height: 260px; max-width: 340px; width: 100%; display: flex; align-items: stretch; }}
.ap-flip-inner {{ position: relative; transition: transform 0.5s; transform-style: preserve-3d; width: 100%; height: 100%; min-width: 0; }}
.ap-flip-container.flipped .ap-flip-inner {{ transform: rotateY(180deg); }}
.ap-card, .ap-card-back {{ backface-visibility: hidden; -webkit-backface-visibility: hidden; border-radius: 10px; padding: 1rem; text-align: center; min-height: 100%; min-width: 0; width: 100%; height: 100%; box-sizing: border-box; display: flex; flex-direction: column; justify-content: center; align-items: center; }}
.ap-card {{ transition: transform 0.15s; overflow: hidden; }}
.ap-flip-container:not(.flipped):hover .ap-card {{ transform: translateY(-2px); }}
.ap-card .ap-icon {{ font-size: 2rem; margin-bottom: 0.25rem; }}
.ap-card .ap-name {{ font-weight: 600; font-size: 0.95rem; margin-bottom: 0.25rem; }}
.ap-card .ap-desc {{ font-size: 0.8rem; color: var(--muted); }}
.ap-card .ap-count {{ display: inline-block; margin-top: 0.5rem; padding: 0.1rem 0.5rem; color: white; border-radius: 10px; font-size: 0.75rem; }}
.ap-card .ap-badge {{ display: inline-block; margin-top: 0.4rem; padding: 0.15rem 0.6rem; border-radius: 10px; font-size: 0.7rem; font-weight: 600; }}
.ap-card .ap-flip-hint {{ display: block; margin-top: 0.5rem; font-size: 0.7rem; color: var(--muted); opacity: 0.6; }}
.ap-card-back {{ position: absolute; top: 0; left: 0; width: 100%; height: 100%; transform: rotateY(180deg); display: flex; flex-direction: column; justify-content: flex-start; align-items: center; text-align: center; overflow: hidden; padding: 1rem; box-sizing: border-box; }}
.ap-card-back .ap-back-title {{ font-weight: 600; font-size: 0.85rem; margin-bottom: 0.5rem; color: var(--text); }}
.ap-card-back .ap-location {{ font-size: 0.82rem; color: var(--text); margin-bottom: 0.25rem; word-break: break-all; }}
.ap-card-back .ap-location .loc-icon {{ margin-right: 0.25rem; }}
.ap-card-back .ap-show-all-btn {{ display: inline-block; margin-top: 0.5rem; padding: 0.3rem 0.75rem; font-size: 0.75rem; font-weight: 600; border-radius: 6px; border: 1px solid var(--border); background: white; color: var(--accent); cursor: pointer; transition: background 0.15s; }}
.ap-card-back .ap-show-all-btn:hover {{ background: var(--bg-info-soft); }}
.ap-card-back .ap-all-locations {{ display: none; margin-top: 0.5rem; max-height: 110px; overflow-y: auto; text-align: left; width: 100%; padding: 0.25rem 0.5rem; font-size: 0.78rem; background: var(--bg); border-radius: 6px; box-sizing: border-box; }}
.ap-card-back .ap-all-locations.visible {{ display: block; }}
.ap-card-back .ap-all-locations li {{ padding: 0.2rem 0; border-bottom: 1px solid var(--border); list-style: none; word-break: break-all; }}
.ap-card-back .ap-all-locations li:last-child {{ border-bottom: none; }}

  /* Complexity Charts */
  .charts-grid {{
    display: grid; grid-template-columns: 1fr 1fr; gap: 1.5rem;
  }}
  .chart-box {{ text-align: center; }}
  .chart-box h3 {{ font-size: 0.95rem; color: var(--muted); margin-bottom: 0.75rem; }}

  /* Tabs */
  .tabs {{
    display: flex; flex-wrap: wrap; gap: 0.25rem;
    border-bottom: 2px solid var(--border); margin-bottom: 1rem;
    overflow-x: auto;
  }}
  .tab-btn {{
    padding: 0.6rem 1.2rem; border: none; background: none;
    cursor: pointer; font-size: 0.9rem; color: var(--muted);
    border-bottom: 3px solid transparent; white-space: nowrap;
    transition: all 0.2s;
  }}
  .tab-btn:hover {{ color: var(--text); }}
  .tab-btn.active {{
    color: var(--accent); border-bottom-color: var(--accent); font-weight: 600;
  }}
  .tab-content {{ display: none; }}
  .tab-content.active {{ display: block; }}

  /* Findings */
  .finding {{
    padding: 1rem 1rem 1rem 1.25rem; margin-bottom: 0.75rem;
    border-radius: 8px; border-left: 4px solid;
  }}
  .finding .title {{ font-weight: 600; margin-bottom: 0.25rem; }}
  .finding .meta {{ font-size: 0.8rem; color: var(--muted); margin-bottom: 0.5rem; }}
  .finding .detail {{ font-size: 0.85rem; color: var(--muted); margin-top: 0.5rem; }}
  .finding .suggestion {{
    font-size: 0.85rem; margin-top: 0.5rem;
    padding: 0.5rem 0.75rem; background: var(--bg-info-soft);
    border-radius: 6px; color: var(--accent);
  }}
  .finding .ap-label {{
    display: inline-block; padding: 0.1rem 0.5rem; margin-left: 0.5rem;
    background: var(--bg-warn-soft); color: var(--yellow); border-radius: 4px;
    font-size: 0.7rem; font-weight: 600;
  }}

  /* Recommendations */
  .rec {{
    padding: 1.25rem; margin-bottom: 1rem; border-radius: 10px;
    background: var(--bg-info-soft);
    border: 1px solid var(--border);
  }}
  .rec .prio {{ font-size: 0.75rem; color: var(--accent); font-weight: 600; text-transform: uppercase; }}
  .rec h3 {{ margin: 0.25rem 0 0.5rem; font-size: 1.05rem; }}
  .rec .reason {{ font-size: 0.85rem; color: var(--muted); margin-bottom: 0.5rem; }}
  .rec .action {{
    font-size: 0.9rem; padding: 0.75rem;
    background: var(--card); border-radius: 8px; border: 1px solid var(--border);
  }}

  .badge {{
    display: inline-block; padding: 0.15rem 0.5rem; border-radius: 4px;
    font-size: 0.75rem; font-weight: 600;
  }}

  .report-actions {{
    display: flex; justify-content: center; gap: 1rem;
    margin: 2rem 0 0.5rem; flex-wrap: wrap;
  }}
  .action-btn {{
    display: inline-flex; align-items: center; gap: 0.5rem;
    padding: 0.75rem 1.5rem; border-radius: 10px;
    font-size: 0.95rem; font-weight: 600; cursor: pointer;
    text-decoration: none; border: none; font-family: inherit;
    transition: background 0.2s, transform 0.1s;
  }}
  .action-btn:hover {{ transform: translateY(-1px); }}
  .action-download {{
    background: var(--accent); color: white;
  }}
  .action-download:hover {{ background: var(--accent-hover); }}
  .action-new {{
    background: var(--bg-ok-soft); color: var(--green); border: 1px solid var(--border);
  }}
  .action-new:hover {{ background: var(--bg-success-soft); }}

  .footer {{
    text-align: center; padding: 1.5rem; color: var(--muted);
    font-size: 0.8rem;
  }}

  /* Per-Sheet Tabelle */
  .sheet-table {{
    width: 100%; border-collapse: collapse; font-size: 0.85rem;
  }}
  .sheet-table th {{
    padding: 0.6rem 0.75rem; text-align: left;
    background: var(--bg); border-bottom: 2px solid var(--border);
    font-weight: 600; font-size: 0.8rem; color: var(--muted);
    text-transform: uppercase; letter-spacing: 0.03em;
  }}
  .sheet-table td {{
    padding: 0.5rem 0.75rem; border-bottom: 1px solid var(--border);
  }}
  .sheet-table tr:hover {{ background: var(--bg); }}
  .sheet-table .num {{ text-align: right; font-variant-numeric: tabular-nums; }}
  .phantom-badge {{
    display: inline-block; padding: 0.1rem 0.4rem; margin-left: 0.4rem;
    background: var(--bg-warn-soft); color: var(--yellow); border-radius: 4px;
    font-size: 0.65rem; font-weight: 600;
  }}
  .db-bar {{
    display: flex; align-items: center; gap: 0.5rem;
  }}
  .db-bar-track {{
    flex: 1; height: 6px; background: var(--border); border-radius: 3px;
    overflow: hidden;
  }}
  .db-bar-fill {{
    height: 100%; border-radius: 3px; transition: width 0.3s;
  }}
  .db-bar-val {{
    font-weight: 700; font-size: 0.8rem; min-width: 2rem; text-align: right;
  }}

  /* Per-Sheet Analyse Cards */
  .sheet-card-header {{
    display: flex; justify-content: space-between; align-items: center;
    margin-bottom: 1rem; flex-wrap: wrap; gap: 0.75rem;
  }}
  .sheet-card-stats {{
    display: grid; grid-template-columns: repeat(auto-fill, minmax(110px, 1fr));
    gap: 0.5rem; margin-bottom: 1rem;
  }}
  .sheet-mini-stat {{
    text-align: center; padding: 0.5rem; background: var(--bg); border-radius: 6px;
  }}
  .sheet-mini-stat .num {{ font-size: 1.1rem; font-weight: 700; color: var(--accent); }}
  .sheet-mini-stat .lbl {{ font-size: 0.7rem; color: var(--muted); }}

  /* Footer Cards */
  .footer-cards {{
    display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
    gap: 1rem; margin: 1.5rem 0;
  }}
  .footer-card {{
    padding: 1.25rem; border-radius: 12px;
    border: 1px solid var(--border); background: var(--card);
    text-align: center; cursor: pointer; text-decoration: none;
    color: var(--text); transition: transform 0.15s, box-shadow 0.15s;
    display: block;
  }}
  .footer-card:hover {{
    transform: translateY(-3px);
    box-shadow: 0 4px 12px rgba(0,0,0,0.08);
  }}
  .footer-card .fc-icon {{ font-size: 2rem; margin-bottom: 0.5rem; }}
  .footer-card .fc-title {{ font-weight: 700; font-size: 1rem; margin-bottom: 0.25rem; }}
  .footer-card .fc-desc {{ font-size: 0.82rem; color: var(--muted); line-height: 1.4; }}

  /* Effizienz-Section */
  .eff-grid {{
    display: grid; grid-template-columns: 1fr 1fr; gap: 1.5rem;
    margin-bottom: 1.5rem;
  }}
  .eff-box {{
    padding: 1.25rem; border-radius: 10px;
    border: 1px solid var(--border);
  }}
  .eff-box h3 {{
    font-size: 0.95rem; margin-bottom: 0.75rem;
    display: flex; align-items: center; gap: 0.5rem;
  }}
  .eff-metric {{
    display: flex; justify-content: space-between; align-items: center;
    padding: 0.5rem 0; border-bottom: 1px solid var(--border);
    font-size: 0.85rem;
  }}
  .eff-metric:last-child {{ border-bottom: none; }}
  .eff-metric .label {{ color: var(--muted); }}
  .eff-metric .value {{ font-weight: 700; font-size: 1rem; }}
  .eff-metric .value.red {{ color: var(--red); }}
  .eff-metric .value.green {{ color: var(--green); }}
  .eff-metric .value.blue {{ color: var(--accent); }}
  .eff-bar {{
    height: 8px; border-radius: 4px; background: #e2e8f0;
    margin-top: 0.25rem; overflow: hidden;
  }}
  .eff-bar-fill {{
    height: 100%; border-radius: 4px;
    transition: width 0.3s;
  }}
  .eff-summary {{
    padding: 1.25rem; border-radius: 10px;
    background: var(--bg-warning-soft);
    border: 1px solid var(--border);
    font-size: 0.9rem; line-height: 1.7;
  }}
  .eff-summary strong {{ color: var(--yellow); }}
  .ai-potential {{
    padding: 1.25rem; border-radius: 10px;
    background: var(--bg-info-soft);
    border: 1px solid var(--border);
    font-size: 0.9rem; line-height: 1.7;
  }}
  .ai-potential strong {{ color: var(--accent); }}
  .ai-blocker-list {{
    list-style: none; padding: 0; margin: 0.5rem 0;
  }}
  .ai-blocker-list li {{
    padding: 0.4rem 0; font-size: 0.85rem; color: var(--text);
    border-bottom: 1px solid var(--border);
  }}
  .ai-blocker-list li:last-child {{ border-bottom: none; }}

  /* ---- Responsive ---- */
  @media (max-width: 768px) {{
    body {{ padding: 0.5rem; }}
    .hero-grid {{ grid-template-columns: 1fr; }}
    .charts-grid {{ grid-template-columns: 1fr; }}
    .stats-grid {{ grid-template-columns: repeat(2, 1fr); }}
    .ap-grid {{ grid-template-columns: repeat(auto-fill, minmax(160px, 1fr)); }}
    .eff-grid {{ grid-template-columns: 1fr; }}
    .card {{ padding: 1rem; }}
    .tabs {{ gap: 0; }}
    .tab-btn {{ padding: 0.5rem 0.8rem; font-size: 0.8rem; }}
  }}
  @media (max-width: 480px) {{
    .stats-grid {{ grid-template-columns: 1fr; }}
    .ap-grid {{ grid-template-columns: 1fr; }}
    .score-ring {{ width: 90px; height: 90px; }}
    .score-ring .value {{ font-size: 2rem; }}
    .header h1 {{ font-size: 1.2rem; }}
  }}
</style>
</head>
<body>
<button class="theme-toggle-report" onclick="toggleTheme()" title="Dark/Light Mode" id="themeToggleBtn">🌙</button>
<div class="container">
<div class="header">
  <h1>⚡ Excel-Reifecheck</h1>
  <div class="subtitle">{_esc(filename)} &middot; {now}</div>
  <div style="font-size:0.8rem; color:var(--muted); margin-top:0.25rem;">Version alpha &middot; bereitgestellt von UCS - Digital Workplace Solutions</div>
</div>"""


def _html_hero(score: int, color: str, label: str,
               report: WorkbookReport, dims: dict[str, float],
               ai_ready: int, ai_ready_color: str, avg_db: int,
               compliance_score: int, compliance_color: str) -> str:
    radar = _radar_chart_svg(dims)

    ai_label = ('Bereit für KI' if ai_ready >= 70
                else 'Nach Bereinigung' if ai_ready >= 40
                else 'Erheblicher Aufwand')
    compliance_label = ('Kein erkennbares Risiko' if compliance_score >= 80
                        else 'Geringes Risiko' if compliance_score >= 60
                        else 'Mittleres Risiko' if compliance_score >= 40
                        else 'Hohes Risiko')

    gauge_quality = _html_score_gauge(
        "Datenqualität", score, color,
        hint=_esc(label),
        tooltip="Datenqualität: wie sauber sind die Daten für Mitarbeiter, die täglich damit arbeiten?",
    )
    gauge_compliance = _html_score_gauge(
        "Compliance", compliance_score, compliance_color,
        hint=compliance_label,
        tooltip="Compliance: Risiko aus sensiblen Daten und undokumentierter Logik. Höher = weniger Risiko.",
    )
    gauge_ai = _html_score_gauge(
        "AI-Readiness", ai_ready, ai_ready_color,
        hint=_esc(ai_label),
        tooltip="AI-Readiness: wie gut eignen sich die Daten für KI-Auswertungen und Datenbank-Migration?",
    )

    return f"""
<div class="card">
  <h2>🎯 Gesamtbewertung</h2>
  <div class="score-gauge-group">
    {gauge_quality}
    {gauge_compliance}
    {gauge_ai}
  </div>
  <div style="text-align:center;padding:0.25rem 0 0.75rem;font-size:0.8rem;color:var(--muted);">
    {report.critical_count} kritisch &middot;
    {report.error_count} wichtig &middot;
    {report.warning_count} Hinweise &middot;
    {report.info_count} Info
  </div>
  <div class="hero-tabs">
    <button class="hero-tab-btn active" onclick="switchHeroTab(this,'hero-radar')">📈 Komplexitätsprofil</button>
    <button class="hero-tab-btn" onclick="switchHeroTab(this,'hero-profil')">📋 Kurzprofil</button>
  </div>
  <div id="hero-radar" class="hero-tab-panel active">
    <div style="text-align:center;">
      {radar}
      <p style="font-size:0.75rem;color:var(--muted);margin-top:0.25rem;line-height:1.4;">
        ⚠️ Je weiter eine Achse nach außen reicht, desto größer der Handlungsbedarf.
      </p>
    </div>
  </div>
  <div id="hero-profil" class="hero-tab-panel">
    <div style="font-size:0.85rem;line-height:2;max-width:500px;margin:0 auto;">
      <div style="display:flex;justify-content:space-between;border-bottom:1px solid #f1f5f9;padding:0.3rem 0;">
        <span>📄 Dateigröße</span>
        <span style="font-weight:600;">{report.file_size_mb:.1f} MB</span>
      </div>
      <div style="display:flex;justify-content:space-between;border-bottom:1px solid #f1f5f9;padding:0.3rem 0;">
        <span>📋 Tabellenblätter</span>
        <span style="font-weight:600;">{report.sheet_count}</span>
      </div>
      <div style="display:flex;justify-content:space-between;border-bottom:1px solid #f1f5f9;padding:0.3rem 0;">
        <span>📊 Zeilen (befüllt)</span>
        <span style="font-weight:600;">{sum(s.row_count for s in report.sheet_stats):,}</span>
      </div>
      <div style="display:flex;justify-content:space-between;border-bottom:1px solid #f1f5f9;padding:0.3rem 0;">
        <span>🔢 Formeln</span>
        <span style="font-weight:600;">{sum(s.formula_count for s in report.sheet_stats):,}</span>
      </div>
      <div style="display:flex;justify-content:space-between;padding:0.3rem 0;">
        <span>🔗 Verbundene Zellen</span>
        <span style="font-weight:600;">{sum(s.merged_regions for s in report.sheet_stats):,}</span>
      </div>
    </div>
  </div>
</div>"""


def _html_score_breakdown(report: WorkbookReport) -> str:
    """Zeigt transparent, wie die Bewertung zustande kommt."""
    # Penalties nach Regel aggregieren
    from collections import Counter
    penalty_by_rule: dict[str, tuple[str, int, int]] = {}  # rule_id -> (name, count, total_penalty)
    for f in report.findings:
        if f.score_penalty > 0:
            if f.rule_id not in penalty_by_rule:
                penalty_by_rule[f.rule_id] = [f.rule_name, 0, 0]
            penalty_by_rule[f.rule_id][1] += 1
            penalty_by_rule[f.rule_id][2] += f.score_penalty

    total_penalty = sum(v[2] for v in penalty_by_rule.values())
    score = max(0, 100 - total_penalty)

    if not penalty_by_rule:
        return f"""
<details class="card">
  <summary><h2>📋 Bewertungs-Aufschlüsselung</h2></summary>
  <p style="color:var(--muted);font-size:0.9rem;">
    ✅ Keine Abzüge – alle geprüften Regeln bestanden. Bewertung: <strong>100/100</strong>
  </p>
</details>"""

    # Nach Penalty sortieren (größter Abzug zuerst)
    sorted_rules = sorted(penalty_by_rule.items(), key=lambda x: -x[1][2])

    rows = []
    for rule_id, (rule_name, count, penalty) in sorted_rules:
        # Grade-Info aus ANTI_PATTERNS
        grade_html = ""
        if rule_id in ANTI_PATTERNS:
            _, _, _, ap_grade = ANTI_PATTERNS[rule_id]
            grade_icon, grade_color, _ = ANTI_PATTERN_GRADES.get(
                ap_grade, ("⚪", "#6b7280", "#f9fafb"))
            grade_html = (
                f'<span style="font-size:0.7rem;padding:0.1rem 0.4rem;'
                f'border-radius:4px;background:{grade_color}15;color:{grade_color};'
                f'font-weight:600;margin-left:0.4rem;">'
                f'{grade_icon} {_esc(ap_grade)}</span>'
            )

        bar_width = min(100, penalty * 2)
        sev_color = "#dc2626" if penalty >= 10 else "#ea580c" if penalty >= 5 else "#ca8a04"
        rows.append(f"""
      <tr>
        <td style="font-size:0.8rem;color:var(--muted);">{_esc(rule_id)}</td>
        <td style="font-weight:500;">{_esc(rule_name)}{grade_html}</td>
        <td class="num" style="font-size:0.85rem;">{count}×</td>
        <td class="num" style="font-weight:700;color:{sev_color};">−{penalty}</td>
        <td style="width:120px;">
          <div style="height:6px;background:#e2e8f0;border-radius:3px;overflow:hidden;">
            <div style="width:{bar_width}%;height:100%;background:{sev_color};border-radius:3px;"></div>
          </div>
        </td>
      </tr>""")

    return f"""
<details class="card">
  <summary><h2>📋 Bewertungs-Aufschlüsselung: So entsteht die Punktzahl</h2></summary>
  <p style="color:var(--muted);font-size:0.85rem;margin-bottom:1rem;line-height:1.5;">
    Die Bewertung startet bei <strong>100 Punkten</strong>. Jede erkannte Regel zieht
    Punkte ab – je schwerwiegender das Problem, desto mehr. So ist der Score jederzeit
    nachvollziehbar.
  </p>
  <div style="overflow-x:auto;">
    <table class="sheet-table">
      <thead>
        <tr>
          <th style="width:70px;">Regel</th>
          <th>Beschreibung</th>
          <th class="num" style="width:50px;">Treffer</th>
          <th class="num" style="width:60px;">Abzug</th>
          <th style="width:120px;"></th>
        </tr>
      </thead>
      <tbody>{''.join(rows)}
      </tbody>
      <tfoot>
        <tr style="border-top:2px solid var(--border);">
          <td colspan="3" style="font-weight:700;padding-top:0.75rem;">Gesamt-Abzug</td>
          <td class="num" style="font-weight:700;color:#dc2626;font-size:1rem;padding-top:0.75rem;">
            −{total_penalty}</td>
          <td style="padding-top:0.75rem;">
            <span style="font-weight:700;font-size:1rem;">→ {score}/100</span>
          </td>
        </tr>
      </tfoot>
    </table>
  </div>
</details>"""


def _html_stats_section(report: WorkbookReport) -> str:
    total_rows = sum(s.row_count for s in report.sheet_stats)
    total_cols = sum(s.col_count for s in report.sheet_stats)
    total_formulas = sum(s.formula_count for s in report.sheet_stats)
    total_merged = sum(s.merged_regions for s in report.sheet_stats)
    size_str = f"{report.file_size_mb:.1f} MB" if report.file_size_mb >= 0.1 else f"{report.file_size_mb * 1024:.0f} KB"
    # Stern-Markierung wenn Stichprobe verwendet wurde (nicht-phantom Sheet > 200 Zeilen)
    sampled = any(
        (s.row_count > 200 or s.col_count > 30) and not s.is_phantom
        for s in report.sheet_stats
    )
    has_phantom = any(s.is_phantom for s in report.sheet_stats)
    star = " *" if sampled else ""
    footnotes = []
    if sampled:
        footnotes.append(
            '* Werte auf Basis einer Stichprobe (200 Zeilen × 30 Spalten pro Sheet) hochgerechnet.'
        )
    if has_phantom:
        footnotes.append(
            '⚠️ Phantom-Zeilen erkannt: Excel meldete bis zu 1.048.576 Zeilen, '
            'aber die tatsächliche Datengrenze wurde durch vollständigen Scan ermittelt.'
        )
    footnote_html = '\n'.join(
        f'  <p style="font-size:0.75rem;color:var(--muted);margin-top:0.5rem;">{fn}</p>'
        for fn in footnotes
    )

    # Per-Sheet Breakdown Table
    sheet_rows = []
    for s in report.sheet_stats:
        db_color = _db_readiness_color(s.db_readiness)
        phantom_tag = '<span class="phantom-badge">Phantom bereinigt</span>' if s.is_phantom else ''
        sheet_rows.append(f"""
      <tr>
        <td style="font-weight:600;">{_esc(s.name)}{phantom_tag}</td>
        <td class="num">{s.row_count:,}</td>
        <td class="num">{s.col_count:,}</td>
        <td class="num">{s.formula_count:,}</td>
        <td class="num">{s.merged_regions}</td>
        <td>
          <div class="db-bar">
            <div class="db-bar-track">
              <div class="db-bar-fill" style="width:{s.db_readiness}%;background:{db_color};"></div>
            </div>
            <span class="db-bar-val" style="color:{db_color};">{s.db_readiness}</span>
          </div>
        </td>
      </tr>""")
    sheet_table = '\n'.join(sheet_rows)

    # Gesamt-DB-Readiness (gewichteter Durchschnitt nach Zeilenzahl)
    if total_rows > 0:
        avg_db = sum(s.db_readiness * max(s.row_count, 1) for s in report.sheet_stats) / max(total_rows, 1)
    else:
        avg_db = sum(s.db_readiness for s in report.sheet_stats) / max(len(report.sheet_stats), 1)
    avg_db = int(avg_db)
    avg_db_color = _db_readiness_color(avg_db)
    avg_db_label = _db_readiness_label(avg_db)

    return f"""
<details class="card">
  <summary><h2>📋 Aufschlüsselung pro Sheet</h2></summary>
  {footnote_html}
  <div style="overflow-x:auto;">
    <table class="sheet-table">
      <thead>
        <tr>
          <th>Sheet</th>
          <th class="num">Zeilen</th>
          <th class="num">Spalten</th>
          <th class="num">Formeln</th>
          <th class="num">Verbundene</th>
          <th style="min-width:120px;">DB-Eignung</th>
        </tr>
      </thead>
      <tbody>{sheet_table}
      </tbody>
    </table>
    <p style="font-size:0.72rem;color:var(--muted);margin-top:0.5rem;">
      🗄️ DB-Eignung: Wie gut lässt sich das Blatt in eine Datenbank überführen?
      100 = direkt übernehmbar · 0 = erhebliche Aufbereitung nötig
    </p>
  </div>
</details>"""


def _db_readiness_color(score: int) -> str:
    if score >= 75:
        return "#16a34a"
    elif score >= 50:
        return "#ca8a04"
    elif score >= 25:
        return "#ea580c"
    return "#dc2626"


def _db_readiness_label(score: int) -> str:
    if score >= 75:
        return "Gut übertragbar"
    elif score >= 50:
        return "Mit Aufbereitung möglich"
    elif score >= 25:
        return "Erhebliche Aufbereitung nötig"
    return "Grundlegende Umstrukturierung nötig"


def _html_per_sheet_section(report: WorkbookReport) -> str:
    """Sektion: Pro-Blatt Detailanalyse mit DB-Eignung und Ergebnissen."""
    if not report.sheet_stats:
        return ""

    tab_id = "sheet"
    parts = [
        '<details class="card">',
        '<summary><h2>🗄️ Blatt-für-Blatt Analyse</h2></summary>',
        '<p style="color:var(--muted);font-size:0.85rem;margin-bottom:1rem;">'
        'Detaillierte Bewertung pro Blatt: Wie gut ist die Datenstruktur für eine '
        'Überführung in eine Datenbank oder ein strukturiertes System geeignet?</p>',
    ]

    # Tab-Buttons
    parts.append(f'<div class="tabs" id="{tab_id}-tabs">')
    for i, s in enumerate(report.sheet_stats):
        active = " active" if i == 0 else ""
        db_color = _db_readiness_color(s.db_readiness)
        parts.append(
            f'<button class="tab-btn{active}" '
            f'onclick="switchTab(\'{tab_id}\', \'{_esc(s.name)}\')">'
            f'{_esc(s.name)} '
            f'<span style="font-size:0.7rem;color:{db_color};font-weight:700;">'
            f'{s.db_readiness}</span></button>'
        )
    parts.append('</div>')

    # Tab-Inhalte
    for i, s in enumerate(report.sheet_stats):
        active = " active" if i == 0 else ""
        db_color = _db_readiness_color(s.db_readiness)
        db_label = _db_readiness_label(s.db_readiness)
        sheet_findings = [f for f in report.findings if f.sheet == s.name]

        parts.append(
            f'<div class="tab-content{active}" '
            f'data-tab-group="{tab_id}" data-tab="{_esc(s.name)}">'
        )

        # Header: DB-Readiness Score prominent
        parts.append(f"""
  <div class="sheet-card-header">
    <div>
      <span style="font-size:1.3rem;font-weight:700;">{_esc(s.name)}</span>
      {'<span class="phantom-badge" style="margin-left:0.5rem;">Phantom-Zeilen bereinigt</span>' if s.is_phantom else ''}
    </div>
    <div style="text-align:center;">
      <div style="font-size:0.7rem;color:var(--muted);text-transform:uppercase;letter-spacing:0.05em;">
        DB-Eignung</div>
      <div style="font-size:2rem;font-weight:700;color:{db_color};line-height:1;">{s.db_readiness}</div>
      <div style="font-size:0.75rem;color:{db_color};">{_esc(db_label)}</div>
    </div>
  </div>""")

        # Mini-Stats
        total_cells = s.formula_count + s.static_count
        formula_pct = f"{s.formula_count / total_cells * 100:.0f}%" if total_cells > 0 else "–"
        parts.append(f"""
  <div class="sheet-card-stats">
    <div class="sheet-mini-stat"><div class="num">{s.row_count:,}</div><div class="lbl">Zeilen</div></div>
    <div class="sheet-mini-stat"><div class="num">{s.col_count:,}</div><div class="lbl">Spalten</div></div>
    <div class="sheet-mini-stat"><div class="num">{s.formula_count:,}</div><div class="lbl">Formeln</div></div>
    <div class="sheet-mini-stat"><div class="num">{s.static_count:,}</div><div class="lbl">Datenzellen</div></div>
    <div class="sheet-mini-stat"><div class="num">{s.merged_regions}</div><div class="lbl">Verbundene</div></div>
    <div class="sheet-mini-stat"><div class="num">{formula_pct}</div><div class="lbl">Formel-Anteil</div></div>
  </div>""")

        # DB-Readiness Bar
        parts.append(f"""
  <div style="margin-bottom:1rem;">
    <div class="db-bar" style="max-width:400px;">
      <div class="db-bar-track" style="height:10px;">
        <div class="db-bar-fill" style="width:{s.db_readiness}%;background:{db_color};"></div>
      </div>
      <span class="db-bar-val" style="color:{db_color};font-size:1rem;">{s.db_readiness}/100</span>
    </div>
  </div>""")

        # Sheet Findings (Phase 2 C.7: Accordion-Cards via _render_finding_card)
        if sheet_findings:
            parts.append(
                f'<h3 style="font-size:0.9rem;color:var(--muted);margin-bottom:0.5rem;">'
                f'📋 Ergebnisse ({len(sheet_findings)})</h3>'
            )
            sev_order = {Severity.CRITICAL: 0, Severity.ERROR: 1,
                         Severity.WARNING: 2, Severity.INFO: 3}
            for f in sorted(sheet_findings, key=lambda x: sev_order.get(x.severity, 9)):
                parts.append(_render_finding_card(f))
        else:
            parts.append(
                '<div style="text-align:center;padding:1.5rem;color:var(--muted);">'
                '✅ Keine Probleme erkannt – dieses Sheet ist gut strukturiert.</div>'
            )

        parts.append('</div>')  # tab-content

    parts.append('</details>')  # card
    return "\n".join(parts)


def _html_anti_patterns(detected: dict[str, tuple[str, str, str, int, str]],
                       findings: list) -> str:
    # Fundorte pro Rule-ID sammeln
    locations_by_rid: dict[str, list[str]] = defaultdict(list)
    for f in findings:
        if f.rule_id in detected:
            loc_parts = []
            if f.sheet:
                loc_parts.append(f"Sheet \"{f.sheet}\"")
            if f.cell:
                loc_parts.append(f"Zelle {f.cell}")
            if not loc_parts and f.message:
                loc_parts.append(f.message[:60])
            if loc_parts:
                locations_by_rid[f.rule_id].append(" · ".join(loc_parts))

    # Nach Schweregrad gruppieren
    by_grade: dict[str, list[tuple[str, str, str, str, int]]] = defaultdict(list)
    for rid, (name, icon, desc, count, grade) in detected.items():
        by_grade[grade].append((rid, name, icon, desc, count))

    # Sortiert nach Schweregrad (Mega → Echt → Schön)
    sorted_grades = sorted(by_grade.keys(), key=lambda g: GRADE_ORDER.get(g, 99))

    parts = [
        '<details class="card" open>',
        f'<summary><h2>⚠️ Erkannte Anti-Patterns ({len(detected)})</h2></summary>',
    ]
    # Mapping vom Grade-Label auf eine CSS-Klasse (theme.css steuert Light+Dark)
    grade_class_map = {
        GRADE_MEGA: "grade-mega",
        GRADE_ECHT: "grade-echt",
        GRADE_SCHOEN: "grade-schoen",
    }
    for grade in sorted_grades:
        items = by_grade[grade]
        grade_icon, grade_color, _ = ANTI_PATTERN_GRADES.get(
            grade, ("⚪", "#6b7280", "#f9fafb"))
        grade_cls = grade_class_map.get(grade, "grade-schoen")

        parts.append(f'<div class="ap-grade-section">')
        parts.append(
            f'<div class="ap-grade-header">'
            f'<span class="grade-icon">{grade_icon}</span>'
            f'<span style="color:{grade_color};">{_esc(grade)}</span>'
            f'<span class="grade-count">({len(items)})</span>'
            f'</div>')
        parts.append('<div class="ap-grid">')
        for rid, name, icon, desc, count in items:
            locs = locations_by_rid.get(rid, [])
            # Rückseite: Fundorte
            back_html = f'<div class="ap-back-title">📍 Fundorte</div>'
            if locs:
                all_items = "".join(f"<li>📄 {_esc(l)}</li>" for l in locs)
                back_html += f'<ul class="ap-all-locations visible">{all_items}</ul>'
            else:
                back_html += '<div class="ap-location">Keine Ortsangabe</div>'
            parts.append(f"""
    <div class="ap-flip-container" onclick="this.classList.toggle('flipped')">
      <div class="ap-flip-inner">
        <div class="ap-card {grade_cls}">
          <div class="ap-icon">{icon}</div>
          <div class="ap-name">{_esc(name)}</div>
          <div class="ap-desc">{_esc(desc)}</div>
          <div class="ap-count" style="background:{grade_color};">{count}× gefunden</div>
          <span class="ap-flip-hint">klicken zum Umdrehen ↻</span>
        </div>
        <div class="ap-card-back {grade_cls}">
          {back_html}
        </div>
      </div>
    </div>""")
        parts.append('</div></div>')

    parts.append('</details>')
    return "\n".join(parts)


def _html_complexity_section(report: WorkbookReport, dims: dict[str, float]) -> str:
    """Abschnitt mit Balkendiagrammen für Zeilen und Formeln pro Sheet."""
    rows_data = {s.name: s.row_count for s in report.sheet_stats if s.row_count > 0}
    formula_data = {s.name: s.formula_count for s in report.sheet_stats if s.formula_count > 0}

    parts = ['<details class="card">', '<summary><h2>📈 Datenverteilung pro Sheet</h2></summary>',
             '<div class="charts-grid">']

    if rows_data:
        parts.append('<div class="chart-box">')
        parts.append('<h3>Zeilen pro Sheet</h3>')
        parts.append(_bar_chart_svg(rows_data, "Zeilen"))
        parts.append('</div>')

    if formula_data:
        parts.append('<div class="chart-box">')
        parts.append('<h3>Formeln pro Sheet</h3>')
        parts.append(_bar_chart_svg(formula_data, "Formeln"))
        parts.append('</div>')

    if not rows_data and not formula_data:
        parts.append('<p style="color:var(--muted);text-align:center;">Keine Sheet-Daten verfügbar.</p>')

    parts.append('</div></details>')
    return "\n".join(parts)


def _html_efficiency_section(report: WorkbookReport,
                             detected_ap: dict[str, tuple],
                             ai_ready: int, ai_ready_color: str,
                             ai_blockers: list[tuple[str, str, str]]) -> str:
    """Sektion: Warum aufgeblähte Excel-Dateien Effizienz und KI-Potenzial kosten."""
    total_rows = sum(s.row_count for s in report.sheet_stats)
    total_formulas = sum(s.formula_count for s in report.sheet_stats)
    total_merged = sum(s.merged_regions for s in report.sheet_stats)
    sheets = report.sheet_count

    # ── Konkrete Schätzungen auf Basis der Datei-Eigenschaften ──
    # Zeitverlust durch manuelle Excel-Pflege (konservativ)
    # Basis: 5 Min/Woche pro 1.000 Zeilen für Pflege, Prüfung, Abgleich
    weekly_maint_min = max(10, total_rows / 1000 * 5)
    # Formeln: jede 100 Formeln = 2 Min/Woche Debugging/Prüfung
    weekly_formula_min = total_formulas / 100 * 2
    # Merged cells & Struktur-Chaos: pauschal 5 Min/Woche pro 10 Merged-Regions
    weekly_struct_min = total_merged / 10 * 5
    # Multi-Sheet-Navigation: 3 Min/Woche pro Sheet über 3
    weekly_nav_min = max(0, sheets - 3) * 3
    # Gesamtzeit pro Woche (für eine Person)
    weekly_total_min = weekly_maint_min + weekly_formula_min + weekly_struct_min + weekly_nav_min
    weekly_total_min = min(weekly_total_min, 600)  # cap bei 10h
    weekly_total_h = weekly_total_min / 60

    # Kostenbasis: ~€65/h Vollkosten (Durchschnitt Sachbearbeiter AT)
    cost_per_hour = 65
    annual_cost_one = weekly_total_h * 48 * cost_per_hour  # 48 Arbeitswochen
    # Typisch nutzen 2-5 Personen dieselbe Excel → Faktor 3 als konservative Mitte
    users_estimate = 3
    annual_cost_team = annual_cost_one * users_estimate

    # Mit Datenbank: 80% Zeitersparnis (konservativ)
    savings_pct = 80
    annual_savings = annual_cost_team * savings_pct / 100

    # ── HTML ──
    parts = [
        '<div class="card">',
        '<h2>💰 Effizienz-Analyse: Was aufgeblähte Excel-Dateien wirklich kosten</h2>',
        '<p style="color:var(--muted);font-size:0.85rem;margin-bottom:1.25rem;line-height:1.6;">'
        'Excel ist das richtige Werkzeug, wenn man schnell loslegen muss – und oft hat man '
        'initial nichts anderes. Aber ab einem gewissen Punkt kosten Workarounds in Excel '
        'mehr als eine saubere Lösung. Diese Schätzung basiert auf den konkreten Eigenschaften '
        'dieser Datei.</p>',
    ]

    # Grid: Zeitverlust links, KI-Readiness rechts
    parts.append('<div class="eff-grid">')

    # ── Linke Box: Zeitverlust ──
    parts.append(
        '<div class="eff-box" style="background:var(--bg-warn-soft);">'
        '<h3>⏱️ Geschätzter Zeitverlust durch manuelle Pflege</h3>'
    )
    parts.append(f"""
      <div class="eff-metric">
        <span class="label">Datenpflege & Abgleich ({total_rows:,} Zeilen)</span>
        <span class="value red">~{weekly_maint_min:.0f} Min/Woche</span>
      </div>""")
    if weekly_formula_min > 0:
        parts.append(f"""
      <div class="eff-metric">
        <span class="label">Formel-Prüfung & Debugging ({total_formulas:,} Formeln)</span>
        <span class="value red">~{weekly_formula_min:.0f} Min/Woche</span>
      </div>""")
    if weekly_struct_min > 0:
        parts.append(f"""
      <div class="eff-metric">
        <span class="label">Strukturprobleme (Merged Cells etc.)</span>
        <span class="value red">~{weekly_struct_min:.0f} Min/Woche</span>
      </div>""")
    if weekly_nav_min > 0:
        parts.append(f"""
      <div class="eff-metric">
        <span class="label">Multi-Sheet-Navigation ({sheets} Blätter)</span>
        <span class="value red">~{weekly_nav_min:.0f} Min/Woche</span>
      </div>""")
    parts.append(f"""
      <div class="eff-metric" style="border-top:2px solid #fde68a;padding-top:0.75rem;margin-top:0.25rem;">
        <span class="label"><strong>Gesamt pro Person</strong></span>
        <span class="value red" style="font-size:1.15rem;">~{weekly_total_h:.1f} h/Woche</span>
      </div>
    </div>""")

    # ── Rechte Box: KI-Readiness ──
    parts.append(
        '<div class="eff-box" style="background:var(--bg-info-soft);">'
        '<h3>🤖 KI-Readiness dieser Datei</h3>'
    )
    parts.append(f"""
      <div style="text-align:center;margin:0.75rem 0;">
        <span style="font-size:2.5rem;font-weight:700;color:{ai_ready_color};">{ai_ready}</span>
        <span style="font-size:0.9rem;color:var(--muted);"> / 100</span>
      </div>
      <div class="eff-bar" style="margin-bottom:0.75rem;">
        <div class="eff-bar-fill" style="width:{ai_ready}%;background:{ai_ready_color};"></div>
      </div>
      <p style="font-size:0.8rem;color:var(--muted);text-align:center;margin-bottom:0.75rem;">
        {'Bereit für KI-Auswertung' if ai_ready >= 70
         else 'Nach Bereinigung KI-tauglich' if ai_ready >= 40
         else 'Erhebliche Aufbereitung nötig'}
      </p>""")
    if ai_blockers:
        parts.append('<ul class="ai-blocker-list">')
        for bicon, btitle, bdesc in ai_blockers:
            parts.append(
                f'<li>{bicon} <strong>{_esc(btitle)}</strong> – {_esc(bdesc)}</li>')
        parts.append('</ul>')
    parts.append('</div>')  # eff-box
    parts.append('</div>')  # eff-grid

    # ── Kosten-Zusammenfassung ──
    parts.append(f"""
    <div class="eff-summary">
      <strong>📊 Konservative Kostenrechnung</strong> (bei ~{users_estimate} Nutzern, €{cost_per_hour}/h Vollkosten)
      <div class="eff-metric" style="border-bottom:none;padding:0.5rem 0 0;">
        <span class="label">Jährliche Kosten für manuelle Excel-Pflege</span>
        <span class="value red">~€{annual_cost_team:,.0f}</span>
      </div>
      <div class="eff-metric" style="border-bottom:none;padding:0 0 0.5rem;">
        <span class="label">Davon einsparbar durch Datenbank/Automatisierung ({savings_pct}%)</span>
        <span class="value green">~€{annual_savings:,.0f}/Jahr</span>
      </div>
      <p style="font-size:0.8rem;color:#92400e;margin-top:0.5rem;">
        💡 Eine typische Migration in eine SharePoint-Liste oder einfache Datenbank kostet
        einmalig €5.000–25.000. Bei diesen Zahlen amortisiert sich das
        <strong>innerhalb von {max(1, round(15000 / max(1, annual_savings) * 12))} Monaten</strong>.
      </p>
    </div>""")

    # ── AI-Potential ──
    parts.append(f"""
    <div class="ai-potential" style="margin-top:1rem;">
      <strong>🚀 Was mit sauberen Daten möglich wäre</strong>
      <table style="width:100%;border-collapse:collapse;font-size:0.85rem;margin-top:0.75rem;">
        <tr style="border-bottom:1px solid rgba(59,130,246,0.15);">
          <td style="padding:0.5rem 0.75rem;">📋 <strong>Automatische Berichte</strong></td>
          <td style="padding:0.5rem 0.75rem;color:var(--muted);">
            Statt manueller Zusammenstellung: Power BI / Dashboards auf Knopfdruck.
            Zeitersparnis ~2-5h/Woche pro Berichtsempfänger.
          </td>
        </tr>
        <tr style="border-bottom:1px solid rgba(59,130,246,0.15);">
          <td style="padding:0.5rem 0.75rem;">🔍 <strong>KI-Anomalie-Erkennung</strong></td>
          <td style="padding:0.5rem 0.75rem;color:var(--muted);">
            KI erkennt Ausreißer, fehlende Werte und Inkonsistenzen automatisch –
            aber nur wenn Daten strukturiert vorliegen. In Excel mit verbundenen Zellen
            und Farblogik: unmöglich.
          </td>
        </tr>
        <tr style="border-bottom:1px solid rgba(59,130,246,0.15);">
          <td style="padding:0.5rem 0.75rem;">💬 <strong>Natürlichsprachliche Abfragen</strong></td>
          <td style="padding:0.5rem 0.75rem;color:var(--muted);">
            „Zeig mir alle Vorfälle im Q3 mit Priorität hoch" – funktioniert mit
            einer Datenbank in Sekunden. In dieser Excel-Datei: manuelle Suche nötig.
          </td>
        </tr>
        <tr>
          <td style="padding:0.5rem 0.75rem;">📈 <strong>Prognosen & Trends</strong></td>
          <td style="padding:0.5rem 0.75rem;color:var(--muted);">
            Vorhersagen auf Basis historischer Daten – aber nur mit sauberen Zeitreihen.
            Gemischte Datentypen und Freitext-IDs machen das unmöglich.
          </td>
        </tr>
      </table>
      <p style="font-size:0.8rem;color:#1e40af;margin-top:0.75rem;">
        ℹ️ Laut Studien verbringen Mitarbeiter in datenintensiven Rollen
        <strong>~30% ihrer Zeit</strong> mit Suchen, Aufbereiten und Prüfen von Daten
        (Quelle: IDC/McKinsey). In strukturierten Systemen sinkt dieser Anteil auf unter 10%.
      </p>
    </div>""")

    parts.append('</div>')
    return "\n".join(parts)


def _html_recommendations_tabs(recs: list[Recommendation]) -> str:
    """Empfehlungen als Tabs nach Typ gruppiert."""
    # Gruppieren nach Typ
    by_type: dict[RecommendationType, list[Recommendation]] = defaultdict(list)
    for r in recs:
        by_type[r.rec_type].append(r)

    tab_id = "rec"
    types = list(by_type.keys())

    parts = ['<details class="card" open>', '<summary><h2>🎯 Strategische Empfehlungen</h2></summary>']

    # Tab-Buttons
    parts.append(f'<div class="tabs" id="{tab_id}-tabs">')
    for i, rt in enumerate(types):
        icon = RECTYPE_ICONS.get(rt, "📌")
        active = " active" if i == 0 else ""
        parts.append(
            f'<button class="tab-btn{active}" '
            f'onclick="switchTab(\'{tab_id}\', \'{rt.value}\')">'
            f'{icon} {_esc(rt.value)}</button>'
        )
    parts.append('</div>')

    # Tab-Inhalte
    for i, rt in enumerate(types):
        active = " active" if i == 0 else ""
        parts.append(f'<div class="tab-content{active}" data-tab-group="{tab_id}" data-tab="{rt.value}">')
        for rec in by_type[rt]:
            icon = RECTYPE_ICONS.get(rec.rec_type, "📌")
            parts.append(f"""
  <div class="rec">
    <div class="prio">Priorität {rec.priority} &middot; {_esc(rec.rec_type.value)}</div>
    <h3>{icon} {_esc(rec.title)}</h3>
    <div class="reason">{_esc(rec.reason)}</div>
    <div class="action">
      <strong>Empfohlene Maßnahme:</strong><br>
      {_esc(rec.action)}
    </div>
  </div>""")
        parts.append('</div>')

    parts.append('</details>')
    return "\n".join(parts)


# ── Audience-Mapping (Phase 2 C.4) ─────────────────────────
# Welche Finding-Kategorie ist für welche Zielgruppe relevant?
# Schlüssel: Category-Value aus models.Category ODER "implicit"/"pii"/"profile"
# als virtuelle Kategorien (per rule_id-Präfix unterschieden).
CATEGORY_TO_AUDIENCES: dict[str, set[str]] = {
    "structure": {"was_tun", "migration"},
    "formula":   {"was_tun", "migration"},
    "volume":    {"was_tun", "migration"},
    "implicit":  {"was_tun", "governance"},
    "pii":       {"governance", "was_tun"},
    "profile":   {"migration"},
}

AUDIENCE_ORDER = ("was_tun", "governance", "migration")
AUDIENCE_LABELS = {
    "was_tun":    ("Was tun?",   "Mitarbeiter — konkrete Fixes"),
    "governance": ("Governance", "Compliance, PII, Audit-Nachweis"),
    "migration":  ("Migration / AI", "Datenbank & KI-Auswertung"),
}


def _finding_virtual_category(f: Finding) -> str:
    """Liefert die virtuelle Audience-Kategorie eines Findings."""
    if f.rule_id.startswith("PII-"):
        return "pii"
    if f.rule_id.startswith("PRF-"):
        return "profile"
    if f.rule_id.startswith("IMP-"):
        return "implicit"
    if f.category == Category.STRUCTURE:
        return "structure"
    if f.category == Category.FORMULA:
        return "formula"
    if f.category == Category.VOLUME:
        return "volume"
    return "structure"


def _audiences_for(f: Finding) -> set[str]:
    return CATEGORY_TO_AUDIENCES.get(_finding_virtual_category(f), {"was_tun"})


# ── Lucide-Inline-Icons (Phase 2 C.8) ──────────────────────
# Minimales Inventar; stroke wird per currentColor gesetzt, damit
# die sev-*-Klassen aus theme.css die Farbe steuern.
_LUCIDE: dict[str, str] = {
    "alert-circle":   '<circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>',
    "alert-triangle": '<path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>',
    "info":           '<circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/>',
    "check-circle":   '<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>',
    "database":       '<ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/>',
    "bar-chart-3":    '<path d="M3 3v18h18"/><path d="M8 17V9"/><path d="M13 17V5"/><path d="M18 17v-3"/>',
    "users":          '<path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/>',
    "shield":         '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>',
    "target":         '<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/>',
    "zap":            '<polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/>',
    "arrow-right":    '<line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/>',
    "copy":           '<rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>',
    "file-spreadsheet": '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><path d="M8 13h8"/><path d="M8 17h8"/>',
}


def _lucide(name: str, cls: str = "") -> str:
    """Rendert ein Lucide-Icon als Inline-SVG."""
    body = _LUCIDE.get(name, _LUCIDE["info"])
    cls_attr = f' class="{cls}"' if cls else ""
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" '
        f'fill="none" stroke="currentColor" stroke-width="2" '
        f'stroke-linecap="round" stroke-linejoin="round"{cls_attr}>{body}</svg>'
    )


_SEVERITY_ICON = {
    Severity.CRITICAL: "alert-triangle",
    Severity.ERROR: "alert-circle",
    Severity.WARNING: "alert-triangle",
    Severity.INFO: "info",
}

_SEVERITY_LABEL_DE = {
    Severity.CRITICAL: "Kritisch",
    Severity.ERROR: "Fehler",
    Severity.WARNING: "Warnung",
    Severity.INFO: "Info",
}


def _format_location(f: Finding) -> str:
    """Erzeugt einen Fundort-String mit klickbarer Cell-Ref (Phase 2 C.6)."""
    parts = []
    if f.sheet:
        parts.append(f'Blatt: <strong>{_esc(f.sheet)}</strong>')
    if f.cell:
        safe_cell = _esc(f.cell)
        parts.append(
            f'Zelle: <code class="cell-ref" data-cell="{safe_cell}" '
            f'onclick="copyCellRef(this)" title="Klicken zum Kopieren">{safe_cell}</code>'
        )
    return ' &middot; '.join(parts)


def _render_finding_card(f: Finding) -> str:
    """Rendert ein Finding als Accordion-Card (Phase 2 C.7)."""
    sev = f.severity
    sev_cls = f"sev sev-{sev.value}"
    sev_label = _SEVERITY_LABEL_DE.get(sev, sev.value.capitalize())
    sev_icon = _lucide(_SEVERITY_ICON.get(sev, "info"))

    title = _humanize(f.rule_id, f.message)
    location = _format_location(f)

    body_parts = [f'<p>{_esc(f.message)}</p>']
    if location:
        body_parts.append(f'<p class="finding-location" style="color:var(--muted); font-size:0.9rem;">{location}</p>')
    if f.detail:
        body_parts.append(f'<p><strong>Details:</strong> {_esc(f.detail)}</p>')
    if f.suggestion:
        body_parts.append(
            f'<p style="margin-top:0.5rem; padding:0.6rem 0.85rem; '
            f'background:var(--bg-info-soft); border-radius:6px;">'
            f'💡 {_esc(f.suggestion)}</p>'
        )
    if f.rule_id:
        body_parts.append(
            f'<p style="font-size:0.75rem; color:var(--muted); margin-top:0.5rem;">Regel: <code>{_esc(f.rule_id)}</code>'
            + (f' &middot; {_esc(f.sample_note)}' if f.sample_note else '')
            + '</p>'
        )

    return (
        f'<details class="finding-card" data-severity="{sev.value}" data-rule-id="{_esc(f.rule_id)}" id="{_finding_anchor(f)}">\n'
        f'  <summary class="finding-summary">'
        f'<span class="{sev_cls}">{sev_icon} {_esc(sev_label)}</span>'
        f'<span>{_esc(title)}</span>'
        f'</summary>\n'
        f'  <div class="finding-body">{"".join(body_parts)}</div>\n'
        f'</details>'
    )


def _severity_filter_pills() -> str:
    """Severity-Filter als Pills über der Accordion-Liste (Phase 2 C.7)."""
    pills = [
        ('all',      'Alle',     'info'),
        ('critical', 'Kritisch', 'alert-triangle'),
        ('error',    'Fehler',   'alert-circle'),
        ('warning',  'Warnung',  'alert-triangle'),
        ('info',     'Info',     'info'),
    ]
    html_pills = []
    for sev, label, icon in pills:
        pressed = 'true' if sev == 'all' else 'false'
        html_pills.append(
            f'<button class="severity-pill" data-severity="{sev}" '
            f'aria-pressed="{pressed}" onclick="filterBySeverity(this)">'
            f'{_lucide(icon)}<span>{_esc(label)}</span></button>'
        )
    return '<div class="severity-pills">' + ''.join(html_pills) + '</div>'


def _html_top_actions(findings: list, limit: int = 3) -> str:
    """Top-Actions mit Score-Impact-Teaser (Phase 2 C.5).

    Sortiert Findings nach score_penalty, dedupliziert per rule_id,
    nimmt die Top-N und rendert sie als Cards.
    """
    sev_order = {Severity.CRITICAL: 0, Severity.ERROR: 1, Severity.WARNING: 2, Severity.INFO: 3}
    seen: set[str] = set()
    impact_per_rule: dict[str, int] = defaultdict(int)
    first_per_rule: dict[str, Finding] = {}
    for f in findings:
        impact_per_rule[f.rule_id] += f.score_penalty
        if f.rule_id not in first_per_rule:
            first_per_rule[f.rule_id] = f

    ranked = sorted(
        first_per_rule.values(),
        key=lambda x: (-impact_per_rule[x.rule_id], sev_order.get(x.severity, 9))
    )

    cards = []
    for f in ranked:
        if f.rule_id in seen:
            continue
        seen.add(f.rule_id)
        if len(cards) >= limit:
            break
        impact = impact_per_rule[f.rule_id]
        if impact <= 0:
            continue
        title = _humanize(f.rule_id, f.message)
        target_id = _finding_anchor(f)
        teaser = f"Fix → +{impact} Punkte Datenqualität" if impact > 0 else "Kein direkter Score-Impact"
        cards.append(
            f'<div class="top-action-card" data-severity="{f.severity.value}">'
            f'  <div class="top-action-title">{_lucide(_SEVERITY_ICON.get(f.severity, "info"))} {_esc(title)}</div>'
            f'  <div class="top-action-impact">{_esc(teaser)}</div>'
            f'  <a class="top-action-link" href="#{target_id}" onclick="scrollToFinding(\'{target_id}\')">Betroffene Zellen zeigen {_lucide("arrow-right")}</a>'
            f'</div>'
        )

    if not cards:
        return ''
    return (
        '<div class="audience-panel-section" style="background:transparent; border:none; box-shadow:none; padding:0;">'
        f'<h3>{_lucide("target")} Top-{len(cards)} Actions</h3>'
        f'<div class="top-actions">{"".join(cards)}</div>'
        '</div>'
    )


def _audience_accordion(findings: list, audience: str) -> str:
    """Rendert alle Findings der gegebenen Audience als sortierte Accordion-Liste."""
    relevant = [f for f in findings if audience in _audiences_for(f)]
    if not relevant:
        return '<p class="panel-placeholder">Keine für diese Zielgruppe relevanten Funde.</p>'
    sev_order = {Severity.CRITICAL: 0, Severity.ERROR: 1, Severity.WARNING: 2, Severity.INFO: 3}
    relevant.sort(key=lambda x: (sev_order.get(x.severity, 9), -x.score_penalty))
    cards = [_render_finding_card(f) for f in relevant]
    return (
        _severity_filter_pills()
        + f'<div class="finding-list" data-audience="{audience}">'
        + "\n".join(cards)
        + '</div>'
    )


def _html_tab_was_tun(findings: list) -> str:
    top = _html_top_actions([f for f in findings if 'was_tun' in _audiences_for(f)])
    body = _audience_accordion(findings, 'was_tun')
    return (
        f'<div class="audience-panel" data-audience="was_tun" data-active="true">'
        f'{top}'
        f'<div class="audience-panel-section">'
        f'<h3>{_lucide("file-spreadsheet")} Alle Funde für Mitarbeiter</h3>'
        f'{body}'
        f'</div>'
        f'</div>'
    )


def _html_pii_panel(findings: list) -> str:
    """Rendert ein Panel mit PII-Findings, gruppiert nach Typ (Phase 2 A.4)."""
    pii = [f for f in findings if f.rule_id.startswith("PII-")]
    if not pii:
        return (
            '<div class="panel-placeholder">'
            'Keine Verdachtstreffer auf E-Mail, IBAN, Sozialversicherungs- '
            'oder Telefonnummer in den geprüften Zellen.'
            '</div>'
        )
    from collections import defaultdict as _dd
    by_rule: dict[str, list[Finding]] = _dd(list)
    for f in pii:
        by_rule[f.rule_id].append(f)
    rows = []
    for rid, items in sorted(by_rule.items()):
        kind = items[0].rule_name
        cells = ", ".join(
            f'<code class="cell-ref" data-cell="{_esc(f.cell or "")}" onclick="copyCellRef(this)">{_esc((f.sheet or "") + "!" + (f.cell or ""))}</code>'
            for f in items[:5] if f.cell
        )
        if len(items) > 5:
            cells += f' <span class="muted">+{len(items) - 5} weitere</span>'
        rows.append(
            f'<li style="padding:0.6rem 0; border-bottom:1px solid var(--border);">'
            f'<strong>{_esc(kind)}</strong> <span class="muted">({len(items)}×)</span><br>'
            f'{cells}'
            f'</li>'
        )
    return (
        '<ul style="list-style:none; padding:0; margin:0;">' + ''.join(rows) + '</ul>'
        + '<p class="muted" style="font-size:0.8rem; margin-top:0.6rem;">'
        + 'Sensible Daten gehören nicht in frei geteilte Excel-Dateien. '
        + 'Entfernen, pseudonymisieren oder in ein geschütztes System verlegen.'
        + '</p>'
    )


def _html_tab_governance(findings: list, compliance_score: int,
                         compliance_risks: list[tuple[str, str, str]]) -> str:
    risk_list = ''.join(
        f'<li style="padding:0.4rem 0; border-bottom:1px solid var(--border);">'
        f'<span style="font-size:1.1rem; margin-right:0.4rem;">{icon}</span>'
        f'<strong>{_esc(title)}</strong> — <span style="color:var(--muted);">{_esc(desc)}</span>'
        f'</li>'
        for icon, title, desc in compliance_risks
    ) or '<li class="panel-placeholder" style="list-style:none;">Keine Governance-Risiken erkannt.</li>'

    pii_html = _html_pii_panel(findings)
    return (
        f'<div class="audience-panel" data-audience="governance">'
        f'<div class="audience-panel-section">'
        f'<h3>{_lucide("shield")} Compliance-Risiko</h3>'
        f'<p class="muted">Score: <strong>{compliance_score}/100</strong> — höher bedeutet weniger Risiko.</p>'
        f'<ul style="list-style:none; padding:0; margin:0.5rem 0;">{risk_list}</ul>'
        f'</div>'
        f'<div class="audience-panel-section">'
        f'<h3>{_lucide("users")} Sensible Daten (PII)</h3>'
        f'{pii_html}'
        f'</div>'
        f'<div class="audience-panel-section">'
        f'<h3>{_lucide("file-spreadsheet")} Governance-relevante Funde</h3>'
        f'{_audience_accordion(findings, "governance")}'
        f'</div>'
        f'</div>'
    )


def _schema_hint(profile: ColumnProfile) -> str:
    """Kurzer Hinweis, wie die Spalte in einer DB aussehen könnte."""
    if profile.total == 0:
        return "—"
    if profile.unique_count >= max(5, int(profile.total * 0.95)) and profile.null_ratio == 0:
        return "Kandidat für PRIMARY KEY"
    if profile.unique_count <= 10 and profile.total >= 20:
        return "Kandidat für ENUM / Dropdown"
    if profile.null_ratio > 0.3:
        return "Hoher Null-Anteil — optional in DB"
    return "—"


def _html_column_profile_table(profiles: list[ColumnProfile]) -> str:
    """Rendert die Column-Profile als Tabelle (Phase 2 A.4)."""
    if not profiles:
        return (
            '<div class="panel-placeholder">'
            'Keine Column-Profile verfügbar (Sheet zu klein oder Sampling aktiv).'
            '</div>'
        )
    rows = []
    for p in profiles:
        types = ", ".join(f"{k} {v}" for k, v in sorted(p.type_mix.items(), key=lambda kv: -kv[1]) if k != "null")
        rows.append(
            f'<tr>'
            f'<td>{_esc(p.sheet)}</td>'
            f'<td><code>{_esc(p.column_letter)}</code></td>'
            f'<td>{_esc(p.header or "—")}</td>'
            f'<td class="num">{p.total}</td>'
            f'<td class="num">{p.null_ratio:.0%}</td>'
            f'<td class="num">{p.unique_count}</td>'
            f'<td>{_esc(p.dominant_type or "—")}</td>'
            f'<td><small class="muted">{_esc(types)}</small></td>'
            f'<td>{_esc(_schema_hint(p))}</td>'
            f'</tr>'
        )
    return (
        '<div style="overflow-x:auto;">'
        '<table class="sheet-table">'
        '<thead><tr>'
        '<th>Blatt</th><th>Spalte</th><th>Header</th>'
        '<th class="num">Zeilen</th><th class="num">Null-%</th><th class="num">Unique</th>'
        '<th>Typ</th><th>Typ-Mix</th><th>Schema-Hint</th>'
        '</tr></thead><tbody>'
        + "".join(rows)
        + '</tbody></table></div>'
    )


def _html_tab_migration(findings: list, report: WorkbookReport, avg_db: int) -> str:
    avg_db_color = _db_readiness_color(avg_db)
    avg_db_label = _db_readiness_label(avg_db)
    profile_html = _html_column_profile_table(report.column_profiles)
    return (
        f'<div class="audience-panel" data-audience="migration">'
        f'<div class="audience-panel-section">'
        f'<h3>{_lucide("database")} DB-Readiness (Durchschnitt)</h3>'
        f'<p><strong style="color:{avg_db_color}; font-size:1.8rem;">{avg_db}</strong> / 100 &middot; '
        f'<span class="muted">{_esc(avg_db_label)}</span></p>'
        f'</div>'
        f'<div class="audience-panel-section">'
        f'<h3>{_lucide("bar-chart-3")} Column-Profile</h3>'
        f'{profile_html}'
        f'</div>'
        f'<div class="audience-panel-section">'
        f'<h3>{_lucide("file-spreadsheet")} Migration-relevante Funde</h3>'
        f'{_audience_accordion(findings, "migration")}'
        f'</div>'
        f'</div>'
    )


def _html_audience_tabs(findings: list, compliance_score: int,
                        compliance_risks: list[tuple[str, str, str]],
                        report: WorkbookReport, avg_db: int) -> str:
    """Ersetzt die alten Kategorie-Tabs durch drei Audience-Tabs."""
    # Tab-Buttons
    tab_btns = []
    for i, aud in enumerate(AUDIENCE_ORDER):
        label, hint = AUDIENCE_LABELS[aud]
        count = sum(1 for f in findings if aud in _audiences_for(f))
        selected = 'true' if i == 0 else 'false'
        tab_btns.append(
            f'<button class="audience-tab" role="tab" '
            f'aria-selected="{selected}" data-audience-tab="{aud}" '
            f'onclick="switchAudience(this)">'
            f'<span>{_esc(label)} ({count})</span>'
            f'<span class="audience-tab-hint">{_esc(hint)}</span>'
            f'</button>'
        )

    panels = [
        _html_tab_was_tun(findings),
        _html_tab_governance(findings, compliance_score, compliance_risks),
        _html_tab_migration(findings, report, avg_db),
    ]
    return (
        '<div class="card">'
        f'<h2>{_lucide("target")} Detaillierte Ergebnisse</h2>'
        '<div class="audience-tabs" role="tablist">'
        + "".join(tab_btns) +
        '</div>'
        + "".join(panels) +
        '</div>'
    )


def generate_llm_section_html(analysis) -> str:
    """Generiert HTML für die LLM-Analyse-Ergebnisse (wird per AJAX eingebaut)."""
    parts = []

    # Summary
    if analysis.summary:
        parts.append(f"""
<div style="padding:1rem 1.25rem;background:var(--bg-info-soft);
     border:1px solid var(--border);border-radius:10px;margin-bottom:1.5rem;">
  <div style="font-size:1.1rem;font-weight:600;margin-bottom:0.5rem;">💬 KI-Zusammenfassung</div>
  <div style="font-size:0.95rem;color:var(--text);line-height:1.6;">{_esc(analysis.summary)}</div>
</div>""")

    # Optimization table
    if analysis.optimization_table:
        parts.append("""
<div style="margin-bottom:1.5rem;">
  <h3 style="font-size:1rem;margin-bottom:0.75rem;">📋 Optimierungsvorschläge</h3>
  <div style="overflow-x:auto;">
  <table style="width:100%;border-collapse:collapse;font-size:0.85rem;">
    <thead>
      <tr style="background:#f1f5f9;text-align:left;">
        <th style="padding:0.6rem 0.75rem;border-bottom:2px solid #e2e8f0;">Bereich</th>
        <th style="padding:0.6rem 0.75rem;border-bottom:2px solid #e2e8f0;">Problem</th>
        <th style="padding:0.6rem 0.75rem;border-bottom:2px solid #e2e8f0;">Vorschlag</th>
        <th style="padding:0.6rem 0.75rem;border-bottom:2px solid #e2e8f0;text-align:center;">Aufwand</th>
        <th style="padding:0.6rem 0.75rem;border-bottom:2px solid #e2e8f0;text-align:center;">Nutzen</th>
      </tr>
    </thead>
    <tbody>""")
        for row in analysis.optimization_table:
            aufwand = _esc(row.get("aufwand", ""))
            nutzen = _esc(row.get("nutzen", ""))
            a_color = {"gering": "#16a34a", "mittel": "#ca8a04", "hoch": "#dc2626"}.get(aufwand.lower(), "#64748b")
            n_color = {"gering": "#dc2626", "mittel": "#ca8a04", "hoch": "#16a34a"}.get(nutzen.lower(), "#64748b")
            parts.append(f"""
      <tr style="border-bottom:1px solid #e2e8f0;">
        <td style="padding:0.6rem 0.75rem;font-weight:600;">{_esc(row.get("bereich", ""))}</td>
        <td style="padding:0.6rem 0.75rem;">{_esc(row.get("problem", ""))}</td>
        <td style="padding:0.6rem 0.75rem;">{_esc(row.get("vorschlag", ""))}</td>
        <td style="padding:0.6rem 0.75rem;text-align:center;">
          <span style="padding:0.15rem 0.5rem;border-radius:4px;font-size:0.75rem;font-weight:600;color:white;background:{a_color};">{aufwand}</span>
        </td>
        <td style="padding:0.6rem 0.75rem;text-align:center;">
          <span style="padding:0.15rem 0.5rem;border-radius:4px;font-size:0.75rem;font-weight:600;color:white;background:{n_color};">{nutzen}</span>
        </td>
      </tr>""")
        parts.append("</tbody></table></div></div>")

    # Formula rewrites
    if analysis.formula_rewrites:
        parts.append("""
<div style="margin-bottom:1.5rem;">
  <h3 style="font-size:1rem;margin-bottom:0.75rem;">🔄 Formel-Verbesserungen</h3>""")
        for fw in analysis.formula_rewrites:
            parts.append(f"""
  <div style="padding:1rem;background:#fefce8;border:1px solid #fde68a;border-radius:8px;margin-bottom:0.75rem;">
    <div style="font-weight:600;color:#92400e;margin-bottom:0.25rem;">Aktuell: {_esc(fw.get("original", ""))}</div>
    <div style="font-weight:600;color:#16a34a;margin-bottom:0.25rem;">→ Besser: {_esc(fw.get("vorschlag", ""))}</div>
    <div style="font-size:0.85rem;color:#64748b;">{_esc(fw.get("erklaerung", ""))}</div>
  </div>""")
        parts.append("</div>")

    # Architecture advice
    if analysis.architecture_advice:
        parts.append(f"""
<div style="padding:1rem 1.25rem;background:var(--bg-warning-soft);
     border:1px solid var(--border);border-radius:10px;margin-bottom:1.5rem;">
  <div style="font-size:1rem;font-weight:600;margin-bottom:0.5rem;">🏗️ Architektur-Empfehlung</div>
  <div style="font-size:0.9rem;color:var(--text);line-height:1.6;">{_esc(analysis.architecture_advice)}</div>
</div>""")

    # Per-Sheet Assessment
    if analysis.per_sheet_assessment:
        parts.append("""
<div style="margin-bottom:1.5rem;">
  <h3 style="font-size:1rem;margin-bottom:0.75rem;">🗄️ KI-Bewertung pro Sheet</h3>""")
        for sa in analysis.per_sheet_assessment:
            prio = sa.get("priority", "").lower()
            prio_color = {"hoch": "#dc2626", "mittel": "#ca8a04", "gering": "#16a34a"}.get(prio, "#64748b")
            parts.append(f"""
  <div style="padding:1rem;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;margin-bottom:0.75rem;">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:0.25rem;">
      <span style="font-weight:600;font-size:0.95rem;">{_esc(sa.get("sheet", ""))}</span>
      <span style="padding:0.15rem 0.5rem;border-radius:4px;font-size:0.75rem;font-weight:600;color:white;background:{prio_color};">
        Priorität {_esc(sa.get("priority", "–"))}</span>
    </div>
    <div style="font-size:0.85rem;color:#334155;line-height:1.5;">{_esc(sa.get("db_verdict", ""))}</div>
  </div>""")
        parts.append("</div>")

    return "\n".join(parts)


def _html_footer() -> str:
    return """
<div class="card" id="llm-section">
  <h2>🤖 KI-Analyse <span style="font-size:0.75rem;color:var(--muted);font-weight:400;">(experimentell)</span></h2>
  <p style="color:var(--muted);font-size:0.9rem;margin-bottom:1rem;">
    Lass dir von einer KI konkrete, freundliche Verbesserungsvorschläge für diese Datei geben.
    Die KI analysiert die Ergebnisse des klassischen Checks.
  </p>
  <div id="llm-config-status" style="margin-bottom:1rem;padding:0.75rem 1rem;border-radius:8px;font-size:0.85rem;"></div>
  <div style="text-align:center;">
    <button onclick="startLlmAnalysis()" id="llmStartBtn" class="action-btn action-download" style="font-size:1rem;">
      🤖 KI-Analyse starten
    </button>
  </div>
  <div id="llm-loading" style="display:none;text-align:center;padding:2rem;">
    <div style="display:inline-block;width:28px;height:28px;border:3px solid var(--border);border-top-color:var(--accent);border-radius:50%;animation:spin 0.6s linear infinite;"></div>
    <p style="margin-top:0.75rem;color:var(--muted);">KI analysiert die Ergebnisse… dauert ca. 10-30 Sekunden</p>
  </div>
  <div id="llm-result"></div>
  <div id="llm-error" style="display:none;padding:0.75rem 1rem;background:#fef2f2;border:1px solid #fecaca;border-radius:8px;color:#dc2626;margin-top:1rem;"></div>
</div>

<div class="report-actions">
  <button onclick="downloadReport()" class="action-btn action-download">
    💾 Report speichern
  </button>
  <a href="/" class="action-btn action-new">
    🔄 Neue Analyse starten
  </a>
</div>

<div class="footer-cards">
  <a href="/learn" class="footer-card">
    <div class="fc-icon">📚</div>
    <div class="fc-title">Wissen aufbauen</div>
    <div class="fc-desc">Warum gehört komplexes Excel in eine Datenbank? Lerne die wichtigsten Konzepte.</div>
  </a>
  <a href="/learn#effizienz" class="footer-card">
    <div class="fc-icon">💰</div>
    <div class="fc-title">Effizienz-Analyse</div>
    <div class="fc-desc">Was kosten aufgeblähte Excel-Dateien wirklich? Kostenrechnung &amp; KI-Readiness.</div>
  </a>
  <a href="mailto:UCS@apg.at" class="footer-card">
    <div class="fc-icon">💬</div>
    <div class="fc-title">Unterstützung anfragen</div>
    <div class="fc-desc">Fragen zur Analyse oder Hilfe bei der Datenbank-Migration? Das UCS-Team hilft gerne!</div>
  </a>
</div>

<div class="footer">
  Excel-Reifecheck · """ + _get_version_info() + """ · Report erstellt am """ + datetime.now().strftime("%d.%m.%Y %H:%M:%S") + """<br>
  UCS - Digital Workplace Solutions
</div>
</div>

<script>
// ── Dark/Light Mode (Track B Phase 1) ──
// Initiale Synchronisierung (Inhalt im Head-Skript, hier nur Button-Icon).
(function() {
  var btn = document.getElementById('themeToggleBtn');
  if (btn && document.documentElement.getAttribute('data-theme') === 'dark') {
    btn.textContent = '☀️';
  }
})();
function toggleTheme() {
  var html = document.documentElement;
  var btn = document.getElementById('themeToggleBtn');
  if (html.getAttribute('data-theme') === 'dark') {
    html.removeAttribute('data-theme');
    if (btn) btn.textContent = '🌙';
    localStorage.setItem('theme', 'light');
  } else {
    html.setAttribute('data-theme', 'dark');
    if (btn) btn.textContent = '☀️';
    localStorage.setItem('theme', 'dark');
  }
}

function downloadReport() {
  const html = document.documentElement.outerHTML;
  const blob = new Blob(['<!DOCTYPE html>' + html], {type: 'text/html;charset=utf-8'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = (document.title || 'Excel-Health-Check-Report') + '.html';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function switchTab(group, tabName) {
  document.querySelectorAll('#' + group + '-tabs .tab-btn').forEach(btn => {
    btn.classList.remove('active');
  });
  document.querySelectorAll('[data-tab-group="' + group + '"]').forEach(c => {
    c.classList.remove('active');
  });
  event.target.classList.add('active');
  document.querySelector('[data-tab-group="' + group + '"][data-tab="' + tabName + '"]').classList.add('active');
}

function switchHeroTab(btn, panelId) {
  btn.parentElement.querySelectorAll('.hero-tab-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  btn.closest('.card').querySelectorAll('.hero-tab-panel').forEach(p => p.classList.remove('active'));
  document.getElementById(panelId).classList.add('active');
}

// ── Audience-Tabs (Phase 2 C.4) ────────────────────────────
function switchAudience(btn) {
  var card = btn.closest('.card');
  if (!card) return;
  var target = btn.getAttribute('data-audience-tab');
  card.querySelectorAll('.audience-tab').forEach(function(b) {
    b.setAttribute('aria-selected', b === btn ? 'true' : 'false');
  });
  card.querySelectorAll('.audience-panel').forEach(function(p) {
    p.setAttribute('data-active', p.getAttribute('data-audience') === target ? 'true' : 'false');
  });
}

// ── Severity-Filter-Pills (Phase 2 C.7) ───────────────────
function filterBySeverity(btn) {
  var panel = btn.closest('.audience-panel');
  if (!panel) return;
  var sev = btn.getAttribute('data-severity');
  panel.querySelectorAll('.severity-pill').forEach(function(p) {
    p.setAttribute('aria-pressed', p === btn ? 'true' : 'false');
  });
  panel.querySelectorAll('.finding-card').forEach(function(card) {
    var cardSev = card.getAttribute('data-severity');
    var show = (sev === 'all') || (cardSev === sev);
    // error + critical zusammenfassen, wenn eine kritische Pille gewählt wird:
    if (sev === 'error' && cardSev === 'critical') show = true;
    card.setAttribute('data-hidden', show ? 'false' : 'true');
  });
}

// ── Clipboard-Copy für Cell-Refs (Phase 2 C.6) ────────────
function copyCellRef(el) {
  var ref = el.getAttribute('data-cell') || el.textContent;
  var done = function() { showToast('Kopiert: ' + ref); };
  var fail = function() { showToast('Kopieren fehlgeschlagen'); };
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(ref).then(done).catch(fail);
  } else {
    // Fallback via temp textarea
    try {
      var ta = document.createElement('textarea');
      ta.value = ref;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
      done();
    } catch(e) { fail(); }
  }
}

function showToast(msg) {
  var toast = document.getElementById('report-toast');
  if (!toast) {
    toast = document.createElement('div');
    toast.className = 'toast';
    toast.id = 'report-toast';
    document.body.appendChild(toast);
  }
  toast.textContent = msg;
  toast.classList.add('is-visible');
  clearTimeout(showToast._t);
  showToast._t = setTimeout(function() {
    toast.classList.remove('is-visible');
  }, 1500);
}

function scrollToFinding(id) {
  var el = document.getElementById(id);
  if (!el) return;
  el.open = true;
  el.scrollIntoView({behavior: 'smooth', block: 'center'});
  el.style.transition = 'box-shadow 0.4s';
  var prev = el.style.boxShadow;
  el.style.boxShadow = '0 0 0 3px var(--accent)';
  setTimeout(function() { el.style.boxShadow = prev; }, 1200);
}

// ── LLM Integration ──
(function() {
  // Check if LLM is configured (from localStorage or .env)
  const savedKey = localStorage.getItem('llm_api_key') || '';
  const savedEndpoint = localStorage.getItem('llm_endpoint') || '';
  const statusEl = document.getElementById('llm-config-status');

  fetch('/api/llm-config').then(r => r.json()).then(cfg => {
    const hasKey = savedKey || cfg.has_key;
    const hasEndpoint = savedEndpoint || cfg.has_endpoint;
    if (hasKey && hasEndpoint) {
      statusEl.style.background = '#f0fdf4';
      statusEl.style.color = '#15803d';
      statusEl.style.border = '1px solid #bbf7d0';
      statusEl.innerHTML = '✅ KI-Verbindung konfiguriert. ' +
        (savedEndpoint ? 'Endpoint: <code>' + savedEndpoint.substring(0, 40) + '…</code>' : 'Endpoint aus .env') +
        ' · <a href="/" style="color:#15803d;">⚙️ Einstellungen ändern</a>';
    } else if (hasKey && !hasEndpoint) {
      statusEl.style.background = '#fefce8';
      statusEl.style.color = '#854d0e';
      statusEl.style.border = '1px solid #fde68a';
      statusEl.innerHTML = '⚠️ API-Key vorhanden, aber <strong>kein Endpoint</strong> konfiguriert. ' +
        'Bitte unter <a href="/" style="color:#854d0e;font-weight:600;">⚙️ KI-Einstellungen</a> die Azure AI Endpoint-URL eingeben.';
    } else {
      statusEl.style.background = '#eff6ff';
      statusEl.style.color = '#1e40af';
      statusEl.style.border = '1px solid #bfdbfe';
      statusEl.innerHTML = '💡 Noch nicht konfiguriert. ' +
        'Gehe zu <a href="/" style="color:#1e40af;font-weight:600;">⚙️ KI-Einstellungen</a> auf der Startseite, ' +
        'um Endpoint und API-Key einzurichten.';
    }
  }).catch(() => {
    statusEl.style.background = '#eff6ff';
    statusEl.style.color = '#1e40af';
    statusEl.style.border = '1px solid #bfdbfe';
    statusEl.textContent = '💡 KI-Einstellungen auf der Startseite unter ⚙️ konfigurieren.';
  });
})();

function getReportId() {
  const m = window.location.pathname.match(/\\/report\\/([a-f0-9]+)/);
  return m ? m[1] : null;
}

function startLlmAnalysis() {
  const reportId = getReportId();
  if (!reportId) { alert('Report-ID nicht gefunden.'); return; }

  const btn = document.getElementById('llmStartBtn');
  const loading = document.getElementById('llm-loading');
  const result = document.getElementById('llm-result');
  const error = document.getElementById('llm-error');

  btn.style.display = 'none';
  loading.style.display = 'block';
  result.innerHTML = '';
  error.style.display = 'none';

  const body = {};
  const savedKey = localStorage.getItem('llm_api_key');
  const savedEndpoint = localStorage.getItem('llm_endpoint');
  const savedModel = localStorage.getItem('llm_model');
  if (savedKey) body.api_key = savedKey;
  if (savedEndpoint) body.endpoint = savedEndpoint;
  if (savedModel) body.model = savedModel;

  // Eingebetteten Report-Kontext mitsenden (Fallback wenn Server-State verloren)
  const ctxEl = document.getElementById('report-context');
  if (ctxEl) {
    try { body.context = JSON.parse(ctxEl.textContent); } catch(e) {}
  }

  fetch('/api/llm-analyze/' + reportId, {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(body),
  }).then(r => r.json()).then(data => {
    loading.style.display = 'none';
    if (data.error) {
      error.textContent = data.error;
      error.style.display = 'block';
      btn.style.display = '';
    } else {
      result.innerHTML = data.html;
      document.getElementById('llm-config-status').style.display = 'none';
    }
  }).catch(err => {
    loading.style.display = 'none';
    error.textContent = 'Netzwerkfehler: ' + err.message;
    error.style.display = 'block';
    btn.style.display = '';
  });
}
</script>
</body>
</html>"""
