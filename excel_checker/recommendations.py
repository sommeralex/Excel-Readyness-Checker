"""Empfehlungs-Modul – leitet strategische Empfehlungen aus Ergebnissen ab."""

from __future__ import annotations

from typing import List

from excel_checker.models import (
    Category, Finding, Recommendation, RecommendationType,
    Severity, WorkbookReport,
)


def generate_recommendations(report: WorkbookReport) -> List[Recommendation]:
    """Analysiert den Bericht und generiert strategische Empfehlungen."""
    recs = []

    # Datenmenge analysieren
    max_rows = max((s.row_count for s in report.sheet_stats), default=0)
    total_formulas = sum(s.formula_count for s in report.sheet_stats)
    sheet_count = report.sheet_count

    # Schweregrad-Zähler
    vol_findings = [f for f in report.findings if f.category == Category.VOLUME]
    struct_findings = [f for f in report.findings if f.category == Category.STRUCTURE]
    formula_findings = [f for f in report.findings if f.category == Category.FORMULA]

    has_external_refs = any(
        f.rule_id == "FRM-005" and "Externe" in f.message for f in report.findings
    )
    has_vlookup_chains = any(f.rule_id == "FRM-003" for f in report.findings)
    has_hidden_sheets = any(f.rule_id == "IMP-002" for f in report.findings)
    has_color_codes = any(f.rule_id == "IMP-001" for f in report.findings)
    has_conditional_fmt = any(f.rule_id == "IMP-004" for f in report.findings)

    # === DATENBANK-MIGRATION ===
    db_triggers = 0
    db_reasons = []

    if max_rows > 50000:
        db_triggers += 3
        db_reasons.append(f"{max_rows:,} Zeilen übersteigen den sinnvollen Excel-Bereich")
    elif max_rows > 10000:
        db_triggers += 1
        db_reasons.append(f"{max_rows:,} Zeilen nähern sich dem Bereich, wo Datenbanken effizienter sind")

    if has_vlookup_chains:
        db_triggers += 2
        db_reasons.append("Exzessive Lookup-Funktionen simulieren Datenbank-JOINs")

    if report.file_size_mb > 20:
        db_triggers += 2
        db_reasons.append(f"Dateigröße ({report.file_size_mb:.1f} MB) verursacht Leistungsprobleme")

    if sheet_count > 15:
        db_triggers += 1
        db_reasons.append(f"{sheet_count} Blätter bilden ein relationales Datenmodell ab")

    if has_external_refs:
        db_triggers += 1
        db_reasons.append("Externe Dateibezüge zeigen verteilte, zusammengehörige Daten")

    if db_triggers >= 3:
        recs.append(Recommendation(
            rec_type=RecommendationType.DATABASE_MIGRATION,
            priority=1,
            title="Diese Daten sind für eine Datenbank-Lösung prädestiniert",
            reason=" | ".join(db_reasons),
            action=(
                "Empfohlener nächster Schritt: Einen Termin mit dem Digital Workplace Team "
                "vereinbaren. Wir analysieren gemeinsam, welche Datenbank-Lösung am besten "
                "passt (SharePoint-Liste, Dataverse, SQL Server) und begleiten die Migration. "
                "Die bestehenden Excel-Auswertungen können als Power BI Dashboard "
                "weitergeführt werden."
            ),
        ))

    # === POWER BI / REPORTING ===
    if total_formulas > 5000 or has_conditional_fmt:
        recs.append(Recommendation(
            rec_type=RecommendationType.POWER_BI,
            priority=2 if db_triggers < 3 else 3,
            title="Berechnungen und Visualisierungen in ein Reporting-Tool auslagern",
            reason=(
                f"{total_formulas:,} Formeln und/oder komplexe bedingte Formatierungen "
                f"deuten auf aufwändige Auswertungslogik hin, die in Power BI oder "
                f"einem vergleichbaren Tool stabiler und flexibler abgebildet werden kann."
            ),
            action=(
                "Power BI kann direkt auf die Excel-Daten zugreifen und die "
                "Auswertungen dynamisch darstellen. Vorteil: Die Berechnungslogik "
                "wird dokumentiert, ist für andere nachvollziehbar und kann ohne "
                "Excel-Expertenwissen angepasst werden."
            ),
        ))

    # === SHAREPOINT-LISTE ===
    if (5 <= max_rows <= 50000 and not has_vlookup_chains
            and db_triggers < 3 and sheet_count <= 5):
        recs.append(Recommendation(
            rec_type=RecommendationType.SHAREPOINT_LIST,
            priority=2,
            title="SharePoint-Liste als niederschwellige Alternative",
            reason=(
                f"Die Datenstruktur ({max_rows:,} Zeilen, {sheet_count} Blätter) "
                f"eignet sich gut für eine SharePoint-Liste – einfach zu erstellen, "
                f"mit Versionshistorie und Berechtigungssteuerung."
            ),
            action=(
                "Eine SharePoint-Liste kann in wenigen Minuten aus der Excel-Tabelle "
                "erstellt werden (Excel → 'Als Tabelle exportieren'). "
                "Vorteile: Gleichzeitiger Zugriff, Änderungshistorie, automatische "
                "Benachrichtigungen und Integration mit Power Automate."
            ),
        ))

    # === DATEN-NORMALISIERUNG ===
    normalization_triggers = sum(1 for f in struct_findings
                                 if f.rule_id in ("STR-001", "STR-002", "STR-003", "STR-004"))
    if normalization_triggers >= 2:
        recs.append(Recommendation(
            rec_type=RecommendationType.NORMALIZATION,
            priority=2 if db_triggers < 3 else 4,
            title="Datenstruktur bereinigen – Grundlage für alle weiteren Schritte",
            reason=(
                f"{normalization_triggers} Strukturprobleme erkannt (verbundene Zellen, "
                f"gemischte Datentypen, fehlende Kopfzeilen, Leerzeilen als Trenner). "
                f"Diese verhindern jede Form der automatisierten Verarbeitung."
            ),
            action=(
                "Erster Schritt: Die Daten in ein tabellarisches Format bringen "
                "(eine Kopfzeile, keine verbundenen Zellen, keine Leerzeilen). "
                "Das Digital Workplace Team kann bei der Bereinigung unterstützen – "
                "oft reicht ein gemeinsamer Workshop-Termin."
            ),
        ))

    # === WORKBOOK AUFTEILEN ===
    if sheet_count > 10 and db_triggers < 3:
        recs.append(Recommendation(
            rec_type=RecommendationType.SPLIT_WORKBOOK,
            priority=3,
            title="Arbeitsmappe aufteilen",
            reason=(
                f"{sheet_count} Blätter in einer Datei sind schwer zu überblicken. "
                f"Oft enthalten verschiedene Blätter verschiedene Themen, die "
                f"besser getrennt verwaltet werden."
            ),
            action=(
                "Prüfen, welche Sheets thematisch zusammengehören und welche "
                "eigenständige Datensätze sind. Eigenständige Bereiche in "
                "separate Dateien oder SharePoint-Listen auslagern."
            ),
        ))

    # === DATENBEREINIGUNG (implizites Wissen) ===
    implicit_count = sum(
        1 for f in report.findings
        if f.rule_id.startswith("IMP-") and f.severity in (Severity.WARNING, Severity.ERROR)
    )
    if implicit_count >= 3:
        recs.append(Recommendation(
            rec_type=RecommendationType.CLEANUP,
            priority=2,
            title="Implizites Wissen dokumentieren und explizit machen",
            reason=(
                f"{implicit_count} Fälle von undokumentiertem Wissen gefunden "
                f"(Farbcodes ohne Legende, versteckte Bereiche, Logik in Kommentaren). "
                f"Dieses Wissen geht verloren, wenn der Ersteller nicht verfügbar ist."
            ),
            action=(
                "Dringend empfohlen: Einen 30-minütigen Termin mit dem Ersteller "
                "vereinbaren, um die impliziten Regeln zu dokumentieren. "
                "Farbcodes → Status-Spalte, Kommentar-Regeln → Datenvalidierung, "
                "Versteckte Blätter → Dokumentation oder Löschung."
            ),
        ))

    # === ALLES OK ===
    if not recs:
        if report.health_score >= 80:
            recs.append(Recommendation(
                rec_type=RecommendationType.OK,
                priority=5,
                title="Gute Arbeit – die Excel-Nutzung ist hier vertretbar",
                reason=(
                    f"Bewertung {report.health_score}/100. "
                    f"Die Datei ist gut strukturiert und im sinnvollen Rahmen für Excel."
                ),
                action=(
                    "Keine dringenden Maßnahmen nötig. "
                    "Für die Zukunft: Regelmäßig prüfen, ob die Datenmenge wächst."
                ),
            ))
        else:
            recs.append(Recommendation(
                rec_type=RecommendationType.CLEANUP,
                priority=3,
                title="Kleinere Optimierungen empfohlen",
                reason=f"Bewertung {report.health_score}/100.",
                action=(
                    "Die einzelnen Hinweise im Bericht beachten – "
                    "die meisten lassen sich mit wenig Aufwand umsetzen."
                ),
            ))

    # Sortiere nach Priorität
    recs.sort(key=lambda r: r.priority)
    return recs
