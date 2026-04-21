"""LLM-Analyse-Modul – Nutzt Claude (Azure AI / Anthropic) für intelligente Excel-Verbesserungsvorschläge."""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field
from typing import Optional

import requests

from excel_checker.models import WorkbookReport


@dataclass
class LLMAnalysis:
    """Ergebnis einer LLM-Analyse."""

    optimization_table: list[dict] = field(default_factory=list)
    formula_rewrites: list[dict] = field(default_factory=list)
    architecture_advice: str = ""
    summary: str = ""
    per_sheet_assessment: list[dict] = field(default_factory=list)
    raw_response: str = ""


def _load_env() -> dict[str, str]:
    """Lädt Variablen aus .env (einfacher Parser, kein python-dotenv nötig)."""
    env_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), ".env")
    env_vars: dict[str, str] = {}
    if os.path.exists(env_path):
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if "=" in line and not line.startswith("#"):
                    k, v = line.split("=", 1)
                    env_vars[k.strip()] = v.strip()
    return env_vars


def get_default_api_key() -> str:
    """Gibt den vorgeladenen API-Key aus .env zurück."""
    env = _load_env()
    return env.get("CLAUDE_API_KEY", "") or os.environ.get("CLAUDE_API_KEY", "")


def get_default_endpoint() -> str:
    """Gibt den Azure-Endpoint aus .env zurück (leer = Anthropic direkt)."""
    env = _load_env()
    return env.get("AZURE_ENDPOINT", "") or os.environ.get("AZURE_ENDPOINT", "")


def get_default_model() -> str:
    """Gibt das Modell aus .env zurück (Fallback: claude-sonnet-4-5-2)."""
    env = _load_env()
    return env.get("CLAUDE_MODEL", "") or os.environ.get("CLAUDE_MODEL", "") or "claude-sonnet-4-5-2"


def mask_key(key: str) -> str:
    """Maskiert den Key für die Anzeige: ••••XXXX (letzte 8 Zeichen sichtbar)."""
    if len(key) <= 8:
        return "••••••••"
    return "••••" + key[-8:]


def _get_api_config(api_key: str, endpoint: str = "") -> tuple[str, dict]:
    """Bestimmt URL und Headers basierend auf Endpoint-Konfiguration."""
    if endpoint:
        url = endpoint.rstrip("/")
        # Wenn die URL schon /messages enthält, direkt verwenden
        if not url.endswith("/messages") and not url.endswith("/chat/completions"):
            url += "/models/chat/completions"

        # Azure Cognitive Services mit /anthropic/ Pfad → Anthropic-kompatible Headers
        if "/anthropic/" in url:
            headers = {
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
                "Content-Type": "application/json",
            }
        else:
            # Sonstiger Azure endpoint (AI Foundry etc.)
            headers = {
                "api-key": api_key,
                "Content-Type": "application/json",
            }
    else:
        # Direct Anthropic API
        url = "https://api.anthropic.com/v1/messages"
        headers = {
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json",
        }
    return url, headers


def test_api_key(api_key: str, endpoint: str = "", model: str = "claude-sonnet-4-5-2") -> tuple[bool, str]:
    """Testet den API-Key mit einem minimalen API-Call.

    Returns (success, message).
    """
    if not api_key:
        return False, "Kein API-Key angegeben."

    if not endpoint:
        return False, "Kein Endpoint angegeben. Bitte die Azure AI Endpoint-URL konfigurieren."

    try:
        url, headers = _get_api_config(api_key, endpoint)
        payload = {
            "model": model,
            "max_tokens": 10,
            "messages": [{"role": "user", "content": "Sag nur OK"}],
        }
        resp = requests.post(url, headers=headers, json=payload, timeout=15)

        if resp.status_code == 200:
            return True, "API-Key funktioniert! ✅"
        elif resp.status_code == 401:
            return False, "API-Key ungültig (401 Unauthorized)."
        elif resp.status_code == 403:
            return False, "Zugriff verweigert (403 Forbidden)."
        else:
            return False, f"API-Fehler: {resp.status_code} – {resp.text[:200]}"
    except requests.exceptions.Timeout:
        return False, "Timeout – Server nicht erreichbar."
    except requests.exceptions.ConnectionError:
        return False, "Verbindungsfehler – Endpoint nicht erreichbar."
    except Exception as e:
        return False, f"Fehler: {str(e)}"


def extract_context(report: WorkbookReport) -> dict:
    """Extrahiert relevanten Kontext aus dem WorkbookReport für die LLM-Analyse."""
    sheets = []
    for s in report.sheet_stats:
        sheets.append({
            "name": s.name,
            "rows": s.row_count,
            "cols": s.col_count,
            "formulas": s.formula_count,
            "static_cells": s.static_count,
            "merged_regions": s.merged_regions,
            "db_readiness": s.db_readiness,
            "is_phantom": s.is_phantom,
        })

    findings_summary = []
    for f in report.findings:
        findings_summary.append({
            "rule_id": f.rule_id,
            "severity": f.severity.value,
            "category": f.category.value,
            "message": f.message,
            "detail": f.detail or "",
            "suggestion": f.suggestion or "",
            "sheet": f.sheet or "",
        })

    return {
        "filename": os.path.basename(report.file_path),
        "file_size_mb": round(report.file_size_mb, 2),
        "sheet_count": report.sheet_count,
        "health_score": report.health_score,
        "sheets": sheets,
        "findings": findings_summary,
        "total_rows": sum(s.row_count for s in report.sheet_stats),
        "total_formulas": sum(s.formula_count for s in report.sheet_stats),
        "total_merged": sum(s.merged_regions for s in report.sheet_stats),
    }


def _build_prompt(context: dict) -> str:
    """Baut den Analyse-Prompt für Claude."""
    context_json = json.dumps(context, ensure_ascii=False, indent=2)

    return f"""Du bist ein freundlicher Excel-Experte und Datenarchitekt bei Austrian Power Grid (APG).
Du analysierst Excel-Dateien und gibst konkrete, hilfreiche Verbesserungsvorschläge.

Dein Ton ist:
- Freundlich und wertschätzend (nie belehrend!)
- Konkret und praxisnah – statt "optimieren" sagst du genau WAS und WIE
- Ermutigend – zeige dass Verbesserung einfach sein kann

WICHTIG für deine Analyse:
- Die Zeilenzahlen und Statistiken im Kontext unten sind bereits bereinigt und spiegeln die tatsächlich relevanten Datenzeilen wider (z.B. ignorierte Phantom-Sheets, leere Zeilen etc.).
- Wenn du große Abweichungen zwischen diesen Zahlen und deiner eigenen Analyse bemerkst, sprich das explizit an und erkläre mögliche Gründe (z.B. unterschiedliche Zählweise, Filter, versteckte Zeilen etc.).
- Beziehe dich bei Empfehlungen immer auf die im Kontext angegebenen Werte.

Hier sind die Analyse-Ergebnisse einer Excel-Datei:

{context_json}

Bitte erstelle eine Analyse im folgenden JSON-Format. Antworte NUR mit validem JSON, kein Markdown, keine Code-Fences:

{{
    "summary": "2-3 kurze Sätze (max 200 Zeichen): Was ist gut, was kann besser werden.",
    "optimization_table": [
                {{
                    "bereich": "z.B. Datenstruktur / Formeln / Zusammenarbeit",
                    "problem": "Max 80 Zeichen – kurz und klar",
                    "vorschlag": "Max 80 Zeichen – konkret und umsetzbar",
                    "aufwand": "gering / mittel / hoch",
                    "nutzen": "gering / mittel / hoch"
                }}
            ],
            "formula_rewrites": [
                {{
                    "original": "Aktuelles Pattern (z.B. verschachtelte SVERWEISe)",
                    "vorschlag": "Bessere Alternative (z.B. INDEX/VERGLEICH)",
                    "erklaerung": "Max 60 Zeichen – warum besser?"
                }}
            ],
            "architecture_advice": "Max 300 Zeichen. Sollte die Datei in eine DB / Power BI migriert werden? Oder ist Excel OK? Konkret begründen.",
            "per_sheet_assessment": [
                {{
                    "sheet": "Name des Sheets",
                    "db_verdict": "Max 80 Zeichen – DB-Tauglichkeit und nächster Schritt",
                    "priority": "hoch / mittel / gering"
                }}
            ]
        }}

        Wichtig:
        - Maximal 5 Einträge in optimization_table (die wichtigsten!)
        - Maximal 3 Einträge in formula_rewrites (oder leer wenn keine Formelprobleme)
        - Ein Eintrag in per_sheet_assessment für jedes Sheet – bewerte die DB-Tauglichkeit konkret
        - Jedes Blatt hat einen db_readiness Score (0-100) – nutze diesen als Ausgangspunkt aber gib deine eigene Einschätzung
        - Wenn Excel für diese Datei OK ist, sag das ehrlich und ermutigend!
        - Alle Texte auf Deutsch
        - Sei konkret: statt "Daten bereinigen" → "Die Spalte 'Status' in Sheet 'Übersicht' enthält gemischte Typen – am besten eine Dropdown-Validierung einrichten"
        - Beziehe dich auf die echten Sheet-Namen und Findings aus der Analyse

        LÄNGENBESCHRÄNKUNGEN (UNBEDINGT EINHALTEN!):
        - summary: Maximal 2-3 kurze Sätze (max 200 Zeichen)
        - optimization_table: Jedes Feld (problem, vorschlag) max 80 Zeichen – kurz und knackig!
        - formula_rewrites: erklaerung max 60 Zeichen
        - architecture_advice: Maximal 3-4 Sätze (max 300 Zeichen)
        - per_sheet_assessment: db_verdict max 80 Zeichen pro Sheet
        - Die gesamte JSON-Antwort muss unter 3000 Zeichen bleiben!
        """


def _extract_json_from_text(raw: str) -> dict | None:
    """Robustes JSON-Extrahieren aus LLM-Antworten.

    Behandelt: Code-Fences (```json ... ```), Text vor/nach JSON,
    abgeschnittenes JSON (fehlende schließende Klammern).
    """
    text = raw.strip()

    # 1) Code-Fences entfernen (```json ... ``` oder ``` ... ```)
    fence_match = re.search(r'```(?:json)?\s*\n?(.*?)```', text, re.DOTALL)
    if fence_match:
        text = fence_match.group(1).strip()

    # 2) Erstes '{' finden und ab da nehmen
    brace_start = text.find('{')
    if brace_start == -1:
        return None
    full_text = text[brace_start:]

    # 3) Letztes '}' finden – versuche erst vollständiges JSON
    brace_end = full_text.rfind('}')
    if brace_end != -1:
        trimmed = full_text[:brace_end + 1]
        try:
            return json.loads(trimmed)
        except json.JSONDecodeError:
            pass

    # 4) Reparatur auf dem VOLLEN Text (nicht abgeschnitten!)
    #    So bleiben Felder erhalten, die NACH dem letzten '}' stehen
    repaired = _repair_truncated_json(full_text)
    if repaired:
        try:
            return json.loads(repaired)
        except json.JSONDecodeError:
            pass

    # 5) Fallback: Reparatur auf dem getrimmten Text
    if brace_end != -1:
        repaired = _repair_truncated_json(trimmed)
        if repaired:
            try:
                return json.loads(repaired)
            except json.JSONDecodeError:
                pass

    return None


def _repair_truncated_json(text: str) -> str | None:
    """Versucht abgeschnittenes JSON zu reparieren.

    Strategie: Progressiv vom Ende her kürzen bis valides JSON entsteht.
    """
    cleaned = text.rstrip()

    # Prüfe ob wir in einem String stecken und schließe ihn
    in_string = False
    i = 0
    while i < len(cleaned):
        c = cleaned[i]
        if in_string:
            if c == '\\':
                i += 2
                continue
            if c == '"':
                in_string = False
        else:
            if c == '"':
                in_string = True
        i += 1

    if in_string:
        cleaned += '"'

    # Versuche progressiv zu reparieren: schneide das letzte
    # unvollständige Element weg und probiere Klammern zu schließen
    # Suche von rechts nach "sicheren Schnittpunkten": Komma, }, ]
    # (außerhalb von Strings)
    attempts = [cleaned]

    # Finde alle möglichen Schnittpunkte von rechts
    in_str = False
    cut_points = []
    i = 0
    while i < len(cleaned):
        c = cleaned[i]
        if in_str:
            if c == '\\':
                i += 2
                continue
            if c == '"':
                in_str = False
        else:
            if c == '"':
                in_str = True
            elif c == ',':
                cut_points.append(i)  # Schnitt VOR Komma
            elif c in ('}', ']'):
                cut_points.append(i + 1)  # Schnitt NACH }]
        i += 1

    # Versuche von der längsten zur kürzesten Version
    for cp in reversed(cut_points):
        candidate = cleaned[:cp].rstrip().rstrip(',')
        attempts.append(candidate)

    for attempt in attempts:
        closed = _close_brackets(attempt)
        try:
            json.loads(closed)
            return closed
        except json.JSONDecodeError:
            continue

    return None


def _close_brackets(text: str) -> str:
    """Zählt offene {/[  und fügt fehlende }/] an."""
    open_braces = 0
    open_brackets = 0
    in_str = False
    i = 0
    while i < len(text):
        c = text[i]
        if in_str:
            if c == '\\':
                i += 2
                continue
            if c == '"':
                in_str = False
        else:
            if c == '"':
                in_str = True
            elif c == '{':
                open_braces += 1
            elif c == '}':
                open_braces -= 1
            elif c == '[':
                open_brackets += 1
            elif c == ']':
                open_brackets -= 1
        i += 1
    return text + ']' * max(0, open_brackets) + '}' * max(0, open_braces)


def analyze_with_llm(
    report_or_context,
    api_key: str,
    endpoint: str = "",
    model: str = "claude-sonnet-4-5-2",
) -> LLMAnalysis:
    """Führt die LLM-Analyse durch. Akzeptiert WorkbookReport oder fertigen Context-Dict."""
    if isinstance(report_or_context, dict):
        context = report_or_context
    else:
        context = extract_context(report_or_context)
    prompt = _build_prompt(context)

    url, headers = _get_api_config(api_key, endpoint)
    payload = {
        "model": model,
        "max_tokens": 4096,
        "messages": [{"role": "user", "content": prompt}],
    }

    resp = requests.post(url, headers=headers, json=payload, timeout=120)
    resp.raise_for_status()

    result = resp.json()

    # Parse response based on API format
    if "content" in result and isinstance(result["content"], list):
        # Anthropic format
        raw_text = result["content"][0]["text"]
    elif "choices" in result:
        # OpenAI/Azure format
        raw_text = result["choices"][0]["message"]["content"]
    else:
        raw_text = json.dumps(result)

    # Erkennen ob die Antwort abgeschnitten wurde
    truncated = False
    if "content" in result and isinstance(result["content"], list):
        stop = result.get("stop_reason", "")
        if stop == "max_tokens":
            truncated = True
    elif "choices" in result:
        finish = result["choices"][0].get("finish_reason", "")
        if finish == "length":
            truncated = True

    # Robustes JSON-Parsing
    data = _extract_json_from_text(raw_text)

    if data is None:
        return LLMAnalysis(
            summary="Die KI-Analyse konnte nicht geparst werden. Bitte versuche es erneut.",
            architecture_advice="",
            raw_response=raw_text,
        )

    analysis = LLMAnalysis(
        optimization_table=data.get("optimization_table", []),
        formula_rewrites=data.get("formula_rewrites", []),
        architecture_advice=data.get("architecture_advice", ""),
        summary=data.get("summary", ""),
        per_sheet_assessment=data.get("per_sheet_assessment", []),
        raw_response=raw_text,
    )

    # Wenn abgeschnitten, Hinweis in die Zusammenfassung
    if truncated and analysis.summary:
        analysis.summary += " ⚠️ Hinweis: Die KI-Antwort wurde gekürzt – einige Details fehlen möglicherweise."

    return analysis
