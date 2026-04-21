"""Test robust JSON parsing for LLM responses."""
from excel_checker.llm_analysis import _extract_json_from_text

# Test 1: Clean JSON
print("Test 1: Clean JSON")
clean = '{"summary": "Test", "optimization_table": []}'
r = _extract_json_from_text(clean)
assert r is not None and r["summary"] == "Test", f"FAIL: {r}"
print("  OK")

# Test 2: Code-fenced JSON
print("Test 2: Code-fenced JSON")
fenced = '```json\n{"summary": "Test", "optimization_table": []}\n```'
r = _extract_json_from_text(fenced)
assert r is not None and r["summary"] == "Test", f"FAIL: {r}"
print("  OK")

# Test 3: Code-fenced without language
print("Test 3: Code-fenced without language")
fenced2 = '```\n{"summary": "Test"}\n```'
r = _extract_json_from_text(fenced2)
assert r is not None and r["summary"] == "Test", f"FAIL: {r}"
print("  OK")

# Test 4: Text before JSON
print("Test 4: Text before JSON")
text_before = 'Here is the analysis:\n{"summary": "Test", "optimization_table": []}'
r = _extract_json_from_text(text_before)
assert r is not None and r["summary"] == "Test", f"FAIL: {r}"
print("  OK")

# Test 5: Truncated JSON - the real-world case!
print("Test 5: Truncated JSON (real-world)")
truncated = '''```json
{
  "summary": "Diese Portfolio-Datei leistet beeindruckende Arbeit.",
  "optimization_table": [
    {"bereich": "Datenidentifikation", "problem": "Keine ID-Spalte", "vorschlag": "IDs einfuegen", "aufwand": "gering", "nutzen": "hoch"},
    {"bereich": "Externe Abh.", "problem": "1.608 externe Bezuege", "vorschlag": "Power Query", "aufwand": "mittel", "nutzen": "hoch"}
  ],
  "formula_rewrites": [
    {"original": "SVERWEIS", "vorschlag": "INDEX/VERGLEICH", "erklaerung": "Schneller"}
  ],
  "architecture_advice": "Migration empfohlen.",
  "per_sheet_assessment": [
    {"sheet": "Portfolio", "db_verdict": "Hoechste Prioritaet fuer DB-'''

r = _extract_json_from_text(truncated)
assert r is not None, "FAIL: got None"
assert r.get("summary") == "Diese Portfolio-Datei leistet beeindruckende Arbeit.", f"FAIL summary: {r.get('summary')}"
assert len(r.get("optimization_table", [])) == 2, f"FAIL opt table: {len(r.get('optimization_table', []))}"
assert len(r.get("formula_rewrites", [])) == 1, f"FAIL rewrites: {len(r.get('formula_rewrites', []))}"
assert r.get("architecture_advice") == "Migration empfohlen.", f"FAIL arch: {r.get('architecture_advice')}"
print(f"  OK - salvaged {len(r.keys())} keys (per_sheet may be partial or missing)")

# Test 6: Truncated mid-string in array
print("Test 6: Truncated mid-value")
trunc2 = '{"summary": "Gut", "optimization_table": [{"bereich": "Test", "problem": "Prob", "vorschlag": "Fix it by doing thi'
r = _extract_json_from_text(trunc2)
assert r is not None, "FAIL: got None"
assert r.get("summary") == "Gut", f"FAIL: {r.get('summary')}"
print("  OK")

# Test 7: Empty/garbage
print("Test 7: Garbage input")
r = _extract_json_from_text("no json here at all")
assert r is None, f"FAIL: should be None, got {r}"
print("  OK")

# Test 8: Real user truncation pattern
print("Test 8: Real-world truncation (from user)")
real = '''```json { "summary": "Diese Portfolio-Datei leistet beeindruckende Arbeit", "optimization_table": [ { "bereich": "Datenidentifikation", "problem": "Keines der 7 Sheets hat eine eindeutige ID-Spalte", "vorschlag": "In jedem Sheet eine erste Spalte ID einfuegen", "aufwand": "gering", "nutzen": "hoch" } ], "architecture_advice": "Migration empfohlen", "per_sheet_assessment": [ { "sheet": "Portfolio", "db_verdict": "Hoechste Prioritaet fuer DB-'''
r = _extract_json_from_text(real)
assert r is not None, "FAIL: got None"
assert r.get("summary") is not None, "FAIL: no summary"
assert len(r.get("optimization_table", [])) == 1, f"FAIL: opt_table count"
print(f"  OK - summary: {r['summary'][:50]}...")

print("\nAll tests passed!")
