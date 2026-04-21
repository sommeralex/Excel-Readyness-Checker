from excel_checker.llm_analysis import _extract_json_from_text, _repair_truncated_json, _close_brackets
import json

# Check: _extract_json_from_text does code-fence removal first
raw = '''```json
{
  "summary": "Test.",
  "architecture_advice": "Migration empfohlen.",
  "per_sheet_assessment": [
    {"sheet": "Portfolio", "db_verdict": "Hoechste Prioritaet fuer DB-'''

text = raw.strip()
# Step 1: Code fence removal
import re
fence_match = re.search(r'```(?:json)?\s*\n?(.*?)```', text, re.DOTALL)
print(f"Fence match: {fence_match}")  # Should be None since no closing ```!

# Step 2: First {
brace_start = text.find('{')
print(f"Brace start: {brace_start}")
text2 = text[brace_start:]

# Step 3: Last }
brace_end = text2.rfind('}')
print(f"Brace end: {brace_end}")
# If -1, text passes through unchanged WITH the code fence prefix

# AH-HA! The fence regex uses .*? (non-greedy) with re.DOTALL
# Since there's no closing ```, the match is None, so the text
# starts with ```json\n{ ... but brace_start finds { inside the fence!
print(f"Text starts with: {text2[:30]}")
