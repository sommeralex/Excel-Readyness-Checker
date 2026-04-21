import re, sys
sys.path.insert(0, '.')

from excel_checker.webapp import app

with app.test_client() as c:
    resp = c.get('/')
    html = resp.data.decode()
    start = html.find('<script>') + 8
    end = html.find('</script>')
    js = html[start:end]

print(f"JS length: {len(js)} chars")
lines = js.split('\n')
print(f"JS lines: {len(lines)}")

# Look for unescaped apostrophes
for i, line in enumerate(lines, 1):
    s = line.rstrip()
    if s.strip().startswith('//'):
        continue
    # Remove properly escaped quotes
    cleaned = s.replace("\\'", "__ESC__")
    count = cleaned.count("'")
    if count % 2 != 0:
        print(f"!!! ODD QUOTES line {i}: {s[:150]}")

print("\n--- Apostrophe patterns ---")
for i, line in enumerate(lines, 1):
    for pattern in ["geht'", "let'", "won'", "we'v", "don'", "isn'"]:
        if pattern in line:
            print(f"Line {i}: {repr(line.rstrip()[:120])}")
            break