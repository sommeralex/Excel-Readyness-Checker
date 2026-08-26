import sys, time
from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    try:
        b = p.chromium.launch(args=["--no-sandbox","--js-flags=--max-old-space-size=4096"])
    except Exception as e:
        print("default launch failed:", str(e)[:120]); 
        b = p.chromium.launch(executable_path="/opt/pw-browsers/chromium-1194/chrome-linux/chrome",
                              args=["--no-sandbox"])
    pg = b.new_page()
    pg.on("console", lambda m: None)
    pg.goto("http://127.0.0.1:8765/bench.html")
    deadline = time.time() + 1500
    while time.time() < deadline:
        if pg.evaluate("window.DONE === true"): break
        time.sleep(3)
    print(pg.inner_text("#log"))
    b.close()
