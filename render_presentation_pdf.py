from pathlib import Path
from playwright.sync_api import sync_playwright

base = Path(__file__).resolve().parent
html = base / "presentation.html"
out = base / "presentation_clickable_v2.pdf"

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page(viewport={"width": 1400, "height": 900})
    page.goto(html.as_uri(), wait_until="networkidle")
    page.emulate_media(media="print")
    page.pdf(
        path=str(out),
        format="A4",
        print_background=True,
        margin={"top": "14mm", "right": "12mm", "bottom": "14mm", "left": "12mm"},
    )
    browser.close()

print(f"Generated: {out}")
