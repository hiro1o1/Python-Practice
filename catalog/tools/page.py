"""Stage 3: write the catalog page (catalog/index.html) from the board config."""

from collections import OrderedDict
from html import escape

CSS = """
:root {
  --bg: #f6f7f9; --card: #ffffff; --ink: #111827; --muted: #5b6472;
  --line: #e3e6eb; --accent: #2563eb; --chip: #eef2ff; --chip-ink: #1e3a8a;
}
@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) {
    --bg: #0f1217; --card: #171b22; --ink: #e8ebf0; --muted: #9aa3b2;
    --line: #2a303b; --accent: #6ea0ff; --chip: #1c2640; --chip-ink: #c7d6ff;
  }
}
:root[data-theme="dark"] {
  --bg: #0f1217; --card: #171b22; --ink: #e8ebf0; --muted: #9aa3b2;
  --line: #2a303b; --accent: #6ea0ff; --chip: #1c2640; --chip-ink: #c7d6ff;
}
* { box-sizing: border-box; }
body { margin: 0; background: var(--bg); color: var(--ink);
  font: 16px/1.55 Inter, system-ui, -apple-system, "Segoe UI", Roboto, sans-serif; }
header, main, footer { max-width: 1180px; margin: 0 auto; padding: 0 16px; }
header { padding-top: 48px; padding-bottom: 8px; }
h1 { font-size: clamp(28px, 4vw, 40px); line-height: 1.15; margin: 0 0 8px; letter-spacing: -0.02em; }
header p { color: var(--muted); margin: 0; max-width: 60ch; }
nav { display: flex; flex-wrap: wrap; gap: 8px; margin: 24px 0 8px; }
nav a { color: var(--ink); text-decoration: none; border: 1px solid var(--line); background: var(--card);
  padding: 6px 12px; border-radius: 999px; font-size: 14px; font-weight: 500; }
nav a:hover { border-color: var(--accent); }
.board { background: var(--card); border: 1px solid var(--line); border-radius: 16px;
  margin: 24px 0; overflow: hidden; }
.board-head { display: flex; flex-wrap: wrap; align-items: baseline; gap: 8px 16px; padding: 24px 24px 0; }
.board h2 { font-size: 24px; margin: 0; letter-spacing: -0.01em; }
.chip { background: var(--chip); color: var(--chip-ink); font-size: 13px; font-weight: 600;
  padding: 2px 10px; border-radius: 999px; }
.shot { display: block; margin: 16px 16px 0; background: #fff; border-radius: 12px; overflow: hidden; }
.shot img { display: block; width: 100%; height: auto; }
.specs { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 8px 32px;
  padding: 20px 24px 8px; }
.specs h3 { font-size: 13px; text-transform: uppercase; letter-spacing: .06em; color: var(--muted);
  margin: 8px 0 6px; font-weight: 600; }
.specs ul { list-style: none; margin: 0 0 12px; padding: 0; }
.specs li { display: flex; justify-content: space-between; gap: 12px; padding: 7px 0;
  border-bottom: 1px solid var(--line); font-size: 15px; }
.specs li span:last-child { color: var(--muted); text-align: right; }
.qty { color: var(--accent); font-weight: 600; margin-left: 4px; }
.downloads { display: flex; flex-wrap: wrap; gap: 8px 16px; padding: 4px 24px 24px; font-size: 14px; }
.downloads a { color: var(--accent); text-decoration: none; font-weight: 500; }
.downloads a:hover { text-decoration: underline; }
footer { color: var(--muted); font-size: 13px; padding-bottom: 48px; }
@media (max-width: 600px) {
  .board-head { padding: 16px 16px 0; }
  .shot { margin: 12px 8px 0; }
  .specs { padding: 16px 16px 4px; }
  .downloads { padding: 4px 16px 16px; }
}
"""


def grouped(callouts):
    """Merge identical (title, sub) entries into one row with a count."""
    rows = OrderedDict()
    for c in callouts:
        key = (c["title"], c["sub"] or "")
        rows[key] = rows.get(key, 0) + 1
    return rows


def spec_list(title, callouts):
    if not callouts:
        return ""
    items = []
    for (name, sub), n in grouped(callouts).items():
        qty = f'<span class="qty">×{n}</span>' if n > 1 else ""
        items.append(f"<li><span>{escape(name)}{qty}</span><span>{escape(sub)}</span></li>")
    return f"<div><h3>{escape(title)}</h3><ul>{''.join(items)}</ul></div>"


def board_section(b):
    edge = [c for c in b["callouts"] if c["side"] != "tag"]
    internal = [c for c in b["callouts"] if c["side"] == "tag"]
    soc = b["name"].split()[0]
    i = b["id"]
    return f"""
<section class="board" id="{i}">
  <div class="board-head"><h2>{escape(b["name"])}</h2><span class="chip">{escape(soc)}</span></div>
  <a class="shot" href="images/{i}.png" title="Open full-resolution image">
    <img src="images/{i}.webp" alt="{escape(b["name"])} with labelled connectors" loading="lazy">
  </a>
  <div class="specs">{spec_list("Connectors & I/O", edge)}{spec_list("On-board headers", internal)}</div>
  <div class="downloads">
    <a href="images/{i}.png" download>Labelled image (PNG)</a>
    <a href="images/{i}-plain.png" download>Board only, no labels (PNG)</a>
  </div>
</section>"""


def write_page(path, boards):
    nav = "".join(f'<a href="#{b["id"]}">{escape(b["name"])}</a>' for b in boards)
    html = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Mainboard Catalog</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap">
<style>{CSS}</style>
</head>
<body>
<header>
  <h1>Mainboard Catalog</h1>
  <p>Android mainboards for interactive displays and signage. Each board shows its connectors and on-board headers.</p>
  <nav>{nav}</nav>
</header>
<main>{"".join(board_section(b) for b in boards)}
</main>
<footer>Images generated by <code>catalog/tools/build.py</code>. Edit labels in <code>catalog/tools/boards.py</code> and rebuild.</footer>
</body>
</html>
"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
