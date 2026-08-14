#!/usr/bin/env python3
"""Wrap the AppRat wireframe fragment in a standalone, self-contained HTML page.

The fragment (src/apprat-ai-wireframe-v2.html) is copied in byte-for-byte.
Everything else in the output is the host shell: doctype, head, design tokens,
component CSS for the classes the fragment expects, and a Lucide loader with an
offline fallback.
"""
import pathlib
import sys

HERE = pathlib.Path(__file__).resolve().parent  # src/
ROOT = HERE.parent  # repo root
FRAGMENT = sys.argv[1] if len(sys.argv) > 1 else str(HERE / "apprat-ai-wireframe-v2.html")
OUT = sys.argv[2] if len(sys.argv) > 2 else str(ROOT / "index.html")

fragment = pathlib.Path(FRAGMENT).read_text(encoding="utf-8").strip()
assert fragment.startswith('<div id="apprat-wireframe">'), "unexpected fragment start"
assert fragment.endswith("</div>"), "unexpected fragment end"

HEAD = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="Static wireframe for AppRat AI, an application rationalization workspace. Executive overview, rationalization workbench and decision detail.">
<meta name="robots" content="noindex">
<title>AppRat AI &mdash; Application Rationalization</title>

<!-- Poppins, with a full system fallback stack if Google Fonts is unavailable. -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&display=swap">

<style>
/* ---------------------------------------------------------------------------
   Host design system.

   The wireframe fragment below is written against a host application's design
   system. It never ships those styles itself, so this block supplies them:
   the CSS custom properties it reads (--foreground, --card, --muted,
   --muted-foreground, --border) and the utility/component classes it uses
   (.card, .btn and variants, .viz-badge, .viz-stat, .table and variants,
   .form-select, and the text helpers). Deliberately neutral -- the intent is
   to make the fragment legible, not to give it a new visual language.
   --------------------------------------------------------------------------- */

:root {
  color-scheme: light;

  --foreground: #1a1a1a;
  --muted-foreground: #6b7280;
  --card: #ffffff;
  --muted: #eef1f4;
  --border: #dfe3e8;

  --page-bg: #f6f8fa;
  --accent: #09375f;

  --font-sans: Poppins, "Segoe UI", Calibri, system-ui, -apple-system,
    "Helvetica Neue", Arial, "Noto Sans", sans-serif;
  --radius: 8px;
}

* { box-sizing: border-box; }

html { -webkit-text-size-adjust: 100%; }

body {
  margin: 0;
  padding: 24px 20px 56px;
  background: var(--page-bg);
  color: var(--foreground);
  font-family: var(--font-sans);
  font-size: 14px;
  line-height: 1.5;
  -webkit-font-smoothing: antialiased;
}

.page-shell { max-width: 1360px; margin: 0 auto; }

h1, h2, h3, h4 { font-weight: 500; line-height: 1.25; }
h2 { font-size: 1.4rem; }
h3 { font-size: 1.05rem; }
p { line-height: 1.55; }

/* --- surfaces ------------------------------------------------------------ */

.card {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: var(--radius);
}

/* --- buttons ------------------------------------------------------------- */

.btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 0.4rem;
  padding: 0.45rem 0.8rem;
  border: 1px solid var(--border);
  border-radius: 6px;
  background: var(--card);
  color: inherit;
  font: inherit;
  line-height: 1.3;
  text-align: center;
  cursor: pointer;
  transition: background-color 0.12s ease, border-color 0.12s ease;
}

.btn:hover { background: var(--muted); }
.btn:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }

.btn-primary { background: var(--accent); border-color: var(--accent); color: #fff; }
.btn-primary:hover { background: #0b4478; }

.btn-ghost { background: transparent; border-color: transparent; }
.btn-ghost:hover { background: var(--muted); }

.btn-block { display: flex; width: 100%; margin-top: 0.75rem; }

/* --- badges and stats ---------------------------------------------------- */

.viz-badge {
  display: inline-flex;
  align-items: center;
  gap: 0.35rem;
  padding: 0.25rem 0.6rem;
  border-radius: 999px;
  background: var(--muted);
  font-size: 12px;
  line-height: 1.4;
  white-space: nowrap;
}

.viz-stat { display: block; }
.viz-stat-value { display: block; font-size: 1.9rem; line-height: 1.15; }

/* --- tables -------------------------------------------------------------- */

.table { width: 100%; border-collapse: collapse; }

.table th,
.table td {
  padding: 0.5rem 0.4rem;
  border-bottom: 1px solid var(--border);
  text-align: left;
  vertical-align: top;
}

.table th {
  font-weight: 500;
  font-size: 12px;
  color: var(--muted-foreground);
  white-space: nowrap;
}

.table tbody tr:last-child td { border-bottom: 0; }

.table-sm th,
.table-sm td { padding: 0.4rem 0.35rem; }

/* Wide tables scroll inside their own box; the page body never scrolls
   sideways. */
.table-responsive {
  overflow-x: auto;
  -webkit-overflow-scrolling: touch;
}

/* --- forms --------------------------------------------------------------- */

.form-select {
  width: 100%;
  padding: 0.4rem 0.5rem;
  border: 1px solid var(--border);
  border-radius: 6px;
  background: var(--card);
  color: inherit;
  font: inherit;
}

.form-select:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px; }

/* --- text helpers -------------------------------------------------------- */

.text-small { font-size: 12px; }
.text-muted { color: var(--muted-foreground); }
.text-end { text-align: right; }

/* --- icons --------------------------------------------------------------- */

/* Lucide swaps each <i data-lucide> for an inline <svg class="lucide">. Both
   states are sized here so the layout is identical whether or not the icon
   script loaded -- no reflow, and no broken-image boxes if it never arrives. */
i[data-lucide],
svg.lucide {
  display: inline-block;
  flex: none;
  width: 16px;
  height: 16px;
  vertical-align: -0.15em;
}

svg.lucide {
  fill: none;
  stroke: currentColor;
  stroke-width: 2;
  stroke-linecap: round;
  stroke-linejoin: round;
}

@media (max-width: 720px) {
  body { padding: 16px 12px 40px; }
}
</style>
</head>
<body>
<main class="page-shell">
"""

TAIL = """</main>

<!-- Icons: Lucide from CDN when reachable. -->
<script src="https://unpkg.com/lucide@0.544.0/dist/umd/lucide.min.js" crossorigin="anonymous"></script>
<script>
/* Offline fallback. If the CDN did not load, stand up a minimal window.lucide
   with the same createIcons() surface, carrying only the ten icons this
   wireframe actually asks for. The fragment's own script calls
   window.lucide.createIcons() when it rebuilds the evidence rail, so the shim
   has to satisfy that call too. Real Lucide always wins when it is present. */
(function () {
  if (!window.lucide || typeof window.lucide.createIcons !== "function") {
    var ICONS = {
      "upload": '<path d="M12 3v12"/><path d="m17 8-5-5-5 5"/><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>',
      "download": '<path d="M12 15V3"/><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m7 10 5 5 5-5"/>',
      "sparkles": '<path d="M9.937 15.5A2 2 0 0 0 8.5 14.063l-6.135-1.582a.5.5 0 0 1 0-.962L8.5 9.936A2 2 0 0 0 9.937 8.5l1.582-6.135a.5.5 0 0 1 .963 0L14.063 8.5A2 2 0 0 0 15.5 9.937l6.135 1.581a.5.5 0 0 1 0 .964L15.5 14.063a2 2 0 0 0-1.437 1.437l-1.582 6.135a.5.5 0 0 1-.963 0z"/><path d="M20 3v4"/><path d="M22 5h-4"/><path d="M4 17v2"/><path d="M5 18H3"/>',
      "shield-check": '<path d="M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72a1.17 1.17 0 0 1 1.52 0C14.51 3.81 17 5 19 5a1 1 0 0 1 1 1z"/><path d="m9 12 2 2 4-4"/>',
      "git-branch": '<line x1="6" x2="6" y1="3" y2="15"/><circle cx="18" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><path d="M18 9a9 9 0 0 1-9 9"/>',
      "circle-alert": '<circle cx="12" cy="12" r="10"/><line x1="12" x2="12" y1="8" y2="12"/><line x1="12" x2="12.01" y1="16" y2="16"/>',
      "circle-check": '<circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/>',
      "triangle-alert": '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/><path d="M12 9v4"/><path d="M12 17h.01"/>',
      "file-text": '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M10 9H8"/><path d="M16 13H8"/><path d="M16 17H8"/>',
      "mouse-pointer-click": '<path d="M14 4.1 12 6"/><path d="m5.1 8-2.9-.8"/><path d="m6 12-1.9 2"/><path d="M7.2 2.2 8 5.1"/><path d="M9.037 9.69a.498.498 0 0 1 .653-.653l11 4.5a.5.5 0 0 1-.074.949l-4.349 1.041a1 1 0 0 0-.74.739l-1.04 4.35a.5.5 0 0 1-.95.074z"/>'
    };

    var SVG_NS = "http://www.w3.org/2000/svg";

    window.lucide = {
      createIcons: function (options) {
        var attrs = (options && options.attrs) || {};
        var nodes = document.querySelectorAll("[data-lucide]");
        for (var i = 0; i < nodes.length; i++) {
          var node = nodes[i];
          var body = ICONS[node.getAttribute("data-lucide")];
          if (!body) continue;

          var svg = document.createElementNS(SVG_NS, "svg");
          svg.setAttribute("xmlns", SVG_NS);
          svg.setAttribute("viewBox", "0 0 24 24");
          svg.setAttribute("fill", "none");
          svg.setAttribute("stroke", "currentColor");
          svg.setAttribute("stroke-width", "2");
          svg.setAttribute("stroke-linecap", "round");
          svg.setAttribute("stroke-linejoin", "round");
          svg.setAttribute("width", attrs.width || 16);
          svg.setAttribute("height", attrs.height || 16);
          svg.setAttribute("aria-hidden", "true");
          svg.setAttribute("class", ("lucide " + (node.getAttribute("class") || "")).trim());
          svg.innerHTML = body;

          if (node.parentNode) node.parentNode.replaceChild(svg, node);
        }
      }
    };
  }

  window.lucide.createIcons({ attrs: { width: 16, height: 16 } });
})();
</script>
</body>
</html>
"""

pathlib.Path(OUT).write_text(HEAD + fragment + "\n" + TAIL, encoding="utf-8")
print("wrote", OUT)
