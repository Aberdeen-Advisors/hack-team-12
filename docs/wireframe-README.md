# AppRat AI wireframe — build notes

Start with the [root README](../README.md) for what the project is, how to run it, what the
interface shows and the scoring model behind the recommendations. This file covers one
narrow thing: why the wireframe is split across two HTML files, and how to rebuild it.

In short: edit `src/apprat-ai-wireframe-v2.html`, then run `python3 src/build-wrapper.py`
to regenerate the root `index.html`. Never hand-edit `index.html`.

## Why there are two files

`src/apprat-ai-wireframe-v2.html` is an HTML **fragment**, not a page. It is a bare
`<div id="apprat-wireframe">` with no doctype, `<head>` or `<body>`, and it assumes it is
being pasted into a host application that already provides:

- design tokens — the CSS custom properties `--foreground`, `--card`, `--muted`,
  `--muted-foreground`, `--border`;
- component classes — `.card`, `.btn` / `.btn-primary` / `.btn-ghost` / `.btn-block`,
  `.viz-badge`, `.viz-stat`, `.table` / `.table-sm` / `.table-responsive`, `.form-select`,
  and the `.text-small` / `.text-muted` / `.text-end` helpers;
- `window.lucide` for icons.

Opened directly in a browser it therefore renders blank or unstyled. `index.html` is a thin
host shell that supplies exactly those missing pieces — nothing more. It contains the
fragment **byte-for-byte**, so the wrapper adds no behaviour and changes no content, markup,
copy or numbers.

Keeping them separate means the fragment stays droppable into the real app unchanged, while
`index.html` gives reviewers a link they can click.

## Regenerating the wrapper

Edit the fragment, then:

```sh
python3 src/build-wrapper.py src/apprat-ai-wireframe-v2.html index.html
```

Run with no arguments it defaults to those same two paths. The script asserts that the input
still starts with `<div id="apprat-wireframe">` and ends with `</div>`; if you restructure the
fragment into a full page, the wrapper is no longer needed and this script should go away.

Only the shell (head, tokens, component CSS, icon loader) lives in `build-wrapper.py`. If a
component looks wrong in the browser but right in the real app, fix the CSS in the script and
regenerate — do not patch `index.html`.

## Icons

Icons are [Lucide](https://lucide.dev). `index.html` loads Lucide from a CDN and then calls
`lucide.createIcons()`. If the CDN is unreachable — offline, a locked-down network, a strict
content-security policy — an inline fallback in the same file supplies the ten icons this
wireframe uses, so the page never shows broken-image boxes and the layout does not shift.
Real Lucide always takes precedence when it loads.

## Fonts

The fragment imports Poppins from Google Fonts and that import is left intact. If it does not
load, the page falls back through `Segoe UI` / Calibri / `system-ui` to a platform sans-serif.
Note that the fragment's own rule (`font-family: Poppins, Calibri, sans-serif`) is what applies
inside the wireframe, so on a machine with neither Poppins nor Calibri it lands on the default
sans-serif — legible, slightly different in feel.

## Known-inert controls

The wireframe is a wireframe: only the navigation is wired. These are deliberately non-functional
and should stay that way until there is something real behind them:

- the three filter selects — Business capability, Recommendation, Critical operation;
- **Upload inventory**;
- **Export decisions**;
- **Generate executive readout**.

What does work: the three top-level tabs, the three cluster **Review** buttons (which jump to the
workbench), clicking an application name in the workbench (which populates the evidence rail),
**Open full decision**, and **Analyze portfolio** (which returns to the overview).
