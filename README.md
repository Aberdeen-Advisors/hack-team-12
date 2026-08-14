# AppRat AI — application rationalization

Team 12, Aberdeen Advisors. Built for hackathon prompt #5, *AI-Powered Application
Rationalization Tool*.

## What this is

AppRat AI takes a healthcare organisation's application portfolio and recommends what to
do with each application. Every application comes out with one of five recommendations:

| Recommendation | Meaning |
| --- | --- |
| `invest` | Keep it, and fund a remediation or enhancement. |
| `retain` | Keep it as-is. Healthy; no new spend. |
| `consolidate` | Fold it into another application covering the same capability. |
| `replace` | Swap it for a different product. |
| `retire` | Decommission it. |

The problem it addresses: a large health system ends up running hundreds of applications
with heavily overlapping capabilities, and is under pressure to cut technology spend
without breaking a clinical workflow that someone depends on. Deciding that by hand takes
a facilitated workshop per application, which is why it rarely gets done at portfolio
scale. This tool scores the portfolio and produces a recommendation per application with
the evidence attached, so the spend comes out without the clinical risk going in.

Applications with insufficient evidence are held as **needs validation** rather than
being given a recommendation. That is a sixth state in the interface, deliberately not
one of the five recommendations.

## How to run it

There is no build step, no package manager, and no dependencies. Open `index.html` in any
browser — double-clicking the file works. That is the whole procedure. If you are looking
for an install command, there isn't one, and nothing is missing.

It also works served statically, if you would rather have a URL:

```sh
python3 -m http.server 8000
# then open http://localhost:8000/
```

A maintainer could also enable GitHub Pages on `main` at the repository root
(Settings → Pages), which would publish it at
`https://aberdeen-advisors.github.io/hack-team-12/`. That is a repository setting nobody
has switched on yet, not a live link.

The page loads Poppins from Google Fonts and icons from a CDN. Both have offline
fallbacks built in, so it renders correctly with no network access.

## What the interface shows

Three views, switched from the tabs at the top:

- **Executive overview** — total annual portfolio spend, savings identified against the
  CIO's target, the mix of recommendations across the portfolio, critical-operation
  guardrails, and the highest-value clusters of redundant applications.
- **Rationalization workbench** — applications listed with their recommendation, the
  avoidable annual saving, and a confidence level. Selecting an application populates an
  evidence panel beside the table.
- **Decision detail** — one application in full: its recommendation and target, cost,
  usage, confidence, the written rationale, and the evidence rows behind it, including
  any precondition that has to be cleared first.

**Upload inventory** and **Analyze portfolio** are real. **Upload inventory** opens a file
dialogue that accepts an Excel workbook (`.xlsx`) or a CSV; **Analyze portfolio** then runs
the scoring model over that file entirely in the browser — no server, no upload, nothing
leaves the machine — and rewrites every figure on the page with the computed result. The
page ships with illustrative numbers describing a fictional 600-application portfolio, and
until an analysis has run it says so: a banner and a `Sample` pill on each figure. After an
analysis both change to name the file the numbers came from.

The scoring is a JavaScript port of `engine/generate_dataset.py`, verified against the
Python engine's own output: **329 of 329 fields matched with zero mismatches** across all
20 applications of `data/synthetic-portfolio/applications-v2.csv`, and identically via the
`.xlsx`. `docs/wireframe-README.md` has the measurement and the column-name expectations.

Still intentionally not wired up: the three filter selects, **Export decisions**, and
**Generate executive readout**. `docs/wireframe-README.md` has the full list.

## How to make changes

Edit `src/apprat-ai-wireframe-v2.html`, then regenerate the root page:

```sh
python3 src/build-wrapper.py
```

**Do not edit `index.html` directly — it is generated, and your changes will be
overwritten.** It is `src/apprat-ai-wireframe-v2.html` copied in byte-for-byte, wrapped in
a page shell (doctype, head, design tokens, component CSS, icon loader) that the source
fragment expects a host application to provide. Keeping them separate means the fragment
stays droppable into the real application unchanged. See `docs/wireframe-README.md` for
the details of that split.

## The scoring model behind the recommendations

Each application is scored on four dimensions:

1. **Business value**
2. **Technical health**
3. **Cost efficiency**
4. **Risk**

Every input is scored 1 to 5 in half steps, with 5 always the favourable end — so a high
risk score means low risk, never high risk. Inputs roll up into their dimension, and each
dimension is gated at **3.0**: at or above passes, below fails. The pass/fail pattern
across the four dimensions gives a four-character key, which is looked up in a table to
return a recommendation and a priority. Every combination is covered, so the key also
serves as the human-readable rationale for the decision. The thresholds and the lookup
table are configuration rather than code, so they can be tuned without a rewrite.

The requirements this was built against are in `docs/requirements/`, as a workbook of 65
requirements tagged to a seven-step application rationalization process — discovery,
inventory, analysis, execution decision, strategy and roadmap, execution, and monitoring —
plus a sheet recording assumptions, data gaps and open questions.

## Repository layout

| Path | What it is |
| --- | --- |
| `index.html` | The runnable page. Generated — do not hand-edit. |
| `src/apprat-ai-wireframe-v2.html` | The wireframe source. This is the file to edit. |
| `src/build-wrapper.py` | Regenerates `index.html` from the source. |
| `docs/wireframe-README.md` | How the two HTML files relate, and how to rebuild. |
| `docs/requirements/` | The requirements workbook and its README. |
| `data/` | The synthetic datasets and the disposition analyses, with its own README. |
| `.gitignore` | Guards against committing client data or licensed vendor templates. |

## Where the data is

The synthetic datasets and the disposition analyses are committed, in `data/` — the team's
own portfolio, the client-intake view of it, and the three Northstar scoring runs, with
`data/README.md` explaining each file and which version is current. Every figure in them is
invented; only the product and vendor names are real, and both portfolios are fictional.

This repository is public, so what stays out stays out: no client data and no licensed
third-party vendor templates, in `data/` or anywhere else. `.gitignore` blocks the known
ones by name.
