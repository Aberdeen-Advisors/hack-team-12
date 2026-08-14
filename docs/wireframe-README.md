# AppRat AI wireframe — build notes

Start with the [root README](../README.md) for what the project is, how to run it, what the
interface shows and the scoring model behind the recommendations. This file covers two
things: why the wireframe is split across two HTML files and how to rebuild it, and what the
**Upload inventory** and **Analyze portfolio** buttons do.

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

The fragment's single `<script>` is self-contained on purpose: the `.xlsx`/`.csv` reader and
the scoring engine are inlined in it, with no imports, no globals and no network calls, so
dropping the fragment into the host application brings the upload-and-analyze behaviour with
it and adds no dependency for the host to satisfy. That script is assembled from three
readable sources rather than written as one 3,000-line block — see the comment at its top.

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

## Upload inventory and Analyze portfolio

These two buttons are real. Everything they do happens in the browser: there is no server, no
build step, no bundler, no npm and no CDN, and the file a reader chooses is never uploaded
anywhere.

**Upload inventory** opens a file dialogue wired to a hidden `<input type="file"
accept=".xlsx,.csv">`. Once a file is chosen the banner under the header reports its name, its
row count, how many of its columns were recognised, and how many of the 18 scored inputs were
found, and says that Analyze is now armed. Nothing on the page changes yet.

**Analyze portfolio** scores that file and rewrites the page: the three KPI figures, the savings
progress bars, the decision-mix bar and its legend, the guardrail counts, the redundancy-cluster
rows, and the workbench table. Selecting a row still opens the evidence rail, now rebuilt from
the computed result — the four dimension scores with their pass or fail, the pattern key, the
priority, the confidence, and a plain-English rationale. **Open full decision** carries the same
application through to the detail view.

### Accepted formats

| Format | Requirement |
|---|---|
| `.csv` | Any modern browser. RFC 4180 quoting, `\r\n` or `\n`. |
| `.xlsx` | Needs `DecompressionStream`: Chrome/Edge 103+, Safari 16.4+, Firefox 113+. |

The `.xlsx` reader parses the ZIP central directory itself and inflates entries with the
browser-native `DecompressionStream('deflate-raw')`, then reads SpreadsheetML with `DOMParser`.
Support is feature-detected on load: in a browser without it, the upload button carries an
explanatory tooltip and the banner says to use a CSV instead, rather than throwing when a
workbook is chosen. In a workbook the sheet that recognises the most columns is used, and the
header row is detected rather than assumed — the project's own `Applications` sheet has a merged
banner row above its real header. Number formats are not applied, so a date stored as an Excel
serial is converted on read; the project's own files store ISO date text.

### Column names

Headers are matched case-insensitively, ignoring spaces, hyphens and underscores, so
`Active Users`, `active-users` and `active_users` are the same column, and a leading or trailing
qualifier is tolerated (`App Inventory: Active Users`). The names are the engine's own, from
`engine/generate_dataset.py`. The 18 scored inputs are `ov_increase_value`,
`ov_reach_consumers`, `ov_reduce_costs_efficiency`, `ov_patient_care_criticality`,
`ov_governance_compliance`, `th_supportability`, `th_architecture_fit`,
`th_operational_stability`, `th_vendor_viability`, `th_customization_debt`,
`c_cost_per_active_user_vs_peers`, `c_unused_licence_waste`, `c_consumption_price_variance`,
`c_absolute_cost_band`, `r_technical_risk`, `r_business_compliance_risk`,
`r_clinical_safety_risk` and `r_end_user_perceived_quality`, each on the 1–5 half-step scale
where 5 is always the favourable end.

Also read when present: `app_id`, `app_name`, `vendor_name`, `primary_capability`,
`lifecycle_stage`, `sourcing_type`, `licences_purchased`, `active_users`, the five `cost_*`
categories and `consumption_based_cost`, `term_end`, `renewal_notice_days`,
`overlap_cluster_id`, `cluster_role`, `replacement_app_id`, `replacement_ongoing_tco`,
`replacement_cost_already_in_baseline`, `retention_obligation_flag`, `retention_expiry_date`,
`residual_archival_cost`, `amortised_one_time_migration_cost`, `realization_lag_months`,
`urg_timeline_sensitivity` and `urg_risk_pain_severity`.

**Degrading gracefully.** A blank cell is a blank cell: `null`, `""` and the literal `unknown`
are all treated as missing and never as zero. Each dimension score is a weighted mean
renormalised over the inputs that are actually populated, so a row with a subset of the 18 still
earns a disposition and a priority, with confidence capped by the engine's own completeness
tiers. A dimension with no populated inputs at all gates as a fail, exactly as the Python does —
but the page says so rather than quietly reporting a retire: the banner names the dimension and
the row count, the guardrail list repeats it, the evidence rail shows `no inputs · fails`, and
the rationale spells it out. If a file has none of the recognised columns nothing is computed at
all: the banner says so, names some expected column names, and the sample figures stay in place,
still flagged as samples.

**One input the engine needs that no file in this repository carries as a column** is the
gross-saving basis (`_gross_saving_basis` is a working key inside the Python, not a column). The
port resolves it in this order: a `gross_saving_annual` column if the file has one (also accepted
as `avoidable_annual_cost`); otherwise a `gross_saving_basis` column; otherwise the run-rate
default, where a retire, consolidate or replace claims the application's recurring annual cost
and a retain or invest claims nothing. The netting itself — successor cost, amortised transition
cost, residual archival cost — is unchanged from the Python either way.

### Honesty about which numbers are which

The figures the page ships with describe a fictional 600-application portfolio and have nothing
to do with any real data. Before an analysis, a banner says so and every headline figure carries
a `Sample` pill. After one, the banner names the file and the date, the pills read `Computed`,
and a one-line note reports what the analysis did: how many rows were scored, how many of the 18
scored inputs were found across how many recognised columns, any dimension that could not be
scored, and how confidence came out. Pressing **Analyze portfolio** with no file chosen explains
that and leaves the sample numbers untouched — mock numbers are never presented as computed.

Both `aria-live="polite"` regions are kept, the banner is one of them, the mix bar and the
savings comparison update their `aria-label` alongside their geometry, and every control stays
keyboard-reachable (the file input is visually hidden but focusable, with an `aria-label`).

### Measured parity with the Python engine

The scoring is a port of the core of `engine/generate_dataset.py` — `CRITERIA`, `PASS_THRESHOLD`,
`DISPOSITION_TABLE`, `PRIORITY_LADDER`, `HML`, `COMPLETENESS_FIELDS`, `dimension_score`, `gate`,
`retain_or_invest`, `step_priority`, `urgency`, `override_priority`, the lifecycle and sourcing
guards, the redundancy override, the retention gate, the successor bump, the savings netting, the
completeness and confidence tiers and the portfolio roll-up. It was verified in real Chromium by
driving the actual file input and scraping the rendered page, not by unit-testing the port in
isolation:

| Input | Fields compared | Matched | Mismatches |
|---|---|---|---|
| `data/synthetic-portfolio/applications-v2.csv` | 329 | 329 | **0** |
| `data/synthetic-portfolio/App-Rationalization-Dummy-Dataset-v2.xlsx` | 329 | 329 | **0** |
| the two against each other | 180 | 180 | **0** |

Per application that is the four dimension scores, the four gate verdicts, the pattern key, the
disposition in both the table and the rail, the priority, the confidence and the net annual
saving. The portfolio figures matched too: **$22,057,000** annual run-rate, **$3,308,550** 15%
target, **$5,818,716** net annual saving (**$3,771,716** available now, excluding the four
constrained or deferred rows), and the spread retain 3 / invest 6 / consolidate 6 / replace 2 /
retire 3.

Python's `round(x, 3)` is half-to-even on an exact tie and JavaScript's is not, so this was
checked rather than assumed: every dimension score is `n / (2 × den)` with `den` between 1 and 7,
which is either non-terminating in binary or has at most three decimal places, so no tie arises.
Rounding to whole numbers does use a half-to-even implementation.

`data/client-intake/client-input.csv` was run as the degradation case: 57 columns, 55 recognised,
**11 of the 18 scored inputs**. All 20 rows still get a disposition and a priority, confidence is
capped at medium on every row (0 high), cost efficiency is reported as unscoreable on all 20 rows
in four places on the page, no cluster columns means the cluster section explains its own
emptiness, and nothing throws.

## Known-inert controls

These are still deliberately non-functional and should stay that way until there is something
real behind them:

- the three filter selects — Business capability, Recommendation, Critical operation;
- **Export decisions**;
- **Generate executive readout**.
