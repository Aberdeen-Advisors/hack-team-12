# TUNED DEMO DATA — NOT A COMPUTED RESULT FROM BINA'S CORRECTED SOURCE

`healthcare_app_rationalization_sample_600_tuned.xlsx` is a **fictional variant** of `healthcare_app_rationalization_sample_600_corrected.xlsx`, built by `engine/tune_northstar_600.py` so that the **unchanged** scoring model returns a net first-year saving of at least 17% of portfolio run cost. The 17% is a property of these input values, not a finding. The un-fitted run stays beside it and is the one to quote: `score_northstar_600_corrected.py` returns $17,409,000 net, 4.9% of $354,330,000.

No weight, band, rubric, gate, lookup-table row, guardrail or savings formula was touched, and no row is special-cased in scoring.

## Result

| | corrected (un-fitted) | tuned |
| --- | --- | --- |
| Portfolio run cost | $354,330,000 | $372,552,000 |
| Gross avoidable claimed | $25,816,000 | $92,845,000 |
| One-time transition | $8,407,000 | $24,032,000 |
| **Net first year** | **$17,409,000** | **$68,813,000** |
| Net as % of run cost | 4.91% | 18.47% |
| retain | 301 | 233 |
| invest | 174 | 129 |
| consolidate | 44 | 129 |
| replace | 0 | 0 |
| retire | 81 | 109 |

## Rows selected, and what protected the rest

- Eligible: APP-021..APP-600, a non-survivor member of a capability cluster, Business Criticality not 'Critical', Patient Care Impact not 'Direct'.
- Selected: **154 of 600**, taken in descending Annual TCO until the projection cleared the target — 68 retire-shaped, 86 consolidate-shaped.
- Left completely alone: APP-001..APP-020 (so the 20-app regression check still means something), every cluster survivor, every 'Critical' application and every application whose Patient Care Impact is 'Direct'.

## Every category of change

| Rows | Annual $ delta | Category |
| --- | --- | --- |
| 462 | — | Capability Map: coverage set to 'Duplicate' (read as Support Role 'Duplicative') on every capability row of a selected duplicate instance |
| 196 | — | Dependencies (consolidate-shaped rows only): 'Required Before Disposition' now names the survivor and the cutover, which is the migration-path evidence the corrected file could only satisfy on 7 rows |
| 178 | — | Capability Map: Capability Criticality stepped down to Medium where the capability is held as a duplicate copy |
| 154 | — | App Inventory: her own Business Value Score -> 2 with a matching rationale, and her (held-out, never scored) Lifecycle Stage label aligned to the story |
| 124 | — | App Inventory: Business Criticality High -> Medium on selected duplicate instances (the capability is held as Primary by the cluster survivor) |
| 86 | $+7,548,000 | TCO: consolidate-shaped rows — six cost components scaled up x1.15, Annual TCO re-derived as their sum, avoidable 55% of the new total, transition 32% of avoidable |
| 68 | — | Performance & Roadmap (retire-shaped rows only): Vendor Support End inside a year, legacy versioned release line, MTTR past the 120-minute band, availability below its own SLA target |
| 68 | $+10,674,000 | TCO: retire-shaped rows — six cost components scaled up x1.25, Annual TCO re-derived as their sum, avoidable 80% of the new total, transition 18% of avoidable |
| 36 | — | Capability Map (cluster survivors only): the survivor is recorded as Primary on the capability its absorbed members hand over, which is the engine's condition (b) for absorbing a duplicate |

## Ratios used, and why they are not 100%

- Retire-shaped: cost components x1.25, avoidable 80% of the new annual total, transition 18% of avoidable. The withheld 20% is internal labour and retained-record archive that does not come out in year one.
- Consolidate-shaped: cost components x1.15, avoidable 55%, transition 32% of avoidable. Interfaces and internal labour persist on the survivor, and the migration cost is not discounted to flatter the net.
- User counts, utilisation and every clinical field were left untouched, so cost per active user stays in a sane range and no clinically critical application changed shape.

## How to check all of this rather than take its word for it

- `python3 engine/audit_tuned_consistency.py` — 15 internal-consistency rules over the fixture and its export: Annual TCO equals the sum of its six components on every row, both App Inventory mirrors match the TCO sheet, Avoidable % and First-Year Net are re-derived rather than stale, transition cost is empty (never zero) on every row that removes no run-rate spend and present on every row that does, the exported absolute-cost band still matches the dollars with the cheapest band at 5, utilisation still equals active / entitled, cost per active user stays inside the range the corrected file already spanned, no row this tuning touched carries an acting disposition alongside Business Criticality 'Critical' / Patient Care Impact 'Direct' / Critical Operation Flag 'Yes', and no retire row carries pilot or beta release wording that would have dodged the lifecycle guard.
- `node engine/verify_tuned_parity.js` — extracts the page's own scoring engine out of `index.html` at run time and replays it over the emitted tool-vocabulary columns, checking disposition parity 600/600, priority parity 600/600, both post-lookup guardrails, and that no row carries a negative net.
- `python3 engine/score_northstar_600_corrected.py` — still returns the corrected run's own figures, untouched, from the corrected input.

## Files

- `healthcare_app_rationalization_sample_600_tuned.xlsx` — this input, with a `Provenance — TUNED` sheet as its first sheet and a banner on `Read Me`
- `engine/score_northstar_600_tuned.py` — the corrected run script with output paths changed only
- `Northstar-Disposition-Analysis-600-tuned.xlsx`, `northstar-dispositions-600-tuned.csv`, `northstar-600-tuned-summary.md`
- `Northstar-600-tuned-tool-vocabulary.xlsx`, `northstar-600-tuned-tool-vocabulary.csv` — for the web page's Upload / Analyze
