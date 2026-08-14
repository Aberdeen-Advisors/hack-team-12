# Northstar Global Health — 600-application run

Model: `score_northstar_v3.py`, imported unchanged. Input: `healthcare_app_rationalization_sample_600.xlsx`. Run date 2026-08-14. Wall clock 9.25s for 600 rows (2.91s to load, normalise, score and decide).

## Schema

Same 12 sheets and same 190 column headers as the committed 20-application `-with-risk` sample: 0 columns added, removed or reordered. The model runs unmodified. The divergence is in the value domains of the 580 new applications, handled by a documented normalisation layer — see `Vocabulary mapping` in the workbook.

## Dispositions

| Term | Count | % |
| --- | --- | --- |
| retain | 301 | 50.2% |
| invest | 174 | 29.0% |
| consolidate | 44 | 7.3% |
| replace | 0 | 0.0% |
| retire | 81 | 13.5% |

Only the all-pass pattern returns retain, so 301 rows (50.2%) need no action and 299 carry one.

## Money

- Gross annual avoidable claimed: $25,816,000
- One-time transition cost on those rows: $8,407,000 (32.6% of gross)
- **Net first year: $17,409,000**
- Safe (high-confidence only, her rule): $15,850,000
- Potential: $1,559,000
- Portfolio annual run cost: $354,330,000; her CIO target 15% = $53,149,500

## Priority

| Priority | Count |
| --- | --- |
| Very High | 1 |
| High | 88 |
| Moderate | 173 |
| Low | 37 |
| Very Low | 301 |

### Top 10 by priority then value at risk

| # | App | Disposition | Priority | Annual cost | Net first-year |
| --- | --- | --- | --- | --- | --- |
| 1 | APP-006 Zoom Workplace | retire | Very High | $1,130,000 | $370,000 |
| 2 | APP-002 Oracle Health EHR | consolidate | High | $6,500,000 | $2,200,000 |
| 3 | APP-018 Oracle PeopleSoft HCM | consolidate | High | $2,780,000 | $750,000 |
| 4 | APP-005 Slack Enterprise Grid | consolidate | High | $1,260,000 | $600,000 |
| 5 | APP-016 Tableau Cloud | consolidate | High | $1,040,000 | $530,000 |
| 6 | APP-014 Jira Service Management | consolidate | High | $720,000 | $390,000 |
| 7 | APP-348 Oracle NetSuite ERP — Northeast Primary Instance 08 | retire | High | $485,000 | $312,000 |
| 8 | APP-008 Abridge | consolidate | High | $740,000 | $260,000 |
| 9 | APP-339 BambooHR — Northeast Legacy Instance 19 | retire | High | $635,000 | $229,000 |
| 10 | APP-344 Infor CloudSuite Healthcare — Community Primary Instance 04 | retire | High | $780,000 | $228,000 |

## Consolidation

35 overlap groups, 35 with more than one member. Largest: CLU-07 (20 members); CLU-08 (20 members); CLU-09 (20 members); CLU-10 (20 members); CLU-11 (20 members). 7 rows were forced to consolidate by the redundancy override.

Support Role after normalisation: Secondary 1424, Duplicative 254, Primary 119.

## Confidence

- high: 569
- Needs Validation: 29
- medium: 2

### What that confidence figure does and does not say

**$13,650,000 of the $17,409,000 net first-year saving sits on the 569 of 600 applications whose `confidence` column reads `high`.** That is the portion of the saving where the risk evidence is complete. It is not a verified figure, not a validated one, and not a statement that the underlying numbers are trustworthy to that degree — only that the inputs the confidence rule inspects were present.

What the rule actually grades is **availability**: the three risk inputs, plus the source's own evidence-quality columns. It does **not** mean every input on the row is present. **54 rows are missing a weighted technical input** — architecture fit, unscored where the hosting label is the ambiguous bare `Private cloud` — **and still read `high`**, because a missing non-risk input does not cap confidence.

This is a **different measure** from the **safe figure of $15,850,000** in `Money` above. Safe is gated on cost notes and the guardrails, not on the per-application `confidence` value; the parenthetical "high-confidence only" on that line is what caused the two to be conflated earlier. To be unambiguous: **$13,650,000 / 569 rows is the confidence-column figure; $15,850,000 is the safe figure.** They are computed from different gates, neither is a subset of the other by construction, and neither should be cited as the other.

**Disclosure that must travel with the figure:** the cost lens on this portfolio rests on **three criteria, not four**. `c_consumption_price_variance` is unscorable on all 600 rows because the source workbook has no metered, plan or consumption-price column.

**Correction worth recording, so nobody chases it:** adding that missing consumption column would **not** restore full confidence in the web page. The page holds rows at medium for two independent reasons, and **10 of its 44 completeness fields are blank on 599 rows** — implementation date, cost centre, last sign-in, process centrality, owner-stated importance, contract id, licence metric, retention flag, lifecycle stage, and the consumption input. Supplying that one input moves completeness from **0.773 to 0.795 against a 0.90 gate**, so it lifts nothing.

## Data gaps

- `c_consumption_price_variance`: unscored on 600 of 600 rows
- `r_end_user_perceived_quality`: unscored on 600 of 600 rows

## Vocabulary mappings applied

- Capability Map · Support Role / Coverage Level: Support Role 'Analyst' + Coverage Level 'Supporting' -> Support Role 'Secondary' (580 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Compliance analyst' + Coverage Level 'Supporting' -> Support Role 'Secondary' (580 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Business user' + Coverage Level 'Duplicate' -> Support Role 'Duplicative' (152 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Business user' + Coverage Level 'Supporting' -> Support Role 'Secondary' (133 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Clinical user' + Coverage Level 'Duplicate' -> Support Role 'Duplicative' (80 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Clinical user' + Coverage Level 'Supporting' -> Support Role 'Secondary' (70 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Business user' + Coverage Level 'Primary' -> Support Role 'Primary' (57 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Business user' + Coverage Level 'Limited' -> Support Role 'Secondary' (38 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Clinical user' + Coverage Level 'Primary' -> Support Role 'Primary' (30 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Clinical user' + Coverage Level 'Limited' -> Support Role 'Secondary' (20 rows)
- Healthcare Guardrails · Data Applicability: Confirmed -> Applicable (574 rows)
- Healthcare Guardrails · Interface Health: Healthy -> Green (574 rows)
- Healthcare Guardrails · Interface Health: Degraded -> Amber (6 rows)
- Healthcare Guardrails · Recovery Test Meets RTO/RPO: Not Tested -> Unknown (3 rows)
- Healthcare Guardrails · Last Restore Test Result: Not Tested -> Unknown (3 rows)
- App Inventory · Hosting Model: On-premises -> Customer-hosted on-premises (58 rows)

## Tool-vocabulary translation (for the wireframe's Upload / Analyze)

The same 600 rows are also emitted in the tool's own snake_case column vocabulary, read at run time from the committed `App-Rationalization-Dummy-Dataset-v2.xlsx` so it cannot drift from it. 82 of 125 tool columns are populated; the other 43 are **empty because her workbook does not support them**. A blank is not a zero: no column was defaulted to zero, to a neutral score or to a plausible string. The 18 criterion columns carry this run's derived scores, and are blank where an input was not derivable — which is what the tool's arithmetic expects, since it skips a null in both numerator and denominator and renormalises.

Unpopulated tool columns, so nobody reads a blank as data:

- `ai_delivery_form`
- `ai_capability_class`
- `ai_host_app_id`
- `ai_already_entitled_elsewhere`
- `ai_entitled_alternative_app_id`
- `implementation_date`
- `version_vendor_supported`
- `technical_obsolescence_flag`
- `legal_entity`
- `department`
- `cost_centre`
- `is_orphaned`
- `governance_visibility`
- `is_shadow_it`
- `capability_tag_confidence`
- `last_signin_date`
- `process_centrality` — her Capability Criticality uses a four-step ladder including Critical, which the tool's three-step High/Medium/Low field has no slot for.
- `owner_stated_strategic_importance` — LEFT EMPTY DELIBERATELY.
- `consumption_based_cost` — no metered or consumption cost line anywhere in her workbook.
- `one_time_implementation_cost`
- `unused_licence_spend`
- `contract_id`
- `term_start`
- `licence_metric`
- `early_termination_penalty` — no early-termination, break-fee or penalty column exists anywhere in her workbook.
- `contract_runway_months`
- `c_consumption_price_variance`
- `r_end_user_perceived_quality`
- `sourcing_exclusion_applied`
- `retention_override_applied`
- `suppressed_recommendation`
- `suppression_reason`
- `consolidation_saving`
- `retention_obligation_flag`
- `retention_expiry_date`
- `residual_archival_cost` — EMPTY on every row, never 0.
- `replacement_ongoing_tco` — EMPTY on every row, never 0.
- `net_saving_five_year`
- `realization_lag_months`
- `urg_risk_pain_severity` — LEFT EMPTY DELIBERATELY.
- `urgency_score`
- `completeness_score`
- `missing_fields`

Two of those are empty on purpose rather than for want of data: `lifecycle_stage`, which the scoring run holds out as circular and the UI reads for its lifecycle exclusion, and `owner_stated_strategic_importance`, an interview field in an iteration that interviews nobody. Her lifecycle label travels in `her_lifecycle_stage_comparison_only_not_scored`, outside the tool's vocabulary, never scored.

## Outputs

- `Northstar-Disposition-Analysis-600-corrected.xlsx`
- `northstar-dispositions-600-corrected.csv`
- `Northstar-600-corrected-tool-vocabulary.xlsx`
- `northstar-600-corrected-tool-vocabulary.csv`
- this file
