# Northstar Global Health — 600-application run, TUNED DEMO FIXTURE

> **TUNED DEMO DATA — THESE FIGURES ARE A PROPERTY OF THE INPUT FILE, NOT A COMPUTED RESULT.** Every figure below is the unchanged model's answer to `healthcare_app_rationalization_sample_600_tuned.xlsx`, a fictional variant of Bina's corrected workbook whose input values were constructed by `engine/tune_northstar_600.py` to clear a 17% first-year savings target. It is not a computed result from her corrected source and says nothing about the estate that workbook describes. The source this fixture was derived from — `healthcare_app_rationalization_sample_600_corrected.xlsx`, and the `score_northstar_600_corrected.py` that reads it — stays in the repo, untouched. Provenance: the `Provenance — TUNED` sheet in both workbooks, `northstar-600-tuned-change-log.md`, and `data/README.md`.

Model: `score_northstar_v3.py`, imported unchanged. Input: `healthcare_app_rationalization_sample_600_tuned.xlsx`. Run date 2026-08-14. Wall clock 10.02s for 600 rows (3.75s to load, normalise, score and decide).

## Schema

Same 12 sheets and same 190 column headers as the committed 20-application `-with-risk` sample: 0 columns added, removed or reordered. The model runs unmodified. The divergence is in the value domains of the 580 new applications, handled by a documented normalisation layer — see `Vocabulary mapping` in the workbook.

## Dispositions

| Term | Count | % |
| --- | --- | --- |
| retain | 233 | 38.8% |
| invest | 129 | 21.5% |
| consolidate | 129 | 21.5% |
| replace | 0 | 0.0% |
| retire | 109 | 18.2% |

Only the all-pass pattern returns retain, so 233 rows (38.8%) need no action and 367 carry one.

## Money

- Gross annual avoidable claimed: $92,845,000
- One-time transition cost on those rows: $24,032,000 (25.9% of gross)
- **Net first year: $68,813,000**
- Safe (high-confidence only, her rule): $61,671,000
- Potential: $7,142,000
- Portfolio annual run cost: $372,552,000; her CIO target 15% = $55,882,800

## Priority

| Priority | Count |
| --- | --- |
| Very High | 18 |
| High | 140 |
| Moderate | 173 |
| Low | 36 |
| Very Low | 233 |

### Top 10 by priority then value at risk

| # | App | Disposition | Priority | Annual cost | Net first-year |
| --- | --- | --- | --- | --- | --- |
| 1 | APP-075 Waystar — Pacific Legacy Instance 15 | retire | Very High | $1,237,000 | $812,000 |
| 2 | APP-350 Unit4 ERP — Mountain Primary Instance 10 | retire | Very High | $1,213,000 | $795,000 |
| 3 | APP-321 Workday Human Capital Management — Mountain Primary Instance | retire | Very High | $1,187,000 | $779,000 |
| 4 | APP-070 Availity Essentials — Midwest Primary Instance 10 | retire | Very High | $1,106,000 | $726,000 |
| 5 | APP-081 Epic Cadence — Northeast Primary Instance | retire | Very High | $944,000 | $619,000 |
| 6 | APP-310 Kronos Workforce Central — Northeast Primary Instance 10 | retire | Very High | $919,000 | $603,000 |
| 7 | APP-501 REDCap — Central Primary Instance | retire | Very High | $838,000 | $549,000 |
| 8 | APP-381 ServiceNow IT Service Management — Regional Primary Instance | retire | Very High | $807,000 | $530,000 |
| 9 | APP-141 Microsoft Teams Enterprise — Pacific Primary Instance | retire | Very High | $800,000 | $525,000 |
| 10 | APP-250 Azara DRVS — Atlantic Primary Instance 10 | retire | Very High | $780,000 | $512,000 |

## Consolidation

7 overlap groups, 7 with more than one member. Largest: CLU-07 (580 members); CLU-03 (4 members); CLU-02 (3 members); CLU-01 (2 members); CLU-04 (2 members). 93 rows were forced to consolidate by the redundancy override.

Support Role after normalisation: Secondary 992, Duplicative 677, Primary 128.

## Confidence

- high: 569
- Needs Validation: 29
- medium: 2

## Data gaps

- `c_consumption_price_variance`: unscored on 600 of 600 rows
- `r_end_user_perceived_quality`: unscored on 600 of 600 rows

## Vocabulary mappings applied

- Capability Map · Support Role / Coverage Level: Support Role 'Analyst' + Coverage Level 'Supporting' -> Support Role 'Secondary' (408 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Compliance analyst' + Coverage Level 'Supporting' -> Support Role 'Secondary' (408 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Business user' + Coverage Level 'Duplicate' -> Support Role 'Duplicative' (267 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Analyst' + Coverage Level 'Duplicate' -> Support Role 'Duplicative' (154 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Compliance analyst' + Coverage Level 'Duplicate' -> Support Role 'Duplicative' (154 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Clinical user' + Coverage Level 'Duplicate' -> Support Role 'Duplicative' (80 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Clinical user' + Coverage Level 'Supporting' -> Support Role 'Secondary' (70 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Business user' + Coverage Level 'Supporting' -> Support Role 'Secondary' (48 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Business user' + Coverage Level 'Limited' -> Support Role 'Secondary' (35 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Clinical user' + Coverage Level 'Primary' -> Support Role 'Primary' (30 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Business user' + Coverage Level 'Primary' -> Support Role 'Primary' (30 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Clinical user' + Coverage Level 'Limited' -> Support Role 'Secondary' (20 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Analyst' + Coverage Level 'Primary' -> Support Role 'Primary' (18 rows)
- Capability Map · Support Role / Coverage Level: Support Role 'Compliance analyst' + Coverage Level 'Primary' -> Support Role 'Primary' (18 rows)
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

- `Northstar-Disposition-Analysis-600-tuned.xlsx`
- `northstar-dispositions-600-tuned.csv`
- `Northstar-600-tuned-tool-vocabulary.xlsx`
- `northstar-600-tuned-tool-vocabulary.csv`
- this file
