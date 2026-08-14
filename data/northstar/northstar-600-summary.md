# Northstar Global Health — 600-application run

Model: `score_northstar_v3.py`, imported unchanged. Input: `healthcare_app_rationalization_sample_600.xlsx`. Run date 2026-08-14. Wall clock 8.09s for 600 rows (2.78s to load, normalise, score and decide).

## Schema

Same 12 sheets and same 190 column headers as the committed 20-application `-with-risk` sample: 0 columns added, removed or reordered. The model runs unmodified. The divergence is in the value domains of the 580 new applications, handled by a documented normalisation layer — see `Vocabulary mapping` in the workbook.

## Dispositions

| Term | Count | % |
| --- | --- | --- |
| retain | 295 | 49.2% |
| invest | 173 | 28.8% |
| consolidate | 44 | 7.3% |
| replace | 7 | 1.2% |
| retire | 81 | 13.5% |

Only the all-pass pattern returns retain, so 295 rows (49.2%) need no action and 305 carry one.

## Money

- Gross annual avoidable claimed: $27,552,000
- One-time transition cost on those rows: $9,136,000 (33.2% of gross)
- **Net first year: $18,416,000**
- Safe (high-confidence only, her rule): $16,857,000
- Potential: $1,559,000
- Portfolio annual run cost: $354,330,000; her CIO target 15% = $53,149,500

## Priority

| Priority | Count |
| --- | --- |
| Very High | 6 |
| High | 92 |
| Moderate | 175 |
| Low | 37 |
| Very Low | 290 |

### Top 10 by priority then value at risk

| # | App | Disposition | Priority | Annual cost | Net first-year |
| --- | --- | --- | --- | --- | --- |
| 1 | APP-006 Zoom Workplace | retire | Very High | $1,130,000 | $370,000 |
| 2 | APP-327 Northstar Human Capital Management Midwest 06 | retire | Very High | $785,000 | $154,000 |
| 3 | APP-367 Northstar Healthcare Supply Chain Northeast 06 | retire | Very High | $655,000 | $119,000 |
| 4 | APP-347 Northstar Financial Management Southeast 06 | retire | Very High | $590,000 | $90,000 |
| 5 | APP-387 Northstar IT Service Management Central 06 | retire | Very High | $510,000 | $86,000 |
| 6 | APP-587 Northstar Administrative AI Agents Central 06 | retire | Very High | $295,000 | $47,000 |
| 7 | APP-002 Oracle Health EHR | consolidate | High | $6,500,000 | $2,200,000 |
| 8 | APP-018 Oracle PeopleSoft HCM | consolidate | High | $2,780,000 | $750,000 |
| 9 | APP-005 Slack Enterprise Grid | consolidate | High | $1,260,000 | $600,000 |
| 10 | APP-016 Tableau Cloud | consolidate | High | $1,040,000 | $530,000 |

## Consolidation

35 overlap groups, 35 with more than one member. Largest: CLU-07 (20 members); CLU-08 (20 members); CLU-09 (20 members); CLU-10 (20 members); CLU-11 (20 members). 7 rows were forced to consolidate by the redundancy override.

Support Role after normalisation: Secondary 1424, Duplicative 254, Primary 119.

## Confidence

- high: 569
- Needs Validation: 29
- medium: 2

## Data gaps

- `th_architecture_fit`: unscored on 54 of 600 rows
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
- App Inventory · Hosting Model: On-premises -> Customer-hosted on-premises (54 rows)
- App Inventory · Hosting Model: Private cloud -> LEFT UNSCORED — deliberately not mapped (54 rows)

## Outputs

- `Northstar-Disposition-Analysis-600.xlsx`
- `northstar-dispositions-600.csv`
- this file
