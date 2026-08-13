# App Rationalization Agent — capability requirements

Working requirements sheet for hackathon prompt #5 (AI-Powered Application Rationalization Tool).

| File | What it is |
| --- | --- |
| `Application Rationalization Agent Requirements.xlsx` | Ryo's original upload, unmodified (Slack, 2026-08-13). Requirements 1–5 only. |
| `Application Rationalization Agent Requirements - filled.xlsx` | Same workbook with requirements 6–50 added in Ryo's structure, plus a `Notes & assumptions` sheet. |

The filled version keeps Ryo's `Sheet1` columns (Requirement ID / Category / AI Tool
Requirement Description) and his rows 1–5 exactly as written; new rows continue his
numbering.

## Organized around slide 13

Per Ryo's suggestion, the added requirements hang off the **Application Rationalization
Seven-Step Process** on slide 13 of Bina's Novant playbook. Every added row opens with a
`(Step N)` marker, and the `Notes & assumptions` sheet carries a full coverage map
including where Ryo's own requirements 1–5 land.

| Step | Requirements |
| --- | --- |
| 1. Discovery | REQ 1 (partial); added 8, 11, 14, 18 |
| 2. Inventory applications & gather data | REQ 1, 2, 3; added 6–18, 20 |
| 3. Analyze & map applications | added 19–26, 50 |
| 4. Execute rationalization | REQ 4; added 25–30, 47, 50 |
| 5. Develop strategy & roadmap | added 16, 29, 31–36, 43 |
| 6. Execute rationalization initiatives | added 36–40 |
| 7. Monitor & optimize results | added 41–45 |
| Cross-cutting | REQ 5; added 46–49 |

Steps 3, 6, and 7 were entirely uncovered by the original five requirements — scoring,
execution tracking, and benefit realization respectively.

## Tagging conventions

MVP vs stretch is tagged inline in each description (`[MVP]` / `[STRETCH]`), because
Ryo's sheet has no priority column and his structure was left intact. Rows that depend
on data the team does not have are tagged `[DATA GAP: …]`, and judgement calls are
tagged `[ASSUMPTION: …]`.

## Open decisions

Listed in full on the `Notes & assumptions` sheet. The two that gate build work:

- **Step 4 vocabulary.** Slide 13 says retain / replace / retire; Ryo's REQ 4 says
  invest / consolidate / replace / retire. Not the same set, and it changes the output
  schema — see the sheet for the recommended reconciliation.
- **Synthetic data.** The Novant extract is missing most cost, contract, risk, and
  capability fields, per Bina. Requirements 10, 16, 19, 20, and 24 all depend on
  filling them with generated data.

The four Infotech reference templates Bina shared could not be read while building this
sheet; the scoring dimensions and TCO components are industry-standard but worth a
field-level cross-check against them.
