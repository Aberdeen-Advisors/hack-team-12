# App Rationalization Agent — capability requirements

Working requirements sheet for hackathon prompt #5 (AI-Powered Application Rationalization Tool).

| File | What it is |
| --- | --- |
| `Application Rationalization Agent Requirements.xlsx` | Ryo's original upload, unmodified (Slack, 2026-08-13). Requirements 1–5 only. |
| `Application Rationalization Agent Requirements - filled.xlsx` | Same workbook with requirements 6–65 added in Ryo's structure, plus a `Notes & assumptions` sheet. |

The filled version keeps Ryo's `Sheet1` columns (Requirement ID / Category / AI Tool
Requirement Description) and his rows 1–5 as written, with one exception: REQ 4's
disposition vocabulary was corrected in place on 2026-08-14 to the five agreed terms
(vocabulary only, no change of meaning). New rows continue his numbering.
**65 requirements total.**

## Organized around slide 13

Per Ryo's suggestion, the added requirements hang off the **Application Rationalization
Seven-Step Process** on slide 13 of Bina's Novant playbook. Every added row opens with a
`(Step N)` marker, and the `Notes & assumptions` sheet carries a full coverage map
including where Ryo's own requirements 1–5 land.

| Step | Requirements |
| --- | --- |
| 1. Discovery | REQ 1 (partial); added 8, 11, 14, 18, 59 |
| 2. Inventory applications & gather data | REQ 1, 2, 3; added 6–18, 20, 51, 53, 55, 57, 59, 61, 63 |
| 3. Analyze & map applications | added 19–26, 50, 52, 55, 59, 60, 62, 63, 65 |
| 4. Execute rationalization | REQ 4; added 25–30, 47, 50, 52, 58, 60, 64 |
| 5. Develop strategy & roadmap | added 16, 29, 31–36, 43, 54, 56 |
| 6. Execute rationalization initiatives | added 36–40, 53, 58, 64 |
| 7. Monitor & optimize results | added 41–45 |
| Cross-cutting | REQ 5; added 46–49 |

Steps 3, 6, and 7 were entirely uncovered by the original five requirements — scoring,
execution tracking, and benefit realization respectively.

## Tagging conventions

MVP vs stretch is tagged inline in each description (`[MVP]` / `[STRETCH]`), because
Ryo's sheet has no priority column and his structure was left intact. Rows that depend
on data the team does not have are tagged `[DATA GAP: …]`, judgement calls are tagged
`[ASSUMPTION: …]`, and requirements 51–65 close with `[REF: …]` naming the reference
framework each is traceable to.

## Decisions settled (2026-08-13, vocabulary superseded 2026-08-14)

- **Step 4 vocabulary — settled at five terms.** Decided by Bina Din, 2026-08-14:
  **retain, invest, consolidate, replace, retire**. `retain` = healthy, leave alone, no
  new spend. `invest` = fund a remediation or enhancement, a deliberate injection of
  money or effort. `consolidate` = fold into another application covering the same
  capability. `replace` = swap for a different product. `retire` = decommission. The
  engine emits the five directly; there is still no mapping layer to any other
  vocabulary. This **supersedes the four-term position of 2026-08-13** (invest,
  consolidate, replace, retire), under which `invest` did double duty for both a healthy
  application kept as-is and one carrying a funded remediation, separated only by
  priority — a distinction that did not survive being read aloud to a steering committee.
  Mechanically it is one row of REQ 52's lookup table: the all-pass key `PPPP`, formerly
  `invest` at Very Low priority, now returns `retain`, and every remaining `invest` row
  fails at least one gate. Requirements carrying it: 4, 27, 52, 64.
- **Scoring questions closed (2026-08-14, all Bina Din).** Risk stays substituted into
  the borrowed engine's end-user dimension slot, with end-user perception still collected
  but at weight 0. Cost moves priority but can never on its own make something a
  `retire`. Patient-care criticality takes the double weight in the business value
  dimension, with governance and compliance dropping to weight 1. Cost efficiency
  continues to use the modelled peer band — modelled from the portfolio itself, not
  measured against any external benchmark, and labelled as such in every output.
- **Disposition engine — specified (REQ 52).** Gate the four scored dimensions
  (business value, technical health, cost efficiency, risk) at 3.0 on a 1–5 scale,
  concatenate the pass/fail results into a `VTCR` key, and look that key up in a 16-row
  editable table returning a disposition plus a priority. The key doubles as the
  human-readable rationale. Thresholds and every table row are configuration, not code.
- **Dependency guardrail — promoted to MVP.** REQ 10 previously sat at `[STRETCH]`
  while the consolidation and disposition requirements that depend on it were `[MVP]`.
  A minimum dependency signal (`has_downstream_dependents` plus an
  enabling/dependent/overlapping flag per application pair) is now MVP, so no
  recommendation claims breakage safety that was not tested.
- **Cost and savings model — corrected.** REQ 20's cost model went from four components
  to eight (adding upgrade/module fees, indirect and training cost, consumption-based
  charges, and one-time implementation cost held separately). REQ 32 now computes *net*
  savings — less the replacement's ongoing cost, less amortised one-time cost, less
  residual archival/retention cost where a retention obligation outlives the
  application.
- **Scoring inputs partitioned.** Each signal now feeds exactly one scoring dimension;
  version currency, EOL and vendor viability previously scored in both technical health
  and risk, which over-recommended retiring old-but-adequate applications.

Remaining open questions (synthetic-data scale and generation approach, capability
taxonomy, demo scope) are listed in full on the `Notes & assumptions` sheet.

## Reference material

All five Info-Tech reference templates Bina shared were read on 2026-08-13 and are cited
on the `Notes & assumptions` sheet. They validated the scoring dimensions, the TCO
component taxonomy, and the disposition mechanism now in REQ 52. Two places where they
are genuinely out of date and we have to invent rather than borrow: there is no line
item anywhere for consumption/usage-based pricing, and every scoring input assumes a
human filling a 1–5 grid in a facilitated workshop — which is precisely the step this
agent removes. None of the reference material mentions AI at all, so the AI-specific
requirements (11, 26) have no external cross-check and need the most internal scrutiny.

The templates are licensed third-party vendor material and this repository is public, so
they are cited here and **never reproduced or committed**. Client operational data is out
of scope (Bina, 2026-08-13): no client application inventory, named individuals, contract
identifiers, cost actuals, server names, or departments appear in this workbook, and none
may be added. The demo runs on a synthetic ~600-application portfolio matching the
scenario in the overarching prompt.
