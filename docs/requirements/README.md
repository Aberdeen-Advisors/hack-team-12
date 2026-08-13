# App Rationalization Agent — capability requirements

Working requirements sheet for hackathon prompt #5 (AI-Powered Application Rationalization Tool).

| File | What it is |
| --- | --- |
| `Application Rationalization Agent Requirements.xlsx` | Ryo's original upload, unmodified (Slack, 2026-08-13). Requirements 1–5 only. |
| `Application Rationalization Agent Requirements - filled.xlsx` | Same workbook with requirements 6–40 added in Ryo's structure, plus a `Notes & assumptions` sheet. |

The filled version keeps Ryo's `Sheet1` columns (Requirement ID / Category / AI Tool
Requirement Description) and his rows 1–5 exactly as written; new rows continue his
numbering. Added requirements cover the full lifecycle:

- **Inventory & discovery** (6–11) — multi-source ingest, app-name canonicalization, shadow IT, dependency mapping, AI-tool inventory
- **Gap detection & consolidation** (12–14) — extends Ryo's requirements 2 and 3
- **Enrichment** (15–20) — vendor/product metadata, contract terms, EOL, ownership, capability tagging, TCO reconstruction
- **Scoring** (21–24) — business value, technical health, cost efficiency, risk (technical and business kept separate)
- **Redundancy & disposition** (25–30) — overlap clustering, invest/consolidate/replace/retire with rationale and confidence
- **Roadmap & business case** (31–35) — wave sequencing, savings quantification, deliverable generation
- **UI and trust** (36–40) — upload surface, human-review queue, field-level provenance

MVP vs stretch is tagged inline in each description (`[MVP]` / `[STRETCH]`), because
Ryo's sheet has no priority column and his structure was left intact. Rows that depend
on data the team does not have are tagged `[DATA GAP: …]`, and judgement calls are
tagged `[ASSUMPTION: …]`.

**Open decisions** are listed on the `Notes & assumptions` sheet — the main one is the
synthetic-data question Bina raised: the Novant extract is missing most cost, contract,
risk, and capability fields, and requirements 16, 19, 20, and 24 all depend on filling
them with generated data.
