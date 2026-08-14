# What we need from the client, and what we work out ourselves

**Answering Bina's question: "which data variables do you expect as raw input from the client?"**

Aberdeen Advisors / hack-team-12 · Application Rationalization · as at 2026-08-14

---

## The short answer

The dataset has **125 columns**. A client is asked for **57 of them**. We derive the other **68**.

Of the 57, only **3 are truly unconditional**, and only **21 are needed for a first run** that still produces a defensible disposition and priority for every application.

And the single most useful thing to say to a client is this: **most of the intake is files, not a form.** We ask for five system extracts and a document folder. The 12 system-sourced columns in the minimum intake below are what we expect to *find in those extracts* — not what we expect somebody to type.

---

## Minimum viable intake — 21 items

The shortest list that still lets a rationalization run return a disposition and a priority for every application.

### A. Twelve columns out of systems (one spreadsheet is fine)

| # | Column | Plain language | Where it comes from |
|---|---|---|---|
| 1 | `app_id` | Any stable row key | Client's own export |
| 2 | `app_name` | What you call it | Client's own export |
| 3 | `vendor_name` | Vendor or publisher | Client's own export |
| 4 | `deployment_model` | SaaS / hosted / on-prem | ITSM or CMDB |
| 5 | `version_installed` | Version actually running | CMDB or software discovery |
| 6 | `implementation_date` | Go-live date | ITSM or CMDB |
| 7 | `business_owner` | Accountable business owner | ITSM/CMDB assignment group |
| 8 | `cost_centre` | Cost centre carrying the spend | General ledger / AP |
| 9 | `licences_purchased` | Entitlements bought | Licence portal or contract |
| 10 | `active_users` | Distinct users with real activity | Identity provider / SSO sign-in logs |
| 11 | `cost_licence_subscription` | Annual licence and subscription spend | General ledger / AP |
| 12 | `term_end` | Current contract term end | Contract register |

### B. One capability field

| # | Column | Plain language | Notes |
|---|---|---|---|
| 13 | `primary_capability` | Primary business capability | Use the client's own capability model if one exists; otherwise we infer it from a description and the confidence is recorded as lower |

### C. Eight answers a person has to give, per application

| # | Column | The question | Who answers |
|---|---|---|---|
| 14 | `ov_increase_value` | Does it carry money in or out? | Application owner, with finance |
| 15 | `ov_patient_care_criticality` | Does clinical work stop without it? | Clinical leadership |
| 16 | `ov_governance_compliance` | Regulatory and trust alignment | Security / privacy / compliance |
| 17 | `process_centrality` | How central to the process it serves | Application owner |
| 18 | `owner_stated_strategic_importance` | What the owner says it is worth | Application owner |
| 19 | `r_technical_risk` | Single points of failure, DR, hardening | Security, with the technical owner |
| 20 | `r_business_compliance_risk` | PHI exposure, residency, SOC 2 / HITRUST | Security / privacy / compliance |
| 21 | `r_clinical_safety_risk` | Patient-safety consequence of an outage | Clinical leadership / patient safety |

Eight answers per application is roughly a ten-minute conversation, and REQ 14 says to ask them as a targeted question set aimed at the named owner rather than as a 200-question survey.

### What degrades if that is all we get

The intake above fully scores three of the four gated dimensions. Everything below is a real, stated loss, not a caveat.

| What breaks | Why | Consequence |
|---|---|---|
| **Technical health rests on one criterion** | `th_supportability` (weight 2) is derivable from the version data; architecture fit (weight 2), operational stability, vendor viability and customization debt are all blank | The T gate becomes a version-currency proxy. It renormalises over 2 of its 7 weight, so a single fact decides a quarter of every disposition |
| **No notice-window warning** | `renewal_notice_days` and `auto_renewal_flag` are absent, so `notice_deadline_date` and `in_notice_window_now` cannot be computed | REQ 41 names "contract auto-renewed through the notice window" as a leading cause of savings that never land, and we would not see it coming |
| **Retire is legally untested** | No retention block, so `retention_override_applied` is always false | REQ 53 is explicit that retention expiry, not contract end date, decides whether an app can be switched off. Every retire recommendation would be unverified |
| **No early-termination check** | `early_termination_penalty` absent | REQ 29 cannot defer a retire whose penalty exceeds the saving |
| **Consolidation cannot claim breakage safety** | No integration inventory, so `has_downstream_dependents` is null | REQ 25 explicitly forbids publishing a cluster recommendation that was not checked against the REQ 10 dependency signal |
| **Cost is licence-only** | Four of the five cost categories fall back to REQ 55 estimation rules | On-prem TCO is understated, so on-prem savings are understated. Every modelled component is labelled *estimated*, not *actual* |
| **AI and metered spend is invisible** | `consumption_based_cost` absent | The dominant cost shape for the REQ 11 AI tools simply does not appear |
| **Nothing can reach high confidence** | The intake populates 23 of the 44 fields the scoring model consumes — a completeness score of about **0.52** | Every row lands at **medium** confidence. "High" requires 40 of 44 fields *and* all 16 weighted criteria populated, so it is unreachable on a minimum intake — by design (REQ 28) |

The guardrails do survive, which is worth saying: `lifecycle_stage` is inferable from `implementation_date` and `sourcing_type` from `deployment_model`, so REQ 51's bar on retiring an application that is merely still ramping stays armed.

---

## Tier counts

| Tier | What it means | Columns | Client effort |
|---|---|---|---|
| **1** | Client must provide — the tool cannot run without it | **3** | Three fields on every row |
| **2** | Provide if you have it — documented fallback exists | **11** | Optional; each one costs accuracy, not the run |
| **3** | Comes out of a source system, not a person | **29** | Five extracts, pulled once |
| **4** | Subject-matter judgement, collected from people | **14** | Interviews or a short form per app |
| **5** | Tool derives — nothing to send | **68** | None |
| | **Total** | **125** | 57 asked for, 68 derived |

### Why this differs from the dataset's own marker

The `Data dictionary` sheet marks **88 input / 37 computed**. That marker means "computed by the generator's engine functions" — it is not the same question as "does the client supply it". Reconciling the two:

- All **37** columns the dictionary calls computed are Tier 5. No exceptions.
- **31 of the 88** columns the dictionary calls *input* are also Tier 5, because in the product they are outputs of the enrichment agents rather than client fields. The dummy generator hand-seeds them only because it has no ingestion pipeline in front of it. Examples: `is_shadow_it` (REQ 8 diffs AP spend and SSO activity against the CMDB), `is_orphaned` (REQ 18 triangulation), `overlap_cluster_id` and `cluster_role` (REQ 25 clustering), `technical_obsolescence_flag` (REQ 17), all six AI columns (REQ 11/26), and `data_source` — which is an *output* of intake and can never be an input to it.
- That leaves **57** client-supplied columns, and the script agrees: nothing in that 57 is written by a function.

`saving_type` is the one genuinely hybrid field. It is seeded by an analyst but the engine overwrites it with "none" wherever gross saving is zero.

### A version note

The workbook on disk is **v1 (124 columns)**. The generator script beside it is **v2 (125 columns)** and is the design authority. This document is written against **v2**. Two differences matter:

- `ov_enhance_services` is renamed `ov_patient_care_criticality`, and the value dimension's double weight moves from governance-and-compliance onto patient-care criticality (Bina's v2 ruling). This changes *who we most need in the room*: the heaviest single criterion in the whole value dimension is now a question only clinical leadership can answer.
- `retain_or_invest_basis` is new, and there are now five disposition terms (retain, invest, consolidate, replace, retire) rather than four. Neither changes the intake.

---

## Tier 1 — Client must provide (3)

Nothing runs without these. Deliberately tiny.

| Column | Plain language | Format | Why it is unconditional |
|---|---|---|---|
| `app_id` | Row key | Any stable unique string | We mint our own canonical key after REQ 7 dedup, but every source row needs a handle so a number in the final deck can be traced back to it (REQ 48) |
| `app_name` | Application name | Free text | **Do not pre-clean it.** Inconsistent spellings are wanted — REQ 7 clusters "MSFT O365", "Office 365 E3" and "Microsoft 365 - Finance" and reports what it merged |
| `vendor_name` | Vendor or publisher | Free text | Load-bearing twice: REQ 7 clusters on name + vendor + domain, and contract and AP matching join on it. Marketing brand is fine; we resolve the legal entity |

---

## Tier 2 — Provide if you have it (11)

Each has a documented fallback, so its absence costs accuracy rather than the run. This is the tier to trade away under time pressure.

| Column | Plain language | Fallback if absent | What accuracy is lost |
|---|---|---|---|
| `description` | One line on what it does | Capability inferred from name + vendor alone | Capability confidence drops; REQ 19 warns that a wrong capability tag makes every downstream redundancy finding wrong |
| `primary_capability` | Primary business capability | REQ 19 inference from the description | The cost-per-user peer group and all redundancy clustering are built on this field |
| `secondary_capabilities` | Other capabilities covered | REQ 19 inference | Partial-overlap clusters are missed; only primary-capability duplicates surface |
| `cost_upgrade_and_modules` | Upgrade and module fees | REQ 55: a percentage of licence for on-prem COTS, zero for SaaS where bundled | Component labelled *estimated*; a CFO challenges the rule instead of the number |
| `cost_maintenance_dev_labour` | Internal support and dev labour | REQ 55: circa 0.25 FTE per COTS/SaaS app, 1.0 per custom app, scaled by user band × loaded rate | Labelled *estimated*, and it is usually the largest modelled component |
| `cost_infrastructure_peripherals` | Servers, storage, network, endpoints | REQ 55 allocates for on-prem and hosted only, off the REQ 57 footprint | On-prem TCO understated, so on-prem savings understated |
| `cost_indirect_and_training` | Help desk, training, staff premium | REQ 55: year-one percentage of licence plus a churn allowance | Labelled *estimated* |
| `one_time_implementation_cost` | Original implementation cost | Omitted | None to the disposition — REQ 20 keeps it out of the run-rate anyway. It only feeds the amortisation line |
| `residual_archival_cost` | Cost surviving switch-off | Modelled from data types and the REQ 57 footprint | **Net saving overstated on every retire** |
| `amortised_one_time_migration_cost` | Migration cost over five years | Implementation cost ÷ 5, or a modelled estimate | **Net saving overstated on every consolidate and replace** |
| `realization_lag_months` | Months to the first saved dollar | A default lag per disposition | Affects *when* the saving lands in the roadmap, not whether it exists |

---

## Tier 3 — Comes out of a source system (29)

This tier matters because it changes **who we ask**. Nobody fills these in by hand; somebody runs an export. REQ 6 expects exactly these four or five files dropped into a watched folder, schema auto-detected, with a load report showing what was rejected and which columns could not be mapped.

### From ITSM / CMDB (8)

`deployment_model`, `sourcing_type`, `lifecycle_stage`, `implementation_date`, `version_installed` (or software discovery — SCCM, Intune), `business_owner`, `technical_owner`, `data_types_held` (with the REQ 57 infrastructure footprint).

`sourcing_type` and `lifecycle_stage` are guardrails rather than scores: REQ 51 uses them to bar dispositions that are not legal (a SaaS product cannot be re-platformed) and to stop the engine retiring an application whose adoption has simply not ramped yet. Both are inferable where not stated.

### From the general ledger / AP (6)

`cost_centre`, `legal_entity`, `business_unit`, `department`, `cost_licence_subscription`, `consumption_based_cost`.

`cost_centre` is the field a saving is allocated against and how a CIO splits the 15% target across the organisation (REQ 59). `legal_entity` is the REQ 60 M&A duplication key — post-merger duplicate stacks are partly invisible without it, because each application is sanctioned within its own entity. `cost_licence_subscription` is the one cost component with **no acceptable fallback**: unused-licence spend and the entire savings baseline are computed from it.

### From the contract or vendor management register (8)

`contract_id`, `annual_contract_value`, `term_start`, `term_end`, `auto_renewal_flag`, `renewal_notice_days`, `licence_metric`, `early_termination_penalty`.

`term_end` is the highest-value field in the block — the renewal calendar, contract runway and the urgency half of priority all read it. `licence_metric` decides whether reclaiming seats saves anything at all: an enterprise agreement returns nothing per seat. The last two usually live in clause text rather than a register, which is what REQ 16's contract-PDF extraction is for, citation to clause included.

### From the identity provider / SSO sign-in logs (2)

`active_users`, `last_signin_date`. Please state the measurement window. With `licences_purchased` these two produce utilisation, unused seats, unused spend and cost per active user — the fastest credible saving available, and one that needs no migration project (REQ 23).

### From licence portals or vendor admin consoles (1)

`licences_purchased` — entitlements bought, deliberately not headcount.

### From integration / middleware inventory (1)

`integration_pattern`. The dependency edges themselves come from the same source but arrive as a file, not a column.

### From records management, legal and data governance (3)

`retention_obligation_flag`, `retention_expiry_date`, `information_classification`.

This block is the one most often forgotten and the one that most often blocks execution. REQ 53 is unambiguous: retention expiry, not contract end date, determines whether an application can be retired now or must first be moved to a read-only archive. `information_classification` accepts "unknown", which is counted as missing on purpose.

---

## Tier 4 — Subject-matter judgement (14)

These are the scored assessment inputs. No system holds them, and the person who answers changes per criterion — which is the practical scheduling problem in a rationalization engagement.

| Column | The question | Who answers | Weight |
|---|---|---|---|
| `ov_increase_value` | V1 — does it carry money in or out? | Application owner, with finance | 1 |
| `ov_patient_care_criticality` | V4 — does clinical work stop without it? | **Clinical leadership** | **2** |
| `ov_governance_compliance` | V5 — regulatory and trust alignment | Security / privacy / compliance, with the owner | 1 |
| `process_centrality` | How central to the process it serves | Application owner | feeds V3 |
| `owner_stated_strategic_importance` | What the owner says it is worth | Application owner | feeds V5 |
| `th_architecture_fit` | T2 — architecture and cloud fit | Technical owner / enterprise architecture | **2** |
| `th_operational_stability` | T3 — how much it breaks | Technical owner | 1 |
| `th_vendor_viability` | T4 — is the vendor going to be here? | Vendor management / sourcing | 1 |
| `th_customization_debt` | T5 — how far we have bent it | Technical owner | 1 |
| `r_technical_risk` | R1 — SPOF, DR, hardening | Security, with the technical owner | 1 |
| `r_business_compliance_risk` | R2 — PHI, residency, SOC 2 / HITRUST | Security / privacy / compliance | 1 |
| `r_clinical_safety_risk` | R3 — patient-safety consequence of an outage | **Clinical leadership / patient safety** | 1 |
| `r_end_user_perceived_quality` | R4 — what users think of it | End-user survey | **0** |
| `urg_risk_pain_severity` | How much pain the current state causes | Application owner / technical owner | feeds urgency |

Three things to flag before anyone fills these in:

- **All risk scores run "5 = controlled", not "5 = risky."** Getting the direction wrong inverts the entire risk gate. Put the direction on the form.
- **`r_end_user_perceived_quality` carries weight 0.** It is retained at the weight the reference template ships, so it never moves a disposition. Leave it off an intake.
- **`owner_stated_strategic_importance` is recorded as owner-attested, never as system-of-record** (REQ 14). Human answers keep their provenance visible so a stakeholder can argue with an input rather than with a black-box number.

Two of the four V criteria and one of the three weighted R criteria need a clinician. Under v2 the heaviest criterion in the value dimension is patient-care criticality, so **clinical leadership is now on the critical path of the intake, not a reviewer at the end of it.**

---

## Tier 5 — Tool derives (68)

Nothing here goes on an intake form. Grouped, with the derivation in one line.

### Enrichment findings — marked "input" in the dataset, derived in the product (31)

| Column | How we derive it |
|---|---|
| `is_shadow_it` | REQ 8: AP or expense spend, or SSO sign-in activity, with no CMDB record |
| `governance_visibility` | REQ 63: Managed / Unmanaged / Unknown / Unsanctioned from the REQ 61 presence flags plus owner and contract resolution |
| `is_orphaned` | REQ 18: neither owner resolvable after triangulating CMDB assignment, cost centre and heaviest SSO users |
| `technical_obsolescence_flag` | REQ 17: installed version behind the supported version, and end-of-support against the analysis date |
| `version_vendor_supported`, `vendor_eos_date` | Vendor lifecycle lookup by us. Not client fields at all |
| `capability_tag_confidence` | Emitted by the REQ 19 tagger; 1.0 only where the client stated the tag |
| `has_downstream_dependents`, `dependency_count` | REQ 10: rollup and edge count from the interface and middleware inventories |
| `overlap_cluster_id`, `cluster_role` | REQ 25: clustering on capability, feature coverage, user overlap and cost, then survivor selection on value, cost, users and contract runway (REQ 60 for M&A pairs) |
| `replacement_app_id`, `replacement_ongoing_tco`, `replacement_cost_already_in_baseline` | The cluster survivor; its own annual run-rate; and whether it is already a paid line here, so its cost is not subtracted twice |
| `action` | REQ 64: proposed from disposition, sourcing type and cluster role, then confirmed by a human |
| `saving_type` | Seeded by the analyst, then forced to "none" wherever gross saving is zero |
| `urg_timeline_sensitivity` | REQ 56: bands over term end, notice deadline and end-of-support |
| `data_source` | REQ 48/61: which sources this row was assembled from. An *output* of intake |
| the six AI columns | See the AI note below |
| plus `ov_reach_consumers`, `ov_reduce_costs_efficiency`, `th_supportability`, `c_cost_per_active_user_vs_peers`, `c_unused_licence_waste`, `c_consumption_price_variance`, `c_absolute_cost_band` | See "scored but not asked" below |

### Scored but not asked — 7 of the 18 criteria are rubric bands over data we already hold

This is the least obvious part of the answer and the part that most reduces what we ask a client for. Seven of the eighteen scored criteria are defined against signals the tool computes, so asking a human to score them would be asking twice:

| Criterion | Rubric band over |
|---|---|
| `ov_reach_consumers` (V2) | `active_users` vs `licences_purchased` |
| `ov_reduce_costs_efficiency` (V3) | the `process_centrality` answer |
| `th_supportability` (T1, weight 2) | installed vs vendor-supported version, and end-of-support proximity |
| `c_cost_per_active_user_vs_peers` (C1, weight 2) | cost per active user, as a percentile within the capability peer group |
| `c_unused_licence_waste` (C2) | licence utilisation rate |
| `c_consumption_price_variance` (C3) | metered spend against the REQ 55 modelled plan |
| `c_absolute_cost_band` (C4, weight 0) | annual run-rate band |

The whole cost-efficiency dimension is machine-scored. A client is never asked to rate their own cost efficiency.

### Arithmetic (14)

`licence_utilisation_rate` (active ÷ purchased) · `unused_licence_count` (purchased − active) · `tco_five_category_subtotal` (sum of the five categories) · `annual_tco_recurring` (+ consumption) · `five_year_cumulative_tco` (× 5, undiscounted) · `cost_per_active_user` · `unused_licence_spend` (licence × unused share) · `notice_deadline_date` (term end − notice days) · `in_notice_window_now` (deadline within 180 days) · `contract_runway_months` · `gross_saving_annual` · `net_saving_annual` (gross − successor run cost − amortised one-time − residual archival) · `net_saving_five_year` · `consolidation_saving` (on absorbed rows only, so nothing is double-counted).

### Scores, gates and the recommendation (23)

The four dimension scores (weighted means renormalised over the criteria that actually carry a value) · the four pass flags at the 3.0 gate · `vtcr_key` · `disposition` · `priority` · `retain_or_invest_basis` · `rationale` · `confidence` · `urgency_score` · the four guard flags (`redundancy_override_applied`, `lifecycle_exclusion_applied`, `sourcing_exclusion_applied`, `retention_override_applied`) · `suppressed_recommendation` and `suppression_reason` · `completeness_score` and `missing_fields`.

Two of these are worth naming to a client because they are the credibility features. `suppressed_recommendation` records every recommendation a guardrail changed, with its reason — REQ 51 requires that a suppressed recommendation with a stated reason is a better artifact than a silently different answer. And `missing_fields` names, per application, exactly which of the 44 model-consumed fields are empty, which is how REQ 28 keeps a thin-evidence recommendation from being presented at the same confidence as a well-evidenced one.

---

## What a realistic health system cannot supply cleanly

The requirements are candid about this. Five requirements carry an explicit `[DATA GAP]` marker recording that the source data available to the team did not contain the field at all — and those five gaps are precisely the tiers a real client will also struggle with.

| Where it breaks | The requirement's own words | What we do about it |
|---|---|---|
| **Contract terms** (Tier 3, 8 columns) | REQ 16: *"[DATA GAP: contract terms missing from the source data available to the team — synthetic contracts needed]"* | REQ 16 extracts them from contract and order-form PDFs with a citation to the clause. This is why the intake asks for a document folder, not just a spreadsheet |
| **Cost components** (Tier 2, 4 columns) | REQ 20: *"[DATA GAP: most cost components missing from the source data available to the team — documented estimation rules required]"* | REQ 55 publishes the estimation rule behind every modelled component, so a CFO challenges the rule rather than the number, and every component stays labelled actual / allocated / estimated |
| **Dependencies** (Tier 3/5) | REQ 10: *"[DATA GAP: no interface or integration inventory in the source data available to the team — needs synthetic dependency data]"* | REQ 10's MVP is deliberately minimal — a per-application `has_downstream_dependents` boolean plus an enabling/dependent/overlapping flag per pair — because REQ 25 and REQ 29 must not claim breakage safety they could not test |
| **Capability tags** (Tier 2) | REQ 19: *"[DATA GAP: no capability field in the source data available to the team]"* | Inferred from product documentation and descriptions, with a confidence score, and REQ 19 flags that if the tagging is wrong every overlap finding is wrong |
| **Risk attributes** (Tier 4) | REQ 24: *"[DATA GAP: technical, compliance and clinical risk attributes missing from the source data available to the team — synthetic risk attributes needed]"* | There is no shortcut. These are Tier 4 because a human has to answer them; the only lever is REQ 14's targeted question set instead of a mass survey |

Beyond the marked gaps, three structural problems are near-universal in a health system:

**Ownership is stale or absent.** REQ 18 does not assume the CMDB has an owner. It triangulates from CMDB assignment groups, the cost centre carrying the spend line, and the heaviest sign-in users, and where no owner can be established it marks the application **orphaned** — noting that orphaned plus low usage is consistently the highest-yield elimination candidate and the easiest to defend. REQ 47 then surfaces orphans in a human-review work list so a consultant clears exceptions in minutes rather than auditing the portfolio.

**Shadow IT means the inventory itself is incomplete.** REQ 8 does not take the CMDB as the population. It diffs SaaS spend in AP and expense data, and sign-in activity in SSO logs, *against* the sanctioned inventory, and outputs what is in use or being paid for but not governed — "typically the largest single block of unmanaged spend in a 15% reduction exercise." REQ 63 turns that into a portfolio-level split with the spend behind each category.

**Reconciliation fails in both directions.** REQ 61 is explicit that the reciprocal exception — a governed, catalogued application whose contract or cost cannot be located — is at least as common as shadow IT, and that it **blocks any credible savings claim on that application**. It is raised as its own work item naming the specific missing source.

And the handling for missing data is designed in rather than bolted on:

- **REQ 12** scores every record for completeness against the 44 fields the scoring model consumes, and ranks the gap list by *how much each missing field would move that application's disposition* — so the team chases the handful of fields that change answers instead of trying to complete every field on every app.
- **REQ 52** renormalises each dimension's weighted average over only the criteria that actually carry a value. This is a deliberate correction to the reference tool: a sparse row is reported as incomplete rather than being silently scored downward toward retire. Without it, missing data would look like a recommendation to switch things off.
- **REQ 28** marks a thin-evidence recommendation low-confidence and names the single piece of missing data that would resolve it.
- **REQ 14** generates a targeted question set for the named owner and folds the answers back marked *owner-attested* rather than *system-of-record*.
- **REQ 7** sends low-confidence name matches to human review instead of merging them silently; **REQ 48** keeps field-level provenance on every value so any number in the deck can be traced the moment a client challenges it.

**The practical read for Bina:** Tier 1 and Tier 3's ITSM block are easy. The contract register block and the records-retention block are where an engagement stalls, and neither sits with IT — they sit with legal, procurement and records management. Tier 4 needs named humans, and in a portfolio with orphaned ownership those humans have to be found before they can be asked. Both of those are scheduling problems as much as data problems, so they belong in the intake conversation on day one.

---

## The AI-agent fields

Six columns describe AI tools and agents: `is_ai_tool`, `ai_delivery_form`, `ai_capability_class`, `ai_host_app_id`, `ai_already_entitled_elsewhere`, `ai_entitled_alternative_app_id`. A seventh, `consumption_based_cost`, is the cost line they need.

**Do not put any of the six on an intake form.** They have no precedent in the reference templates and no client will have them in any system today. REQ 11 says so directly: AI tools and agents "rarely appear in the CMDB and overlap heavily." REQ 55 makes the matching point about the money: consumption and usage-based pricing "has no equivalent in the reference cost model and must be modelled from first principles, since it is the dominant cost structure for SaaS renewals and the primary one for the AI tools in REQ 11."

So all six are Tier 5 — we classify them. What we ask the client for instead is the **evidence that lets us find them**:

1. **AP and expense-report line items**, including corporate-card and expense-claim data, not just PO-backed spend. This is where standalone copilots and chatbots bought on a departmental card actually show up, and it is the input REQ 8 diffs against the CMDB.
2. **The SSO application list with sign-in activity**, which catches AI tools that were adopted without ever being purchased centrally.
3. **Enterprise licence entitlement documents** — the order forms and product-edition detail for the large platform agreements already in place. This is the input REQ 26 needs to prove that a capability being bought standalone is *already entitled* inside a licence the organisation holds. That finding is a saving with no functional loss at all, which makes it the most persuasive item in the deck, and it is uncomputable without the entitlement paperwork.
4. **Vendor billing portal exports** for anything metered, so `consumption_based_cost` is an actual figure rather than a modelled one.

The `ai_delivery_form` distinction — standalone product versus AI feature inside an app already licensed — is the whole point of items 1 and 3, and it cannot be answered from a client's inventory. It has to be assembled from spend plus entitlements.

---

## Two schema notes for the team

Neither is in scope to fix here; both are worth a decision before the next dataset build.

1. **`th_operational_stability` has no underlying data column.** The criterion is defined as "incident and ticket volume; proactive vs reactive maintenance" — which is an ITSM extract, not a human judgement. There is no incident-count or ticket-volume column anywhere in the 125, so today it can only be collected as a Tier 4 opinion from the technical owner. Adding an incident-volume field would move a weighted criterion out of the interview burden and into the extract, which is a straight win for intake effort.

2. **The intake asks for things that are not columns.** REQ 6 expects four to five structured extracts in a watched folder and REQ 9 expects a document repository — architecture diagrams, contracts, SOWs, security reviews, support handbooks. A per-column checklist cannot express either, so the checklist workbook that accompanies this document should be handed over *with* a file-level request, not instead of one.
