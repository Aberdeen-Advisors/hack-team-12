# Scoring model review — response to Bina Din's five points

**Scope.** Everything below was read out of `engine/generate_dataset.py` (the `CRITERIA` block and
`gate()`; the line numbers cited throughout are those of the revision reviewed, and have since
shifted) and cross-checked against the `Scoring model` sheet of
`App-Rationalization-Dummy-Dataset-v2.xlsx`. The two agree exactly. The sensitivity run was done
in a scratch copy of the generator; `generate_dataset.py`, the workbook and the CSV are unchanged.

**How to read this.** Part 1 is the arithmetic — what the weights actually are. Part 2 answers the
question behind the review: *would re-weighting change any recommendation?* Part 3 is the honest
list of weak points.

**The one-line answer up front.** No. Not one of the six re-weightings Bina's comments point at
changes a single disposition or a single priority on any of the 20 applications. The dimension
scores move — up to 0.30 — but every application stays on the same side of the 3.0 gate. The
weights she is questioning are, on this roster, not load-bearing. Two weights that *are*
load-bearing are the two currently set to zero.

---

## Part 1 — the weights, as the file actually has them

Each dimension is a weighted average of its own inputs. The *normalised* weight is the raw weight
divided by that dimension's weight sum, and is the share of the dimension the input actually
controls. Weight 0 means the input is collected and stored but contributes nothing.

Each of our eighteen inputs sits in a slot that corresponds to a criterion of the licensed
Info-Tech reference tool whose structure the gated model follows. That tool is cited, not
reproduced — it is licensed third-party material and this repository is public — so its
criterion names are not restated in the tables below, and the four dimensions are named in our
own terms throughout. Where the correspondence matters to an argument, it is described rather
than named.

### Business value (V) — weight sum 6

| Input | Raw weight | Normalised |
|---|---|---|
| `ov_increase_value` | 1 | **0.1667** |
| `ov_reach_consumers` | 1 | 0.1667 |
| `ov_reduce_costs_efficiency` | 1 | 0.1667 |
| `ov_patient_care_criticality` | **2** | **0.3333** |
| `ov_governance_compliance` | 1 | 0.1667 |

### Technical health (T) — weight sum 7

| Input | Raw weight | Normalised |
|---|---|---|
| `th_supportability` | **2** | 0.2857 |
| `th_architecture_fit` | **2** | 0.2857 |
| `th_operational_stability` | 1 | 0.1429 |
| `th_vendor_viability` | 1 | 0.1429 |
| `th_customization_debt` | 1 | 0.1429 |

Two of the five technical-health slots are re-bound: the reference tool ships something else in
them, and we carry version currency and architecture fit there instead.

### Cost efficiency (C) — weight sum 4

| Input | Raw weight | Normalised |
|---|---|---|
| `c_cost_per_active_user_vs_peers` | **2** | **0.5000** |
| `c_unused_licence_waste` | 1 | **0.2500** |
| `c_consumption_price_variance` | 1 | 0.2500 |
| `c_absolute_cost_band` | **0** | 0.0000 |

### Risk posture (R) — weight sum 3

| Input | Raw weight | Normalised |
|---|---|---|
| `r_technical_risk` | 1 | 0.3333 |
| `r_business_compliance_risk` | 1 | 0.3333 |
| `r_clinical_safety_risk` | 1 | 0.3333 |
| `r_end_user_perceived_quality` | **0** | 0.0000 |

The whole R block is re-bound: the reference tool ships an end-user-perspective lens in this
position, and we carry risk here instead. Only the fourth input, end-user perceived quality,
still measures what that lens measures — and it is the one sitting at weight 0.

**Both of Bina's figures are confirmed.** `c_cost_per_active_user_vs_peers` is raw weight 2 out of a
cost weight sum of 4, so its normalised weight is exactly **0.5** — it controls half the cost
dimension on its own. `c_unused_licence_waste` is raw weight 1 of 4, so **0.25** is right. Note the
weight sum is 4, not 3: `c_absolute_cost_band` is counted in the denominator even though it carries
weight 0, which is how the generator computes it (line 2441) and how the workbook displays it.

A note on direction, because it inverts twice: on the cost inputs a **high score means low cost**
(5 = cheapest, on plan, almost no waste). On the risk inputs a **high score means low risk**
(5 = controlled). So a *high* number is always the *good* number, on every one of the 18 inputs.

---

## Part 2 — Bina's five points, answered

### 1. `ov_increase_value` — "not a priority for Healthcare"

**Current weight 1, normalised 0.1667 — one sixth of the value dimension.**

What it measures in our implementation, verbatim from `CRITERIA` (line 100):

> "REQ 21 criticality to revenue: does the app carry money in or out?"

The data dictionary phrases it the same way: *"V1. Criticality to revenue: does the app carry money
in or out?"*

The reference tool's own wording for the slot is broader than ours. Its prompt asks, in general
terms, how much financial value the application delivers to the organisation and to its customers,
and it scores that on a generic five-step correlation ladder shared with two of the other value
criteria. That wording is the licensed tool's and is not reproduced here; what matters for the
argument is only that it is a general commercial-value question rather than a health-system one.

**Assessment.** Bina is right that a general "increase value" axis — financial value and
customer value, as the reference tool frames it — is not the axis a health system leads with. But our implementation has already
narrowed it to something a health system does care about and does not score anywhere else:
*revenue capture*. Waystar (claims and revenue-cycle) and Solventum 360 Encompass (coding) are the
rows that depend on it. If the criterion goes to zero, the portfolio loses its only signal for
"this application is how money arrives", and revenue-cycle applications would be judged purely on
patient-care criticality, where they legitimately score lower than clinical systems. The safer
change, if Bina wants the emphasis moved, is to **rename it to what it actually measures**
(`ov_revenue_capture_criticality`) rather than to zero it. That is a labelling change with no
arithmetic consequence — and S1 below shows zeroing it has no consequence either.

### 2. `ov_governance_compliance` — does it reference HIPAA and PII?

**Current weight 1, normalised 0.1667. No — it names no specific regulatory regime.**

Our definition (line 109):

> "REQ 21 regulatory/trust alignment plus owner-stated strategic importance."

The reference tool's wording for the same slot is likewise generic: it asks how well the application
aligns to regulation, builds trust and reputation and mitigates audit risk, and its five anchors run
from fully compliant and verified down to critical or undefined compliance gaps. No specific regime
is named anywhere in it. (That wording belongs to the licensed tool and is deliberately paraphrased
rather than quoted here.)

So it scores *generic compliance maturity plus what the owner says the application is worth
strategically*. It does not name HIPAA, PHI, PII, HITRUST, SOC 2 or data residency anywhere.

**HIPAA and PHI are already scored — in the risk dimension, once.** `r_business_compliance_risk`
(line 135) is defined as:

> "REQ 24 PHI/HIPAA exposure, data residency, SOC 2 / HITRUST, lock-in. 5 = controlled."

REQ 24 assigns those facts to risk exclusively: *"REQ 24 owns single points of failure, absent DR or
backup, single-vendor concentration, unhardened configuration, PHI and HIPAA exposure, data
residency, SOC 2 / HITRUST posture and contractual lock-in."*

**This matters, because it is the defect the team already fixed once.** REQ 22 and REQ 24 carry an
explicit warning that scoring one underlying fact in two dimensions *"fails two of four gates on a
single finding and systematically over-recommends retiring old-but-adequate applications."* That is
exactly what trap case T1 (Sunquest CoPath Plus) exists to catch. So the recommendation is: **do not
re-point `ov_governance_compliance` at HIPAA/PHI, and do not raise its weight on the grounds that it
covers HIPAA.** It does not, and if it were made to, the double-count returns. If Bina wants HIPAA
weighted more heavily, the correct lever is `r_business_compliance_risk` inside the risk dimension —
which, per the leave-one-out test in Part 3, is already the single most powerful criterion in the
whole model.

On the weight history Bina refers to: correct. `ov_governance_compliance` held weight 2 in v1 and
dropped to 1 in v2 when `ov_patient_care_criticality` took the double-weight slot, on Bina's own Q3
ruling. The value weight sum is 6 either way, so v1 and v2 value scores are directly comparable, and
the generator's note records that the largest single move was 0.17 and no pass/fail result changed.
S2 below re-tests putting governance back to 2.

### 3. `c_cost_per_active_user_vs_peers` — why is it the dominant cost input?

**Weight 2 of 4, normalised 0.5 — it alone controls half the dimension. `c_absolute_cost_band`
sits at weight 0.**

The rationale is that REQ 23 defines the dimension as cost efficiency, not cost:

> "Score cost efficiency as TCO against value delivered: cost per active user benchmarked against
> peer applications in the same capability, and licensed-vs-active seat waste."

The generator's own justification for the weight-0 on absolute spend (the `Scoring model` sheet):

> "the absolute-dollar cost band (because absolute cost fails every large enterprise
> system on sight, while the requirement defines cost EFFICIENCY relatively — cost per active user
> against peers)."

Epic Hyperspace is the proof. It is the largest line item in the portfolio, and its
`c_absolute_cost_band` is **1.0** — the worst possible score, because the band is scored on a flat
dollar ladder whose worst step is simply "above a million dollars a year". Its cost *per active user* is 3.5,
mid-band for its peer group, because 17,240 of 18,500 seats are active. On absolute spend Epic looks
like the portfolio's biggest problem; on cost efficiency it is unremarkable. Trap case T12 exists
specifically to catch the wrong answer here — *"Recommending action against the largest line item."*
Part 3 shows that enabling the absolute-cost band at weight 2 does exactly that.

**One caveat that must travel with the number.** The peer band is **modelled, not measured**
(assumption A5). This portfolio has one laboratory application and one ERP, so those two rows have
no internal peer group at all and their scores are set against an assumed external band. Bina's
answer to Q4 was "continue to use the model", so the band stands — but the single most influential
input in the cost dimension is currently a modelled comparison, and sourcing a real benchmark is
still open (O2). That is a bigger exposure than the weight itself.

### 4. `c_unused_licence_waste` and `c_consumption_price_variance` — should they change?

Both are **weight 1, normalised 0.25**.

**`c_unused_licence_waste`** — bears on REQ 23 (which requires unused-licence spend as its own
explicit dollar line), REQ 42 (licence counts creeping back up) and REQ 43 (KPI baseline). REQ 23
calls this *"the fastest credible saving available against a 15% reduction target"* and notes it
*"needs no migration project."* There is a genuine argument for raising it: it is the one cost signal
that converts to cash without a project. The argument against is that it is already fully reported
as a dollar line item (`unused_licence_spend`) and as its own saving type, so its influence on the
*score* is not what makes it actionable. **View: leave at 1.** Raising it changes nothing (S4) and
would slightly dilute the peer-cost comparison, which is the input REQ 23 names first.

**`c_consumption_price_variance`** — bears on REQ 20 and REQ 55, which add consumption/usage-based
pricing as a cost component *with no precedent in any Info-Tech template*: *"Consumption / usage-based
pricing has no equivalent in the reference cost model and must be modelled from first principles,
since it is the dominant cost structure for SaaS renewals and the primary one for the AI tools in
REQ 11."* This is the team's own differentiator, it occupies a slot the reference tool uses for a
conventional recurring-cost line, and the AI rows (REQ 11, REQ 26) are priced this way. **View: this is the stronger candidate for a raise
to 2** — not because it changes any answer today (it does not, per S3), but because on a real 600-app
portfolio with consumption-priced AI tooling it is the input most likely to be doing real work, and
weight 1 of 4 understates a cost structure the reference model does not even have a line for. Waystar
(APP-005) is the row that demonstrates it: its evidence text says per-transaction charges are running
well above the peer benchmark and *"cost efficiency fails on both cost per active user and
consumption variance."*

### 5. Bina's open question — "are there other fields we should reconsider?"

Answered in Part 3.

---

## Part 3 — sensitivity run

Six scenarios, each recomputing all 20 rows from the same input scores with only the weights changed.

| Scenario | Change | Dispositions changed | Priorities changed | Which apps | Resulting spread |
|---|---|---|---|---|---|
| **Baseline** | as shipped | — | — | — | invest 6, retain 3, consolidate 6, replace 2, retire 3 |
| **S1** | `ov_increase_value` → 0 | **0** | 0 | none | invest 6, retain 3, consolidate 6, replace 2, retire 3 |
| **S2** | S1 + `ov_governance_compliance` → 2 | **0** | 0 | none | invest 6, retain 3, consolidate 6, replace 2, retire 3 |
| **S3** | `c_consumption_price_variance` → 2 | **0** | 0 | none | invest 6, retain 3, consolidate 6, replace 2, retire 3 |
| **S4** | `c_unused_licence_waste` → 2 | **0** | 0 | none | invest 6, retain 3, consolidate 6, replace 2, retire 3 |
| **S5** | cost flattened to 1/1/1 | **0** | 0 | none | invest 6, retain 3, consolidate 6, replace 2, retire 3 |
| **S6** | S1 + S3 together | **0** | 0 | none | invest 6, retain 3, consolidate 6, replace 2, retire 3 |

**The spread is identical in all six scenarios.** No application changes disposition, no application
changes priority, and no application's four-character VTCR key changes.

The scores genuinely do move — this is not a harness that failed to apply the change:

| Scenario | Rows whose dimension score moved | Largest single move |
|---|---|---|
| S1 | 19 | 0.083 |
| S2 | 12 | 0.167 |
| S3 | 17 | **0.300** |
| S4 | 17 | 0.150 |
| S5 | 16 | 0.167 |
| S6 | 36 | 0.300 |

**Why nothing flips.** The margin to the 3.0 gate is far larger than any of these moves. The
smallest margin anywhere in the value dimension is Spok Mobile at +0.333; the smallest in cost is
±0.500. A re-weight would have to move a score by more than that, and the largest move available
from any of these six changes is 0.300.

**How hard would you have to push?** Each of the 18 weights was swept from 0 to 12 individually.
For **fourteen** of them, including every input in the value dimension and all three non-zero cost
inputs, **no weight anywhere in 0–12 flips a single gate on a single row.** Setting
`ov_patient_care_criticality` as the *sole* value input, or zeroing
`c_cost_per_active_user_vs_peers` entirely, also changes nothing. Only four weights can move an
answer at all, and three of those are in the risk dimension or currently at zero (see Part 4).

### Trap cases — all six scenarios clear

| Trap | Requirement | S1 | S2 | S3 | S4 | S5 | S6 |
|---|---|---|---|---|---|---|---|
| T1 Sunquest CoPath Plus (APP-017) | must be `invest`, not retire *and not retain* | PFPP invest / Moderate | same | same | same | same | same |
| T4 shadow IT, Otter.ai (APP-010) | must stay `retire` / Very High | FFFF retire / Very High | same | same | same | same | same |
| T14 cost-only failures (APP-005/011/019) | all three PPFP `invest` | holds | holds | holds | holds | holds | holds |
| T12 Epic Hyperspace (APP-001) | must be `retain` / Very Low | holds | holds | holds | holds | holds | holds |

No scenario breaks a trap. Two *out-of-scenario* pushes do, and are recorded here as guardrails
rather than as findings against the current model:

- `th_operational_stability` at weight **5** (from 1) flips Sunquest CoPath Plus from `PFPP invest`
  to `PPPP retain` — breaking T1 in the direction the trap sheet also forbids ("must be invest and
  NOT retain"). Sunquest's operational stability is 4.0 precisely because the application is stable;
  over-weighting stability would let stability mask the version-currency failure that *is* the
  remediation plan.
- `c_absolute_cost_band` at weight **2** (from 0) flips Epic Hyperspace from `retain` to `invest`,
  breaking T12. See Part 4, item 1.

---

## Part 4 — what else should be reconsidered

Ordered by how likely a health-system reviewer is to challenge it.

### 1. The two weight-0 inputs are not inert — they are the only weights that can change an answer

This is the most surprising finding in the review. Fourteen of the eighteen weights cannot change an
answer at any weight from 0 to 12. Both weight-0 inputs can change an answer **at weight 1**.

- **`c_absolute_cost_band` at weight 1** flips Otter.ai's cost gate (FFFF → FFPF) and drops its
  priority from Very High to High — because Otter.ai is cheap in absolute terms. At weight 2 it flips
  **Epic Hyperspace from `retain` to `invest`**, which is precisely the wrong answer trap T12 was
  written to catch: recommending action against the largest line item because it is the largest line
  item. The weight-0 decision here is correct and load-bearing, and it should be defended in the
  deck rather than buried.
- **`r_end_user_perceived_quality` at weight 1** flips Luma Health from `invest` to `retain`. Luma's
  end-user score is 4.0 while its `r_business_compliance_risk` is **1.5**. Turning the satisfaction
  input on would let a good satisfaction number dilute a genuine compliance failure into a pass. This
  is a real hazard: the input is the one Info-Tech itself ships at weight 0, and it is sitting in the
  *risk* dimension because that slot was re-bound. A reviewer who reads "end-user perceived quality"
  next to three risk criteria will reasonably ask why it is there at all. **Suggestion: move it out of
  the R block entirely** and hold it as an unscored reported field, rather than as a weight-0 member
  of the risk dimension where switching it on silently weakens risk.

### 2. `th_operational_stability` is defined by data that does not exist

Its definition is *"REQ 22 incident and ticket volume, proactive vs reactive maintenance"* and the
data dictionary repeats *"Incident and ticket volume."* **There is no incident-count or ticket-volume
column anywhere in the 125 columns** of `applications-v2.csv`. Nor is one collected on the client
intake. So the criterion is an expert judgement wearing the label of a measurement. This is the
easiest point for a health-system reviewer to puncture — an operations director will ask which
ITSM queue the number came from. Two honest options: add an `incident_count_12mo` /
`ticket_volume_12mo` intake field and derive the score from it, or restate the definition as a
qualitative maintenance-posture assessment and label it as such. Note also that the reference tool's
own anchor for this slot is quantitative and specific — it asks for a proactive-maintenance
percentage and current-release status — which makes the missing denominator more conspicuous, not
less.

### 3. The premise "cost only moves priority" is not quite what the engine does

Worth correcting before it is said out loud to a client. What is true, and is `retain_or_invest()`'s
explicit design (line 1528), is that a **cost-only** failure cannot kill an application: PPFP
resolves to `invest`, never to `retire`. But cost is not priority-only. Forcing every cost input to
pass changes **9 of 20 rows** — four move `invest → retain`, two move `replace → invest`, and Aidoc
moves **`retire → consolidate`** (FPFP → FPPP). Forcing every cost input to fail moves three rows
`retain → invest` and Sunquest CoPath Plus `invest → replace`. So cost efficiency does change
dispositions whenever it is combined with a value failure, and the accurate statement is: *a cost
failure on its own buys a remediation; a cost failure on top of a value failure is what separates
retire from consolidate.*

### 4. One criterion decides one gate on its own — and it is the HIPAA one

A leave-one-out test (zero each criterion in turn, recompute all 20 rows) produces exactly **one**
gate flip across the entire model: dropping **`r_business_compliance_risk`** flips Luma Health from
`PPPF invest` to `PPPP retain`. Luma's risk score is 2.667, built from 3.0 / **1.5** / 3.5 — the
compliance criterion alone is what fails the gate. With three equally weighted inputs and a 3.0 gate,
any single score of 1.5 or below fails the risk dimension almost unaided. That is arguably correct
behaviour for a PHI exposure, but it should be stated as a deliberate design property rather than
discovered by a reviewer, and it is the concrete reason why Bina's instinct about HIPAA weighting is
right while the *field* she pointed at (`ov_governance_compliance`) is the wrong lever.

Separately, Otter.ai's cost dimension rests on a **single populated input**
(`c_consumption_price_variance`); its other two cost scores are null. Its cost score is therefore
weight-invariant, which is why S3, S4 and S5 cannot touch it — and why enabling the absolute-cost
band changes it immediately. The renormalise-over-populated-inputs rule (REQ 52's correction to the
template) is doing its job, but on a sparse row it can leave a dimension resting on one number, and
the row does not currently say so beyond its `confidence: medium` flag.

### 5. The 3.0 gate is not discriminating on three of four dimensions

Actual distribution of the 20 rows, and how many clear the gate:

| Dimension | Min | Median | Max | Pass | Fail | Within 0.5 of 3.0 | Empty band straddling 3.0 |
|---|---|---|---|---|---|---|---|
| Business value | 1.583 | 4.083 | 4.833 | 16 | 4 | 2 | **2.500 → 3.333** (0.83 wide) |
| Technical health | 1.643 | 4.000 | 4.714 | 14 | 6 | 3 | 2.750 → 3.143 (0.39 wide) |
| Cost efficiency | 1.250 | 2.250 | 4.375 | **6** | **14** | 4 | **2.500 → 3.500** (1.00 wide) |
| Risk posture | 2.167 | 3.500 | 4.167 | 16 | 4 | 10 | 2.667 → 3.167 (0.50 wide) |

Two readings, and both should be said out loud:

- **Cost efficiency is the only dimension where the gate bites.** 14 of 20 rows fail it, and it is
  the reason so many rows land on `invest` and `consolidate`. Value and risk pass 16 of 20; technical
  health passes 14 of 20. If the deck claims four independent lenses, the honest caveat is that cost
  is carrying most of the discrimination.
- **Every dimension has an empty band straddling 3.0.** No row sits between 2.500 and 3.333 on
  value, or between 2.500 and 3.500 on cost. Sweeping the threshold itself confirms it: **the
  disposition set is byte-identical for every threshold from 2.75 to 3.5.** It only starts moving at
  2.5 (two rows change) and at 4.0 (four rows change). So the 3.0 gate has roughly ±0.3 of slack in
  either direction and is neither too easy nor too harsh — but that robustness is partly a property
  of how this synthetic dataset was authored, with clearly-good and clearly-bad rows and nothing
  ambiguous in between. **On a real 600-application portfolio the mass will sit in that empty band,
  and the sensitivity result in Part 3 will not hold.** The correct claim to make on Friday is "the
  model is stable against re-weighting *on this roster*", not "the weights do not matter."

### 6. Smaller points worth a line each

- **`ov_increase_value` is misnamed for the audience.** It measures revenue-capture criticality but
  its name still echoes the generic commercial-value label the reference tool uses for the slot,
  which is exactly what triggered Bina's comment. Renaming costs nothing and removes the objection.
- **`ov_governance_compliance` bundles two unlike things** — regulatory alignment *and*
  owner-stated strategic importance. One is an audit fact, the other is an opinion the owner
  volunteered. They are averaged into one 1–5 score with no way to tell which drove it.
- **The re-bound slots are a presentation risk.** All four inputs in the R block sit in slots the
  reference tool uses for end-user-perception criteria rather than for risk. That is a documented
  and supported customisation, but anyone reading that tool alongside our workbook will see
  clinical safety risk sitting in a slot its own documentation describes in end-user terms, and
  will want an explanation.
- **The cost weight sum is 4 with only three live inputs**, because the weight-0 band is counted in
  the denominator. The normalised weights (0.5 / 0.25 / 0.25) are internally consistent and sum to
  1.0, so nothing is wrong — but a reviewer adding 2+1+1 and getting 4 while seeing three inputs will
  ask, and the answer should be ready.

---

## Recommendation

**Do not re-weight for this Friday.** Six scenarios, zero changed recommendations. Re-weighting would
cost the team a re-run of every artifact and buy no different answer, and S2 in particular would
re-open a double-count the team has already closed once. Bina's instincts are sound but they point at
inputs that are not load-bearing.

**Do make three changes that cost nothing arithmetically:**

1. Rename `ov_increase_value` to say that it measures revenue-capture criticality.
2. Say explicitly, in the `Scoring model` sheet, that HIPAA/PHI is scored once, in
   `r_business_compliance_risk`, and that `ov_governance_compliance` is deliberately generic — with
   the double-count warning attached.
3. Add the caveat from Part 4 item 5 wherever the sensitivity result is quoted: stable *on this
   roster*, because this roster has an empty band around the gate.

**Two things to fix before a real engagement:** the missing incident/ticket data behind
`th_operational_stability`, and the modelled-not-measured peer band behind the single most
influential cost input.
