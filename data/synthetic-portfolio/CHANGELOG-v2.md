# What changed in v2, and why

`App-Rationalization-Dummy-Dataset-v2.xlsx` | `applications-v2.csv` | v2 (2026-08-14). Replaces `App-Rationalization-Dummy-Dataset-v1.xlsx`.

Bina Din, the subject-matter expert, reviewed v1 and answered the five open questions it carried. Two of her answers changed the file. Her exact words are on the 'Notes & assumptions' sheet, each next to the change it produced.

## 1. There are five recommendation words now, not four

Bina's words: *"No, separate by invest, retain, consolidate, replace, and retire."*

In v1 the word **invest** had to do two jobs. It meant "this application is healthy, leave it alone" AND it meant "we are putting money into this application". The only thing telling them apart was the priority column, which is a lot to ask of a reader in a steering committee. So v1 said *invest, priority Very Low* about Epic Hyperspace - the single largest line item in the portfolio - when what it meant was *leave it alone*.

v2 has a separate word for each:

| Word | What it means |
|---|---|
| **retain** | Healthy. Leave it alone. Spend nothing. |
| **invest** | Deliberately put money or effort in - fund a remediation or an enhancement. |
| **consolidate** | Fold it into another application that keeps the capability alive. |
| **replace** | Swap in a different product for the same capability. |
| **retire** | Switch it off. The capability goes away, or is already covered elsewhere. |

### How the tool decides between retain and invest

One rule, read straight off the four pass/fail gates:

- **All four dimensions pass** -> there is no failing dimension to fund, so nothing needs money. **retain**.
- **Any dimension fails** -> that dimension is exactly what the money would buy. **invest**, and the recommendation names it.

That is why an invest can always answer "invest in what?" - risk, or cost efficiency, or technical health. It also means only one of the sixteen possible pass/fail patterns changed meaning between v1 and v2: the all-pass one. Every other pattern already had a failing dimension to point at, so invest was always the honest word for it.

Priority no longer carries any part of this distinction. Priority is only about urgency now.

### Which applications moved

Three of the twenty, all of them from invest to retain, and no priority changed:

| Application | v1 said | v2 says | Why |
|---|---|---|---|
| APP-001 Epic Hyperspace | invest / Very Low | **retain** / Very Low | PPPP: all four gates pass and nothing is being funded, so v1's invest / Very Low was really retain all along. The largest line item in the portfolio, and the right answer is still to leave it alone. |
| APP-012 TigerConnect | invest / Very Low | **retain** / Very Low | PPPP: cluster survivor, all four gates pass. Keeps its gate disposition under D2, and that disposition is now retain rather than a Very Low invest. |
| APP-014 Microsoft Power BI | invest / Very Low | **retain** / Very Low | PPPP: cluster survivor, all four gates pass. Same as APP-012 - retain reads correctly where 'invest in Power BI, priority Very Low' did not. |

Two of those three are cluster survivors - the application a group of overlapping products gets folded into. v1 called them *invest, priority Very Low*, which sounds like a funding request against an application nobody was proposing to spend money on. **retain** says what is actually meant: this is the one we are keeping, and the money is being spent on moving the others onto it. That reads considerably better than v1 did.

The seventeen other applications say exactly what they said in v1.

## 2. Patient-care criticality now counts double in the value score

Bina's answer to the question of which value signal a health system should weigh twice was **yes** - patient-care criticality, not the scoring engine's own 'governance and compliance' criterion. So:

- Patient-care criticality moves from weight 1 to **weight 2**.
- Governance and compliance moves from weight 2 to **weight 1**.
- The column that holds the score is renamed from `ov_enhance_services` to `ov_patient_care_criticality`, so it says what it measures. This is the one column name that changed between v1 and v2.

Every business value score was recalculated. The largest single change was 0.17 of a point, and **no application's pass/fail result changed** - the applications that are clinically critical already scored well on both criteria. So this is the right principle without being a disruptive change on a 20-row roster. On a 600-application portfolio it will matter more, which is the argument for making it now.

## 3. Bina's other three answers confirmed v1 as it stood

- **Risk in the end-user slot** - keep it. Risk stays as the fourth scoring dimension and end-user perception is still collected at zero weight.
- **Cost moves priority but never on its own makes something a retire** - accepted. From v2 this is enforced in the generator rather than just described: if an application fails only on cost and the engine ever returns retire, the script refuses to write the file.
- **The modelled peer cost band** - continue to use the model. It stands, and it is labelled as modelled rather than measured everywhere it appears.

## What did NOT change

Deliberately identical to v1: the same 20 applications, the same real product names with every other value invented, the same overlap clusters and survivors, the same trap cases, and the same cost and savings arithmetic. Savings still net off a replacement's run cost and any residual archival cost, and a cluster survivor still keeps its own gate result rather than being stamped consolidate.

The money is unchanged to the dollar: portfolio run-rate $22,057,000, net annual saving identified $5,818,716, of which $3,771,716 is available once the blocked and deferred rows are set aside.

## New in the reviewer's answer key

The 'Trap cases' sheet went from twelve rows to fourteen. One was rewritten and two are new, all because of the vocabulary change:

- **T12 rewritten.** In v1 this tested that two applications with different problems both came out invest and were told apart by priority. It now tests that they come out on *different words*: Epic Hyperspace retain, Luma Health invest.
- **T13 new.** A cluster survivor must come out retain - not consolidate (which would read as 'consolidate Power BI into Power BI') and not invest (nothing about the survivor is being funded).
- **T14 new.** An application that fails only on cost must come out invest. Not retire, which Bina's answer rules out, and not retain, which would ignore a failing gate.

Trap case T1 also gained a new wrong answer to catch: Sunquest CoPath Plus is old but adequate, and it must still come out **invest** rather than retire *or* retain. Its technical health genuinely fails, so there is genuinely something to fund.
