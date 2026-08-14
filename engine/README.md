# Scoring engine and analysis scripts

The Python behind the application rationalization tool: the scoring engine, the synthetic
dataset generator, and the runs that score the Northstar Global Health sample through it.

## What is what

| Script | What it does |
| --- | --- |
| `generate_dataset.py` | **The scoring engine.** Generates the team's synthetic 20-application dataset and computes every scored dimension, gate, disposition, priority and saving. Everything else here reuses this model. |
| `score_northstar.py` | Run 1 — scores the separately-built 20-application Northstar sample through the engine. |
| `score_northstar_v2.py` | Run 2 — same derivations as run 1, with the three risk inputs rebuilt from the revised `Healthcare Guardrails` evidence and risk promoted to a first-class gated dimension. |
| `score_northstar_v3.py` | **Run 3 — current.** Applies the answers to the ten open items left by v2 (Sev-1/Sev-2 double-count netted out, lifecycle guard armed off `Current Release / Version`, and the rest). Use this one. |
| `score_northstar_600.py` | **A dataset run, not a model version.** Scores the 600-application Northstar portfolio on the v3 model, which it imports from `score_northstar_v3.py` rather than restating. Adds one layer v3 has none of: a documented, audited value-vocabulary normalisation, because the 600-row workbook matches v3's schema header for header but populates its 580 new applications under a different convention. |
| `build_client_input.py` | Derives the client-supplied-columns-only view of the dataset by projecting the built rows — no value is regenerated or perturbed. |

v1 and v2 of the scoring runs are kept deliberately: the progression from v1 to v3 is part of
how the model was arrived at, and each file documents in its own header exactly what changed
and why. For any current number, read v3.

## Running them

Python 3 (developed on 3.11). One third-party dependency across all five scripts:

```sh
pip install openpyxl        # tested against openpyxl 3.1.5
```

Everything else they import is standard library: `csv`, `datetime`, `os`, `statistics`,
`sys`, `warnings`, `json`, `collections`, `inspect`.

```sh
python3 engine/generate_dataset.py
python3 engine/score_northstar_v3.py
python3 engine/build_client_input.py
```

### Paths — read before running

Every path is now resolved relative to the script's own location, so all five scripts run from a
fresh clone with no editing. The scripts previously pointed at a working directory that no longer
exists; those constants have been repointed at the committed copies of the same inputs.

- `generate_dataset.py` needs nothing, but writes into the directory holding the script itself
  (`OUT_DIR`) — four files, including a generated `README.md` that **would overwrite this file**.
  Run it in a scratch copy of the directory, then move the two dataset files into
  `data/synthetic-portfolio/`. See the rerun warning in `data/README.md` first: the committed
  workbook is ahead of the generator, and a plain rerun drops a sheet and three data-dictionary
  columns.
- `score_northstar*.py` read their sample workbook from `data/northstar/` (v1 reads
  `healthcare_app_rationalization_sample_20.xlsx`, v2 and v3 read `…-with-risk.xlsx`) and write
  their workbook and CSV back into the same directory, replacing the committed outputs. The samples
  are opened read-only and are never written back to. Run them in order — v2 diffs against v1's CSV
  and v3 against v2's.
- `build_client_input.py` imports the engine from beside itself, reads the generated v2 dataset from
  `data/synthetic-portfolio/`, and reads `column-tiers.json` and writes both its outputs in
  `data/client-intake/`. Run `generate_dataset.py` first if the v2 dataset has changed.

### What each produces

| Script | Outputs |
| --- | --- |
| `generate_dataset.py` | `App-Rationalization-Dummy-Dataset-v2.xlsx` (7 sheets), `applications-v2.csv`, plus a generated `README.md` and `CHANGELOG-v2.md` describing the run |
| `score_northstar.py` | `Northstar-Disposition-Analysis-v1.xlsx`, `northstar-dispositions.csv` |
| `score_northstar_v2.py` | `Northstar-Disposition-Analysis-v2.xlsx`, `northstar-dispositions-v2.csv` |
| `score_northstar_v3.py` | `Northstar-Disposition-Analysis-v3.xlsx`, `northstar-dispositions-v3.csv` |
| `score_northstar_600.py` | `Northstar-Disposition-Analysis-600.xlsx`, `northstar-dispositions-600.csv`, `northstar-600-summary.md`, plus the same rows in the tool's own column vocabulary as `Northstar-600-tool-vocabulary.xlsx` and `northstar-600-tool-vocabulary.csv` |
| `build_client_input.py` | `Client-Input-Dataset-v1.xlsx`, `client-input.csv` |

## Where the datasets are

The datasets these scripts generate and consume are committed, in `data/` — including the two
Northstar sample workbooks the `score_northstar*.py` runs read, which is why `SOURCE_XLSX` now
resolves to `data/northstar/healthcare_app_rationalization_sample_20.xlsx` (v1) or
`…-with-risk.xlsx` (v2 and v3) without anyone editing it. `data/README.md` explains every file. All of it is synthetic:
invented figures for fictional organisations, with only the product and vendor names real.

This repository is public, and client-supplied material and licensed third-party vendor
templates must never land in it; `.gitignore` blocks the known ones by name.

## The scoring model in brief

Four dimensions:

- **business value** (V)
- **technical health** (T)
- **cost efficiency** (C)
- **risk** (R)

Eighteen inputs feed them, each scored **1 to 5 in half steps** (nine possible values), with
**5 always the favourable end** — including for the cost and risk inputs, where 5 means cheap
and controlled respectively. Each dimension is a weighted arithmetic mean of its inputs.

Every dimension is then **gated at 3.0**: at or above 3.0 passes, below fails. The four
pass/fail results in V, T, C, R order form a four-character key (`PPPP` … `FFFF`), and that
key is looked up in a 16-row table which yields one of five dispositions —

`invest` · `retain` · `consolidate` · `replace` · `retire`

— together with a priority on a five-step ladder from Very Low to Very High. The table is the
mapping decision, expressed as configuration rather than as branching code, so it can be
argued about and changed in one place. `generate_dataset.py` holds it (`DISPOSITION_TABLE`)
along with the per-input weights (`CRITERIA`) and the threshold (`PASS_THRESHOLD`); each
scoring run restates the same constants so it can execute standalone, and self-checks that
they still match the engine.

The mechanics are adapted from a licensed third-party portfolio-rationalization tool, cited in
the script headers. The template itself is not in this repository and must not be added.
