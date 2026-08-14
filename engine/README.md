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

### Path assumptions — read before running

The scripts were written and run in a working directory that no longer exists, and their
paths have deliberately **not** been repointed, so that the committed code is byte-for-byte
the code whose output was verified. To run one, you will need to supply the inputs it expects:

- `generate_dataset.py` needs nothing. It writes into the directory holding the script
  itself (`OUT_DIR`), so run it somewhere you are happy to have four files appear.
- `score_northstar*.py` each read one workbook at a hardcoded absolute `SOURCE_XLSX` near the
  top of the file — an ephemeral Slack upload path under `/mnt/user-data/uploads/slack/…`.
  Point that constant at the Northstar sample workbook (v1 reads the original, v2 and v3 read
  the revised file) or copy the workbook to that path. The workbook is opened read-only and is
  never written back to. Outputs land next to the script.
- `build_client_input.py` has a hardcoded `OUT` directory near the top and expects to find
  `generate_dataset.py`, `column-tiers.json` and the generator's own
  `App-Rationalization-Dummy-Dataset-v2.xlsx` / `applications-v2.csv` in it. Easiest path is
  to run `generate_dataset.py` first into that directory.

### What each produces

| Script | Outputs |
| --- | --- |
| `generate_dataset.py` | `App-Rationalization-Dummy-Dataset-v2.xlsx` (7 sheets), `applications-v2.csv`, plus a generated `README.md` and `CHANGELOG-v2.md` describing the run |
| `score_northstar.py` | `Northstar-Disposition-Analysis-v1.xlsx`, `northstar-dispositions.csv` |
| `score_northstar_v2.py` | `Northstar-Disposition-Analysis-v2.xlsx`, `northstar-dispositions-v2.csv` |
| `score_northstar_v3.py` | `Northstar-Disposition-Analysis-v3.xlsx`, `northstar-dispositions-v3.csv` |
| `build_client_input.py` | `Client-Input-Dataset-v1.xlsx`, `client-input.csv` |

## Where the datasets are

The datasets these scripts generate and consume are committed, in `data/` — including the two
Northstar sample workbooks the `score_northstar*.py` runs read, so point `SOURCE_XLSX` at
`data/northstar/healthcare_app_rationalization_sample_20.xlsx` (v1) or
`…-with-risk.xlsx` (v2 and v3). `data/README.md` explains every file. All of it is synthetic:
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
