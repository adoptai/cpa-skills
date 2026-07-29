---
name: audit-sampling
description: Select an audit sample from a population using monetary unit (PPS), stratified, or random attribute sampling — with the sample size derived from stated parameters, a recorded random seed so the selection can be re-performed exactly, and a prepared testing worksheet. Use this whenever the user mentions audit sampling, selecting a sample, sample size, MUS or PPS or monetary unit sampling, statistical sampling, haphazard selection, stratifying a population, testing a sample of transactions, projecting misstatement, or asks how many items to test. Also trigger on "pick a sample of invoices," "select items for testing," "how big should my sample be," "select 40 disbursements," "sample the AR balances," "extrapolate the error," or when a population export is provided for testing. Runs fully local — no client population data leaves the machine.
---

# Audit Sampling

Sampling is where audit files are most often indefensible, and almost never because the sample
was too small. It is because **the basis was never written down.** A reviewer three years later
cannot tell why 40 items were chosen, how they were chosen, whether the population was complete,
or what the result implies about the balance as a whole.

This skill produces a selection that can be **re-performed exactly** and a worksheet that states
its own basis. Everything else is secondary.

## The gate

No selection is produced unless:

1. **The population ties to a stated control total.** You supply the general ledger or trial
   balance figure the population should equal. If the population does not agree to it, sampling
   stops — a sample drawn from an incomplete population supports nothing, and the projection
   will understate misstatement by exactly the amount that was missing.
2. **A random seed is recorded.** Every selection is reproducible from the seed, the method, and
   the parameters. This is what makes the workpaper defensible.
3. **Sample size is derived, not chosen.** It follows from tolerable misstatement, expected
   misstatement, and risk of incorrect acceptance — all of which are recorded on the workpaper.
4. **Negative and zero balances are handled explicitly**, not silently dropped. See below.

## Standing rule on parameters

Sample sizes depend on firm methodology. This skill **computes** what can be computed and
**requires you to supply** what is a matter of methodology.

- The **reliability factor** for zero expected misstatements is computed exactly as
  `-ln(risk of incorrect acceptance)` — mathematics, not a remembered table.
- The **expansion factor** used when misstatement is expected varies by methodology. Supply your
  firm's factor with `--expansion-factor`. The script will not invent one, and says so on the
  workpaper if none was given.
- **Tolerable misstatement** and **risk of incorrect acceptance** come from your planning. The
  script records them; it does not suggest them.

If your firm has sampling tables, use them and enter the resulting size with
`--sample-size` — the script will still perform the selection, the reproducibility, and the
projection, and will document that the size came from firm tables.

## Choosing a method

| Method | Use when | Watch out for |
|---|---|---|
| **MUS / PPS** (`mus`) | Testing for overstatement of a balance made of many positive items — receivables, inventory, additions, disbursements. The default for substantive testing. | Cannot handle negative balances. Poor at detecting understatement. Zero balances have no chance of selection. |
| **Stratified** (`stratified`) | The population has a few large items and many small ones — most real populations. Reduces sample size for the same assurance. | Strata boundaries must be stated and justified. |
| **Random attribute** (`attribute`) | Testing whether a control operated, not a dollar amount — approvals, matching, authorization. | The conclusion is a deviation rate, not a dollar projection. Do not project dollars from an attribute sample. |

MUS examines every item at or above the sampling interval with certainty. That property is why
it is preferred for substantive testing: the largest items are not left to chance, and the
sample naturally concentrates where the money is.

## Step 1 — Prove the population

```bash
python3 scripts/select_sample.py --population pop.csv \
  --control-total 4128455.19 --validate-only
```

Reports population count, total, negative and zero items, the largest items, and whether the
total agrees to the control total. **Resolve any difference before selecting.**

Negative balances in a MUS population are a genuine professional problem, not a data issue.
Credit balances in receivables, credit memos in disbursements, and contra items cannot be
sampled proportionally to size. Handle them one of three ways, and record which:

- Test them **100% separately** as their own population — usually correct when few
- **Exclude and disclose**, testing them by another means
- Sample them separately with their own parameters

The script requires an explicit `--negatives` choice and refuses to guess.

## Step 2 — Select

```bash
python3 scripts/select_sample.py \
  --population pop.csv \
  --control-total 4128455.19 \
  --method mus \
  --tolerable-misstatement 125000 \
  --expected-misstatement 25000 \
  --expansion-factor 1.6 \
  --risk 0.05 \
  --negatives separate \
  --seed 20260729 \
  --client "Harbour Freight Systems Inc" \
  --assertion "Existence of trade receivables at 12/31/2025" \
  --out "Harbour Freight - AR Existence Sample.xlsx"
```

The seed is yours to choose and must be recorded before selection — not after. Choosing a seed
after seeing results is not random selection. Use the engagement date or any documented value.

## Step 3 — Test, then project

Fill in the worksheet: for each selected item record the evidence examined, the audited amount,
and the difference. Then:

```bash
python3 scripts/select_sample.py --project results.csv \
  --population pop.csv --control-total 4128455.19 \
  --method mus --tolerable-misstatement 125000 --risk 0.05 \
  --seed 20260729 \
  --out "Harbour Freight - AR Existence Conclusion.xlsx"
```

For MUS the projection uses **tainting**: an item examined in full (at or above the interval)
projects its actual difference; an item below the interval projects
`(difference ÷ book value) × sampling interval`. The script reports projected misstatement and
compares it to tolerable misstatement.

**The comparison is the conclusion, and it is not a formality.** Projected misstatement below
tolerable supports the balance. Projected misstatement approaching tolerable means the balance
may be materially misstated even though every individual difference looked small — that is
precisely the situation sampling exists to reveal, and the one most often waved away.

Basic precision and an allowance for sampling risk depend on methodology; supply your firm's
factors or compute the upper limit under your own approach. The script reports the components so
either can be applied, and states which were supplied.

## Step 4 — Deliver

**Workbook tabs:**

1. **Sampling Plan** — assertion tested, population and how it was obtained, control total and
   the tie-out, method, every parameter, computed sample size and the arithmetic behind it,
   seed, negative-balance treatment, and the sign-off block. This is the tab that has to answer
   a reviewer's questions without anyone present to explain.
2. **Selection** — every selected item with its row in the population, book value, selection
   basis (top stratum / interval hit / random), and blank columns for evidence examined, audited
   amount, difference, and disposition.
3. **Population Summary** — count, total, distribution by size band, coverage achieved in items
   and in dollars, and the top-stratum items examined in full.
4. **Not Selected** — the population that was not selected, so the untested remainder is
   visible and quantified rather than implicit.
5. **Projection** — once results are entered: differences found, tainting per item, projected
   misstatement, comparison to tolerable, and the conclusion.
6. **Reproducibility** — the exact command, seed, parameters, and a hash of the population file
   so a reviewer can confirm the same population produces the same sample.

## Things that quietly invalidate a sample

Worth checking before you rely on the result:

- **The population was filtered before you received it.** Ask what the export excluded. Date
  ranges, posted-only flags, and excluded accounts are the usual culprits.
- **The control total is the wrong total** — a subtotal, a net figure, or a balance after
  reclassification.
- **Items were added after selection.** The selection binds to the population as it stood;
  re-perform if it changed.
- **A selected item was swapped** because it was inconvenient to test. Replacing an item without
  documenting why destroys the statistical basis. If an item cannot be tested, that is a scope
  limitation, not a substitution.
- **The sample tests a different assertion than the one planned.** A sample selected
  proportional to recorded value tests overstatement of what is recorded — it says very little
  about completeness, because unrecorded items are not in the population at all.
- **Deviations treated as isolated.** A deviation is isolated only with evidence that it is;
  otherwise it projects.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry. The
population file is read-only and hashed for reproducibility, never transmitted. Output filenames
carry client name and assertion only.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
