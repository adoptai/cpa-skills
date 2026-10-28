---
name: debt-covenant-check
description: Computes covenant ratios from the trial balance using the credit agreement's own component mappings and formulas, tracks headroom over time against prior compliance certificates, and flags approaching breaches — but only after confirming the trial balance itself balances. Use this whenever the user mentions debt covenants, covenant compliance, covenant headroom, leverage ratio, fixed charge coverage ratio, compliance certificate, loan covenant testing, or "are we in compliance with our credit agreement." Trigger on "check our covenants," "compute the leverage ratio," "are we close to breaching a covenant," "build the compliance certificate," or when a trial balance and covenant definitions are supplied together. Pairs with trial-balance-integrity, whose balance check this skill reuses before computing anything — if the TB doesn't balance yet, run that skill first.
---

# Debt Covenant Check

This is a quarterly ritual with real career risk attached if it's wrong, so it starts from the
same discipline as `trial-balance-integrity`: nothing is computed on a trial balance that doesn't
balance. Every covenant's formula, threshold, and account mapping comes from the client's actual
credit agreement — none of it is assumed or defaulted by the script.

## The gate

The script refuses to compute anything, or refuses a specific covenant, when:

1. **The trial balance doesn't balance.** Debits must equal credits before any covenant is
   touched. If it doesn't, the whole run is blocked — run `trial-balance-integrity` first.
2. **A component's mapped account doesn't exist in the TB.** That covenant becomes an exception,
   never computed with the missing account treated as zero.
3. **A covenant has no threshold or comparison operator supplied.** The script refuses that
   covenant by name rather than assuming a "typical" leverage or coverage threshold.
4. **A formula can't be evaluated** from the components actually supplied — divide-by-zero, an
   unrecognized component label, or unsupported syntax are all reported, never guessed around.

Early-warning headroom flags are opt-in: supply `--warning-band-pct` or the Approaching Breach tab
stays intentionally empty rather than silently applying an assumed band.

## Inputs

1. **Trial balance** — `account, debit, credit`.
2. **Covenant components** — `covenant_name, component_label, account, operator` (`add` or
   `subtract`). This is how the client's own credit agreement defines what rolls into, say,
   "EBITDA" or "Fixed Charges" for that specific facility — addbacks and exclusions vary by
   agreement and are never assumed.
3. **Covenant definitions** — `covenant_name, formula_description, formula, threshold,
   comparison`. `formula` is written using the component labels exactly as they appear in the
   components file (e.g. `Total Debt / EBITDA`). `comparison` is `>=`, `<=`, `>`, or `<`.
4. **Prior compliance certificates** (optional) — `covenant_name, period, actual_value`, for the
   headroom trend.
5. **`--warning-band-pct`** (optional) — the percentage-of-threshold band that counts as
   "approaching" a breach. No default; omitting it means no early-warning test runs.

## Step 1 — Get blank templates

```bash
python3 scripts/debt_covenant_check.py --template components --out components_template.csv
python3 scripts/debt_covenant_check.py --template definitions --out definitions_template.csv
```

## Step 2 — Run the check

```bash
python3 scripts/debt_covenant_check.py --tb trial_balance.csv \
  --components covenant_components.csv --definitions covenant_definitions.csv \
  --certificates prior_certificates.csv --warning-band-pct 15 \
  --client "Northwind Systems Inc" --period "FY2026 Q1" \
  --out "Northwind - Covenant Compliance.xlsx"
```

A covenant failing its test (e.g., leverage above the cap) is a real finding, not a gate failure —
the workbook still writes and reports the failure clearly. The gate is about whether the
computation itself is trustworthy, not whether the client happens to be in compliance.

## Deliver

**Workbook tabs:**

1. **Covenant Compliance** — each covenant's computed value, threshold, comparison, pass/fail,
   headroom, whether it's in the approaching-breach band, and the prior-period trend beneath it.
2. **Component Trace** — every account feeding every covenant, with its TB value and how it rolled
   into the component (add/subtract), so the computation can be audited line by line.
3. **Exceptions** — covenants that couldn't be computed and why (missing account, missing
   threshold, unevaluable formula).
4. **Approaching Breach** — only populated when `--warning-band-pct` is supplied; otherwise it
   states plainly that no early-warning test was run, rather than appearing as a clean empty tab.

**Then, in chat:** whether the TB balanced, the pass/fail count, and any covenant currently failing
or approaching its band — those are the only numbers that matter to the reader in the first
sentence.

## What to escalate

- **Any covenant that fails.** Report the computed value, the threshold, and the shortfall — do
  not characterize materiality or likely lender response.
- **A covenant approaching its warning band** for the first time, or trending toward it across the
  certificate history.
- **A component mapping that looks inconsistent with the credit agreement's actual definitions**
  discovered while building the components file — that's a conversation with whoever negotiated
  the facility, not a script decision.
- **A trial balance that doesn't balance.** This blocks the entire check and needs to be resolved
  before covenant compliance can be assessed at all.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry. Trial
balance and credit agreement terms never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
