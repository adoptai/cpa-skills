---
name: budget-vs-actual
description: Compares actuals to budget by department and account, applies dollar and percentage thresholds, and refuses to call the report complete until every variance above its applicable threshold has an explanation whose components sum to the variance exactly. Use this whenever the user mentions budget vs actual, BvA report, variance analysis against budget, flux against budget, department spend vs budget, or "why did we go over budget." Trigger on "build the BvA," "run budget vs actual for this quarter," "which departments are over budget," "explain this variance," or when actuals, a budget export, and threshold config are supplied together. Reuses the enforced-explanation gate pattern from return-yoy-variance, applied to budget rather than prior-year comparisons — use return-yoy-variance for year-over-year variance, this one for actual-vs-budget variance.
---

# Budget vs Actual Report

A BvA report that lists variances without forcing an explanation for the ones that matter is a
list, not a report. This skill applies the same enforced-explanation discipline as
`return-yoy-variance`, scoped to actual-vs-budget instead of year-over-year.

## The gate

The script refuses to call the report complete when:

1. **A variance above its threshold has no explanation.** It lands on the Unexplained Variance
   tab, and the workbook summary states plainly the report is not complete while that tab is
   non-empty.
2. **An explanation's components don't sum to the variance.** If component rows are supplied, they
   must add up exactly to the variance they claim to explain — a partial explanation is an
   exception, not a pass.
3. **Aggregation drops a row.** Total actual and total budget on the report must equal the sums of
   the source files exactly.

## Inputs

1. **Actuals** — `department, account, actual_amount`.
2. **Budget export** — `department, account, budget_amount`.
3. **Threshold config** — `department, account, dollar_threshold, pct_threshold`. Either field may
   be blank to disable that test, and `department`/`account` may be blank to scope a threshold more
   broadly — blank-both is the global default. The most specific threshold that matches wins:
   account+department beats account-only beats department-only beats global. **No threshold is
   ever hardcoded** — a department/account with no matching threshold row at all is flagged for
   review rather than silently passed.
4. **Explanations** — `department, account, reason_category, narrative, component_description,
   component_amount`. One row per component; an account can have several rows explaining pieces
   of one variance (a travel overage split into conference airfare and hotel, say).

## Step 1 — Get a blank explanations worksheet for what currently exceeds threshold

```bash
python3 scripts/budget_vs_actual.py --actuals actuals.csv --budget budget.csv \
  --thresholds thresholds.csv --template --out explanations_template.csv
```

This evaluates thresholds first and emits one row per department/account pair found in the
actuals or budget files, ready for explanation narratives to be filled in.

## Step 2 — Build the report

```bash
python3 scripts/budget_vs_actual.py --actuals actuals.csv --budget budget.csv \
  --thresholds thresholds.csv --explanations explanations.csv \
  --period "FY2026 Q1" --client "Northwind Systems Inc" \
  --out "Northwind - Q1 BvA.xlsx"
```

## Deliver

**Workbook tabs:**

1. **BvA Report** — department, account, budget, actual, variance $, variance %, whether it
   exceeds its threshold, whether it's explained, and the explanation narrative.
2. **Unexplained Variance** — the gate. Empty means the report is complete; non-empty means it
   isn't, and the workbook summary says so.
3. **Threshold Config** — which threshold basis (account+department / account / department /
   global) actually applied to each line, so a reviewer can audit the thresholds themselves.

**Then, in chat:** total budget and total actual first, then the count exceeding threshold and the
count still unexplained, then the two or three largest unexplained variances by dollar amount.

## What to escalate

- **A department/account pair with no matching threshold row at all.** The script flags it rather
  than inventing a default — the threshold config is incomplete and needs a decision, not a guess.
- **A variance explained the same way every period** (e.g., "timing" every month for a year) — a
  pattern worth a conversation about whether the budget itself needs revisiting.
- **A component total that keeps falling short of the variance** across periods for the same
  account — suggests an unidentified driver, not a reporting error.
- **Large unexplained variances concentrated in one department.** Report the fact; don't
  characterize whether it reflects a budgeting miss or an operational one.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.
Department spend and budget data never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
