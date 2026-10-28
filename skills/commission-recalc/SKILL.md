---
name: commission-recalc
description: Recomputes sales commissions from bookings against the comp plan's own tiered rate schedule — splitting a deal across tier boundaries where it straddles one — and ties the recomputed total to what payroll actually paid, by rep and period, with every discrepancy reported individually and never netted away. Use this whenever the user mentions recalculating commissions, auditing the commission spreadsheet, verifying sales comp payouts, checking commission plan tiers, or reconciling commissions to payroll. Also trigger on "recompute commissions from the comp plan," "check if commissions were calculated correctly," "verify the commission tiers," "does this commission payout match the plan," or when bookings, a tiered comp plan, and a payroll register are supplied together. Pairs with payroll-tax-reconciliation downstream. Every controller privately doubts the commission spreadsheet; a full-population recompute settles it. Runs fully local — no compensation data leaves the machine.
---

# Commission Recalculation

This recomputes commission from the comp plan's own tiers — nothing here is a flat or assumed
rate. Three conventions are stated here because they're load-bearing and otherwise invisible:

- `rate_pct` is a percentage (`5` means 5%), applied as `rate_pct / 100` times the dollar amount
  apportioned to that tier.
- A blank `tier_max_amount` means "and above" — unbounded, not a guessed ceiling.
- Each deal's commission is attributed to the payroll period matching the **calendar month**
  (`YYYY-MM`) of its `close_date`. The payroll register's own `period` column must use that same
  format to tie out — a register using a different period convention shows up as non-matching
  rather than being forced to align.

Commission tiers reset to zero at the start of each comp plan period (`plan_effective_date`) and
accumulate chronologically, by `close_date`, within that period.

## The gate

**Hard gate** — blocks the workbook unless `--force`:

1. **No usable comp plan tiers supplied.** Without a tiered schedule there is nothing to
   recompute a rate from, and the script never substitutes a flat or "standard" rate.

**Soft exceptions** — routed to the Exceptions tab, workbook still produced:

- A deal whose `close_date` matches zero or more than one comp plan period for that rep — never
  defaulted to the nearest plan.
- A plan period where **any** tier has a blank `rate_pct`. The entire plan period is treated as
  incomplete — every deal that falls in it becomes an exception, because once a rate is undefined
  mid-schedule, cumulative-based tier placement for later deals in that period can no longer be
  trusted.
- A deal whose amount span is not fully covered by contiguous tiers — a gap between tiers, or an
  amount beyond the rep's highest defined tier. The deal is **never partially rated**; it's
  excluded from the recomputed total entirely, with the uncovered dollar range named.

**Recomputed vs. paid** is reported per rep, per period, on its own row — every discrepancy
stands on its own dollar effect, never netted against another rep or period.

## Inputs

1. **Bookings or invoice export** — `rep, deal_id, amount, close_date`.
2. **Comp plan rules** — `rep, plan_effective_date, plan_expiry_date, tier_min_amount,
   tier_max_amount, rate_pct`. One row per tier; rows sharing the same rep and plan dates form
   one plan period.
3. **Payroll register** — `rep, period, commission_paid`.

## Step 1 — Emit a blank worksheet

```bash
python3 scripts/commission_recalc.py --template --out bookings_template.csv
```

## Step 2 — Recompute

```bash
python3 scripts/commission_recalc.py \
  --bookings bookings.csv --plan comp_plan.csv --payroll payroll.csv \
  --client "Northwind Systems Inc" \
  --out "Northwind - Commission Recalc.xlsx"
```

## Step 3 — How a straddling deal is split

A deal that crosses a tier boundary is apportioned across every tier it touches, at that tier's
own rate, based on the rep's cumulative booked amount in the plan period *before* this deal plus
the deal's own amount. The Recalculated Schedule tab shows the cumulative range before and after
each deal and the exact tier breakdown that produced its commission — so a disputed number can be
walked back to the specific tier math, not just a final figure.

## Deliver

**Workbook tabs:** Recalculated Schedule (every deal, cumulative range, tier breakdown,
commission, payroll period), Payroll Tie-Out (recomputed vs. paid by rep/period with a tie/break
flag), Discrepancies (every break, sorted by dollar effect, descending), Exceptions (unmatched or
overlapping plans, incomplete plan periods, uncovered tier gaps).

**Then, in chat:** the gate result, total deals recomputed and exceptions, then the discrepancy
list with dollar effects — that's the number that settles the argument about the spreadsheet.

## What to escalate

- A plan period with a blank-rate tier that's been in that state for more than one recalc cycle —
  likely a plan document that was never finished being entered, not a one-off typo.
- A rep with a cluster of "no plan found" exceptions around a plan renewal date — usually means
  the new plan period wasn't loaded in time, not that the rep has no plan.
- A discrepancy that's consistently in the rep's favor (paid > recomputed) across periods — worth
  a direct conversation with whoever enters the comp plan tiers.
- A deal whose amount alone exceeds every tier a rep has ever been assigned — may indicate a
  data-entry error on the deal amount rather than a plan gap.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.
Compensation, rep, and deal data never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
