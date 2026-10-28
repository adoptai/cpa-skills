---
name: thirteen-week-cash-forecast
description: Builds a 13-week direct cash forecast from AR aging, AP aging, the payroll calendar, and recurring items, with a weekly actual-vs-forecast reforecast mode. Every input row lands in a specific week or an explicit Beyond Horizon bucket, and beginning balance plus net flows must tie exactly to the week-13 ending balance before the workbook is produced. Use this whenever the user mentions a 13-week cash forecast, direct cash forecast, weekly cash forecast, cash runway, cash flow projection, liquidity forecast, or reforecasting cash against actuals. Trigger on "build a 13-week forecast," "how much cash will we have in 8 weeks," "forecast our cash position," "reforecast this week's cash," "what's our cash runway," or when AR aging, AP aging, a payroll calendar, and recurring items are supplied together. The canonical CFO forecasting artifact, recurring weekly, not a one-shot. Pairs with `cash-flow-tieout`, which reconciles a historical statement rather than projecting forward.
---

# 13-Week Cash Forecast

A direct cash forecast is only as credible as its placement logic. Every dollar in AR, AP,
payroll, and recurring items has to land in a specific week — or in an explicit catch-all for
anything beyond the 13-week window — and the script proves that placement rather than asserting
it.

## The gate

The script refuses to produce a clean workbook when:

1. **A row cannot be placed.** Every AR, AP, payroll, and recurring-generated occurrence must land
   in one of the 13 weeks or the Beyond Horizon bucket. A receivable whose aging bucket has no
   matching row in the collection-timing assumptions is an exception, not a guess.
2. **The arithmetic identity breaks.** Beginning balance plus the sum of all 13 weeks' net flows
   must equal the week-13 ending balance exactly. If it doesn't, the week-by-week rollforward has
   a bug and nobody should be looking at the output yet.
3. **A row is missing a required field.** Amount, date, or direction fields that can't be parsed
   are reported by row number, never defaulted.
4. **An actual is compared that was never supplied.** In reforecast mode, only weeks with a
   genuine actual in `--actuals` are compared — a missing week is left out of the Reforecast
   Variance tab, never shown as a zero variance.

## Inputs

1. **AR aging** — `customer, amount, due_date`.
2. **AP aging** — `vendor, amount, due_date`.
3. **Payroll calendar** — `pay_date, amount`.
4. **Recurring items** — `description, amount, frequency, next_date, direction`. `frequency` is
   `weekly`, `biweekly`, or `monthly`; `direction` is `inflow` or `outflow`.
5. **Collection-timing assumptions** — `aging_bucket, week_offset, pct_collected_in_week`. This is
   the client's own judgment about when each aging bucket actually converts to cash, not something
   the script invents. Every bucket that appears in the AR file needs a row here.
6. **`--beginning-balance`** and **`--forecast-start`** — required.
7. **`--buckets`** (optional, default `30,60,90`) — the aging bucket boundaries in days past due,
   same convention as `ar-aging-tie-out`. Change it if the client ages receivables differently.
8. **`--actuals`** (optional) — `week_ending, actual_net_flow`, for the weekly reforecast cadence
   once a week has closed.

## Step 1 — Get blank templates for anything you don't have yet

```bash
python3 scripts/cash_forecast_13wk.py --template ar --out ar_template.csv
python3 scripts/cash_forecast_13wk.py --template assumptions --out assumptions_template.csv
```

`--template` accepts `ar`, `ap`, `payroll`, `recurring`, `assumptions`, or `actuals`.

## Step 2 — Build the forecast

```bash
python3 scripts/cash_forecast_13wk.py \
  --ar ar_aging.csv --ap ap_aging.csv --payroll payroll_calendar.csv \
  --recurring recurring_items.csv --assumptions collection_timing.csv \
  --beginning-balance 850000.00 --forecast-start 2026-01-05 \
  --client "Northwind Systems Inc" \
  --out "Northwind - 13-Week Cash Forecast.xlsx"
```

AR is spread across weeks per its aging bucket's assumption rows (a bucket can split across
multiple weeks — the percentages don't have to sum to 100, and whatever's left over at week 13 is
swept into Beyond Horizon with the shortfall shown, never silently absorbed). AP, payroll, and
recurring items land in the week of their due date, pay date, or projected occurrence.

## Step 3 — The weekly reforecast

Once a week closes, run again with `--actuals` to see where the forecast missed:

```bash
python3 scripts/cash_forecast_13wk.py \
  --ar ar_aging.csv --ap ap_aging.csv --payroll payroll_calendar.csv \
  --recurring recurring_items.csv --assumptions collection_timing.csv \
  --actuals week1_actuals.csv \
  --beginning-balance 850000.00 --forecast-start 2026-01-05 \
  --client "Northwind Systems Inc" \
  --out "Northwind - 13-Week Reforecast W1.xlsx"
```

This is a recurring-use skill by design — re-run it every week against the same forecast horizon,
and the Reforecast Variance tab shows how assumption drift is accumulating, which is the signal to
revisit the collection-timing table itself rather than patch the forecast by hand.

## Deliver

**Workbook tabs:**

1. **13-Week Forecast** — weekly beginning balance, AR collections, AP disbursements, payroll,
   net recurring, and ending balance, plus the gate status.
2. **Assumptions** — every collection-timing row actually used, so a reviewer can see the basis
   for the AR placement without opening the script.
3. **Beyond Horizon** — AR, AP, payroll, and recurring amounts that fall outside the 13-week
   window, by category.
4. **Source Detail** — every input row, its amount and date, and exactly which week (or Beyond
   Horizon) it was placed in — the row-count reconciliation that proves nothing was dropped.
5. **Reforecast Variance** (only when `--actuals` supplied) — forecast vs. actual net flow per
   week with an actual on record.

**Then, in chat:** the beginning and week-13 ending balance first, then any gate failures, then
the Beyond Horizon totals so the reader knows what the 13-week window doesn't cover.

## What to escalate

- **A collection-timing assumption that doesn't match recent actuals** once the reforecast mode
  has a few weeks of history — that's a conversation about the assumption, not a script fix.
- **A week where AP or payroll alone exceeds the beginning balance** — a liquidity event, not a
  forecasting detail.
- **Recurring items whose `next_date` is stale** (far in the past) — likely means the item already
  occurred and the source file wasn't updated.
- **Large, unscheduled AR remainders landing in Beyond Horizon** — usually means the
  collection-timing table doesn't cover that bucket's full percentage and needs a review.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry. AR,
AP, payroll, and vendor data never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
