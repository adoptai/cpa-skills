---
name: board-package-builder
description: Assembles the monthly board reporting package — P&L, balance sheet, KPI page, flux commentary, and a cash-flow movement rollforward — directly from the trial balance, with every figure traced to its source TB account and an unbalanced TB or an unclassifiable account blocking the draft outright. Use this whenever the user mentions building a board package, board deck numbers, monthly reporting package, assembling the board book, or pulling together P&L/balance sheet/KPI/flux commentary for a board or investor update. Also trigger on "build the board package," "assemble this month's board deck financials," "pull the board numbers together," "draft the board package from the TB," or when a trial balance and KPI definitions are supplied together. Pairs with flux-narrative (supplies the commentary this skill passes through) and consolidation-rollup (consolidated TB for multi-entity closes) upstream, audit-workpaper downstream. Runs fully local — no client data leaves the machine.
---

# Board Package Builder

This is a mechanical compilation, not a presentation-judgment tool. It sums the trial balance
into the P&L and balance sheet exactly as recorded — debit minus credit per account, with no
revenue/expense or asset/liability sign flip and no equity rollforward applied. That
presentation judgment belongs to the preparer; review sign conventions before issuing the
package. What this skill guarantees instead is that every number in the package can be traced
back to a specific TB account, and that nothing is estimated, netted, or silently invented.

## The gate

**Hard gate** — blocks the workbook unless `--force`:

1. **The trial balance does not balance.** Total debits must equal total credits exactly before
   anything downstream is built.
2. **A TB row has an unclassifiable `account_type`.** Only `P&L` and `BS` are recognized. An
   account that can't be placed on either statement is not silently dropped or guessed onto one
   — it blocks the build.

**Soft exceptions** — reported on the Exceptions tab, workbook still produced:

- A KPI whose account mapping references an account not present in the TB, or an unrecognized
  `operator`. That KPI is marked **not computed** — never partially computed from the accounts
  that *did* resolve.
- A flux line item with no commentary supplied — marked `(commentary not supplied)`, never
  filled in with invented narrative.
- A cash-flow account with no prior-period balance — its full current balance is shown as the
  movement, flagged, rather than guessing a prior figure.
- A duplicate account name in the TB — balances are summed and the duplication is flagged, not
  silently merged without a trace.

## Inputs

1. **Trial balance** — `account, account_type (P&L|BS), debit, credit`.
2. **KPI definitions** — `kpi_name, account, operator (add|subtract)`, one row per account that
   rolls into a KPI. A KPI is computed **only** from the accounts named here — never from a
   formula the script invents.
3. **Flux commentary** (optional) — `line_item, commentary`, typically the output of
   `flux-narrative`.
4. **Prior-period figures** (optional) — `line_item, prior_amount`, one row per account, for the
   cash-flow rollforward.

## Step 1 — Emit a blank worksheet

```bash
python3 scripts/board_package.py --template --out trial_balance_template.csv
```

## Step 2 — Build the package

```bash
python3 scripts/board_package.py \
  --trial-balance tb.csv --kpi-defs kpis.csv \
  --flux flux_commentary.csv --prior prior_period.csv \
  --client "Northwind Systems Inc" --period-end 2025-12-31 \
  --out "Northwind - December 2025 Board Package.xlsx"
```

## Step 3 — What the cash flow tab actually is

If `--prior` is supplied, the Cash Flow tab is an **account-level movement rollforward**:
current balance minus prior balance, per balance-sheet account, plus net income as recorded on
the P&L. It does **not** classify movements into operating, investing, or financing activities
— that classification is a professional judgment call the script does not make, and the tab
says so. If `--prior` is not supplied, the tab states plainly that cash flow could not be built,
rather than fabricating one from a single period.

## Deliver

**Workbook tabs:** Summary (gate status, TB totals, P&L/BS net balances), P&L, Balance Sheet,
KPI (value, components, status — computed or not computed), Cash Flow (rollforward or
not-built placeholder), Flux Commentary (verbatim or "not supplied"), Trace - Tie-Out (every
account mapped to the statement and KPI(s) it feeds), Exceptions (every soft exception above).

**Then, in chat:** the gate result first, then the KPI/flux/cash-flow exception counts, then
flag any account whose `account_type` looked ambiguous so the preparer can confirm the
classification before this goes anywhere near the board.

## What to escalate

- Any account where sign presentation (debit/credit as positive) looks backwards for its
  account type — this script does not correct it, and an inverted sign in a board deck is the
  kind of error that gets noticed in the room.
- A KPI repeatedly landing on "not computed" month over month — usually means the KPI mapping
  was never updated for a chart-of-accounts change.
- Cash-flow accounts with no prior balance in a period where the entity is not new — likely a
  broken link to the prior TB, not a real new account.
- Flux commentary coverage that's thin in a period with a large P&L swing.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.
Trial balance and KPI data never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
