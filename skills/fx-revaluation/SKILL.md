---
name: fx-revaluation
description: Revalues monetary GL accounts at period-end FX rates, computes realized and unrealized gain/loss, proposes a balanced revaluation JE, and rolls the cumulative translation adjustment (CTA) forward — refusing to deliver when a monetary account has no period-end rate, the monetary/nonmonetary flag is unstated, or the proposed JE doesn't balance. Use this whenever the user mentions FX revaluation, period-end currency revaluation, unrealized FX gain or loss, remeasurement of foreign-currency balances, CTA rollforward, or multi-currency close. Also trigger on "revalue our foreign currency accounts," "what's our unrealized FX gain/loss," "roll the CTA forward," "propose the FX revaluation JE," or when GL detail by currency and a period-end rate table are supplied together. Ships alongside `consolidation-rollup`, and shares the recompute-and-tie shape of `depreciation-tie-out` applied to monetary balances. Runs fully local — no GL or currency data leaves the machine.
---

# FX Revaluation

## The gate

The script refuses to produce a clean revaluation workbook when:

1. **A row has no explicit monetary flag.** `monetary` must be `y` or `n` per row, stated by the
   client. A blank is an exception, never assumed to be "no" — guessing wrong here means a real
   monetary balance never gets revalued.
2. **A row flagged `monetary=y` has no period-end rate for its currency.** The rate table is
   client-supplied; a missing currency is a visible exception, never defaulted to 1.00 and never
   interpolated from a nearby currency or period.
3. **The proposed revaluation JE doesn't balance.** Each row's entry — the account moving to its
   new functional-currency value, offset by FX Gain/(Loss) — is self-balancing by construction; a
   failure here means something upstream is broken and the workbook should not go out.

Nonmonetary accounts (`monetary=n`) are never silently dropped. They're logged to their own tab so
a reviewer can see exactly what was excluded and why, rather than wondering whether an account was
missed.

## Inputs

1. **GL detail by currency** — account, currency, the foreign-currency amount, the functional-
   currency amount currently recorded, and an explicit monetary/nonmonetary flag per row.
2. **Period-end spot rates** by currency.
3. **Settlement detail** (optional) — foreign-currency items that settled during the period, so
   their gain/loss is classified realized rather than unrealized.
4. **The beginning CTA balance**, to roll forward.

## Step 1 — Build the worksheet

```bash
python3 scripts/fx_revaluation.py --gl gl_detail.csv --template --out gl_worksheet.csv
```

## Step 2 — Run the revaluation

```bash
python3 scripts/fx_revaluation.py --gl gl_worksheet.csv --rates rates.csv \
  --beginning-cta -42500.00 --settlements settlements.csv \
  --client "Northwind Systems Inc" --period-end 2025-12-31 \
  --out "Northwind - Q4 2025 FX Revaluation.xlsx"
```

## Deliver

**Workbook tabs:**

1. **Workpaper Summary** — unrealized and realized gain/loss totals, the CTA rollforward, whether
   the proposed JE balances, and the conclusion.
2. **Revaluation Workbook** — every monetary row with its FC amount, old and new functional
   equivalents, the period-end rate applied, the gain/loss, and realized/unrealized status.
3. **Proposed JEs** — the balanced two-line entry per account with a nonzero revaluation effect:
   the account itself moving to its new value, offset by FX Gain/(Loss).
4. **CTA Rollforward** — beginning CTA, the period's unrealized translation adjustment, ending
   CTA.
5. **Nonmonetary Accounts Excluded** — the transparency log of every account flagged `monetary=n`.
6. **Exceptions** — every missing rate, unstated flag, or unparseable amount.

**Then, in chat:** total unrealized and realized gain/loss first, then the CTA rollforward, then
any account whose gain/loss is unusually large relative to its balance.

## What to escalate

- **A large gain/loss relative to the account's recorded balance** — the script flags these as
  judgment items; confirm the rate, the FC amount, and whether the account was already partially
  revalued elsewhere before accepting the figure.
- **An account the client flagged `monetary=n` that functions like a monetary balance** (an
  intercompany loan booked to a balance-sheet-only account, for instance) — report what the flag
  says and what the account appears to be, and let the engagement decide the classification.
- **A settlement whose settlement_rate differs sharply from the nearby spot rate**, which can
  indicate a forward contract, a hedge, or a data entry error — not something this skill resolves.
- **Any currency with a rate the engagement hasn't independently verified** for the period — this
  skill applies whatever rate you give it and recomputes faithfully; it does not validate the rate
  itself.

Report each as a fact with the account and currency. Do not characterize intent.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry. GL and
currency data never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
