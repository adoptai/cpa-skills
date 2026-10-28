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

---

# Appendix — bundled files

If you installed the `.skill` package these files are already in place and you can ignore this appendix. If you copied the skill as text, create the files below at the paths shown, alongside your `SKILL.md`. The skill will not run without them.

## `scripts/fx_revaluation.py`

```python
#!/usr/bin/env python3
"""
Revalue monetary accounts at period-end FX rates, compute realized and
unrealized gain/loss, propose the revaluation JE, and roll the CTA forward.

Gates:
  * every row must have an explicit monetary flag (y/n) - a blank is an
    exception, never assumed "no"
  * every row flagged monetary=y must have a period-end rate for its currency -
    never defaulted to 1.00 and never interpolated
  * proposed revaluation JEs must balance (debits = credits) before being
    included
  * nonmonetary accounts are logged to their own tab for transparency, never
    silently skipped without a record

Usage - emit a blank worksheet:
    python3 fx_revaluation.py --gl gl_detail.csv --template --out gl_worksheet.csv

Usage - run:
    python3 fx_revaluation.py --gl gl_worksheet.csv --rates rates.csv \
        --beginning-cta -42500.00 --settlements settlements.csv \
        --client "Northwind Systems Inc" --period-end 2025-12-31 \
        --out "Northwind - Q4 2025 FX Revaluation.xlsx"

--gl CSV (GL detail by currency):
    account, currency, fc_amount, functional_amount_recorded, monetary

  monetary: y | n - stated per row by the client, never inferred from the
  account name or account type.

--rates CSV (period-end spot rates):
    currency, period_end_rate

--settlements CSV (optional - for realized gain/loss):
    account, currency, fc_amount, settlement_date, settlement_rate
"""

from __future__ import annotations

import argparse
import csv
import sys
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is required.  pip install openpyxl")

ZERO = Decimal("0.00")
TOL = Decimal("0.01")
MONEY = '#,##0.00;[Red](#,##0.00)'
RATE = '0.000000'

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")

GL_COLS = ["account", "currency", "fc_amount", "functional_amount_recorded", "monetary"]
RATE_COLS = ["currency", "period_end_rate"]
SETTLE_COLS = ["account", "currency", "fc_amount", "settlement_date", "settlement_rate"]


def dec(raw):
    if raw is None:
        return None
    s = str(raw).strip()
    if s in ("", "-", "--", "n/a", "N/A", "None"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace(",", "").replace("$", "").strip()
    if s.endswith("-"):
        neg, s = True, s[:-1]
    try:
        v = Decimal(s)
    except InvalidOperation:
        raise ValueError(f"Unparseable amount: {raw!r}")
    return -v if neg else v


def clean(s) -> str:
    return " ".join(str(s or "").split())


def pdate(raw):
    s = clean(raw)
    if not s:
        return None
    for f in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d-%b-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s.split(" ")[0], f).date()
        except ValueError:
            continue
    raise ValueError(f"Unrecognized date {raw!r}. Prefer ISO YYYY-MM-DD.")


def load(path: Path, cols: list[str]) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No rows in {path}")
    out = []
    for i, r in enumerate(rows, start=1):
        x = {k: clean(r.get(k)) for k in cols}
        x["row"] = i
        out.append(x)
    return out


def write_template(rows: list[dict], out: Path) -> None:
    with out.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(GL_COLS)
        for r in rows:
            w.writerow([r.get(c, "") for c in GL_COLS])
    print(f"Wrote worksheet with {len(rows)} row(s) and required columns: {out}")


def evaluate(gl_rows, rate_rows, settle_rows, beginning_cta):
    errors, exceptions = [], []
    rates = {}
    for r in rate_rows:
        rate = dec(r["period_end_rate"])
        if rate is None:
            exceptions.append(f"row {r['row']}: rate for {r['currency']!r} is "
                               f"missing or unparseable")
            continue
        rates[r["currency"]] = rate

    settled_keys = set()
    settle_detail = []
    for s in settle_rows:
        fc = dec(s["fc_amount"])
        srate = dec(s["settlement_rate"])
        if fc is None or srate is None:
            exceptions.append(f"settlement row {s['row']}: unparseable amount or rate")
            continue
        settled_keys.add((s["account"], s["currency"]))
        settle_detail.append({**s, "fc_amount": fc, "settlement_rate": srate})

    revalued, nonmonetary = [], []
    je_debits = je_credits = ZERO

    for x in gl_rows:
        acct = x["account"]
        cur = x["currency"]
        monetary_raw = x["monetary"].strip().lower()
        fc = dec(x["fc_amount"])
        recorded = dec(x["functional_amount_recorded"])

        if monetary_raw not in ("y", "n"):
            errors.append(f"{acct} ({cur}): monetary flag is {x['monetary']!r}, "
                           f"not 'y' or 'n' - never assumed")
            continue
        if fc is None or recorded is None:
            errors.append(f"{acct} ({cur}): fc_amount or functional_amount_recorded "
                           f"is missing or unparseable")
            continue

        if monetary_raw == "n":
            nonmonetary.append({"account": acct, "currency": cur, "fc_amount": fc,
                                 "functional_amount_recorded": recorded})
            continue

        rate = rates.get(cur)
        if rate is None:
            errors.append(f"{acct} ({cur}): flagged monetary but no period-end "
                           f"rate supplied for {cur!r} - never defaulted to 1.00")
            continue

        new_equiv = (fc * rate).quantize(TOL)
        gain_loss = (new_equiv - recorded).quantize(TOL)
        is_realized = (acct, cur) in settled_keys
        revalued.append({
            "account": acct, "currency": cur, "fc_amount": fc,
            "old_functional": recorded, "rate": rate, "new_functional": new_equiv,
            "gain_loss": gain_loss,
            "status": "realized" if is_realized else "unrealized",
        })
        # each row's revaluation is a self-balancing 2-line entry: the account
        # itself moves to its new functional value, offset by FX gain/loss.
        if gain_loss > ZERO:
            je_debits += gain_loss      # Dr the account, up to its new value
            je_credits += gain_loss     # Cr FX Gain
        elif gain_loss < ZERO:
            je_debits += -gain_loss     # Dr FX Loss
            je_credits += -gain_loss    # Cr the account, down to its new value

    je_balanced = (je_debits - je_credits).copy_abs() <= TOL
    if not je_balanced:
        errors.append(f"Proposed revaluation JE does not balance: debits "
                       f"{je_debits:,.2f} vs credits {je_credits:,.2f}")

    unrealized_total = sum((r["gain_loss"] for r in revalued
                             if r["status"] == "unrealized"), ZERO)
    realized_total = sum((r["gain_loss"] for r in revalued
                           if r["status"] == "realized"), ZERO)

    ending_cta = None
    if beginning_cta is not None:
        ending_cta = (beginning_cta + unrealized_total).quantize(TOL)

    judgment = [r for r in revalued if abs(r["gain_loss"]) >
                abs(r["old_functional"]) * Decimal("0.20") and r["old_functional"] != ZERO]

    return (revalued, nonmonetary, errors, exceptions, je_debits, je_credits,
            je_balanced, unrealized_total, realized_total, ending_cta, judgment)


def style_header(ws, headers):
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"


def autosize(ws):
    for i, col in enumerate(ws.columns, start=1):
        width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
        ws.column_dimensions[get_column_letter(i)].width = min(width + 2, 46)


def sheet_summary(wb, meta, unrealized_total, realized_total, beginning_cta,
                   ending_cta, je_balanced, errors, valid):
    ws = wb.active
    ws.title = "Workpaper Summary"
    ws.append(["FX Revaluation"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])
    for k, v in meta.items():
        ws.append([k, v])
    ws.append([])
    for label, val in [
        ("Unrealized gain/(loss)", unrealized_total),
        ("Realized gain/(loss)", realized_total),
        ("Beginning CTA", beginning_cta),
        ("Ending CTA", ending_cta),
    ]:
        ws.append([label, val])
        cell = ws.cell(row=ws.max_row, column=2)
        if isinstance(val, Decimal):
            cell.number_format = MONEY
    ws.append(["Proposed JE balances", "yes" if je_balanced else "NO"])
    ws.append([])
    ws.append(["CONCLUSION", "CLEAN" if valid else "GATE NOT MET"])
    ws[f"B{ws.max_row}"].font = OK_FONT if valid else BAD_FONT
    if errors:
        ws.append([])
        ws.append(["Gate failures"])
        ws["A" + str(ws.max_row)].font = BAD_FONT
        for e in errors:
            ws.append([e])
    autosize(ws)


def sheet_revaluation(wb, revalued):
    ws = wb.create_sheet("Revaluation Workbook")
    headers = ["Account", "Currency", "FC Amount", "Old Functional",
               "Period-End Rate", "New Functional", "Gain/(Loss)", "Status"]
    style_header(ws, headers)
    for r in revalued:
        ws.append([r["account"], r["currency"], r["fc_amount"], r["old_functional"],
                   r["rate"], r["new_functional"], r["gain_loss"], r["status"]])
        row = ws.max_row
        for col in (3, 4, 6, 7):
            ws.cell(row=row, column=col).number_format = MONEY
        ws.cell(row=row, column=5).number_format = RATE
        if r["gain_loss"] < ZERO:
            ws.cell(row=row, column=7).font = BAD_FONT
        elif r["gain_loss"] > ZERO:
            ws.cell(row=row, column=7).font = OK_FONT
    autosize(ws)


def sheet_je(wb, revalued):
    ws = wb.create_sheet("Proposed JEs")
    style_header(ws, ["Account", "Debit", "Credit", "Memo"])
    for r in revalued:
        if r["gain_loss"] == ZERO:
            continue
        memo = (f"Revalue {r['currency']} balance to period-end rate "
                f"{r['rate']} ({r['status']})")
        if r["gain_loss"] > ZERO:
            ws.append([r["account"], r["gain_loss"], "", memo])
            ws.append(["FX Gain/(Loss)", "", r["gain_loss"], memo])
        else:
            loss = -r["gain_loss"]
            ws.append(["FX Gain/(Loss)", loss, "", memo])
            ws.append([r["account"], "", loss, memo])
        for row in (ws.max_row - 1, ws.max_row):
            for col in (2, 3):
                cell = ws.cell(row=row, column=col)
                if cell.value != "":
                    cell.number_format = MONEY
    autosize(ws)


def sheet_cta(wb, beginning_cta, unrealized_total, ending_cta):
    ws = wb.create_sheet("CTA Rollforward")
    style_header(ws, ["Line", "Amount"])
    for label, val in [("Beginning CTA", beginning_cta),
                        ("Current-period unrealized translation adjustment",
                         unrealized_total),
                        ("Ending CTA", ending_cta)]:
        ws.append([label, val])
        cell = ws.cell(row=ws.max_row, column=2)
        if isinstance(val, Decimal):
            cell.number_format = MONEY
    autosize(ws)


def sheet_nonmonetary(wb, nonmonetary):
    ws = wb.create_sheet("Nonmonetary Accounts Excluded")
    style_header(ws, ["Account", "Currency", "FC Amount", "Functional Amount Recorded"])
    for r in nonmonetary:
        ws.append([r["account"], r["currency"], r["fc_amount"],
                   r["functional_amount_recorded"]])
        row = ws.max_row
        ws.cell(row=row, column=3).number_format = MONEY
        ws.cell(row=row, column=4).number_format = MONEY
    autosize(ws)


def sheet_exceptions(wb, errors, exceptions):
    ws = wb.create_sheet("Exceptions")
    style_header(ws, ["Type", "Detail"])
    for e in errors:
        ws.append(["Gate failure", e])
        ws.cell(row=ws.max_row, column=1).font = BAD_FONT
    for e in exceptions:
        ws.append(["Data exception", e])
        ws.cell(row=ws.max_row, column=1).font = BAD_FONT
    autosize(ws)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--gl", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--rates")
    ap.add_argument("--settlements")
    ap.add_argument("--beginning-cta")
    ap.add_argument("--period-end")
    ap.add_argument("--client", default="")
    ap.add_argument("--template", action="store_true",
                     help="emit a blank GL worksheet instead of running")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    gl_rows = load(Path(args.gl), GL_COLS)

    if args.template:
        write_template(gl_rows, Path(args.out))
        return 0

    rate_rows = load(Path(args.rates), RATE_COLS) if args.rates else []
    settle_rows = load(Path(args.settlements), SETTLE_COLS) if args.settlements else []
    if not args.rates:
        sys.exit("--rates is required to run (not needed for --template)")
    beginning_cta = dec(args.beginning_cta)

    (revalued, nonmonetary, errors, exceptions, je_debits, je_credits,
     je_balanced, unrealized_total, realized_total, ending_cta,
     judgment) = evaluate(gl_rows, rate_rows, settle_rows, beginning_cta)

    valid = not errors

    meta = {
        "Client": args.client or "(not stated)",
        "Period end": args.period_end or "(not stated)",
        "Monetary rows revalued": len(revalued),
        "Nonmonetary rows excluded": len(nonmonetary),
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    print("=" * 76)
    print("FX REVALUATION")
    print("=" * 76)
    print(f"Client : {meta['Client']}    Period end: {meta['Period end']}")
    print(f"Revalued {len(revalued)} monetary row(s)   excluded "
          f"{len(nonmonetary)} nonmonetary row(s)")
    print(f"Unrealized {unrealized_total:,.2f}   Realized {realized_total:,.2f}")
    if ending_cta is not None:
        print(f"CTA: beginning {beginning_cta:,.2f} -> ending {ending_cta:,.2f}")
    print(f"Proposed JE: Dr {je_debits:,.2f} / Cr {je_credits:,.2f}  "
          f"({'balanced' if je_balanced else 'NOT BALANCED'})")
    if exceptions:
        print(f"\nDATA EXCEPTIONS ({len(exceptions)}):")
        for e in exceptions[:15]:
            print(f"  ! {e}")
    if errors:
        print(f"\nGATE FAILURES ({len(errors)}):")
        for e in errors[:15]:
            print(f"  ! {e}")
    if judgment:
        print(f"\nJUDGMENT ITEMS - large moves, worth a second look ({len(judgment)}):")
        for j in judgment[:10]:
            print(f"  ? {j['account']} ({j['currency']}): gain/loss {j['gain_loss']:,.2f}")

    if not valid and not args.force:
        print("\n" + "=" * 76)
        print("WORKBOOK NOT WRITTEN.")
        print("A rate is missing, a monetary flag is unstated, or the proposed JE does")
        print("not balance. Resolve the exceptions above, or re-run with --force to see")
        print("a workbook marked [GATE NOT MET].")
        print("=" * 76)
        return 1

    wb = Workbook()
    sheet_summary(wb, meta, unrealized_total, realized_total, beginning_cta,
                  ending_cta, je_balanced, errors, valid)
    sheet_revaluation(wb, revalued)
    sheet_je(wb, revalued)
    sheet_cta(wb, beginning_cta, unrealized_total, ending_cta)
    sheet_nonmonetary(wb, nonmonetary)
    sheet_exceptions(wb, errors, exceptions)
    wb.active = 0

    out = Path(args.out)
    if not valid:
        out = out.with_name(out.stem + " [GATE NOT MET]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\n{'CLEAN' if valid else 'NOT MET (forced)'}.  Workbook: {out}")
    return 0 if valid else 1


if __name__ == "__main__":
    sys.exit(main())
```

