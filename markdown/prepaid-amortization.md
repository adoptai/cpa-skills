---
name: prepaid-amortization
description: Build and verify a prepaid expense amortization schedule, recompute straight-line amortization from each item's own start date, end date, cost, and convention, tie the ending balance and current-period amortization to the GL, and propose catch-up or reversing journal entries for any gap - never posting, only proposing. Use whenever the user mentions prepaid expenses, the prepaid schedule, prepaid amortization, prepaid insurance, prepaid rent, a 12-month rule item, unamortized prepaid balances, or checking whether the prepaid schedule ties to the GL. Also trigger on "does the prepaid schedule tie," "roll forward prepaids," "build the prepaid schedule," "what's the catch-up on this prepaid," "amortize this invoice," or when a prepaid schedule or invoice list is supplied with GL detail for the prepaid accounts. Same tie-out pattern as depreciation-tie-out, applied to the other recurring schedule every close carries. Runs fully local - no client data leaves the machine.
---

# Prepaid Amortization Schedule

A prepaid schedule rots the same way a fixed asset register does: it is rolled forward every
close, rarely rebuilt, and whatever was wrong last month is still wrong this month plus
whatever broke since. Three failures account for most of what this tie-out finds:

1. **An item whose service period already ended, still carrying a balance.** The vendor was
   paid, the coverage or service expired, and nobody ran the last entry. The balance sheet
   still shows an asset that no longer represents future economic benefit.
2. **A beginning balance that doesn't agree to what was actually closed last period.** Usually
   a spreadsheet overwritten instead of rolled, or a formula that didn't drag down for a new
   row. It misstates every period after it until someone notices.
3. **An item amortizing on the wrong cadence** — too fast, too slow, or not started at all —
   so this period's entry doesn't match what straight-line amortization of the item's own
   cost, start date, and end date actually requires.

None of these show up by reading the amortization expense line. They show up in a rollforward
and a recomputation, which is why this skill runs both before proposing anything.

## The gate

No clean workpaper unless all of these hold:

1. **The schedule rollforward foots.** Beginning balance + additions − current-period
   amortization = ending balance, in total and by GL account.
2. **The beginning balance agrees to the prior period as closed** — supplied from the prior
   period's schedule or GL, not from whatever the current file happens to show, which is what
   introduced the error in the first place.
3. **The schedule agrees to the GL** for the ending balance and the current-period
   amortization expense.
4. **No item has cumulative amortization exceeding its total cost.**
5. **Recomputed straight-line amortization agrees to what the schedule states** for every item
   with enough data to recompute.
6. **Every item is free of integrity findings** — missing dates or cost, a period that ended
   with a balance still outstanding, amortization running before the service period begins, or
   a duplicate.

## Standing rule on capitalization

**This skill does not decide whether a payment should be capitalized as a prepaid asset at
all.** That question turns on the 12-month rule and the entity's own capitalization policy
threshold, both of which change and neither of which belongs in a schedule-arithmetic tool from
memory. Where it matters, the workpaper says *"verify whether this item qualifies for current
deduction against current instructions for Reg. 1.263(a)-4(f) and the entity's capitalization
policy."* It assumes every item already on the schedule was properly capitalized, and checks
only whether the arithmetic of amortizing it is right.

What the script does instead:

- **Recomputes straight-line amortization exactly**, from the item's own cost, start date, end
  date, and convention — whole-month (even by calendar month, the common prepaid convention)
  or daily (prorated by exact days). This is arithmetic, not judgment, and it is checked
  precisely.
- **Computes the expected cumulative amortization through period end** and nets out what the
  schedule already shows as amortized through the prior period. The difference is this period's
  proposed entry — the ordinary monthly amount for a clean item, or the full catch-up for an
  item that fell behind (or a reversal for one that ran ahead) in a period before this one.
- **Never guesses a start date, end date, or cost.** An item missing any of these becomes a
  visible exception, excluded from the recomputation, not silently zeroed or carried forward.

## Inputs

1. **Prepaid schedule or invoice list**, per item: ID, vendor, description, the prepaid (GL)
   account, the expense account it amortizes into, payment date, service start and end dates,
   total cost, convention, cumulative amortization through the prior period, and — only if the
   item is already on an existing schedule — this period's amortization and ending balance as
   currently shown. Leave the last two blank for a new item and the script builds them.
2. **Prior-period ending balances**, by GL account, **as actually closed** — from the prior
   period's schedule or the GL, not the current file's own idea of where it started.
3. **GL or trial balance figures** to tie to: the prepaid account ending balance and the
   current-period amortization expense.
4. **Book and tax schedules separately**, if they differ — for example because an item
   qualifies for the 12-month rule for tax but is amortized over its service period for book.
   Do not reconcile one to the other and call the difference an error; it's a deferred item.

## Step 1 — Build or verify, and tie out

```bash
python3 scripts/prepaid_amortization.py \
  --schedule prepaid_schedule.csv \
  --prior-period prior_balances.csv \
  --tb-balance 184500.00 --tb-amortization 42300.00 \
  --period-end 2026-09-30 \
  --client "Kestrel Fabrication LLC" \
  --out "Kestrel - September 2026 Prepaid Amortization Tie-Out.xlsx"
```

`--period-start` defaults to the first of the `--period-end` month; pass it explicitly for a
quarter-end or other non-monthly close. Every argument except `--schedule`, `--period-end`, and
`--out` is optional — supply what you have. Fewer inputs means fewer tests run (each says so
explicitly) and never a guessed number filling the gap.

Six tests run:

- **Test 1** — Schedule rollforward foots, in total and by GL account
- **Test 2** — Beginning balance agrees to the prior period as closed
- **Test 3** — Schedule agrees to the GL
- **Test 4** — No item amortized beyond its total cost
- **Test 5** — Recomputed amortization agrees to the schedule (see below)
- **Test 6** — Item-level integrity (see below)

## Step 2 — Item-level review

Test 6 checks each item for the conditions that indicate a stale or broken schedule:

- **Service period ended, balance remains** — the most common real finding. An unrecorded
  expense sitting on the balance sheet as if it still had future value.
- **Amortizing before the service period begins** — a cutoff problem, not an amortization
  problem.
- **Over-amortized** — cumulative amortization exceeding cost, arithmetically impossible.
- **Missing start date, end date, or cost** — the item cannot be recomputed or built, which is
  itself a finding, not a reason to assume zero.
- **Missing GL or expense account** — nothing can be tied out or posted for this item.
- **Duplicate items** — same vendor, description, cost, and start date under different IDs. A
  double-recorded prepaid doubles amortization and is invisible in the totals.

## Step 3 — Work the proposed entries, not apply them

Every item that disagrees with its recomputation, and every item the schedule left blank,
produces one line on the **Proposed JEs** tab: a first-time build, a catch-up for a prior-period
shortfall, or a reversal for a prior-period excess. Each line states its reason. **Nothing is
posted.** This is a batch for a Controller or preparer to approve or reject line by line, the
same discipline `accrual-proposals` applies to accrual entries.

## Step 4 — Deliver

**Workbook tabs:**

1. **Tie-Out Summary** — the six tests with amounts and differences, the verdict, and the
   capitalization-rule note. The signable page.
2. **Rollforward** — beginning through ending balance by GL account, with the prior-period
   agreement shown.
3. **Schedule Detail** — every item with its recomputed figure, whether it was built or
   verified, the difference, and flags.
4. **Exceptions** — every item-level finding with the item and the condition.
5. **Proposed JEs** — the debit, credit, amount, and reason for every entry needed to bring the
   schedule's current-period figure to what the recomputation says it should be.

**Then, in chat:** whether the rollforward foots, whether the beginning balance agrees to the
prior period, the current-period amortization figure and whether it ties, and the proposed
entries worth acting on. Lead with any item whose service period already ended but still
carries a balance — it is unrecorded expense sitting on the balance sheet right now.

## What to escalate

- **Beginning balance not agreeing to the prior period as closed.** Determine whether the prior
  period, the current period, or both are wrong before proceeding — this misstates every period
  going forward and does not resolve itself.
- **A large or repeated catch-up on the same item.** One missed month is a timing slip; several
  months in a row on the same vendor is a process gap worth fixing at the source, not just
  correcting in the schedule.
- **An item whose capitalization is questionable** — paid in a lump sum, short service period,
  or small dollar amount relative to the entity's policy threshold. Route the "should this even
  be a prepaid" question to the signer; this skill only amortizes what's already on the
  schedule.
- **Book and tax schedules diverging** for the same item. That's a deferred tax item to track,
  not an error to force into agreement.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry. The
schedule is read-only. Output filenames carry client name and period only.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*

---

# Appendix — bundled files

If you installed the `.skill` package these files are already in place and you can ignore this appendix. If you copied the skill as text, create the files below at the paths shown, alongside your `SKILL.md`. The skill will not run without them.

## `scripts/prepaid_amortization.py`

```python
#!/usr/bin/env python3
"""
Prepaid expense schedule builder and tie-out: rollforward, prior-period
agreement, GL tie-out, straight-line recomputation, and proposed catch-up
journal entries.

Six tests. No clean workpaper unless the schedule rollforward foots, beginning
balances agree to the prior period as closed, the schedule agrees to the GL,
no item is amortized beyond its cost, and the straight-line recomputation
agrees to what the schedule states.

Per item, the schedule may supply a current-period amortization figure (verify
mode) or leave it blank (build mode). Either way the script recomputes the
expected cumulative amortization from the item's own start date, end date,
cost, and convention, and proposes a journal-entry line for whatever gap
exists between that recomputed figure and what the schedule currently shows —
a first-time build, or a catch-up for a prior-period shortfall or excess.
Nothing is posted; every line is a proposal for a reviewer to approve.

Usage:
    python3 prepaid_amortization.py --schedule prepaid_schedule.csv \
        --prior-period prior_balances.csv \
        --tb-balance 184500.00 --tb-amortization 42300.00 \
        --period-end 2026-09-30 --client "Kestrel Fabrication LLC" \
        --out "Kestrel - September 2026 Prepaid Amortization Tie-Out.xlsx"

--schedule CSV, one row per prepaid item:
    prepaid_id, vendor, description, gl_account, expense_account,
    payment_date, start_date, end_date, total_cost, convention,
    prior_amortized, current_amortization, ending_balance

  convention: "whole-month" (default) amortizes evenly by calendar month
  regardless of the day within the month. "daily" prorates by exact days in
  the service period. Both are pure arithmetic — no statutory figure is
  involved in prepaid amortization itself.

  prior_amortized: cumulative amortization recognized through the end of the
  PRIOR period, per the schedule as last closed. Blank defaults to 0 (a new
  item).

  current_amortization: this period's amortization per the schedule. Leave
  blank to have the tool build it from the recomputation (a brand-new item
  or a schedule being constructed from an invoice list for the first time).

  ending_balance: blank is computed as total_cost - prior_amortized -
  current_amortization (using the built figure when current_amortization is
  blank).

--prior-period CSV (the prior period's ending balances, AS CLOSED):
    gl_account, ending_balance
    (a row with gl_account "TOTAL" is accepted for totals only)
"""

from __future__ import annotations

import argparse
import csv
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is required.  pip install openpyxl")

ZERO = Decimal("0.00")
CENT = Decimal("0.01")
MONEY = '#,##0.00;[Red](#,##0.00)'
DATEF = "yyyy-mm-dd"

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
TOP = Border(top=Side(style="thin"))
DBL = Border(top=Side(style="thin"), bottom=Side(style="double"))

DAILY = {"daily", "day", "days", "d"}


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


def d0(raw):
    v = dec(raw)
    return v if v is not None else ZERO


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


def month_diff(a: date, b: date) -> int:
    return (b.year - a.year) * 12 + (b.month - a.month)


def load_schedule(path: Path) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No prepaid items in {path}")
    out = []
    for i, r in enumerate(rows, start=1):
        a = {
            "row": i,
            "prepaid_id": clean(r.get("prepaid_id")) or f"PP{i:05d}",
            "vendor": clean(r.get("vendor")),
            "description": clean(r.get("description")),
            "gl_account": clean(r.get("gl_account")) or "(unspecified)",
            "expense_account": clean(r.get("expense_account")),
            "payment_date": pdate(r.get("payment_date")),
            "start_date": pdate(r.get("start_date")),
            "end_date": pdate(r.get("end_date")),
            "total_cost": d0(r.get("total_cost")),
            "convention": clean(r.get("convention")),
            "prior_amortized": d0(r.get("prior_amortized")),
            "current_amortization": dec(r.get("current_amortization")),
            "ending_balance": dec(r.get("ending_balance")),
            "flags": [],
        }
        out.append(a)
    return out


def load_prior(path: Path | None) -> dict:
    out = {"by_account": {}, "total_balance": None}
    if not path or not path.exists():
        return out
    with path.open(newline="", encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            acct = clean(r.get("gl_account"))
            bal = dec(r.get("ending_balance"))
            if acct.upper() == "TOTAL":
                out["total_balance"] = bal
            elif acct:
                out["by_account"][acct] = bal or ZERO
    if out["total_balance"] is None and out["by_account"]:
        out["total_balance"] = sum(out["by_account"].values(), ZERO)
    return out


# ------------------------------------------------------------------ recompute

def recompute_cumulative(a: dict, as_of: date) -> Decimal | None:
    """Expected straight-line amortization from start_date through as_of,
    inclusive, capped at total_cost. None if the item lacks the basis to
    recompute at all."""
    if not a["start_date"] or not a["end_date"] or a["total_cost"] <= ZERO:
        return None
    if a["end_date"] < a["start_date"]:
        return None
    if as_of < a["start_date"]:
        return ZERO
    conv = a["convention"].lower() if a["convention"] else ""
    if conv in DAILY:
        term_days = (a["end_date"] - a["start_date"]).days + 1
        elapsed_end = min(as_of, a["end_date"])
        elapsed_days = (elapsed_end - a["start_date"]).days + 1
        frac = Decimal(elapsed_days) / Decimal(term_days)
        return min(a["total_cost"], a["total_cost"] * frac).quantize(CENT)
    term_months = month_diff(a["start_date"], a["end_date"]) + 1
    monthly = a["total_cost"] / term_months
    elapsed_months = min(month_diff(a["start_date"], min(as_of, a["end_date"])) + 1,
                         term_months)
    return min(a["total_cost"], monthly * elapsed_months).quantize(CENT)


def item_flags(items: list[dict], period_start: date, period_end: date) -> list[str]:
    problems = []
    dupes = defaultdict(list)

    for a in items:
        f = a["flags"]
        missing = not a["start_date"] or not a["end_date"] or a["total_cost"] <= ZERO

        if missing:
            f.append("missing start date, end date, or cost - cannot recompute or "
                     "build amortization for this item")
        elif a["end_date"] < a["start_date"]:
            f.append(f"end date {a['end_date']} is before start date {a['start_date']}")

        if a["total_cost"] < ZERO:
            f.append(f"negative total cost {a['total_cost']:,.2f}")
        if a["total_cost"] == ZERO and (a["prior_amortized"] != ZERO or
                                        (a["current_amortization"] not in (None, ZERO))):
            f.append("amortization recorded against zero cost")

        if not a["gl_account"] or a["gl_account"] == "(unspecified)":
            f.append("missing GL (prepaid asset) account - cannot tie to the GL")
        if not a["expense_account"]:
            f.append("missing expense account - cannot post the proposed entry")

        if a["end_date"] and a["end_date"] < period_start and a["ending_used"] != ZERO:
            f.append(f"service period ended {a['end_date']}, before this period began, "
                     f"but {a['ending_used']:,.2f} remains unamortized on the schedule - "
                     f"should be fully expensed")
        if a["start_date"] and a["start_date"] > period_end and a["current_used"] != ZERO:
            f.append(f"service period does not begin until {a['start_date']}, but "
                     f"{a['current_used']:,.2f} of amortization is recorded this period - "
                     f"premature recognition")

        if a["built"]:
            if a["total_cost"] > ZERO and a["prior_amortized"] > a["total_cost"]:
                f.append(f"prior amortization {a['prior_amortized']:,.2f} already exceeds "
                         f"total cost {a['total_cost']:,.2f} before this period's entry "
                         f"is even built")
        else:
            cum_actual = a["prior_amortized"] + a["current_used"]
            if a["total_cost"] > ZERO and cum_actual > a["total_cost"]:
                f.append(f"OVER-AMORTIZED: prior {a['prior_amortized']:,.2f} + current "
                         f"{a['current_used']:,.2f} = {cum_actual:,.2f}, which exceeds "
                         f"total cost {a['total_cost']:,.2f} by "
                         f"{cum_actual - a['total_cost']:,.2f} - arithmetically impossible")

        if a["recomp_diff"] not in (None, ZERO):
            f.append(f"recomputed amortization differs from the schedule by "
                     f"{a['recomp_diff']:,.2f} - proposed as a catch-up entry")

        dupes[(a["vendor"].lower(), clean(a["description"]).lower(), a["total_cost"],
               a["start_date"])].append(a)

        for msg in f:
            problems.append(f"{a['prepaid_id']} {a['description'][:36]}: {msg}")

    for key, group in dupes.items():
        if len(group) > 1 and key[2] > ZERO and key[1]:
            ids = ", ".join(x["prepaid_id"] for x in group)
            for x in group:
                x["flags"].append(f"possible DUPLICATE item: same vendor, description, "
                                  f"cost and start date as {ids} - a double-recorded "
                                  f"prepaid doubles amortization and is invisible in totals")
            problems.append(f"Possible duplicate items ({ids}): {group[0]['description']} "
                            f"at {key[2]:,.2f}")
    return problems


# ---------------------------------------------------------------------- tests

def rollforward(items: list[dict], prior: dict, period_start: date):
    by_account = defaultdict(lambda: {"begin_balance": ZERO, "additions": ZERO,
                                       "current": ZERO, "end_balance": ZERO})
    for a in items:
        c = by_account[a["gl_account"]]
        is_addition = (a["prior_amortized"] == ZERO and a["start_date"] and
                      period_start <= a["start_date"])
        if is_addition:
            c["additions"] += a["total_cost"]
        else:
            c["begin_balance"] += a["total_cost"] - a["prior_amortized"]
        c["current"] += a["current_used"]
        c["end_balance"] += a["ending_used"]

    rows = []
    for acct, c in sorted(by_account.items()):
        computed_end = c["begin_balance"] + c["additions"] - c["current"]
        rows.append({
            "gl_account": acct, **c,
            "computed_end": computed_end,
            "diff": computed_end - c["end_balance"],
            "prior_balance": prior["by_account"].get(acct),
        })
    return rows


def totals_of(rows, key):
    return sum((r[key] for r in rows), ZERO)


# -------------------------------------------------------------------- workbook

def hdr(ws, n, row=1):
    for c in range(1, n + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def widths(ws, w):
    for c, v in w.items():
        ws.column_dimensions[get_column_letter(c)].width = v


def sheet_summary(wb, meta, tests, overall, escalations, notes):
    ws = wb.active
    ws.title = "Tie-Out Summary"
    widths(ws, {1: 4, 2: 52, 3: 18, 4: 18, 5: 62})
    r = 1

    def line(label, a=None, b=None, *, bold=False, size=11, fill=None, note="",
             border=None):
        nonlocal r
        c = ws.cell(row=r, column=2, value=label)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c.fill = fill
        for col, val in ((3, a), (4, b)):
            if val is None:
                continue
            cc = ws.cell(row=r, column=col,
                         value=float(val) if isinstance(val, Decimal) else val)
            if isinstance(val, Decimal):
                cc.number_format = MONEY
            cc.font = Font(bold=bold)
            if border:
                cc.border = border
        if note:
            n = ws.cell(row=r, column=5, value=note)
            n.font = Font(italic=True, size=9)
            n.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    line("PREPAID EXPENSE AMORTIZATION TIE-OUT", bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    v = ws.cell(row=r, column=2,
                value="TIED OUT - schedule foots, beginning balances agree, GL ties"
                if overall else
                "NOT TIED OUT - see the failing test(s) below")
    v.font = OK_FONT if overall else BAD_FONT
    if not overall:
        v.fill = BAD_FILL
    r += 2

    for t in tests:
        line(f"TEST {t['num']} - {t['name']}", bold=True, size=12, fill=SUB_FILL)
        if t.get("skipped"):
            line("  " + t["note"], fill=WARN_FILL)
            r += 1
            continue
        counts = set(t.get("counts", []))
        for label, val in t["lines"]:
            line("  " + label, int(val) if label in counts else val)
        line("  Items flagged" if t.get("is_count") else "  Difference",
             int(t["difference"]) if t.get("is_count") else t["difference"],
             bold=True, border=TOP, note=t.get("note", ""))
        c = ws.cell(row=r - 1, column=3)
        c.font = OK_FONT if t["passed"] else BAD_FONT
        if not t["passed"]:
            c.fill = BAD_FILL
        r += 1

    if escalations:
        line("ESCALATE", bold=True, size=12, fill=BAD_FILL)
        for e in escalations:
            line("  " + e)
        r += 1

    line("NOTES", bold=True, size=12, fill=WARN_FILL)
    for n in notes:
        line("  " + n)
    r += 1
    line("This workpaper performs no judgment about whether a prepayment should be "
         "capitalized at all. Whether a given item is subject to the 12-month rule "
         "under Reg. 1.263(a)-4(f), or falls under the entity's own capitalization "
         "policy threshold, must be verified separately - this script assumes items "
         "on the schedule are already properly capitalized and recomputes only the "
         "arithmetic of the amortization itself.", bold=True)
    r += 2
    line("Prepared by / date:  __________________  ____________", bold=True)
    line("Reviewed by / date:  __________________  ____________", bold=True)


def sheet_rollforward(wb, rows):
    ws = wb.create_sheet("Rollforward")
    heads = ["GL account", "Begin balance", "Additions", "Current amortization",
             "Computed end balance", "End balance per schedule", "Diff",
             "PY/prior-period end balance", "Beginning diff vs prior period"]
    ws.append(heads)
    hdr(ws, len(heads))
    for x in rows:
        pd = (x["begin_balance"] - x["prior_balance"]) if x["prior_balance"] is not None else None
        ws.append([x["gl_account"], float(x["begin_balance"]), float(x["additions"]),
                   float(x["current"]), float(x["computed_end"]), float(x["end_balance"]),
                   float(x["diff"]),
                   float(x["prior_balance"]) if x["prior_balance"] is not None else None,
                   float(pd) if pd is not None else None])
    last = ws.max_row
    ws.append(["TOTAL"] + [f"=SUM({get_column_letter(c)}2:{get_column_letter(c)}{last})"
                           for c in range(2, len(heads) + 1)])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in range(1, len(heads)):
            row[i].number_format = MONEY
        for i in (6, 8):
            if isinstance(row[i].value, (int, float)) and row[i].value:
                row[i].font, row[i].fill = BAD_FONT, BAD_FILL
    for c in range(1, len(heads) + 1):
        cell = ws.cell(row=ws.max_row, column=c)
        cell.font, cell.border = Font(bold=True), TOP
    ws.freeze_panes = "B2"
    widths(ws, {i: 18 for i in range(1, len(heads) + 1)} | {1: 24})


def sheet_items(wb, items):
    ws = wb.create_sheet("Schedule Detail")
    heads = ["Prepaid ID", "Vendor", "Description", "GL account", "Expense account",
             "Payment date", "Start", "End", "Total cost", "Convention",
             "Prior amortized", "Current (used)", "Built?", "Recomputed diff",
             "Ending balance", "Flags"]
    ws.append(heads)
    hdr(ws, len(heads))
    for a in items:
        ws.append([a["prepaid_id"], a["vendor"], a["description"], a["gl_account"],
                   a["expense_account"] or None, a["payment_date"], a["start_date"],
                   a["end_date"], float(a["total_cost"]), a["convention"] or "whole-month",
                   float(a["prior_amortized"]), float(a["current_used"]),
                   "Yes" if a["built"] else "No",
                   float(a["recomp_diff"]) if a["recomp_diff"] is not None else None,
                   float(a["ending_used"]), "; ".join(a["flags"])])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in (5, 6, 7):
            row[i].number_format = DATEF
        for i in (8, 10, 11, 13, 14):
            row[i].number_format = MONEY
        if row[13].value not in (None, 0):
            row[13].font, row[13].fill = BAD_FONT, BAD_FILL
        if row[15].value:
            row[15].fill = BAD_FILL
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:P{max(ws.max_row, 2)}"
    widths(ws, {1: 12, 2: 22, 3: 32, 4: 18, 5: 18, 6: 12, 7: 12, 8: 12, 9: 13,
                10: 13, 11: 14, 12: 14, 13: 8, 14: 14, 15: 14, 16: 80})


def sheet_list(wb, title, rows, heads, fmt_map, empty_msg, footer=""):
    ws = wb.create_sheet(title)
    ws.append(heads)
    hdr(ws, len(heads))
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i, f in fmt_map.items():
            if i < len(row):
                row[i].number_format = f
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value=empty_msg)
    elif footer:
        rr = ws.max_row + 2
        c = ws.cell(row=rr, column=1, value=footer)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.font = Font(italic=True)
    ws.freeze_panes = "A2"
    widths(ws, {i: 18 for i in range(1, len(heads) + 1)} | {2: 34, len(heads): 70})


# ------------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--schedule", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--prior-period")
    ap.add_argument("--tb-balance")
    ap.add_argument("--tb-amortization")
    ap.add_argument("--period-end", required=True, help="YYYY-MM-DD")
    ap.add_argument("--period-start", help="YYYY-MM-DD, defaults to the 1st of the "
                                            "period-end's month")
    ap.add_argument("--client", default="")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    period_end = pdate(args.period_end)
    period_start = pdate(args.period_start) if args.period_start else period_end.replace(day=1)

    items = load_schedule(Path(args.schedule))
    prior = load_prior(Path(args.prior_period) if args.prior_period else None)

    for a in items:
        expected_cum = recompute_cumulative(a, period_end)
        expected_current = (expected_cum - a["prior_amortized"]) if expected_cum is not None else None
        a["expected_current"] = expected_current
        a["built"] = a["current_amortization"] is None
        if a["built"]:
            a["current_used"] = expected_current if expected_current is not None else ZERO
            a["recomp_diff"] = None
        else:
            a["current_used"] = a["current_amortization"]
            a["recomp_diff"] = ((expected_current - a["current_amortization"]).quantize(CENT)
                                if expected_current is not None else None)
        a["ending_used"] = (a["ending_balance"] if a["ending_balance"] is not None else
                            (a["total_cost"] - a["prior_amortized"] - a["current_used"]))

    problems = item_flags(items, period_start, period_end)
    rows = rollforward(items, prior, period_start)

    tot = {k: totals_of(rows, k) for k in
           ("begin_balance", "additions", "current", "end_balance", "computed_end")}
    tb_balance, tb_amort = dec(args.tb_balance), dec(args.tb_amortization)

    tests = []

    tests.append({"num": 1, "name": "Schedule rollforward foots",
                  "lines": [("Beginning balance", tot["begin_balance"]),
                            ("Add: additions (new items)", tot["additions"]),
                            ("Less: current-period amortization", -tot["current"]),
                            ("= Computed ending balance", tot["computed_end"]),
                            ("Ending balance per schedule", tot["end_balance"])],
                  "difference": tot["computed_end"] - tot["end_balance"],
                  "passed": tot["computed_end"] == tot["end_balance"]})

    if prior["total_balance"] is not None:
        d = tot["begin_balance"] - prior["total_balance"]
        tests.append({"num": 2,
                      "name": "Beginning balance agrees to the prior period as closed",
                      "lines": [("Beginning balance per schedule", tot["begin_balance"]),
                                ("Prior-period ending balance as closed",
                                 prior["total_balance"])],
                      "difference": d, "passed": d == ZERO,
                      "note": "" if d == ZERO else
                              f"off by {d:,.2f}. This silently misstates every "
                              f"subsequent period and does not resolve itself."})
    else:
        tests.append({"num": 2, "name": "Beginning balance agrees to the prior period",
                      "skipped": True, "passed": True,
                      "note": "NOT PERFORMED - prior-period ending balances not supplied. "
                              "Obtain the prior period's schedule or GL balance as closed.",
                      "lines": [], "difference": ZERO})

    if tb_balance is not None or tb_amort is not None:
        lines, diff = [], ZERO
        for label, sched, tb in (("Ending balance", tot["end_balance"], tb_balance),
                                 ("Current-period amortization", tot["current"], tb_amort)):
            if tb is None:
                continue
            lines += [(f"{label} per schedule", sched), (f"{label} per GL", tb)]
            diff += sched - tb
        tests.append({"num": 3, "name": "Schedule agrees to the GL",
                      "lines": lines, "difference": diff, "passed": diff == ZERO})
    else:
        tests.append({"num": 3, "name": "Schedule agrees to the GL",
                      "skipped": True, "passed": True,
                      "note": "NOT PERFORMED - no GL balance or amortization figure "
                              "supplied", "lines": [], "difference": ZERO})

    over = [a for a in items if not a["built"] and a["total_cost"] > ZERO
            and (a["prior_amortized"] + a["current_used"]) > a["total_cost"]]
    tests.append({"num": 4, "name": "No item amortized beyond its total cost",
                  "counts": ["Items over-amortized"],
                  "lines": [("Items over-amortized", Decimal(len(over))),
                            ("Excess amortization",
                             sum(((a["prior_amortized"] + a["current_used"]) - a["total_cost"]
                                  for a in over), ZERO))],
                  "difference": sum(((a["prior_amortized"] + a["current_used"]) - a["total_cost"]
                                     for a in over), ZERO),
                  "passed": not over})

    tested = [a for a in items if a["recomp_diff"] is not None]
    bad = [a for a in tested if a["recomp_diff"] != ZERO]
    tests.append({"num": 5, "name": "Recomputed amortization agrees to the schedule",
                  "counts": ["Items recomputed", "Items disagreeing"],
                  "lines": [("Items recomputed", Decimal(len(tested))),
                            ("Items disagreeing", Decimal(len(bad))),
                            ("Net catch-up proposed",
                             sum((a["recomp_diff"] for a in bad), ZERO))],
                  "difference": sum((a["recomp_diff"] for a in bad), ZERO),
                  "passed": not bad,
                  "note": "Each disagreement is proposed as a catch-up journal entry on "
                          "the Proposed JEs tab, not applied automatically."})

    n_flagged = len([a for a in items if a["flags"]])
    tests.append({"num": 6, "name": "Item-level integrity",
                  "counts": ["Items in schedule", "Items with findings"],
                  "is_count": True,
                  "lines": [("Items in schedule", Decimal(len(items))),
                            ("Items with findings", Decimal(n_flagged))],
                  "difference": Decimal(n_flagged),
                  "passed": n_flagged == 0,
                  "note": "" if not n_flagged else
                          f"{n_flagged} item(s) carry findings - see Exceptions and "
                          f"Schedule Detail. The count is items, not dollars."})

    overall = all(t["passed"] for t in tests)

    escalations = []
    t2 = tests[1]
    if not t2.get("skipped") and not t2["passed"]:
        escalations.append(
            "Beginning balance does not agree to the prior period as closed. Determine "
            "whether the prior period, the current period, or both are wrong before "
            "proceeding.")
    expired = [a for a in items if a["end_date"] and a["end_date"] < period_start
               and a["ending_used"] != ZERO]
    if expired:
        escalations.append(
            f"{len(expired)} item(s) have a service period that already ended but still "
            f"carry a balance totalling {sum((a['ending_used'] for a in expired), ZERO):,.2f} "
            f"- an unrecorded expense sitting on the balance sheet.")
    if over:
        escalations.append(
            f"{len(over)} item(s) amortized beyond cost by "
            f"{sum(((a['prior_amortized'] + a['current_used']) - a['total_cost'] for a in over), ZERO):,.2f}.")
    premature = [a for a in items if a["start_date"] and a["start_date"] > period_end
                 and a["current_used"] != ZERO]
    if premature:
        escalations.append(
            f"{len(premature)} item(s) show amortization before their service period "
            f"begins - a cutoff issue to resolve before the batch is approved.")

    notes = [
        "Whether a prepayment must be capitalized at all, or may be currently deducted "
        "under the 12-month rule, is not evaluated here - verify against current "
        "instructions for Reg. 1.263(a)-4(f) and the entity's capitalization policy.",
        "Book and tax prepaid schedules should be maintained and tied out separately - "
        "do not reconcile a book schedule to a tax figure and call the difference an "
        "error.",
    ]

    je_rows = []
    for a in items:
        amount = a["current_used"] if a["built"] else a["recomp_diff"]
        if amount in (None, ZERO):
            continue
        reason = ("first-time build - no amortization figure was on the schedule" if a["built"]
                  else ("catch-up for a recomputation gap" if amount > ZERO else
                        "reversal - prior amortization overstated"))
        if amount > ZERO:
            debit, credit = a["expense_account"], a["gl_account"]
        else:
            debit, credit = a["gl_account"], a["expense_account"]
        je_rows.append([a["prepaid_id"], a["vendor"], a["description"], debit, credit,
                        float(abs(amount)), reason])

    # ---- console
    print("=" * 74)
    print("PREPAID EXPENSE AMORTIZATION TIE-OUT")
    print("=" * 74)
    print(f"Client : {args.client or '(not stated)'}    Period: {period_start} to "
          f"{period_end}")
    print(f"Items  : {len(items)}")
    print()
    for t in tests:
        if t.get("skipped"):
            print(f"Test {t['num']}: NOT PERFORMED - {t['note'][:80]}")
            continue
        if t.get("is_count"):
            print(f"Test {t['num']}: {'PASS' if t['passed'] else '*** FAIL ***':<14} "
                  f"{int(t['difference']):>10} item(s) flagged   {t['name']}")
        else:
            print(f"Test {t['num']}: {'PASS' if t['passed'] else '*** FAIL ***':<14} "
                  f"diff {t['difference']:>16,.2f}   {t['name']}")
    print()
    print(f"Balance   : begin {tot['begin_balance']:>14,.2f} + add {tot['additions']:>13,.2f} "
          f"- current {tot['current']:>13,.2f} = {tot['computed_end']:>14,.2f}")
    print(f"Proposed JE lines: {len(je_rows)}")

    if problems:
        print(f"\nITEM-LEVEL FINDINGS ({len(problems)}):")
        for p in problems[:20]:
            print(f"  ! {p}")
        if len(problems) > 20:
            print(f"  ... and {len(problems) - 20} more - see Schedule Detail")
    if escalations:
        print("\nESCALATE:")
        for e in escalations:
            print(f"  ! {e}")

    if not overall and not args.force:
        print("\n" + "=" * 74)
        print("WORKBOOK NOT WRITTEN. The schedule does not tie out.")
        print("A wrong beginning balance misstates every subsequent period - resolve "
              "before relying on the amortization figure.")
        print("=" * 74)
        return 1

    meta = {
        "Client": args.client or "(not stated)",
        "Period start": str(period_start),
        "Period end": str(period_end),
        "Items in schedule": len(items),
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    wb = Workbook()
    sheet_summary(wb, meta, tests, overall, escalations, notes)
    sheet_rollforward(wb, rows)
    sheet_items(wb, items)
    sheet_list(wb, "Exceptions",
               [[i + 1, p] for i, p in enumerate(problems)],
               ["#", "Finding"], {}, "None - no item-level findings.")
    sheet_list(wb, "Proposed JEs", je_rows,
               ["Prepaid ID", "Vendor", "Description", "Debit account", "Credit account",
                "Amount", "Reason"],
               {5: MONEY},
               "No journal entries proposed - every item already agrees to its "
               "recomputed amortization.",
               footer="Nothing here is posted. Each line is a proposal for a reviewer "
                      "to approve or reject before it reaches the GL.")
    wb.active = 0
    out = Path(args.out)
    if not overall:
        out = out.with_name(out.stem + " [DOES NOT TIE]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\n{'TIED OUT' if overall else 'DOES NOT TIE (forced)'}.  Workbook: {out}")
    return 0 if overall else 1


if __name__ == "__main__":
    sys.exit(main())
```

