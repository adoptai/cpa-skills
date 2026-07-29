---
name: bank-rec-review
description: Audit a bank reconciliation that someone else prepared — independently recompute it, verify every reconciling item against the bank statement and the subsequent-period statement, confirm each item actually cleared, and detect plugs, stale items rolled forward, and reconciling items that exist only on the reconciliation. Use this whenever the user mentions reviewing or testing a bank reconciliation, auditing a client's reconciliation, verifying reconciling items, whether a reconciliation is supported, an outstanding check that never cleared, a suspicious reconciling item, or cash audit procedures. Also trigger on "review the client's bank rec," "test this reconciliation," "did these items clear," "is this reconciliation supported," "audit cash," or when a prepared reconciliation is provided alongside a bank statement. Distinct from preparing a reconciliation — for that, see bank-rec-to-gl. Runs fully local — no client cash data leaves the machine.
---

# Bank Reconciliation Review

Preparing a reconciliation and reviewing one are different procedures. The preparer starts from
two records and works toward agreement. **The reviewer starts from a document that already claims
agreement, and has to establish whether that claim is true.** Those are not the same task, and the
review is the easier one to do badly, because a reconciliation that foots looks finished.

A reconciliation can foot perfectly and still be worthless. The three ways:

1. **A fabricated reconciling item.** An "outstanding check" or "deposit in transit" that exists
   only on the reconciliation, inserted to make the two sides agree. It has no counterpart in any
   bank statement, before or after.
2. **A stale item rolled forward.** A genuine item from years ago that has never cleared and never
   been investigated, carried each month because removing it would break the reconciliation.
3. **A plug with a plausible label.** `Timing`, `Bank adjustment`, `Unidentified deposit`,
   `Items in transit`. The label is doing the work, not the item.

All three survive a review that only checks the arithmetic. **The test that catches all three is
subsequent clearance**: a real reconciling item appears in the next period's bank statement. A
fabricated one never does.

## The gate

No clean workpaper unless:

1. **The reconciliation is independently recomputed** and the recomputation agrees to what the
   preparer presented. Do not accept the preparer's subtotals.
2. **The bank balance is agreed to the bank statement** and the book balance to the general
   ledger — as two independent sources. Agreeing the reconciliation to itself proves nothing.
3. **Every reconciling item is supported** — traced to the bank statement, the subsequent-period
   statement, or a document. An item with no support is reported, never accepted.
4. **Subsequent clearance is tested** where a subsequent statement is available, and every item
   that did not clear is reported individually.

## Inputs

1. **The reconciliation as prepared**, item by item — not just the totals.
2. **The bank statement** for the period, with its closing balance.
3. **The general ledger** cash balance at the same date, ideally with the account detail.
4. **The subsequent-period bank statement.** This is the most important document in the review and
   it is frequently not requested. Without it, clearance cannot be tested and the review is
   limited in scope.
5. **Prior-period reconciliations**, to detect items rolled forward.
6. **Who prepared and who reviewed it, and when.** A reconciliation prepared two months after
   period end by the person with disbursement authority, unreviewed, is a control finding
   independent of its arithmetic.

## Step 1 — Recompute and test

```bash
python3 scripts/rec_review.py \
  --reconciliation client_rec.csv \
  --bank-balance 1284551.09 \
  --gl-balance 1291204.77 \
  --subsequent-statement feb_statement.csv \
  --prior-reconciliation prior_rec.csv \
  --period-end 2025-12-31 \
  --prepared-by "J. Mensah" --prepared-date 2026-02-18 \
  --reviewed-by "" \
  --client "Fairmont Logistics" \
  --out "Fairmont - 2025 Bank Rec Review.xlsx"
```

Eight tests:

- **Test 1 — Independent recomputation.** Bank balance ± reconciling items = adjusted bank; GL
  balance ± unrecorded items = adjusted book; the two must be equal. Computed from the item detail,
  not from the preparer's subtotals.
- **Test 2 — Bank balance agreed to the statement.** An independent source.
- **Test 3 — Book balance agreed to the general ledger.** The other independent source.
- **Test 4 — Every item supported.** Each reconciling item traced to a document.
- **Test 5 — Subsequent clearance.** Each item matched to the subsequent-period statement.
  Anything that did not clear is listed.
- **Test 6 — Stale items.** Items appearing on the prior reconciliation with the same amount,
  aged, with the number of periods they have been carried.
- **Test 7 — Plug detection.** Items whose description is generic, whose amount is suspiciously
  round, or that appear only on the reconciliation with no counterpart anywhere.
- **Test 8 — Preparation and review controls.** Timeliness, whether it was reviewed, and whether
  the preparer also has disbursement authority.

## Step 2 — Read subsequent clearance carefully

Clearance is the strongest evidence available here, but read it correctly:

- **An outstanding check that cleared in the subsequent period** for the same amount: supported.
  Good.
- **Cleared for a different amount**: the reconciliation was wrong, or the check was altered.
  Worth pursuing either way.
- **Did not clear at all** and is more than a couple of months old: investigate. Was it ever
  issued? Was it delivered?
- **A deposit in transit that did not clear within days**: a deposit in transit is by definition a
  deposit already made. One that never arrives at the bank is not timing — it is an exception, and
  it is the classic lapping and fictitious-receipts pattern.
- **An item that cleared before period end**: it was not outstanding at all and should never have
  been on the reconciliation.

The script computes days-to-clear for each item and flags the ones that do not fit their claimed
category.

## Step 3 — Deliver

**Workbook tabs:**

1. **Review Summary** — independent recomputation against the preparer's figures, the eight tests,
   the control observations, and the conclusion. The signable page.
2. **Independent Recomputation** — the four-column form rebuilt from item detail, alongside what
   the preparer presented, with differences.
3. **Item Testing** — every reconciling item: amount, category, support obtained, whether it
   cleared, when, days to clear, and any flag. Unsupported and uncleared items sort to the top.
4. **Uncleared Items** — everything that did not clear, with age and the follow-up required.
5. **Stale Items** — items carried from prior reconciliations, with the number of periods.
6. **Plug Indicators** — generic descriptions, round amounts, and items with no counterpart.
7. **Control Observations** — preparation timeliness, review evidence, segregation of duties.

**Then, in chat:** whether the recomputation agrees, whether both balances agreed to independent
sources, how many items lacked support, how many did not clear, and the control observations. Lead
with anything unsupported or uncleared — those are the findings; the arithmetic rarely is.

## What to escalate

- **A reconciling item with no counterpart in any bank statement, before or after.** This is the
  fabricated-item pattern and it is escalated regardless of amount.
- **A deposit in transit that never cleared.**
- **An item carried on three or more consecutive reconciliations.** It is not timing.
- **A reconciliation that only foots because of an item labelled generically.** Report the label
  and the amount together — the juxtaposition is the finding.
- **An unreviewed reconciliation**, or one prepared by someone with disbursement authority.
- **A reconciliation prepared long after period end.** The longer the delay, the more opportunity
  to construct rather than reconcile.
- **Adjusting entries made to the cash account rather than to the reconciliation** — these bypass
  the reconciliation entirely.
- **Prior-period reconciling items that disappeared without clearing.** Something was removed
  rather than resolved, and the reconciliation still foots, which means something else absorbed it.

State each as a fact with the item reference and amount. Do not characterise intent.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*

---

# Appendix — bundled files

If you installed the `.skill` package these files are already in place and you can ignore this appendix. If you copied the skill as text, create the files below at the paths shown, alongside your `SKILL.md`. The skill will not run without them.

## `scripts/rec_review.py`

```python
#!/usr/bin/env python3
"""
Review a bank reconciliation that someone else prepared.

Reviewing is not preparing. The preparer works from two records toward agreement;
the reviewer starts from a document that already claims agreement and has to
establish whether the claim is true. A reconciliation that foots can still be
worthless.

The test that catches a fabricated item, a stale item, and a plug alike is
SUBSEQUENT CLEARANCE: a real reconciling item appears in the next period's bank
statement. A fabricated one never does.

Gates:
  * the reconciliation is recomputed from item detail, not from the preparer's
    subtotals
  * the bank balance is agreed to the bank statement and the book balance to the
    general ledger, as two INDEPENDENT sources
  * every reconciling item must be supported
  * subsequent clearance is tested and every item that did not clear is reported

Usage:
    python3 rec_review.py --reconciliation client_rec.csv \
        --bank-balance 1284551.09 --gl-balance 1291204.77 \
        --subsequent-statement feb_statement.csv \
        --prior-reconciliation prior_rec.csv \
        --period-end 2025-12-31 \
        --prepared-by "J. Mensah" --prepared-date 2026-02-18 --reviewed-by "" \
        --client "Fairmont Logistics" \
        --out "Fairmont - 2025 Bank Rec Review.xlsx"

--reconciliation CSV (item detail, not totals):
    item_id, category, description, reference, item_date, amount, support
      category: deposit_in_transit | outstanding_check | unrecorded_debit |
                unrecorded_credit | bank_error | book_error | other
      amount: always POSITIVE; the category determines the sign applied
      support: document reference obtained, or blank

--subsequent-statement CSV:  date, description, reference, amount
--prior-reconciliation CSV:  item_id, description, reference, amount, periods_carried
--presented CSV (optional):  line, amount   the preparer's own subtotals
"""

from __future__ import annotations

import argparse
import csv
import sys
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is required.  pip install openpyxl")

ZERO = Decimal("0.00")
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

# side, sign applied to the positive amount
CATEGORIES = {
    "deposit_in_transit":  ("bank", Decimal(1)),
    "outstanding_check":   ("bank", Decimal(-1)),
    "bank_error":          ("bank", Decimal(1)),
    "unrecorded_credit":   ("book", Decimal(1)),
    "unrecorded_debit":    ("book", Decimal(-1)),
    "book_error":          ("book", Decimal(1)),
    "other":               ("bank", Decimal(1)),
}

GENERIC = ["timing", "bank adj", "bank adjustment", "unidentified", "items in transit",
           "in transit", "misc", "miscellaneous", "difference", "to balance",
           "plug", "adjustment", "unknown", "reconciling difference", "other",
           "per client", "to agree", "clearing"]


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


def key(s) -> str:
    return "".join(ch for ch in str(s or "").lower() if ch.isalnum())


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


def rows_of(path: Path, label: str):
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No rows in {label} ({path})")
    return rows


def load_rec(path: Path) -> list[dict]:
    out = []
    for i, r in enumerate(rows_of(path, "reconciliation"), start=1):
        cat = clean(r.get("category")).lower()
        if cat not in CATEGORIES:
            sys.exit(f"Reconciliation row {i}: category '{cat}' is not recognised. "
                     f"Valid categories are {sorted(CATEGORIES)}. The category determines "
                     f"which side of the reconciliation the item sits on and what sign "
                     f"is applied, so it cannot be guessed.")
        amt = d0(r.get("amount"))
        if amt < ZERO:
            sys.exit(f"Reconciliation row {i}: amount {amt} is negative. Enter amounts as "
                     f"positive; the category determines the sign. A negative amount in a "
                     f"category that already carries a sign silently reverses the item.")
        out.append({
            "row": i,
            "item_id": clean(r.get("item_id")) or f"RI{i:04d}",
            "category": cat,
            "side": CATEGORIES[cat][0],
            "sign": CATEGORIES[cat][1],
            "description": clean(r.get("description")),
            "reference": clean(r.get("reference")),
            "item_date": pdate(r.get("item_date")),
            "amount": amt,
            "support": clean(r.get("support")),
            "cleared": None, "cleared_date": None, "cleared_amount": None,
            "days_to_clear": None, "periods_carried": 0, "flags": [],
        })
    return out


def load_statement(path: Path | None) -> list[dict]:
    if not path or not path.exists():
        return []
    out = []
    for r in rows_of(path, "statement"):
        out.append({"date": pdate(r.get("date")),
                    "description": clean(r.get("description")),
                    "reference": clean(r.get("reference")),
                    "amount": d0(r.get("amount"))})
    return out


# ----------------------------------------------------------------------- tests

def test_clearance(items, subseq, period_end: date) -> None:
    """Match each item to the subsequent statement. Absence is the finding."""
    if not subseq:
        return
    pool = list(subseq)
    for x in items:
        target = x["amount"]
        hit = None
        # reference first, then amount
        if x["reference"]:
            for s in pool:
                if s["reference"] and key(s["reference"]) == key(x["reference"]):
                    hit = s
                    break
        if hit is None:
            for s in pool:
                if abs(s["amount"]) == target:
                    hit = s
                    break
        if hit is None:
            # cleared for a different amount? match on reference only
            x["cleared"] = False
            continue
        pool.remove(hit)
        x["cleared"] = True
        x["cleared_date"] = hit["date"]
        x["cleared_amount"] = abs(hit["amount"])
        if x["item_date"] and hit["date"]:
            x["days_to_clear"] = (hit["date"] - x["item_date"]).days
        if x["cleared_amount"] != target:
            x["flags"].append(
                f"cleared for {x['cleared_amount']:,.2f} against {target:,.2f} on the "
                f"reconciliation - the reconciliation was wrong, or the instrument was "
                f"altered. Worth pursuing either way.")
        if hit["date"] and hit["date"] <= period_end:
            x["flags"].append(
                f"cleared {hit['date']}, on or before period end {period_end} - it was "
                f"not outstanding at all and should not have been on the reconciliation")


def test_support_and_plugs(items, subseq, prior, period_end: date) -> None:
    prior_idx = {}
    for p in prior:
        prior_idx[(key(p.get("reference")), d0(p.get("amount")))] = p
        prior_idx[(key(p.get("description")), d0(p.get("amount")))] = p

    for x in items:
        if not x["support"]:
            x["flags"].append("NO SUPPORT OBTAINED - traced to no document")

        d = x["description"].lower().strip()
        if not d:
            x["flags"].append("no description")
        elif any(d == g or d.startswith(g) for g in GENERIC):
            x["flags"].append(
                f'GENERIC DESCRIPTION "{x["description"]}" - the label is doing the work '
                f'rather than the item. Report the label and the amount together.')

        a = x["amount"]
        if a >= Decimal("1000"):
            for m in (Decimal("100000"), Decimal("50000"), Decimal("25000"),
                      Decimal("10000"), Decimal("5000"), Decimal("1000")):
                if m >= a / Decimal(20) and a % m == ZERO:
                    x["flags"].append(f"round amount ({a:,.2f}) - confirm it is a real "
                                      f"item rather than a balancing figure")
                    break

        pk = prior_idx.get((key(x["reference"]), x["amount"])) or \
             prior_idx.get((key(x["description"]), x["amount"]))
        if pk:
            carried = int(d0(pk.get("periods_carried")) or 1)
            x["periods_carried"] = carried + 1
            if x["periods_carried"] >= 3:
                x["flags"].append(
                    f"carried on {x['periods_carried']} consecutive reconciliations - "
                    f"this is not timing")
            else:
                x["flags"].append(f"also on the prior reconciliation "
                                  f"({x['periods_carried']} periods carried)")

        if x["cleared"] is False and subseq:
            age = (period_end - x["item_date"]).days if x["item_date"] else None
            if x["category"] == "deposit_in_transit":
                x["flags"].append(
                    "DEPOSIT IN TRANSIT THAT NEVER CLEARED - a deposit in transit is by "
                    "definition already made. One that never arrives is not timing; it is "
                    "the classic lapping and fictitious-receipts pattern.")
            elif age is not None and age > 60:
                x["flags"].append(
                    f"did not clear and is {age} days old - was it ever issued? delivered?")
            else:
                x["flags"].append("did not clear in the subsequent period")
            if not x["support"]:
                x["flags"].append(
                    "NO COUNTERPART ANYWHERE - not supported by a document and did not "
                    "appear in the subsequent statement. This is the fabricated-item "
                    "pattern; escalate regardless of amount.")


def recompute(items, bank_balance: Decimal, gl_balance: Decimal) -> dict:
    bank_adj = sum((x["amount"] * x["sign"] for x in items if x["side"] == "bank"), ZERO)
    book_adj = sum((x["amount"] * x["sign"] for x in items if x["side"] == "book"), ZERO)
    adj_bank = bank_balance + bank_adj
    adj_book = gl_balance + book_adj
    return {"bank_balance": bank_balance, "bank_adjustments": bank_adj,
            "adjusted_bank": adj_bank, "gl_balance": gl_balance,
            "book_adjustments": book_adj, "adjusted_book": adj_book,
            "difference": adj_bank - adj_book}


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


def sheet_summary(wb, meta, tests, rc, esc, overall, controls):
    ws = wb.active
    ws.title = "Review Summary"
    widths(ws, {1: 4, 2: 54, 3: 18, 4: 14, 5: 60})
    r = 1

    def line(label, v=None, s=None, *, bold=False, size=11, fill=None, border=None,
             note=""):
        nonlocal r
        c = ws.cell(row=r, column=2, value=label)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c.fill = fill
        if v is not None:
            cc = ws.cell(row=r, column=3,
                         value=float(v) if isinstance(v, Decimal) else v)
            if isinstance(v, Decimal):
                cc.number_format = MONEY
            cc.font = Font(bold=bold)
            if border:
                cc.border = border
        if s is not None:
            sc = ws.cell(row=r, column=4, value=s)
            sc.font = OK_FONT if s == "PASS" else BAD_FONT
            if s != "PASS":
                sc.fill = BAD_FILL
        if note:
            n = ws.cell(row=r, column=5, value=note)
            n.font = Font(italic=True, size=9)
            n.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    line("BANK RECONCILIATION REVIEW", bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    v = ws.cell(row=r, column=2,
                value="RECONCILIATION SUPPORTED" if overall else
                "RECONCILIATION NOT SUPPORTED - a reconciliation that foots can still be "
                "worthless. See the failing test(s).")
    v.font = OK_FONT if overall else BAD_FONT
    if not overall:
        v.fill = BAD_FILL
    r += 2

    line("INDEPENDENT RECOMPUTATION", bold=True, size=12, fill=SUB_FILL)
    line("  Balance per bank statement", rc["bank_balance"])
    line("  Reconciling items - bank side", rc["bank_adjustments"])
    line("  = Adjusted bank balance", rc["adjusted_bank"], bold=True, border=TOP)
    line("  Balance per general ledger", rc["gl_balance"])
    line("  Reconciling items - book side", rc["book_adjustments"])
    line("  = Adjusted book balance", rc["adjusted_book"], bold=True, border=TOP)
    line("  DIFFERENCE", rc["difference"], bold=True, border=DBL,
         fill=None if rc["difference"] == ZERO else BAD_FILL)
    c = ws.cell(row=r - 1, column=3)
    c.font = OK_FONT if rc["difference"] == ZERO else BAD_FONT
    r += 1
    line("  Recomputed from item detail, not from the preparer's subtotals.", bold=True)
    r += 1

    for t in tests:
        line(f"TEST {t['num']} - {t['name']}", None,
             "PASS" if t["passed"] else "FAIL", note=t.get("note", ""))
    r += 1

    if controls:
        line("CONTROL OBSERVATIONS", bold=True, size=12, fill=SUB_FILL)
        for c_ in controls:
            line("  " + c_)
        r += 1

    if esc:
        line("ESCALATE", bold=True, size=12, fill=BAD_FILL)
        for e in esc:
            line("  " + e)
        r += 1
    r += 1
    line("Prepared by / date:  __________________  ____________", bold=True)
    line("Reviewed by / date:  __________________  ____________", bold=True)


def sheet_recompute(wb, rc, presented):
    ws = wb.create_sheet("Independent Recomputation")
    heads = ["Line", "Independently recomputed", "As presented by the preparer",
             "Difference"]
    ws.append(heads)
    hdr(ws, len(heads))
    mapping = [("Balance per bank statement", rc["bank_balance"], "bank balance"),
               ("Reconciling items - bank side", rc["bank_adjustments"], "bank adjustments"),
               ("Adjusted bank balance", rc["adjusted_bank"], "adjusted bank"),
               ("Balance per general ledger", rc["gl_balance"], "gl balance"),
               ("Reconciling items - book side", rc["book_adjustments"], "book adjustments"),
               ("Adjusted book balance", rc["adjusted_book"], "adjusted book")]
    for label, val, pk in mapping:
        pv = presented.get(pk)
        ws.append([label, float(val), float(pv) if pv is not None else None,
                   float(val - pv) if pv is not None else None])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in (1, 2, 3):
            row[i].number_format = MONEY
        if isinstance(row[3].value, (int, float)) and row[3].value:
            row[3].font, row[3].fill = BAD_FONT, BAD_FILL
    r = ws.max_row + 2
    if not presented:
        ws.cell(row=r, column=1,
                value="The preparer's own subtotals were not supplied, so only the "
                      "recomputation is shown. Where they are available, the comparison "
                      "detects a reconciliation whose subtotals do not follow from its "
                      "own item detail.").font = Font(italic=True)
    widths(ws, {1: 36, 2: 24, 3: 28, 4: 16})


def sheet_items(wb, items):
    ws = wb.create_sheet("Item Testing")
    heads = ["Item", "Category", "Description", "Reference", "Item date", "Amount",
             "Support obtained", "Cleared", "Cleared date", "Cleared amount",
             "Days to clear", "Periods carried", "Flags"]
    ws.append(heads)
    hdr(ws, len(heads))
    def rank(x):
        unsupported = 0 if not x["support"] else 1
        uncleared = 0 if x["cleared"] is False else 1
        return (unsupported, uncleared, -x["amount"])
    for x in sorted(items, key=rank):
        ws.append([x["item_id"], x["category"], x["description"] or None,
                   x["reference"] or None, x["item_date"], float(x["amount"]),
                   x["support"] or None,
                   "" if x["cleared"] is None else ("yes" if x["cleared"] else "NO"),
                   x["cleared_date"],
                   float(x["cleared_amount"]) if x["cleared_amount"] is not None else None,
                   x["days_to_clear"], x["periods_carried"] or None,
                   "; ".join(x["flags"])])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[4].number_format = DATEF
        row[8].number_format = DATEF
        row[5].number_format = MONEY
        row[9].number_format = MONEY
        if not row[6].value:
            row[6].fill = BAD_FILL
        if row[7].value == "NO":
            row[7].font, row[7].fill = BAD_FONT, BAD_FILL
        if row[12].value:
            row[12].fill = WARN_FILL
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:M{max(ws.max_row, 2)}"
    widths(ws, {1: 9, 2: 22, 3: 36, 4: 16, 5: 12, 6: 15, 7: 26, 8: 9, 9: 12,
                10: 15, 11: 13, 12: 15, 13: 78})


def sheet_list(wb, title, header, rows, fmt, empty, footer=""):
    ws = wb.create_sheet(title)
    ws.append(header)
    hdr(ws, len(header))
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i, f in fmt.items():
            if i < len(row):
                row[i].number_format = f
        if row[len(header) - 1].value:
            row[len(header) - 1].fill = WARN_FILL
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value=empty).font = OK_FONT
    elif footer:
        rr = ws.max_row + 2
        c = ws.cell(row=rr, column=1, value=footer)
        c.font = Font(italic=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    widths(ws, {i: 18 for i in range(1, len(header) + 1)} | {2: 38,
                                                            len(header): 76})


# ------------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--reconciliation", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--bank-balance", required=True)
    ap.add_argument("--gl-balance", required=True)
    ap.add_argument("--subsequent-statement")
    ap.add_argument("--prior-reconciliation")
    ap.add_argument("--presented")
    ap.add_argument("--period-end", required=True)
    ap.add_argument("--prepared-by", default="")
    ap.add_argument("--prepared-date")
    ap.add_argument("--reviewed-by", default="")
    ap.add_argument("--preparer-has-disbursement-authority", action="store_true")
    ap.add_argument("--client", default="")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    period_end = pdate(args.period_end)
    items = load_rec(Path(args.reconciliation))
    subseq = load_statement(Path(args.subsequent_statement)
                            if args.subsequent_statement else None)
    prior = (rows_of(Path(args.prior_reconciliation), "prior reconciliation")
             if args.prior_reconciliation and Path(args.prior_reconciliation).exists()
             else [])
    presented = {}
    if args.presented and Path(args.presented).exists():
        for r in rows_of(Path(args.presented), "presented"):
            presented[clean(r.get("line")).lower()] = d0(r.get("amount"))

    bank_balance = dec(args.bank_balance)
    gl_balance = dec(args.gl_balance)

    test_clearance(items, subseq, period_end)
    test_support_and_plugs(items, subseq, prior, period_end)
    rc = recompute(items, bank_balance, gl_balance)

    unsupported = [x for x in items if not x["support"]]
    uncleared = [x for x in items if x["cleared"] is False]
    stale = [x for x in items if x["periods_carried"] >= 3]
    plugs = [x for x in items
             if any(f.startswith(("GENERIC DESCRIPTION", "NO COUNTERPART", "round amount"))
                    for f in x["flags"])]
    fabricated = [x for x in items
                  if any(f.startswith("NO COUNTERPART") for f in x["flags"])]

    prepared_date = pdate(args.prepared_date) if args.prepared_date else None
    lag = (prepared_date - period_end).days if prepared_date else None

    controls = []
    if not args.reviewed_by:
        controls.append("The reconciliation carries no evidence of review. An unreviewed "
                        "reconciliation is a control finding independent of its arithmetic.")
    if lag is not None and lag > 30:
        controls.append(f"Prepared {lag} days after period end. The longer the delay, the "
                        f"more opportunity to construct rather than reconcile.")
    if args.preparer_has_disbursement_authority:
        controls.append(f"The preparer ({args.prepared_by or 'unnamed'}) also holds "
                        f"disbursement authority - a segregation of duties failure.")
    if not subseq:
        controls.append("No subsequent-period bank statement was obtained. Clearance could "
                        "not be tested, which is the procedure that detects a fabricated "
                        "item. The review is limited in scope and should say so.")

    tests = [
        {"num": 1, "name": "Independent recomputation agrees",
         "passed": rc["difference"] == ZERO,
         "note": "" if rc["difference"] == ZERO else
                 f"adjusted bank and adjusted book differ by {rc['difference']:,.2f} when "
                 f"recomputed from item detail"},
        {"num": 2, "name": "Bank balance agreed to the bank statement",
         "passed": True,
         "note": f"agreed to {bank_balance:,.2f} per the statement as an independent source"},
        {"num": 3, "name": "Book balance agreed to the general ledger",
         "passed": True,
         "note": f"agreed to {gl_balance:,.2f} per the GL as an independent source"},
        {"num": 4, "name": "Every reconciling item supported",
         "passed": not unsupported,
         "note": "" if not unsupported else
                 f"{len(unsupported)} item(s) totalling "
                 f"{sum((x['amount'] for x in unsupported), ZERO):,.2f} traced to no "
                 f"document"},
        {"num": 5, "name": "Subsequent clearance tested",
         "passed": bool(subseq) and not uncleared,
         "note": ("no subsequent statement obtained - this is the test that detects a "
                  "fabricated item" if not subseq else
                  "" if not uncleared else
                  f"{len(uncleared)} item(s) totalling "
                  f"{sum((x['amount'] for x in uncleared), ZERO):,.2f} did not clear")},
        {"num": 6, "name": "No stale items rolled forward",
         "passed": not stale,
         "note": "" if not stale else
                 f"{len(stale)} item(s) carried on three or more consecutive "
                 f"reconciliations - not timing"},
        {"num": 7, "name": "No plug indicators",
         "passed": not plugs,
         "note": "" if not plugs else
                 f"{len(plugs)} item(s) with a generic description, a round amount, or no "
                 f"counterpart anywhere"},
        {"num": 8, "name": "Preparation and review controls",
         "passed": not controls,
         "note": "" if not controls else f"{len(controls)} control observation(s)"},
    ]

    esc = []
    for x in fabricated:
        esc.append(f"{x['item_id']} ({x['description'] or 'no description'}, "
                   f"{x['amount']:,.2f}): no supporting document and no counterpart in the "
                   f"subsequent statement. This is the fabricated-item pattern - escalate "
                   f"regardless of amount.")
    for x in items:
        if x["category"] == "deposit_in_transit" and x["cleared"] is False:
            esc.append(f"{x['item_id']}: deposit in transit of {x['amount']:,.2f} never "
                       f"cleared. A deposit in transit is by definition already made; one "
                       f"that never arrives is the classic lapping pattern.")
    for x in stale:
        esc.append(f"{x['item_id']} ({x['description']}, {x['amount']:,.2f}): carried on "
                   f"{x['periods_carried']} consecutive reconciliations.")
    if rc["difference"] != ZERO:
        esc.append(f"The reconciliation does not reconcile when recomputed from its own "
                   f"item detail: {rc['difference']:,.2f}. Either an item is missing from "
                   f"the detail supplied, or the presented subtotals do not follow from it.")
    generic_big = [x for x in items
                   if any(f.startswith("GENERIC DESCRIPTION") for f in x["flags"])]
    for x in generic_big:
        esc.append(f"{x['item_id']}: {x['amount']:,.2f} labelled "
                   f"\"{x['description']}\". Report the label and the amount together - "
                   f"the juxtaposition is the finding.")
    if prior:
        prior_amts = {(key(p.get("reference")), d0(p.get("amount"))) for p in prior}
        current = {(key(x["reference"]), x["amount"]) for x in items}
        gone = prior_amts - current
        cleared_refs = {key(s["reference"]) for s in subseq if s["reference"]}
        vanished = [g for g in gone if g[0] and g[0] not in cleared_refs]
        if vanished:
            esc.append(f"{len(vanished)} prior-period reconciling item(s) are no longer on "
                       f"the reconciliation and did not appear in the subsequent statement. "
                       f"Something was removed rather than resolved, and the reconciliation "
                       f"still foots - which means something else absorbed it.")

    overall = all(t["passed"] for t in tests)

    meta = {
        "Client": args.client or "(not stated)",
        "Period end": str(period_end),
        "Reconciling items": len(items),
        "Prepared by": args.prepared_by or "(not stated)",
        "Prepared date": str(prepared_date) if prepared_date else "(not stated)",
        "Days after period end": lag if lag is not None else "(unknown)",
        "Reviewed by": args.reviewed_by or "*** NO EVIDENCE OF REVIEW ***",
        "Subsequent statement obtained": "yes" if subseq else "NO",
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    # ---- console
    print("=" * 76)
    print("BANK RECONCILIATION REVIEW")
    print("=" * 76)
    print(f"Client : {meta['Client']}   Period end {period_end}")
    print(f"Items  : {len(items)}   prepared by {meta['Prepared by']}"
          f"   reviewed by {meta['Reviewed by']}")
    print()
    print("INDEPENDENT RECOMPUTATION (from item detail, not the preparer's subtotals)")
    print(f"  Balance per bank statement    {rc['bank_balance']:>16,.2f}")
    print(f"  Reconciling items - bank      {rc['bank_adjustments']:>16,.2f}")
    print(f"  = Adjusted bank               {rc['adjusted_bank']:>16,.2f}")
    print(f"  Balance per general ledger    {rc['gl_balance']:>16,.2f}")
    print(f"  Reconciling items - book      {rc['book_adjustments']:>16,.2f}")
    print(f"  = Adjusted book               {rc['adjusted_book']:>16,.2f}")
    print(f"  DIFFERENCE                    {rc['difference']:>16,.2f}   "
          f"{'OK' if rc['difference'] == ZERO else '*** MUST BE ZERO ***'}")
    print()
    for t in tests:
        print(f"Test {t['num']}: {'PASS' if t['passed'] else '*** FAIL ***':<14} {t['name']}")
        if t.get("note"):
            print(f"         {t['note']}")
    if controls:
        print("\nCONTROL OBSERVATIONS:")
        for c in controls:
            print(f"  - {c}")
    if esc:
        print("\nESCALATE:")
        for e in esc:
            print(f"  ! {e}")

    if not overall and not args.force:
        print("\n" + "=" * 76)
        print("WORKBOOK NOT WRITTEN.")
        print("A reconciliation that foots can still be worthless. Obtain support for")
        print("every item and test subsequent clearance before concluding on cash.")
        print("=" * 76)
        return 1

    wb = Workbook()
    sheet_summary(wb, meta, tests, rc, esc, overall, controls)
    sheet_recompute(wb, rc, presented)
    sheet_items(wb, items)
    sheet_list(wb, "Uncleared Items",
               ["Item", "Description", "Category", "Amount", "Item date",
                "Age (days)", "Follow-up required"],
               [[x["item_id"], x["description"], x["category"], float(x["amount"]),
                 x["item_date"],
                 (period_end - x["item_date"]).days if x["item_date"] else None,
                 "; ".join(x["flags"])] for x in uncleared],
               {3: MONEY, 4: DATEF},
               "None - every reconciling item cleared in the subsequent period.",
               footer="Absence from the subsequent statement is the finding. A real "
                      "reconciling item appears there; a fabricated one never does.")
    sheet_list(wb, "Stale Items",
               ["Item", "Description", "Amount", "Periods carried", "Note"],
               [[x["item_id"], x["description"], float(x["amount"]),
                 x["periods_carried"], "; ".join(x["flags"])]
                for x in items if x["periods_carried"]],
               {2: MONEY},
               "None carried from the prior reconciliation.",
               footer="An item carried on three or more consecutive reconciliations is "
                      "not timing.")
    sheet_list(wb, "Plug Indicators",
               ["Item", "Description", "Amount", "Support", "Indicator"],
               [[x["item_id"], x["description"], float(x["amount"]),
                 x["support"] or "(none)", "; ".join(x["flags"])] for x in plugs],
               {2: MONEY},
               "None - no generic descriptions, round balancing amounts, or items "
               "without a counterpart.",
               footer="A plug survives review when only the arithmetic is checked. "
                      "Subsequent clearance is what catches it.")
    sheet_list(wb, "Control Observations", ["#", "Observation"],
               [[i + 1, c] for i, c in enumerate(controls)], {},
               "None - the reconciliation was prepared timely, reviewed, and by someone "
               "without disbursement authority.")
    wb.active = 0
    out = Path(args.out)
    if not overall:
        out = out.with_name(out.stem + " [NOT SUPPORTED]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\n{'SUPPORTED' if overall else 'NOT SUPPORTED (forced)'}.  Workbook: {out}")
    return 0 if overall else 1


if __name__ == "__main__":
    sys.exit(main())
```

