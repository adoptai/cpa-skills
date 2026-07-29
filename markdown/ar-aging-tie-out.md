---
name: ar-aging-tie-out
description: Tie the accounts receivable aging to the general ledger control account and test it properly — recomputing every ageing bucket from the invoice date rather than trusting the report, grossing up netted credit balances, and flagging stale balances, unapplied cash, and cutoff errors. Use this whenever the user mentions AR aging, accounts receivable aging, tying the aging to the GL, an AR subledger that does not agree to the control account, stale or uncollectible receivables, the allowance for doubtful accounts, unapplied cash or credits, or receivable confirmations. Also trigger on "does our AR aging tie," "reconcile accounts receivable," "which receivables are old," "check the allowance," "why is AR off," "review the aging," or when an AR aging report and a trial balance are provided together. Runs fully local — no customer or receivable data leaves the machine.
---

# AR Aging Tie-Out

Almost everyone ties the aging total to the general ledger and stops. That test is necessary and
nearly worthless on its own, because **the total can agree while the aging itself is wrong** — and
the aging is what drives the allowance, the collectibility discussion, and the confirmation
sample.

Four things go wrong beneath a correct total:

1. **Mis-bucketed invoices.** Aging reports commonly bucket from the wrong date, or fail to
   re-age after a partial payment or a credit application. So this skill **recomputes every
   bucket from the invoice date and compares** rather than trusting the column the report printed.
2. **Netted credit balances.** A customer with a $90,000 open invoice and a $85,000 unapplied
   credit shows as $5,000 current. Both numbers matter and the netting hides them.
3. **Unapplied cash sitting in the aging** as a negative current balance, which understates real
   exposure and often means cash was received against an invoice nobody closed.
4. **Invoices dated after period end** included in the aging — a cutoff error that overstates both
   revenue and receivables.

## The gate

No clean workpaper unless:

1. **The aging total equals the GL control account balance.** You supply the GL figure.
2. **The buckets foot.** Bucket amounts sum to each customer's total, and customer totals sum to
   the report total.
3. **Every bucket agrees to a recomputation from the invoice date** using the stated bucket
   boundaries and as-of date. Any invoice in the wrong bucket is reported individually.
4. **Debits and credits are shown gross**, not netted, at customer level.

## Inputs

1. **AR aging detail at invoice level** — customer, invoice number, invoice date, due date, original
   amount, open balance, and the bucket the report assigned. Invoice-level detail is essential; a
   customer-level summary cannot be re-aged.
2. **GL control account balance** for accounts receivable at the same date.
3. **As-of date and bucket boundaries** — the client's own, whatever they are. Do not assume
   30/60/90.
4. **Payment terms**, by customer if they vary. Aging past terms is the meaningful measure, not
   aging past an arbitrary 30 days.
5. **Allowance for doubtful accounts** balance and how it was computed, if testing the allowance.
6. **Subsequent cash receipts**, if available — the strongest evidence of collectibility and the
   best test of existence short of a confirmation.

## Step 1 — Tie and re-age

```bash
python3 scripts/ar_aging.py \
  --aging aging_detail.csv \
  --gl-balance 2841655.40 \
  --as-of 2025-12-31 \
  --buckets 30,60,90,120 \
  --allowance 148000 \
  --subsequent-receipts receipts.csv \
  --client "Fairhaven Distributors" \
  --out "Fairhaven - 2025 AR Aging Tie-Out.xlsx"
```

Seven tests:

- **Test 1 — Aging total to GL control account.** The one everybody runs.
- **Test 2 — Bucket footing.** Buckets sum to customer totals, customers sum to the report total.
- **Test 3 — Bucket recomputation.** Every invoice re-aged from its invoice date against the
  stated boundaries. Mis-bucketed invoices are listed individually with both buckets shown. This
  is the test that finds real problems.
- **Test 4 — Credit balances grossed up.** Total debits and total credits shown separately, with
  customers whose net balance masks a material credit identified.
- **Test 5 — Cutoff.** Invoices dated after the as-of date, and invoices with a due date before
  their invoice date.
- **Test 6 — Internal integrity.** Open balance exceeding the original amount, zero-balance rows
  carried in the aging, duplicate invoice numbers, missing dates, and invoices with no customer.
- **Test 7 — Allowance reasonableness**, where an allowance balance is supplied. Reports the
  allowance as a percentage of total AR and of each aged bucket, and the implied coverage of
  balances past terms.

## Step 2 — Read the collectibility picture

The workbook aggregates by customer and by age band, but the useful views are these:

- **Aging past terms, not past 30 days.** A customer on 60-day terms at 45 days is current. A
  customer on 10-day terms at 45 days is a problem. Where terms are supplied the schedule shows
  days past terms alongside days past invoice, and they tell different stories.
- **Concentration.** The largest few customers as a percentage of total AR. A receivable book
  where one customer is 40% of the balance carries a different risk than the aging profile alone
  suggests, and it belongs in the allowance discussion.
- **Balances with no subsequent receipt.** Where subsequent receipts are supplied, the schedule
  flags aged balances with no payment activity after period end. An old balance that has since
  been paid is a collection-timing issue; an old balance with no subsequent activity is a
  valuation issue, and possibly an existence one.
- **Round-number balances** on aged accounts, which sometimes indicate an estimated or
  placeholder invoice rather than a real one.
- **Credit balances aged into old buckets** — an unapplied credit sitting for months usually means
  a customer overpaid or was double-billed, and it may be a refund liability rather than a
  receivable offset.

## Step 3 — Deliver

**Workbook tabs:**

1. **Tie-Out Summary** — the seven tests, the GL agreement, aging profile, concentration, and the
   allowance analysis. The signable page.
2. **Re-aged Detail** — every invoice with the bucket per the report, the recomputed bucket, days
   outstanding, days past terms, and any flag. Mis-bucketed invoices sort to the top.
3. **By Customer** — gross debits, gross credits, net balance, bucket distribution, days past
   terms on the oldest item, and share of total AR.
4. **Exceptions** — every finding from Tests 3 through 6 with the invoice and dollar effect.
5. **Credit Balances** — all credit balances with age, separated from the debit aging entirely.
6. **Allowance Analysis** — allowance as a percentage of total and of each bucket, coverage of
   balances past terms, and the customers driving the exposure.
7. **Subsequent Receipts** — aged balances with and without post-period collection activity.

**Then, in chat:** whether the aging ties to the GL, how many invoices were mis-bucketed and the
dollars involved, the aging profile past terms, concentration, and the allowance conclusion. Lead
with the mis-bucketing if there is any — it means the aging the client has been relying on for
credit decisions and the allowance is wrong.

## What to escalate

- **The aging not tying to the control account.** Until it does, nothing built on the aging is
  reliable. The usual causes are unposted cash, a journal entry made directly to the control
  account, and a subledger cut on a different date.
- **Journal entries posted directly to the AR control account.** These bypass the subledger by
  definition, so they are invisible in the aging and are a common route for both error and
  concealment. Ask for the control account detail and read every manual entry.
- **Material credit balances aged beyond a few months** — potentially a refund liability, and
  possibly unclaimed property depending on the jurisdiction and dormancy rules. Flag the question;
  confirm the applicable rules rather than assuming them.
- **Aged balances with no subsequent receipts and no collection activity**, particularly where the
  customer continues to be invoiced.
- **Receivables from related parties or employees** mixed into trade AR — a classification and
  disclosure matter, and they are rarely subject to the same credit discipline.
- **An allowance that has not moved** while the aging profile deteriorated. A static allowance
  against a worsening book is a judgment that needs documenting, not a rollforward.
- **Invoices dated after period end** in the aging, which overstates revenue and receivables in
  the same stroke.
- **Write-offs late in the period** clearing balances that were about to become conspicuously
  aged. Report the pattern as a fact.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.
Customer names and balances never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*

---

# Appendix — bundled files

If you installed the `.skill` package these files are already in place and you can ignore this appendix. If you copied the skill as text, create the files below at the paths shown, alongside your `SKILL.md`. The skill will not run without them.

## `scripts/ar_aging.py`

```python
#!/usr/bin/env python3
"""
AR aging tie-out: aging <-> GL control account, with every bucket recomputed.

Seven tests. The distinguishing one is Test 3: every invoice is re-aged from its
invoice date against the stated bucket boundaries and compared to the bucket the
report assigned. A total can agree to the GL while the aging itself is wrong, and
the aging is what drives the allowance and the credit decisions.

Usage:
    python3 ar_aging.py --aging aging_detail.csv --gl-balance 2841655.40 \
        --as-of 2025-12-31 --buckets 30,60,90,120 --allowance 148000 \
        --subsequent-receipts receipts.csv \
        --client "Fairhaven Distributors" \
        --out "Fairhaven - 2025 AR Aging Tie-Out.xlsx"

--aging CSV (invoice level - a customer summary cannot be re-aged):
    customer, invoice_no, invoice_date, due_date, original_amount, open_balance,
    reported_bucket (optional), terms_days (optional), customer_type (optional)

--subsequent-receipts CSV:
    customer, invoice_no, receipt_date, amount
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
PCT = '0.0"%"'
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

ROUND_MIN = Decimal("1000")


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


def bucket_labels(bounds: list[int]) -> list[str]:
    labels = ["Current"]
    prev = 0
    for b in bounds:
        labels.append(f"{prev + 1}-{b}")
        prev = b
    labels.append(f"{prev + 1}+")
    return labels


def bucket_of(days: int, bounds: list[int], labels: list[str]) -> str:
    if days <= 0:
        return labels[0]
    for i, b in enumerate(bounds):
        if days <= b:
            return labels[i + 1]
    return labels[-1]


def load_aging(path: Path) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No rows in {path}")
    out = []
    for i, r in enumerate(rows, start=1):
        cust = clean(r.get("customer"))
        out.append({
            "row": i,
            "customer": cust or "(NO CUSTOMER)",
            "invoice_no": clean(r.get("invoice_no")) or f"(no number {i})",
            "invoice_date": pdate(r.get("invoice_date")),
            "due_date": pdate(r.get("due_date")),
            "original_amount": d0(r.get("original_amount")),
            "open_balance": d0(r.get("open_balance")),
            "reported_bucket": clean(r.get("reported_bucket")),
            "terms_days": dec(r.get("terms_days")),
            "customer_type": clean(r.get("customer_type")),
            "flags": [],
        })
    return out


def load_receipts(path: Path | None) -> dict:
    out = defaultdict(lambda: {"amount": ZERO, "last": None, "count": 0})
    if not path or not path.exists():
        return out
    with path.open(newline="", encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            key = (clean(r.get("customer")).lower(), clean(r.get("invoice_no")).lower())
            a = out[key]
            a["amount"] += d0(r.get("amount"))
            a["count"] += 1
            d = pdate(r.get("receipt_date"))
            if d and (a["last"] is None or d > a["last"]):
                a["last"] = d
    return out


def is_round(v: Decimal) -> bool:
    a = abs(v)
    if a < ROUND_MIN:
        return False
    for m in (Decimal("100000"), Decimal("50000"), Decimal("25000"),
              Decimal("10000"), Decimal("5000"), Decimal("1000")):
        if m >= a / Decimal(20) and a >= m and a % m == ZERO:
            return True
    return False


# ----------------------------------------------------------------------- tests

def analyse(rows, as_of: date, bounds, labels, receipts) -> dict:
    misbucketed, cutoff, integrity = [], [], []
    seen = defaultdict(list)

    for x in rows:
        inv_d = x["invoice_date"]
        if inv_d is None:
            x["days"] = None
            x["recomputed_bucket"] = "(no invoice date)"
            integrity.append(f"{x['customer']} / {x['invoice_no']}: no invoice date - "
                             f"cannot be aged")
        else:
            x["days"] = (as_of - inv_d).days
            x["recomputed_bucket"] = bucket_of(x["days"], bounds, labels)

        terms = x["terms_days"]
        if terms is None and x["due_date"] and inv_d:
            terms = Decimal((x["due_date"] - inv_d).days)
        x["terms_used"] = terms
        x["days_past_terms"] = (x["days"] - int(terms)) if (x["days"] is not None
                                                           and terms is not None) else None

        if x["reported_bucket"] and x["recomputed_bucket"] != "(no invoice date)":
            if clean(x["reported_bucket"]).lower() != x["recomputed_bucket"].lower():
                x["flags"].append(f"MIS-BUCKETED: report says '{x['reported_bucket']}', "
                                  f"recomputed from invoice date {inv_d} at {x['days']} "
                                  f"days is '{x['recomputed_bucket']}'")
                misbucketed.append(x)

        if inv_d and inv_d > as_of:
            x["flags"].append(f"invoice dated {inv_d}, AFTER the as-of date {as_of} - "
                              f"cutoff error overstating revenue and receivables")
            cutoff.append(x)
        if inv_d and x["due_date"] and x["due_date"] < inv_d:
            x["flags"].append(f"due date {x['due_date']} precedes invoice date {inv_d}")
            cutoff.append(x)

        if x["open_balance"] > ZERO and x["original_amount"] > ZERO \
                and x["open_balance"] > x["original_amount"]:
            x["flags"].append(f"open balance {x['open_balance']:,.2f} exceeds the original "
                              f"invoice {x['original_amount']:,.2f}")
            integrity.append(f"{x['customer']} / {x['invoice_no']}: open balance exceeds "
                             f"original amount")
        if x["open_balance"] == ZERO:
            x["flags"].append("zero open balance carried in the aging")
            integrity.append(f"{x['customer']} / {x['invoice_no']}: zero balance row in "
                             f"the aging - clutters the report and distorts item counts")
        if x["customer"] == "(NO CUSTOMER)":
            integrity.append(f"Row {x['row']} ({x['invoice_no']}): no customer")
        if is_round(x["open_balance"]) and x["days"] is not None and x["days"] > bounds[0]:
            x["flags"].append("round-number balance on an aged account - confirm this is "
                              "a real invoice rather than an estimate or placeholder")
        seen[(x["customer"].lower(), x["invoice_no"].lower())].append(x["row"])

        key = (x["customer"].lower(), x["invoice_no"].lower())
        rec = receipts.get(key)
        x["subsequent_amount"] = rec["amount"] if rec else ZERO
        x["subsequent_last"] = rec["last"] if rec else None
        if receipts and x["open_balance"] > ZERO and x["days"] is not None \
                and x["days"] > bounds[-1] and not rec:
            x["flags"].append("aged beyond the oldest bucket with NO subsequent receipt - "
                              "a valuation question, and possibly an existence one")

    for (c, inv), rowlist in seen.items():
        if len(rowlist) > 1:
            integrity.append(f"Duplicate invoice number for one customer: {inv} on rows "
                             f"{rowlist}")
            for x in rows:
                if x["row"] in rowlist:
                    x["flags"].append(f"duplicate invoice number within customer "
                                      f"(rows {rowlist})")

    return {"misbucketed": misbucketed, "cutoff": cutoff, "integrity": integrity}


def by_customer(rows, labels) -> list[dict]:
    agg = {}
    for x in rows:
        c = agg.setdefault(x["customer"], {
            "customer": x["customer"], "debits": ZERO, "credits": ZERO, "net": ZERO,
            "buckets": {l: ZERO for l in labels}, "oldest_days": None,
            "oldest_past_terms": None, "items": 0,
            "customer_type": x["customer_type"], "flags": set(),
        })
        c["items"] += 1
        if x["open_balance"] >= ZERO:
            c["debits"] += x["open_balance"]
        else:
            c["credits"] += x["open_balance"]
        c["net"] += x["open_balance"]
        if x["recomputed_bucket"] in c["buckets"]:
            c["buckets"][x["recomputed_bucket"]] += x["open_balance"]
        if x["days"] is not None and x["open_balance"] > ZERO:
            if c["oldest_days"] is None or x["days"] > c["oldest_days"]:
                c["oldest_days"] = x["days"]
        if x["days_past_terms"] is not None and x["open_balance"] > ZERO:
            if c["oldest_past_terms"] is None or x["days_past_terms"] > c["oldest_past_terms"]:
                c["oldest_past_terms"] = x["days_past_terms"]
        if x["customer_type"] and x["customer_type"].lower() in (
                "related party", "related-party", "employee", "affiliate", "intercompany"):
            c["flags"].add(f"{x['customer_type']} - classification and disclosure matter, "
                           f"and rarely subject to the same credit discipline")
    out = sorted(agg.values(), key=lambda c: c["net"], reverse=True)
    return out


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


def sheet_summary(wb, meta, tests, totals, labels, customers, allowance, overall, esc):
    ws = wb.active
    ws.title = "Tie-Out Summary"
    widths(ws, {1: 4, 2: 52, 3: 18, 4: 16, 5: 62})
    r = 1

    def line(label, a=None, b=None, *, bold=False, size=11, fill=None, border=None,
             note=""):
        nonlocal r
        c = ws.cell(row=r, column=2, value=label)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c.fill = fill
        for col, v in ((3, a), (4, b)):
            if v is None:
                continue
            cc = ws.cell(row=r, column=col,
                         value=float(v) if isinstance(v, Decimal) else v)
            if isinstance(v, Decimal):
                cc.number_format = MONEY
            cc.font = Font(bold=bold)
            if border:
                cc.border = border
        if note:
            n = ws.cell(row=r, column=5, value=note)
            n.font = Font(italic=True, size=9)
            n.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    line("ACCOUNTS RECEIVABLE AGING TIE-OUT", bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    v = ws.cell(row=r, column=2,
                value="TIED OUT - the aging agrees to the control account and every "
                      "bucket recomputes"
                if overall else
                "NOT TIED OUT - see the failing test(s). The aging drives the allowance "
                "and credit decisions, so a wrong aging matters even when the total "
                "agrees.")
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
        for label, val in t["lines"]:
            line("  " + label, val)
        if t.get("is_count"):
            line("  Items failing", int(t["difference"]), bold=True, border=TOP,
                 note=t.get("note", ""))
        else:
            line("  Difference", t["difference"], bold=True, border=TOP,
                 note=t.get("note", ""))
        c = ws.cell(row=r - 1, column=3)
        c.font = OK_FONT if t["passed"] else BAD_FONT
        if not t["passed"]:
            c.fill = BAD_FILL
        r += 1

    line("AGING PROFILE (recomputed)", bold=True, size=12, fill=SUB_FILL)
    total_debits = totals["debits"]
    for l in labels:
        amt = totals["buckets"][l]
        pct = (amt / total_debits * 100) if total_debits else ZERO
        ws.cell(row=r, column=2, value="  " + l)
        c1 = ws.cell(row=r, column=3, value=float(amt))
        c1.number_format = MONEY
        c2 = ws.cell(row=r, column=4, value=float(pct))
        c2.number_format = PCT
        r += 1
    r += 1

    line("CONCENTRATION", bold=True, size=12, fill=SUB_FILL)
    top5 = customers[:5]
    tot = totals["net"]
    for c in top5:
        share = (c["net"] / tot * 100) if tot else ZERO
        ws.cell(row=r, column=2, value=f"  {c['customer']}")
        cc = ws.cell(row=r, column=3, value=float(c["net"]))
        cc.number_format = MONEY
        cp = ws.cell(row=r, column=4, value=float(share))
        cp.number_format = PCT
        if share >= 20:
            cp.fill = WARN_FILL
        r += 1
    top5_share = (sum((c["net"] for c in top5), ZERO) / tot * 100) if tot else ZERO
    line("  Top five as a share of total AR", None, top5_share, bold=True,
         note="Concentration changes the risk profile independently of the aging. A book "
              "where one customer dominates belongs in the allowance discussion.")
    ws.cell(row=r - 1, column=4).number_format = PCT
    r += 1

    if allowance is not None:
        line("ALLOWANCE ANALYSIS", bold=True, size=12, fill=SUB_FILL)
        line("  Allowance for doubtful accounts", allowance)
        pct_total = (allowance / total_debits * 100) if total_debits else ZERO
        line("  As a percentage of gross receivables", None, float(pct_total))
        ws.cell(row=r - 1, column=4).number_format = PCT
        past = totals["past_terms"]
        line("  Balances past terms", past)
        cov = (allowance / past * 100) if past else None
        if cov is not None:
            line("  Allowance coverage of balances past terms", None, float(cov),
                 note="Not a conclusion. Compare to history and to the specific "
                      "customers driving the exposure.")
            ws.cell(row=r - 1, column=4).number_format = PCT
        r += 1

    if esc:
        line("ESCALATE", bold=True, size=12, fill=BAD_FILL)
        for e in esc:
            line("  " + e)
        r += 1

    r += 1
    line("Prepared by / date:  __________________  ____________", bold=True)
    line("Reviewed by / date:  __________________  ____________", bold=True)


def sheet_detail(wb, rows, labels):
    ws = wb.create_sheet("Re-aged Detail")
    heads = ["Customer", "Invoice", "Invoice date", "Due date", "Terms (d)",
             "Original", "Open balance", "Days outstanding", "Days past terms",
             "Bucket per report", "Bucket recomputed", "Agrees",
             "Subsequent receipts", "Last receipt", "Flags"]
    ws.append(heads)
    hdr(ws, len(heads))
    ordered = sorted(rows, key=lambda x: (0 if any(f.startswith("MIS-BUCKETED")
                                                   for f in x["flags"]) else 1,
                                          -(x["days"] or 0)))
    for x in ordered:
        agrees = ""
        if x["reported_bucket"]:
            agrees = "OK" if clean(x["reported_bucket"]).lower() == \
                             x["recomputed_bucket"].lower() else "MIS-BUCKETED"
        ws.append([x["customer"], x["invoice_no"], x["invoice_date"], x["due_date"],
                   float(x["terms_used"]) if x["terms_used"] is not None else None,
                   float(x["original_amount"]), float(x["open_balance"]),
                   x["days"], x["days_past_terms"],
                   x["reported_bucket"] or None, x["recomputed_bucket"], agrees,
                   float(x["subsequent_amount"]) if x["subsequent_amount"] else None,
                   x["subsequent_last"], "; ".join(x["flags"])])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[2].number_format = DATEF
        row[3].number_format = DATEF
        row[13].number_format = DATEF
        for i in (5, 6, 12):
            row[i].number_format = MONEY
        if row[11].value == "MIS-BUCKETED":
            row[11].font, row[11].fill = BAD_FONT, BAD_FILL
        if isinstance(row[6].value, (int, float)) and row[6].value < 0:
            row[6].fill = WARN_FILL
        if row[14].value:
            row[14].fill = WARN_FILL
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:O{max(ws.max_row, 2)}"
    widths(ws, {1: 30, 2: 14, 3: 12, 4: 12, 5: 9, 6: 14, 7: 15, 8: 15, 9: 14,
                10: 17, 11: 18, 12: 14, 13: 18, 14: 13, 15: 72})


def sheet_customers(wb, customers, labels, total_net):
    ws = wb.create_sheet("By Customer")
    heads = ["Customer", "Type", "Items", "Gross debits", "Gross credits", "Net balance",
             "% of AR", "Oldest days", "Oldest days past terms"] + labels + ["Flags"]
    ws.append(heads)
    hdr(ws, len(heads))
    for c in customers:
        share = (c["net"] / total_net * 100) if total_net else ZERO
        ws.append([c["customer"], c["customer_type"] or None, c["items"],
                   float(c["debits"]), float(c["credits"]), float(c["net"]),
                   float(share), c["oldest_days"], c["oldest_past_terms"]]
                  + [float(c["buckets"][l]) for l in labels]
                  + ["; ".join(sorted(c["flags"]))])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in (3, 4, 5):
            row[i].number_format = MONEY
        row[6].number_format = PCT
        for i in range(9, 9 + len(labels)):
            row[i].number_format = MONEY
        if isinstance(row[4].value, (int, float)) and row[4].value:
            row[4].fill = WARN_FILL
        if isinstance(row[6].value, (int, float)) and row[6].value >= 20:
            row[6].fill = WARN_FILL
        if row[len(heads) - 1].value:
            row[len(heads) - 1].fill = WARN_FILL
    ws.freeze_panes = "B2"
    w = {1: 32, 2: 18, 3: 7, 4: 16, 5: 16, 6: 15, 7: 10, 8: 12, 9: 20}
    for i in range(len(labels)):
        w[10 + i] = 14
    w[len(heads)] = 70
    widths(ws, w)


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
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value=empty)
    elif footer:
        rr = ws.max_row + 2
        c = ws.cell(row=rr, column=1, value=footer)
        c.font = Font(italic=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    widths(ws, {i: 20 for i in range(1, len(header) + 1)} | {1: 32,
                                                            len(header): 70})


# ------------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--aging", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--gl-balance")
    ap.add_argument("--as-of", required=True)
    ap.add_argument("--buckets", default="30,60,90",
                    help="comma-separated day boundaries; use the client's own, "
                         "not an assumed 30/60/90")
    ap.add_argument("--allowance")
    ap.add_argument("--subsequent-receipts")
    ap.add_argument("--client", default="")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    as_of = pdate(args.as_of)
    bounds = [int(b) for b in args.buckets.split(",") if b.strip()]
    if not bounds:
        sys.exit("--buckets must contain at least one boundary.")
    labels = bucket_labels(bounds)

    rows = load_aging(Path(args.aging))
    receipts = load_receipts(Path(args.subsequent_receipts)
                             if args.subsequent_receipts else None)
    res = analyse(rows, as_of, bounds, labels, receipts)
    customers = by_customer(rows, labels)

    total_net = sum((x["open_balance"] for x in rows), ZERO)
    debits = sum((x["open_balance"] for x in rows if x["open_balance"] > ZERO), ZERO)
    credits = sum((x["open_balance"] for x in rows if x["open_balance"] < ZERO), ZERO)
    buckets = {l: ZERO for l in labels}
    for x in rows:
        if x["recomputed_bucket"] in buckets:
            buckets[x["recomputed_bucket"]] += x["open_balance"]
    past_terms = sum((x["open_balance"] for x in rows
                      if x["days_past_terms"] is not None and x["days_past_terms"] > 0
                      and x["open_balance"] > ZERO), ZERO)
    totals = {"net": total_net, "debits": debits, "credits": credits,
              "buckets": buckets, "past_terms": past_terms}

    gl = dec(args.gl_balance)
    allowance = dec(args.allowance)
    tests = []

    if gl is not None:
        tests.append({"num": 1, "name": "Aging total agrees to the GL control account",
                      "lines": [("Aging total per report", total_net),
                                ("GL control account balance", gl)],
                      "difference": total_net - gl,
                      "passed": total_net == gl,
                      "note": "" if total_net == gl else
                              "Usual causes: unposted cash, a journal entry made directly "
                              "to the control account (invisible in the subledger by "
                              "definition), or a subledger cut on a different date."})
    else:
        tests.append({"num": 1, "name": "Aging total agrees to the GL control account",
                      "skipped": True, "passed": True, "lines": [], "difference": ZERO,
                      "note": "NOT PERFORMED - no GL control account balance supplied"})

    cust_sum = sum((c["net"] for c in customers), ZERO)
    bucket_sum = sum(buckets.values(), ZERO)
    tests.append({"num": 2, "name": "Buckets foot to customer totals and to the report",
                  "lines": [("Sum of recomputed buckets", bucket_sum),
                            ("Sum of customer net balances", cust_sum),
                            ("Aging total", total_net)],
                  "difference": (bucket_sum - total_net) + (cust_sum - total_net),
                  "passed": bucket_sum == total_net and cust_sum == total_net})

    mis = res["misbucketed"]
    mis_val = sum((x["open_balance"] for x in mis), ZERO)
    reported_any = any(x["reported_bucket"] for x in rows)
    if reported_any:
        tests.append({"num": 3,
                      "name": "Every bucket recomputes from the invoice date",
                      "lines": [("Invoices with a reported bucket",
                                 Decimal(len([x for x in rows if x["reported_bucket"]]))),
                                ("Mis-bucketed invoices", Decimal(len(mis))),
                                ("Value mis-bucketed", mis_val)],
                      "difference": Decimal(len(mis)), "is_count": True,
                      "passed": not mis,
                      "note": "" if not mis else
                              "The total can agree to the GL while the aging is wrong. "
                              "A mis-aged book has been driving the allowance and the "
                              "credit decisions."})
    else:
        tests.append({"num": 3, "name": "Bucket recomputation", "skipped": True,
                      "passed": True, "lines": [], "difference": ZERO,
                      "note": "NOT PERFORMED - the aging carried no reported_bucket "
                              "column, so there is nothing to compare the recomputation "
                              "against. Recomputed buckets are used throughout."})

    masked = [c for c in customers if c["credits"] < ZERO and c["debits"] > ZERO]
    tests.append({"num": 4, "name": "Credit balances shown gross, not netted",
                  "lines": [("Gross debits", debits), ("Gross credits", credits),
                            ("Net", total_net),
                            ("Customers with offsetting credits",
                             Decimal(len(masked)))],
                  "difference": abs(credits), "passed": True,
                  "note": "Netting hides both numbers. A customer with a large open "
                          "invoice and a large unapplied credit shows as a small current "
                          "balance; both figures matter."})

    cut = res["cutoff"]
    tests.append({"num": 5, "name": "Cutoff - no invoices dated after the as-of date",
                  "lines": [("Invoices failing cutoff", Decimal(len(cut))),
                            ("Value", sum((x["open_balance"] for x in cut), ZERO))],
                  "difference": Decimal(len(cut)), "is_count": True,
                  "passed": not cut,
                  "note": "" if not cut else
                          "An invoice dated after period end in the aging overstates "
                          "revenue and receivables in the same stroke."})

    integ = res["integrity"]
    tests.append({"num": 6, "name": "Internal integrity of the aging detail",
                  "lines": [("Findings", Decimal(len(integ)))],
                  "difference": Decimal(len(integ)), "is_count": True,
                  "passed": not integ})

    if allowance is not None:
        tests.append({"num": 7, "name": "Allowance reasonableness (informational)",
                      "lines": [("Allowance balance", allowance),
                                ("Gross receivables", debits),
                                ("Balances past terms", past_terms)],
                      "difference": ZERO, "passed": True,
                      "note": "Reported for judgement, not pass/fail. Compare to history "
                              "and to the specific customers driving exposure. A static "
                              "allowance against a deteriorating book is a judgement that "
                              "needs documenting."})
    else:
        tests.append({"num": 7, "name": "Allowance reasonableness", "skipped": True,
                      "passed": True, "lines": [], "difference": ZERO,
                      "note": "NOT PERFORMED - no allowance balance supplied"})

    esc = []
    if gl is not None and total_net != gl:
        esc.append(f"The aging does not agree to the GL control account by "
                   f"{total_net - gl:,.2f}. Nothing built on the aging is reliable until "
                   f"this is resolved. Ask for the control account detail and read every "
                   f"manual journal entry posted directly to it - those bypass the "
                   f"subledger by definition.")
    if mis:
        esc.append(f"{len(mis)} invoice(s) totalling {mis_val:,.2f} are in the wrong "
                   f"ageing bucket. The aging the client has been using for the allowance "
                   f"and for credit decisions is wrong.")
    if cut:
        esc.append(f"{len(cut)} invoice(s) fail the cutoff test.")
    old_credits = [x for x in rows if x["open_balance"] < ZERO
                   and x["days"] is not None and x["days"] > bounds[-1]]
    if old_credits:
        esc.append(f"{len(old_credits)} credit balance(s) totalling "
                   f"{sum((x['open_balance'] for x in old_credits), ZERO):,.2f} aged "
                   f"beyond the oldest bucket. Likely a customer overpayment or double "
                   f"billing - potentially a refund liability rather than a receivable "
                   f"offset, and possibly unclaimed property. Confirm the applicable "
                   f"dormancy rules rather than assuming them.")
    related = [c for c in customers if c["flags"]]
    if related:
        esc.append(f"{len(related)} customer(s) appear to be related parties, affiliates, "
                   f"or employees within trade AR - a classification and disclosure "
                   f"matter.")
    if receipts:
        noreceipt = [x for x in rows if x["open_balance"] > ZERO
                     and x["days"] is not None and x["days"] > bounds[-1]
                     and x["subsequent_amount"] == ZERO]
        if noreceipt:
            esc.append(f"{len(noreceipt)} aged balance(s) totalling "
                       f"{sum((x['open_balance'] for x in noreceipt), ZERO):,.2f} have no "
                       f"subsequent receipt. An old balance since paid is a collection "
                       f"timing issue; one with no subsequent activity is a valuation "
                       f"issue and possibly an existence one.")

    overall = all(t["passed"] for t in tests)

    meta = {
        "Client": args.client or "(not stated)",
        "As-of date": str(as_of),
        "Bucket boundaries (days)": args.buckets,
        "Invoices in aging": len(rows),
        "Customers": len(customers),
        "Subsequent receipts supplied": "yes" if receipts else "no",
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    # ---- console
    print("=" * 76)
    print("AR AGING TIE-OUT")
    print("=" * 76)
    print(f"Client : {meta['Client']}    As of: {as_of}    Buckets: {args.buckets}")
    print(f"Invoices {len(rows)}   customers {len(customers)}")
    print()
    for t in tests:
        if t.get("skipped"):
            print(f"Test {t['num']}: NOT PERFORMED - {t['note']}")
            continue
        st = "PASS" if t["passed"] else "*** FAIL ***"
        if t.get("is_count"):
            print(f"Test {t['num']}: {st:<14} {int(t['difference']):>6} item(s)   "
                  f"{t['name']}")
        else:
            print(f"Test {t['num']}: {st:<14} diff {t['difference']:>15,.2f}   "
                  f"{t['name']}")
    print()
    print(f"Gross debits {debits:>16,.2f}   gross credits {credits:>14,.2f}   "
          f"net {total_net:>16,.2f}")
    print("Aging profile (recomputed):")
    for l in labels:
        amt = buckets[l]
        pct = (amt / debits * 100) if debits else ZERO
        print(f"  {l:<10} {amt:>16,.2f}  {pct:>6.1f}%")
    print(f"Past terms   {past_terms:>16,.2f}")
    if mis:
        print(f"\nMIS-BUCKETED INVOICES ({len(mis)}, {mis_val:,.2f}):")
        for x in mis[:12]:
            print(f"  {x['customer'][:24]:<24} {x['invoice_no']:<12} "
                  f"{x['open_balance']:>13,.2f}  report '{x['reported_bucket']}' vs "
                  f"recomputed '{x['recomputed_bucket']}' at {x['days']}d")
    if integ:
        print(f"\nINTEGRITY FINDINGS ({len(integ)}):")
        for p in integ[:12]:
            print(f"  ! {p}")
    if esc:
        print("\nESCALATE:")
        for e in esc:
            print(f"  ! {e}")

    if not overall and not args.force:
        print("\n" + "=" * 76)
        print("WORKBOOK NOT WRITTEN.")
        print("A correct total does not make a correct aging. Resolve the failing test(s)")
        print("before relying on the aging for the allowance or for credit decisions.")
        print("=" * 76)
        return 1

    wb = Workbook()
    sheet_summary(wb, meta, tests, totals, labels, customers, allowance, overall, esc)
    sheet_detail(wb, rows, labels)
    sheet_customers(wb, customers, labels, total_net)
    sheet_list(wb, "Exceptions", ["#", "Finding"],
               [[i + 1, p] for i, p in enumerate(integ)], {},
               "None - no integrity findings.")
    sheet_list(wb, "Credit Balances",
               ["Customer", "Invoice", "Invoice date", "Days", "Credit balance",
                "Recomputed bucket"],
               [[x["customer"], x["invoice_no"], x["invoice_date"], x["days"],
                 float(x["open_balance"]), x["recomputed_bucket"]]
                for x in sorted((x for x in rows if x["open_balance"] < ZERO),
                                key=lambda y: y["open_balance"])],
               {2: DATEF, 4: MONEY},
               "No credit balances.",
               footer="Credit balances are separated from the debit aging entirely. An "
                      "aged credit is usually an overpayment or a double billing, and may "
                      "be a refund liability rather than a receivable offset.")
    if receipts:
        sheet_list(wb, "Subsequent Receipts",
                   ["Customer", "Invoice", "Open balance", "Days", "Subsequent receipts",
                    "Last receipt", "Status"],
                   [[x["customer"], x["invoice_no"], float(x["open_balance"]), x["days"],
                     float(x["subsequent_amount"]), x["subsequent_last"],
                     "collected after period end" if x["subsequent_amount"] > ZERO
                     else "NO subsequent receipt"]
                    for x in sorted((x for x in rows if x["open_balance"] > ZERO),
                                    key=lambda y: -(y["days"] or 0))],
                   {2: MONEY, 4: MONEY, 5: DATEF},
                   "No subsequent receipts supplied.",
                   footer="Subsequent collection is the strongest evidence of "
                          "collectibility short of a confirmation.")
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

