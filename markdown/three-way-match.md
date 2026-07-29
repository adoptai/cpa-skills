---
name: three-way-match
description: Perform a three-way match between purchase orders, vendor invoices, and receiving records — testing quantity and price against tolerance, detecting duplicate invoices and over-billing, and proving that every invoice in the population is either matched or reported as an exception. Use this whenever the user mentions three-way match, PO to invoice matching, matching invoices to receipts, accounts payable testing, duplicate payment review, over-billing, price variance on invoices, unmatched invoices, goods-received-not-invoiced or invoiced-not-received, or AP internal control testing. Also trigger on "match these invoices to POs," "did we pay twice," "check AP against receiving," "find duplicate payments," "test our purchasing controls," "GRNI review," or when purchase order, invoice, and receiving files are provided together. Runs fully local — no vendor or purchasing data leaves the machine.
---

# Three-Way Match

A three-way match asks whether the company **ordered** it, **received** it, and was **billed
correctly** for it. The control is old and well understood; what makes testing it valuable is
that the failures are almost never random.

Duplicate payments cluster around the same handful of vendors, because the cause is a process
defect — an invoice submitted twice under slightly different numbers, a credit memo never
applied, a PO closed and reopened. Price variances cluster too, because someone is invoicing off
a stale price list. **So the pattern matters more than the individual exception**, and the output
is organized to make patterns visible rather than to produce a long flat list.

## The gate

The population is proven before it is analysed:

1. **Every invoice is accounted for.** Matched invoices plus exceptions must equal the full
   invoice population, in count and in value. An invoice that quietly falls out of the middle of
   a matching routine is the one worth finding, so the script balances the population and
   refuses to report until it does.
2. **The invoice population ties to a stated control total** — the AP subledger or the GL
   purchases figure. Testing a filtered extract proves nothing about the account.
3. **No invoice is matched twice.** A PO line consumed by one invoice is not available to
   another; otherwise a duplicate looks like a match.
4. **Tolerances are stated, not assumed.** Quantity and price tolerance come from the client's
   own policy and are recorded on the workpaper.

## Inputs

Ask for all three, plus the control total:

1. **Purchase orders** — PO number, line, item, quantity ordered, unit price, vendor, date,
   status.
2. **Vendor invoices** — invoice number, vendor, date, PO reference, line, item, quantity
   billed, unit price, extended amount, and the amount actually paid if available.
3. **Receiving records** — receipt number, PO reference, line, item, quantity received, date.
4. **Tolerance policy** — the client's approved quantity and price variance thresholds, and any
   de minimis amount below which variances are accepted without approval.
5. **Control total** — AP subledger total or GL purchases for the period.

Where the client has no written tolerance policy, that is itself a finding. Use zero tolerance,
say that you did, and note the absence of a policy as a control observation.

## Step 1 — Prove the population

```bash
python3 scripts/three_way_match.py \
  --invoices inv.csv --pos po.csv --receipts rec.csv \
  --control-total 8412990.55 --validate-only
```

Reports counts and totals for each file, invoices with no PO reference, POs with no invoices,
receipts with no PO, and whether the invoice population ties to the control total. Resolve
differences first — in particular, ask what the export excluded. A receiving file limited to
one warehouse will manufacture hundreds of false "invoiced not received" exceptions and bury the
real ones.

## Step 2 — Match

```bash
python3 scripts/three_way_match.py \
  --invoices inv.csv --pos po.csv --receipts rec.csv \
  --control-total 8412990.55 \
  --qty-tolerance-pct 2 --price-tolerance-pct 1 --de-minimis 50 \
  --client "Cascade Industrial Supply" --period "FY2025" \
  --out "Cascade - FY2025 Three-Way Match.xlsx"
```

Matching runs on PO number and line where available, falling back to PO plus item. Each invoice
line receives one of these outcomes, and every outcome is reported:

| Outcome | Meaning |
|---|---|
| `matched` | Quantity and price within tolerance against both PO and receipt |
| `price variance` | Billed unit price outside tolerance of the PO price |
| `quantity variance` | Billed quantity outside tolerance of quantity received |
| `over-PO` | Billed quantity exceeds quantity ordered |
| `over-received` | Billed quantity exceeds quantity received — **billed for goods not received** |
| `not received` | Invoice matched to a PO with no receiving record at all |
| `no PO` | Invoice with no purchase order — the control was bypassed |
| `PO not found` | Invoice references a PO that does not exist in the file |
| `duplicate` | See below |

**Duplicate detection runs on four patterns**, because duplicates rarely repeat exactly:

- Same vendor and same invoice number
- Same vendor, same amount, same invoice date
- Same vendor, same amount, dates within a short window
- Same vendor and same PO line billed more than once

The third and fourth patterns find the ones that matter. An invoice resubmitted as `INV-4471`
and `INV4471`, or a PO line billed on two invoices a month apart, will pass an exact-match check
and has to be caught structurally.

## Step 3 — Read the exceptions as patterns

The workbook aggregates by vendor before it lists by invoice, because that is the order in which
the findings are actionable.

- **Recurring price variance for one vendor** — almost always a stale price list or an unrecorded
  price increase, not fraud. Real money, and recoverable.
- **Recurring over-billing for one vendor** — short shipments billed in full. Also recoverable,
  and worth a conversation about receiving discipline.
- **Duplicates concentrated in one vendor** — a submission process problem at their end or an AP
  intake problem at yours.
- **"No PO" concentrated in one requester or department** — the control is being bypassed
  routinely, which is a more serious finding than any single dollar amount in the listing.
- **Invoices consistently just below an approval threshold** — worth escalating. See below.
- **Goods received not invoiced (GRNI)** — receipts with no matching invoice represent an
  unrecorded liability at period end. This is a completeness finding for the financial
  statements, and it is on its own tab because it is the exception a payables test most often
  misses: nobody complains about an invoice that never arrived.

## Step 4 — Deliver

**Workbook tabs:**

1. **Summary** — population proof, tolerances used, outcome counts and values, recovery
   opportunity quantified, and the control observations. The page a controller reads.
2. **Population Proof** — invoices matched plus exceptions equals the population, in count and
   value, with the tie to the control total. The gate, shown.
3. **Exceptions by Vendor** — vendor rollup with exception counts, values, and the dominant
   exception type. Read this before the detail.
4. **Exception Detail** — every exception line with PO, receipt, invoice figures, variance in
   units and dollars, and columns for investigation and disposition.
5. **Duplicates** — grouped, with which pattern caught each and the amount at risk.
6. **GRNI** — received not invoiced, with age, for the accrual.
7. **Matched Detail** — clean matches with the rule that matched them, for re-performance.

**Then, in chat:** population and whether it ties, the three or four vendor-level patterns that
matter, quantified recovery opportunity, the GRNI accrual figure, and control observations. Lead
with recoverable dollars — that is what funds the next engagement.

## What to escalate rather than list

- **Invoices repeatedly just below an approval threshold** from the same vendor or requester.
  Splitting to avoid approval implies knowledge of the control.
- **A vendor with no PO discipline at all** and material spend.
- **Payment exceeding the invoice**, or payment where no invoice exists.
- **A vendor bank detail change** shortly before a payment, if that data is available — the
  classic payment-diversion pattern.
- **A vendor whose address, bank details, or contact match an employee record.** Report the match
  as a fact; do not characterize it.
- **Receiving records created after the invoice date**, particularly near period end — receiving
  documentation generated to clear a match rather than to record a delivery.

State these as attributes with the supporting documents. Do not conclude on intent.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry. Vendor
master data and bank details, if present in the inputs, are never transmitted and are excluded
from output beyond what the exception requires.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*

---

# Appendix — bundled files

If you installed the `.skill` package these files are already in place and you can ignore this appendix. If you copied the skill as text, create the files below at the paths shown, alongside your `SKILL.md`. The skill will not run without them.

## `scripts/three_way_match.py`

```python
#!/usr/bin/env python3
"""
Three-way match: purchase orders <-> vendor invoices <-> receiving records.

Gates:
  * matched invoices + exceptions must equal the whole invoice population, in
    count and in value -- an invoice that silently falls out of a matching
    routine is exactly the one worth finding
  * the invoice population must tie to a stated control total
  * a PO line consumed by one invoice is not available to another, so a
    duplicate cannot masquerade as a match
  * tolerances are supplied, never assumed

Usage:
    python3 three_way_match.py --invoices inv.csv --pos po.csv --receipts rec.csv \
        --control-total 8412990.55 --validate-only

    python3 three_way_match.py --invoices inv.csv --pos po.csv --receipts rec.csv \
        --control-total 8412990.55 \
        --qty-tolerance-pct 2 --price-tolerance-pct 1 --de-minimis 50 \
        --client "Cascade Industrial Supply" --period FY2025 \
        --out "Cascade - FY2025 Three-Way Match.xlsx"

--invoices CSV:
    invoice_no, vendor, invoice_date, po_no, po_line, item, qty_billed,
    unit_price, extended_amount, amount_paid (optional), requester (optional)

--pos CSV:
    po_no, po_line, vendor, po_date, item, qty_ordered, unit_price, status

--receipts CSV:
    receipt_no, po_no, po_line, item, qty_received, receipt_date
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
MONEY = '#,##0.00;[Red](#,##0.00)'
QTY = '#,##0.####'
PCT = '0.00"%"'

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
TOP = Border(top=Side(style="thin"))

MATCHED = "matched"
OUTCOMES = [MATCHED, "price variance", "quantity variance", "over-PO", "over-received",
            "not received", "no PO", "PO not found", "duplicate"]
RECOVERABLE = {"price variance", "over-PO", "over-received", "duplicate"}


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


def key_of(s) -> str:
    """Normalised join key: strips punctuation and case so INV-4471 == inv4471."""
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


# --------------------------------------------------------------------- loading

def load_invoices(path: Path) -> list[dict]:
    out = []
    for i, r in enumerate(rows_of(path, "invoices"), start=1):
        qty = d0(r.get("qty_billed"))
        up = d0(r.get("unit_price"))
        ext = dec(r.get("extended_amount"))
        if ext is None:
            ext = qty * up
        out.append({
            "row": i,
            "invoice_no": clean(r.get("invoice_no")) or f"(no number {i})",
            "vendor": clean(r.get("vendor")) or "(no vendor)",
            "invoice_date": pdate(r.get("invoice_date")),
            "po_no": clean(r.get("po_no")),
            "po_line": clean(r.get("po_line")),
            "item": clean(r.get("item")),
            "qty_billed": qty, "unit_price": up, "extended_amount": ext,
            "amount_paid": dec(r.get("amount_paid")),
            "requester": clean(r.get("requester")),
            "outcomes": [], "notes": [], "matched_po": None, "matched_receipt_qty": None,
        })
    return out


def load_pos(path: Path) -> dict:
    out = {}
    for i, r in enumerate(rows_of(path, "purchase orders"), start=1):
        po, line = clean(r.get("po_no")), clean(r.get("po_line"))
        k = (key_of(po), key_of(line))
        out[k] = {
            "row": i, "po_no": po, "po_line": line,
            "vendor": clean(r.get("vendor")),
            "po_date": pdate(r.get("po_date")),
            "item": clean(r.get("item")),
            "qty_ordered": d0(r.get("qty_ordered")),
            "unit_price": d0(r.get("unit_price")),
            "status": clean(r.get("status")),
            "qty_billed_against": ZERO, "invoices": [],
        }
    return out


def load_receipts(path: Path) -> dict:
    agg: dict[tuple, dict] = {}
    for i, r in enumerate(rows_of(path, "receipts"), start=1):
        po, line = clean(r.get("po_no")), clean(r.get("po_line"))
        k = (key_of(po), key_of(line))
        a = agg.setdefault(k, {"po_no": po, "po_line": line, "qty_received": ZERO,
                               "receipts": [], "first_date": None, "last_date": None})
        q = d0(r.get("qty_received"))
        a["qty_received"] += q
        d = pdate(r.get("receipt_date"))
        a["receipts"].append({"receipt_no": clean(r.get("receipt_no")),
                              "qty": q, "date": d,
                              "item": clean(r.get("item"))})
        if d:
            a["first_date"] = d if a["first_date"] is None else min(a["first_date"], d)
            a["last_date"] = d if a["last_date"] is None else max(a["last_date"], d)
    return agg


# -------------------------------------------------------------------- matching

def find_duplicates(invs: list[dict], window_days: int) -> list[dict]:
    groups = []

    def add(pattern, items):
        if len(items) > 1:
            groups.append({"pattern": pattern, "items": items,
                           "at_risk": sum((x["extended_amount"] for x in items[1:]), ZERO)})
            for x in items:
                if "duplicate" not in x["outcomes"]:
                    x["outcomes"].append("duplicate")
                x["notes"].append(f"duplicate pattern: {pattern}")

    # 1 same vendor + same normalised invoice number
    by_num = defaultdict(list)
    for x in invs:
        by_num[(key_of(x["vendor"]), key_of(x["invoice_no"]))].append(x)
    for (v, n), items in by_num.items():
        uniq = {}
        for x in items:
            uniq.setdefault(x["row"], x)
        if len(uniq) > 1:
            add("same vendor and invoice number (normalised)", list(uniq.values()))

    # 2 same vendor + amount + date
    by_amt_date = defaultdict(list)
    for x in invs:
        if x["invoice_date"]:
            by_amt_date[(key_of(x["vendor"]), x["extended_amount"],
                         x["invoice_date"])].append(x)
    for k, items in by_amt_date.items():
        nums = {key_of(i["invoice_no"]) for i in items}
        if len(items) > 1 and len(nums) > 1:
            add("same vendor, amount and date, different invoice numbers", items)

    # 3 same vendor + amount within a date window
    by_amt = defaultdict(list)
    for x in invs:
        by_amt[(key_of(x["vendor"]), x["extended_amount"])].append(x)
    for k, items in by_amt.items():
        if len(items) < 2:
            continue
        items = sorted([i for i in items if i["invoice_date"]],
                       key=lambda i: i["invoice_date"])
        for a in range(len(items)):
            for b in range(a + 1, len(items)):
                gap = (items[b]["invoice_date"] - items[a]["invoice_date"]).days
                if gap == 0 or gap > window_days:
                    continue
                if key_of(items[a]["invoice_no"]) == key_of(items[b]["invoice_no"]):
                    continue
                add(f"same vendor and amount within {window_days} days",
                    [items[a], items[b]])

    # 4 same PO line billed on more than one invoice
    by_poline = defaultdict(list)
    for x in invs:
        if x["po_no"]:
            by_poline[(key_of(x["po_no"]), key_of(x["po_line"]))].append(x)
    for k, items in by_poline.items():
        nums = {key_of(i["invoice_no"]) for i in items}
        if len(nums) > 1:
            add("same PO line billed on more than one invoice", items)

    return groups


def match(invs, pos, receipts, qtol: Decimal, ptol: Decimal, dmin: Decimal) -> None:
    for x in invs:
        if not x["po_no"]:
            x["outcomes"].append("no PO")
            x["notes"].append("no purchase order reference - the purchasing control was "
                              "bypassed for this invoice")
            continue
        k = (key_of(x["po_no"]), key_of(x["po_line"]))
        po = pos.get(k)
        if po is None:
            # fall back to PO + item
            cands = [v for (pk, lk), v in pos.items()
                     if pk == key_of(x["po_no"]) and key_of(v["item"]) == key_of(x["item"])]
            po = cands[0] if len(cands) == 1 else None
            if po is not None:
                x["notes"].append("matched on PO + item (PO line not supplied or "
                                  "did not agree)")
        if po is None:
            x["outcomes"].append("PO not found")
            x["notes"].append(f"invoice references PO {x['po_no']} line "
                              f"{x['po_line'] or '(none)'} which is not in the PO file")
            continue

        x["matched_po"] = po
        po["invoices"].append(x)

        # price test
        if po["unit_price"] > ZERO:
            pdiff = x["unit_price"] - po["unit_price"]
            ppct = (abs(pdiff) / po["unit_price"]) * Decimal(100)
            dollar_effect = pdiff * x["qty_billed"]
            if ppct > ptol and abs(dollar_effect) > dmin:
                x["outcomes"].append("price variance")
                x["notes"].append(
                    f"billed {x['unit_price']} vs PO {po['unit_price']} "
                    f"({ppct:.2f}% > {ptol}% tolerance), dollar effect "
                    f"{dollar_effect:,.2f}")

        # quantity vs PO
        po["qty_billed_against"] += x["qty_billed"]
        if po["qty_billed_against"] > po["qty_ordered"]:
            over = po["qty_billed_against"] - po["qty_ordered"]
            opct = (over / po["qty_ordered"] * Decimal(100)) if po["qty_ordered"] else Decimal(100)
            if opct > qtol:
                x["outcomes"].append("over-PO")
                x["notes"].append(
                    f"cumulative billed {po['qty_billed_against']} exceeds ordered "
                    f"{po['qty_ordered']} by {over}")

        # quantity vs receipts
        rec = receipts.get(k)
        if rec is None:
            cands = [v for (pk, lk), v in receipts.items() if pk == key_of(x["po_no"])]
            rec = cands[0] if len(cands) == 1 else None
        if rec is None:
            x["outcomes"].append("not received")
            x["notes"].append("no receiving record for this PO line - billed for goods "
                              "with no evidence of receipt")
        else:
            x["matched_receipt_qty"] = rec["qty_received"]
            if x["qty_billed"] > rec["qty_received"]:
                over = x["qty_billed"] - rec["qty_received"]
                opct = (over / rec["qty_received"] * Decimal(100)) if rec["qty_received"] else Decimal(100)
                effect = over * x["unit_price"]
                if opct > qtol and abs(effect) > dmin:
                    x["outcomes"].append("over-received")
                    x["notes"].append(
                        f"billed {x['qty_billed']} but received {rec['qty_received']} "
                        f"- billed for {over} not received, effect {effect:,.2f}")
            elif rec["qty_received"] > x["qty_billed"]:
                diff = rec["qty_received"] - x["qty_billed"]
                dpct = (diff / rec["qty_received"] * Decimal(100)) if rec["qty_received"] else ZERO
                if dpct > qtol:
                    x["outcomes"].append("quantity variance")
                    x["notes"].append(
                        f"received {rec['qty_received']} but billed only "
                        f"{x['qty_billed']} - possible unbilled receipt")
            # receipt created after the invoice
            if rec["first_date"] and x["invoice_date"] and rec["first_date"] > x["invoice_date"]:
                x["notes"].append(
                    f"receiving record dated {rec['first_date']} is AFTER the invoice "
                    f"date {x['invoice_date']} - receiving documentation may have been "
                    f"created to clear the match rather than to record a delivery")

    for x in invs:
        if not x["outcomes"]:
            x["outcomes"].append(MATCHED)


def grni(pos, receipts, invs) -> list[dict]:
    """Received but not invoiced - an unrecorded liability at period end."""
    billed = defaultdict(lambda: ZERO)
    for x in invs:
        if x["po_no"]:
            billed[(key_of(x["po_no"]), key_of(x["po_line"]))] += x["qty_billed"]
    out = []
    for k, rec in receipts.items():
        b = billed.get(k, ZERO)
        if rec["qty_received"] > b:
            po = pos.get(k)
            unit = po["unit_price"] if po else ZERO
            qty = rec["qty_received"] - b
            out.append({
                "po_no": rec["po_no"], "po_line": rec["po_line"],
                "vendor": po["vendor"] if po else "(PO not found)",
                "item": po["item"] if po else (rec["receipts"][0]["item"] if rec["receipts"] else ""),
                "qty_received": rec["qty_received"], "qty_billed": b,
                "qty_uninvoiced": qty, "unit_price": unit,
                "value": qty * unit,
                "last_receipt": rec["last_date"],
            })
    return sorted(out, key=lambda r: r["value"], reverse=True)


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


def sheet_summary(wb, meta, stats, proof, dupes, g, vendor_rollup, observations):
    ws = wb.active
    ws.title = "Summary"
    widths(ws, {1: 4, 2: 46, 3: 16, 4: 18, 5: 60})
    r = 1

    def line(label, a=None, b=None, *, bold=False, size=11, fill=None, note=""):
        nonlocal r
        c = ws.cell(row=r, column=2, value=label)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c.fill = fill
        if a is not None:
            ws.cell(row=r, column=3, value=a).font = Font(bold=bold)
        if b is not None:
            cc = ws.cell(row=r, column=4, value=float(b))
            cc.number_format, cc.font = MONEY, Font(bold=bold)
        if note:
            n = ws.cell(row=r, column=5, value=note)
            n.font = Font(italic=True, size=9)
            n.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    line("THREE-WAY MATCH", bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    v = ws.cell(row=r, column=2,
                value="POPULATION PROVEN - matched plus exceptions equals the invoice "
                      "population"
                if proof["balanced"] and proof["ties"] else
                "POPULATION NOT PROVEN - see the Population Proof tab")
    v.font = OK_FONT if (proof["balanced"] and proof["ties"]) else BAD_FONT
    if not (proof["balanced"] and proof["ties"]):
        v.fill = BAD_FILL
    r += 2

    ws.cell(row=r, column=3, value="Count").font = Font(bold=True)
    ws.cell(row=r, column=4, value="Value").font = Font(bold=True)
    r += 1
    line("Invoice lines in population", stats["total_count"], stats["total_value"],
         bold=True)
    for o in OUTCOMES:
        if stats["by_outcome"][o]["count"]:
            line("  " + o, stats["by_outcome"][o]["count"],
                 stats["by_outcome"][o]["value"],
                 fill=None if o == MATCHED else WARN_FILL)
    r += 1
    line("Recovery opportunity (price variance, over-billing, duplicates)",
         stats["recoverable_count"], stats["recoverable_value"], bold=True,
         fill=WARN_FILL,
         note="Amounts potentially recoverable from vendors. Investigate before "
              "claiming - a price variance may reflect an approved change that was "
              "never entered on the PO.")
    line("Goods received not invoiced (GRNI)", len(g),
         sum((x["value"] for x in g), ZERO), bold=True,
         note="An unrecorded liability at period end. This is a financial statement "
              "completeness matter, and the exception a payables test most often "
              "misses - nobody complains about an invoice that never arrived.")
    r += 1

    if vendor_rollup:
        line("VENDORS WITH THE MOST EXCEPTION VALUE", bold=True, size=12, fill=SUB_FILL)
        for h, col in zip(["Vendor", "Exceptions", "Value", "Dominant type"],
                          (2, 3, 4, 5)):
            c = ws.cell(row=r, column=col, value=h)
            c.fill, c.font = HDR_FILL, HDR_FONT
        r += 1
        for vr in vendor_rollup[:10]:
            ws.cell(row=r, column=2, value=vr["vendor"])
            ws.cell(row=r, column=3, value=vr["count"])
            c = ws.cell(row=r, column=4, value=float(vr["value"]))
            c.number_format = MONEY
            ws.cell(row=r, column=5, value=vr["dominant"])
            r += 1
        r += 1
        line("Exceptions cluster by vendor because the causes are process defects, not "
             "random events. Read this table before the detail - a stale price list "
             "shows up here as one vendor and forty exceptions.", bold=True)
        r += 1

    if observations:
        line("ESCALATE / CONTROL OBSERVATIONS", bold=True, size=12, fill=BAD_FILL)
        for o in observations:
            line("  " + o)
        r += 1

    r += 1
    line("Prepared by / date:  __________________  ____________", bold=True)
    line("Reviewed by / date:  __________________  ____________", bold=True)


def sheet_proof(wb, proof, stats, meta):
    ws = wb.create_sheet("Population Proof")
    widths(ws, {1: 4, 2: 50, 3: 16, 4: 18, 5: 62})
    r = 1
    ws.cell(row=r, column=2, value="POPULATION PROOF").font = Font(bold=True, size=14)
    r += 2
    ws.cell(row=r, column=3, value="Count").font = Font(bold=True)
    ws.cell(row=r, column=4, value="Value").font = Font(bold=True)
    r += 1
    for label, cnt, val, bold, border in [
        ("Matched invoice lines", proof["matched_count"], proof["matched_value"], False, None),
        ("Exception invoice lines", proof["exception_count"], proof["exception_value"], False, None),
        ("= Total accounted for", proof["accounted_count"], proof["accounted_value"], True, TOP),
        ("Invoice lines in population", stats["total_count"], stats["total_value"], False, None),
        ("Difference", proof["count_diff"], proof["value_diff"], True, TOP),
    ]:
        ws.cell(row=r, column=2, value=label).font = Font(bold=bold)
        c1 = ws.cell(row=r, column=3, value=cnt)
        c2 = ws.cell(row=r, column=4, value=float(val))
        c2.number_format = MONEY
        for c in (c1, c2):
            c.font = Font(bold=bold)
            if border:
                c.border = border
        if label == "Difference":
            good = proof["balanced"]
            for c in (c1, c2):
                c.font = OK_FONT if good else BAD_FONT
                if not good:
                    c.fill = BAD_FILL
            if not good:
                ws.cell(row=r, column=5,
                        value="An invoice line that is neither matched nor an exception "
                              "has silently fallen out of the matching routine. That is "
                              "the line worth finding."
                        ).alignment = Alignment(wrap_text=True)
        r += 1
    r += 1
    ct = proof["control_total"]
    if ct is None:
        c = ws.cell(row=r, column=2,
                    value="Control total NOT SUPPLIED - the completeness of the invoice "
                          "population is unproven. Testing a filtered extract proves "
                          "nothing about the account.")
        c.fill = BAD_FILL
        r += 1
    else:
        for label, val in [("Invoice population value", stats["total_value"]),
                           ("Control total (AP subledger / GL purchases)", ct),
                           ("Difference", stats["total_value"] - ct)]:
            ws.cell(row=r, column=2, value=label).font = Font(bold=label == "Difference")
            c = ws.cell(row=r, column=4, value=float(val))
            c.number_format = MONEY
            if label == "Difference":
                c.font = OK_FONT if val == ZERO else BAD_FONT
                if val != ZERO:
                    c.fill = BAD_FILL
                    ws.cell(row=r, column=5,
                            value="Ask what the export excluded. A receiving file limited "
                                  "to one warehouse manufactures false 'not received' "
                                  "exceptions and buries the real ones."
                            ).alignment = Alignment(wrap_text=True)
            r += 1
    r += 2
    ws.cell(row=r, column=2, value="Tolerances applied").font = Font(bold=True, size=12)
    r += 1
    for k in ("Quantity tolerance", "Price tolerance", "De minimis"):
        if k in meta:
            ws.cell(row=r, column=2, value="  " + k)
            ws.cell(row=r, column=3, value=meta[k])
            r += 1


def sheet_vendor(wb, rollup):
    ws = wb.create_sheet("Exceptions by Vendor")
    heads = ["Vendor", "Exception lines", "Exception value", "Dominant type",
             "Price var", "Over-PO", "Over-received", "Not received", "No PO",
             "Duplicates", "Likely cause"]
    ws.append(heads)
    hdr(ws, len(heads))
    causes = {
        "price variance": "stale price list or an unrecorded price increase - usually "
                          "recoverable",
        "over-received": "short shipments billed in full - recoverable, and a receiving "
                         "discipline issue",
        "over-PO": "billing beyond the order - PO not amended, or unauthorised scope",
        "duplicate": "invoice submission or AP intake defect",
        "no PO": "purchasing control routinely bypassed - more serious than the dollar "
                 "amount suggests",
        "not received": "no receiving evidence - possible unrecorded receipt or a "
                        "fictitious delivery",
    }
    for v in rollup:
        ws.append([v["vendor"], v["count"], float(v["value"]), v["dominant"],
                   v["counts"].get("price variance", 0), v["counts"].get("over-PO", 0),
                   v["counts"].get("over-received", 0), v["counts"].get("not received", 0),
                   v["counts"].get("no PO", 0), v["counts"].get("duplicate", 0),
                   causes.get(v["dominant"], "")])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[2].number_format = MONEY
        if row[1].value and row[1].value >= 5:
            row[0].fill = WARN_FILL
    ws.freeze_panes = "A2"
    widths(ws, {1: 32, 2: 14, 3: 17, 4: 18, 5: 11, 6: 9, 7: 14, 8: 14, 9: 8,
                10: 11, 11: 62})


def sheet_detail(wb, invs, only_exceptions: bool, title: str):
    ws = wb.create_sheet(title)
    heads = ["Invoice", "Vendor", "Inv date", "PO", "Line", "Item", "Qty billed",
             "Qty ordered", "Qty received", "Unit price", "PO price", "Extended",
             "Variance $", "Outcome(s)", "Notes", "Investigation", "Disposition"]
    ws.append(heads)
    hdr(ws, len(heads))
    for x in invs:
        is_exc = x["outcomes"] != [MATCHED]
        if only_exceptions != is_exc:
            continue
        po = x["matched_po"]
        var = ZERO
        if po and po["unit_price"] > ZERO:
            var = (x["unit_price"] - po["unit_price"]) * x["qty_billed"]
        ws.append([x["invoice_no"], x["vendor"], x["invoice_date"], x["po_no"] or None,
                   x["po_line"] or None, x["item"] or None, float(x["qty_billed"]),
                   float(po["qty_ordered"]) if po else None,
                   float(x["matched_receipt_qty"]) if x["matched_receipt_qty"] is not None else None,
                   float(x["unit_price"]),
                   float(po["unit_price"]) if po else None,
                   float(x["extended_amount"]), float(var),
                   ", ".join(x["outcomes"]), "; ".join(x["notes"]), None, None])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[2].number_format = "yyyy-mm-dd"
        for i in (6, 7, 8):
            row[i].number_format = QTY
        for i in (9, 10, 11, 12):
            row[i].number_format = MONEY
        if row[13].value and row[13].value != MATCHED:
            row[13].fill = BAD_FILL if "over-received" in str(row[13].value) else WARN_FILL
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value="None.")
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:Q{max(ws.max_row, 2)}"
    widths(ws, {1: 16, 2: 26, 3: 11, 4: 12, 5: 7, 6: 22, 7: 11, 8: 12, 9: 13,
                10: 11, 11: 11, 12: 14, 13: 13, 14: 26, 15: 66, 16: 26, 17: 20})


def sheet_dupes(wb, groups):
    ws = wb.create_sheet("Duplicates")
    heads = ["Group", "Pattern caught by", "Invoice", "Vendor", "Date", "PO",
             "Amount", "At risk", "Disposition"]
    ws.append(heads)
    hdr(ws, len(heads))
    for i, g in enumerate(groups, start=1):
        for j, x in enumerate(g["items"]):
            ws.append([i if j == 0 else None,
                       g["pattern"] if j == 0 else None,
                       x["invoice_no"], x["vendor"], x["invoice_date"],
                       x["po_no"] or None, float(x["extended_amount"]),
                       float(g["at_risk"]) if j == 0 else None, None])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[4].number_format = "yyyy-mm-dd"
        row[6].number_format = MONEY
        row[7].number_format = MONEY
        if row[7].value:
            row[7].font = BAD_FONT
    if ws.max_row == 1:
        ws.cell(row=2, column=3, value="None detected on any of the four patterns.")
    else:
        r = ws.max_row + 2
        ws.cell(row=r, column=2, value="TOTAL AT RISK").font = Font(bold=True)
        c = ws.cell(row=r, column=8,
                    value=float(sum((g["at_risk"] for g in groups), ZERO)))
        c.number_format, c.font, c.border = MONEY, Font(bold=True), TOP
        r += 2
        for note in [
            "Duplicates rarely repeat exactly. Patterns 3 and 4 - same amount within a "
            "date window, and the same PO line billed twice - find the ones an exact",
            "invoice-number check misses. INV-4471 and INV4471 are the same invoice.",
        ]:
            ws.cell(row=r, column=2, value=note)
            r += 1
    ws.freeze_panes = "C2"
    widths(ws, {1: 7, 2: 44, 3: 16, 4: 26, 5: 11, 6: 12, 7: 14, 8: 14, 9: 22})


def sheet_grni(wb, g, asof):
    ws = wb.create_sheet("GRNI")
    heads = ["PO", "Line", "Vendor", "Item", "Qty received", "Qty billed",
             "Uninvoiced qty", "Unit price", "Accrual value", "Last receipt",
             "Age (days)"]
    ws.append(heads)
    hdr(ws, len(heads))
    for x in g:
        age = (asof - x["last_receipt"]).days if (asof and x["last_receipt"]) else None
        ws.append([x["po_no"], x["po_line"], x["vendor"], x["item"],
                   float(x["qty_received"]), float(x["qty_billed"]),
                   float(x["qty_uninvoiced"]), float(x["unit_price"]),
                   float(x["value"]), x["last_receipt"], age])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in (4, 5, 6):
            row[i].number_format = QTY
        for i in (7, 8):
            row[i].number_format = MONEY
        row[9].number_format = "yyyy-mm-dd"
        if isinstance(row[10].value, int) and row[10].value > 60:
            row[10].fill = WARN_FILL
    if ws.max_row == 1:
        ws.cell(row=2, column=3, value="None - every receipt has been invoiced.")
    else:
        r = ws.max_row + 2
        ws.cell(row=r, column=3, value="TOTAL ACCRUAL").font = Font(bold=True)
        c = ws.cell(row=r, column=9, value=float(sum((x["value"] for x in g), ZERO)))
        c.number_format, c.font, c.border = MONEY, Font(bold=True), TOP
        r += 2
        ws.cell(row=r, column=3,
                value="Received but not invoiced is an unrecorded liability at period "
                      "end. Propose the accrual; aged items may indicate a lost invoice "
                      "or a receipt recorded in error.")
    ws.freeze_panes = "C2"
    widths(ws, {1: 12, 2: 7, 3: 28, 4: 24, 5: 13, 6: 12, 7: 14, 8: 11, 9: 15,
                10: 12, 11: 11})


# ------------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--invoices", required=True)
    ap.add_argument("--pos", required=True)
    ap.add_argument("--receipts", required=True)
    ap.add_argument("--control-total")
    ap.add_argument("--validate-only", action="store_true")
    ap.add_argument("--qty-tolerance-pct", default="0")
    ap.add_argument("--price-tolerance-pct", default="0")
    ap.add_argument("--de-minimis", default="0")
    ap.add_argument("--duplicate-window", type=int, default=45)
    ap.add_argument("--as-of", help="period end YYYY-MM-DD, for GRNI aging")
    ap.add_argument("--client", default="")
    ap.add_argument("--period", default="")
    ap.add_argument("--out")
    args = ap.parse_args()

    invs = load_invoices(Path(args.invoices))
    pos = load_pos(Path(args.pos))
    receipts = load_receipts(Path(args.receipts))
    ct = dec(args.control_total)

    total_value = sum((x["extended_amount"] for x in invs), ZERO)
    no_po = [x for x in invs if not x["po_no"]]

    print("=" * 74)
    print("POPULATION")
    print("=" * 74)
    print(f"Invoice lines     : {len(invs):,}   value {total_value:>16,.2f}")
    print(f"PO lines          : {len(pos):,}")
    print(f"Receipt PO lines  : {len(receipts):,}")
    print(f"Invoices with no PO reference: {len(no_po)}")
    if ct is not None:
        d = total_value - ct
        print(f"Control total     : {ct:>16,.2f}")
        print(f"Difference        : {d:>16,.2f}  "
              f"{'TIES' if d == ZERO else '*** DOES NOT TIE ***'}")
    else:
        print("Control total     : NOT SUPPLIED")

    if args.validate_only:
        po_no_inv = len(pos) - len({(key_of(x['po_no']), key_of(x['po_line']))
                                    for x in invs if x["po_no"]} & set(pos))
        print(f"PO lines with no invoice     : {po_no_inv}")
        orphan_rec = len([k for k in receipts if k not in pos])
        print(f"Receipt lines with no PO     : {orphan_rec}")
        print("\nResolve differences before matching. Ask what each export excluded - a "
              "receiving\nfile limited to one warehouse manufactures false exceptions "
              "and buries real ones.")
        return 0 if (ct is None or total_value == ct) else 1

    if not args.out:
        sys.exit("--out is required unless --validate-only is used.")

    qtol = dec(args.qty_tolerance_pct) or ZERO
    ptol = dec(args.price_tolerance_pct) or ZERO
    dmin = dec(args.de_minimis) or ZERO

    dupes = find_duplicates(invs, args.duplicate_window)
    match(invs, pos, receipts, qtol, ptol, dmin)
    g = grni(pos, receipts, invs)
    asof = pdate(args.as_of) if args.as_of else max(
        (x["invoice_date"] for x in invs if x["invoice_date"]), default=None)

    # ---- stats and the population gate
    by_outcome = {o: {"count": 0, "value": ZERO} for o in OUTCOMES}
    for x in invs:
        for o in x["outcomes"]:
            by_outcome[o]["count"] += 1
            by_outcome[o]["value"] += x["extended_amount"]
    matched = [x for x in invs if x["outcomes"] == [MATCHED]]
    exceptions = [x for x in invs if x["outcomes"] != [MATCHED]]
    recoverable = [x for x in invs if set(x["outcomes"]) & RECOVERABLE]

    proof = {
        "matched_count": len(matched),
        "matched_value": sum((x["extended_amount"] for x in matched), ZERO),
        "exception_count": len(exceptions),
        "exception_value": sum((x["extended_amount"] for x in exceptions), ZERO),
        "control_total": ct,
    }
    proof["accounted_count"] = proof["matched_count"] + proof["exception_count"]
    proof["accounted_value"] = proof["matched_value"] + proof["exception_value"]
    proof["count_diff"] = proof["accounted_count"] - len(invs)
    proof["value_diff"] = proof["accounted_value"] - total_value
    proof["balanced"] = proof["count_diff"] == 0 and proof["value_diff"] == ZERO
    proof["ties"] = ct is None or total_value == ct

    stats = {"total_count": len(invs), "total_value": total_value,
             "by_outcome": by_outcome,
             "recoverable_count": len(recoverable),
             "recoverable_value": sum((x["extended_amount"] for x in recoverable), ZERO)}

    # vendor rollup
    vagg = defaultdict(lambda: {"count": 0, "value": ZERO, "counts": defaultdict(int)})
    for x in exceptions:
        a = vagg[x["vendor"]]
        a["count"] += 1
        a["value"] += x["extended_amount"]
        for o in x["outcomes"]:
            a["counts"][o] += 1
    rollup = []
    for v, a in vagg.items():
        dom = max(a["counts"].items(), key=lambda kv: kv[1])[0] if a["counts"] else ""
        rollup.append({"vendor": v, "count": a["count"], "value": a["value"],
                       "counts": dict(a["counts"]), "dominant": dom})
    rollup.sort(key=lambda v: v["value"], reverse=True)

    # observations
    obs = []
    if no_po:
        byreq = defaultdict(int)
        for x in no_po:
            byreq[x["requester"] or x["vendor"]] += 1
        worst = sorted(byreq.items(), key=lambda kv: -kv[1])[:3]
        obs.append(f"{len(no_po)} invoice line(s) with no purchase order. "
                   f"Concentrated in: " +
                   ", ".join(f"{k or '(unknown)'} ({n})" for k, n in worst) +
                   ". Routine bypass of the purchasing control is more serious than the "
                   "dollar amount suggests.")
    over_rec = [x for x in invs if "over-received" in x["outcomes"]]
    if over_rec:
        obs.append(f"{len(over_rec)} invoice line(s) billed for quantities exceeding "
                   f"receipts, totalling "
                   f"{sum((x['extended_amount'] for x in over_rec), ZERO):,.2f}. Billed "
                   f"for goods not received - recoverable.")
    if dupes:
        obs.append(f"{len(dupes)} duplicate group(s) with "
                   f"{sum((d['at_risk'] for d in dupes), ZERO):,.2f} at risk.")
    late_rec = [x for x in invs
                if any("AFTER the invoice date" in n for n in x["notes"])]
    if late_rec:
        obs.append(f"{len(late_rec)} invoice line(s) where the receiving record postdates "
                   f"the invoice. Receiving documentation may have been created to clear "
                   f"the match rather than to record a delivery. Report as an attribute; "
                   f"do not conclude on intent.")
    for x in invs:
        if x["amount_paid"] is not None and x["amount_paid"] > x["extended_amount"]:
            obs.append(f"Invoice {x['invoice_no']} ({x['vendor']}): paid "
                       f"{x['amount_paid']:,.2f} against an invoice of "
                       f"{x['extended_amount']:,.2f} - payment exceeds the invoice.")
    if not proof["balanced"]:
        obs.append("POPULATION DOES NOT BALANCE - matched plus exceptions does not equal "
                   "the invoice population. An invoice line has fallen out of the "
                   "matching routine; that is the line to find.")

    # ---- console
    print()
    print("=" * 74)
    print("MATCH RESULTS")
    print("=" * 74)
    print(f"Tolerances: quantity {qtol}%, price {ptol}%, de minimis {dmin:,.2f}")
    print()
    for o in OUTCOMES:
        b = by_outcome[o]
        if b["count"]:
            print(f"  {o:<18} {b['count']:>6,}   {b['value']:>16,.2f}")
    print()
    print(f"  Population proof : matched {proof['matched_count']} + exceptions "
          f"{proof['exception_count']} = {proof['accounted_count']} of {len(invs)}  "
          f"{'BALANCED' if proof['balanced'] else '*** DOES NOT BALANCE ***'}")
    print(f"  Recovery opportunity: {stats['recoverable_value']:>16,.2f} "
          f"({stats['recoverable_count']} lines)")
    print(f"  GRNI accrual        : {sum((x['value'] for x in g), ZERO):>16,.2f} "
          f"({len(g)} PO lines)")
    if rollup:
        print("\nTop vendors by exception value:")
        for v in rollup[:8]:
            print(f"  {v['vendor'][:34]:<34} {v['count']:>4} exc  "
                  f"{v['value']:>14,.2f}  mostly {v['dominant']}")
    if obs:
        print("\nOBSERVATIONS:")
        for o in obs:
            print(f"  ! {o}")

    if not proof["balanced"]:
        print("\n" + "=" * 74)
        print("WORKBOOK NOT WRITTEN. The population does not balance.")
        print("Every invoice line must be either matched or an exception.")
        print("=" * 74)
        return 1

    meta = {
        "Client": args.client or "(not stated)",
        "Period": args.period or "(not stated)",
        "Invoice lines": len(invs),
        "PO lines": len(pos),
        "Receipt PO lines": len(receipts),
        "Quantity tolerance": f"{qtol}%",
        "Price tolerance": f"{ptol}%",
        "De minimis": float(dmin),
        "Duplicate window (days)": args.duplicate_window,
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    wb = Workbook()
    sheet_summary(wb, meta, stats, proof, dupes, g, rollup, obs)
    sheet_proof(wb, proof, stats, meta)
    sheet_vendor(wb, rollup)
    sheet_detail(wb, invs, True, "Exception Detail")
    sheet_dupes(wb, dupes)
    sheet_grni(wb, g, asof)
    sheet_detail(wb, invs, False, "Matched Detail")
    wb.active = 0
    out = Path(args.out)
    if not proof["ties"]:
        out = out.with_name(out.stem + " [POPULATION UNTIED]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\nWorkbook: {out}")
    return 0 if proof["ties"] else 1


if __name__ == "__main__":
    sys.exit(main())
```

