#!/usr/bin/env python3
"""
Month-end accrual proposals.

This is the inverse of the cutoff-and-unrecorded-liabilities search: instead of
proving a search population is complete and finding what was missed, this tool
proposes what should be booked, from three sources -

  * open purchase orders with goods or services received but not yet invoiced
    (GRNI), cross-checked against the AP register so a stale "invoiced" figure
    on the PO report cannot produce a duplicate accrual
  * recurring vendor/category spend, accrued at the trailing average of prior
    periods, but only when enough history exists and only when the actual
    invoice for the current period is not already in the AP register
  * manual one-time items the preparer supplies, each requiring its own basis

It never posts anything. The output is a proposed journal entry batch, every
line carrying a basis and a reversal flag, for a reviewer (typically the
Controller) to approve before it is booked. Nothing is proposed without a
stated, data-backed reason - a candidate that cannot be verified becomes a
visible exception, never a guess.

Gates:
  * every open-PO line, every recurring vendor/category candidate, and every
    manual item must be classified - proposed, excluded (with a reason), or
    an exception (with a reason). None may be silently dropped.
  * every proposed line has a non-blank basis and an explicit reversal flag;
    a non-reversing item requires a stated rationale.
  * the proposed batch must balance - total debits equal total credits.

Usage:
    python3 accrual_proposal.py \
        --open-pos open_pos.csv --ap-register ap_register.csv \
        --vendor-history vendor_history.csv --prior-accruals prior_accruals.csv \
        --manual-items manual_items.csv \
        --period-end 2026-08-31 --materiality 25000 \
        --client "Ardenway Manufacturing" \
        --out "Ardenway - August 2026 Accrual Proposals.xlsx"

--open-pos CSV:        po_no, vendor, description, gl_account, received_amount,
                       invoiced_amount, receipt_date
--ap-register CSV:     invoice_date, vendor, po_no, category, invoice_ref, amount
--vendor-history CSV:  vendor, category, gl_account, period, amount
--prior-accruals CSV:  vendor, category, period, amount, reversed, reversal_date
--manual-items CSV:    vendor, description, gl_account, amount, basis, reversing,
                       reversal_rationale     (all optional except vendor/amount)
"""

from __future__ import annotations

import argparse
import calendar
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


# --------------------------------------------------------------------- parsing

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


def yn(raw, default=True) -> bool:
    s = clean(raw).lower()
    if s in ("y", "yes", "true", "1"):
        return True
    if s in ("n", "no", "false", "0"):
        return False
    return default


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


def month_start(d: date) -> date:
    return date(d.year, d.month, 1)


def next_month_start(d: date) -> date:
    if d.month == 12:
        return date(d.year + 1, 1, 1)
    return date(d.year, d.month + 1, 1)


def rows_of(path: Path, label: str):
    import csv
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No rows in {label} ({path})")
    return rows


def load_simple(path: Path | None, label: str) -> list[dict]:
    if not path or not path.exists():
        return []
    return rows_of(path, label)


# ----------------------------------------------------------------------- POs

def process_open_pos(rows, ap_by_po, period_end) -> list[dict]:
    out = []
    for i, r in enumerate(rows, start=1):
        received = d0(r.get("received_amount"))
        invoiced = d0(r.get("invoiced_amount"))
        receipt_date = pdate(r.get("receipt_date"))
        po_no = clean(r.get("po_no"))
        item = {
            "row": i, "po_no": po_no, "vendor": clean(r.get("vendor")) or "(no vendor)",
            "description": clean(r.get("description")), "gl_account": clean(r.get("gl_account")),
            "received_amount": received, "invoiced_amount": invoiced,
            "receipt_date": receipt_date, "ap_tie": None,
            "outstanding": received - invoiced, "classification": "", "proposed": ZERO,
            "notes": [],
        }
        ap_amt = ap_by_po.get(key(po_no)) if po_no else None
        item["ap_tie"] = ap_amt

        if item["outstanding"] < ZERO:
            item["classification"] = "exception - invoiced exceeds received"
            item["notes"].append(
                f"PO report shows {invoiced:,.2f} invoiced against {received:,.2f} "
                f"received - review before proposing anything on this PO")
        elif item["outstanding"] == ZERO:
            item["classification"] = "excluded - fully invoiced"
        elif receipt_date is None:
            item["classification"] = "exception - no receipt date"
            item["notes"].append(
                "cannot confirm goods or services were received before period end "
                "without a receipt date - not proposed")
        elif receipt_date > period_end:
            item["classification"] = "excluded - received after period end"
        elif ap_amt is not None and ap_amt != invoiced:
            item["classification"] = "exception - PO report does not tie to AP register"
            item["notes"].append(
                f"AP register shows {ap_amt:,.2f} invoiced against this PO through "
                f"period end, PO report shows {invoiced:,.2f} - reconcile before "
                f"proposing an accrual")
        else:
            item["classification"] = "proposed"
            item["proposed"] = item["outstanding"]
            item["notes"].append(
                f"received {received:,.2f} (receipt dated {receipt_date}), invoiced "
                f"{invoiced:,.2f} as of {period_end} - accrue the difference")
        out.append(item)
    return out


# ------------------------------------------------------------------ recurring

def process_recurring(rows, ap_current, prior_not_reversed, period_start, period_end,
                       min_history) -> list[dict]:
    groups = defaultdict(list)
    meta = {}
    for r in rows:
        vendor, category = clean(r.get("vendor")), clean(r.get("category")) or "(uncategorized)"
        gk = (key(vendor), key(category))
        per = pdate(r.get("period"))
        if per is None or per >= period_start:
            continue  # only prior history feeds the average
        groups[gk].append((per, d0(r.get("amount"))))
        meta[gk] = (vendor, category, clean(r.get("gl_account")))

    out = []
    for gk, points in groups.items():
        vendor, category, gl_account = meta[gk]
        points.sort(key=lambda p: p[0], reverse=True)
        n = len(points)
        avg = (sum((v for _, v in points), ZERO) / n) if n else ZERO
        item = {
            "vendor": vendor, "category": category, "gl_account": gl_account,
            "periods_used": [str(p) for p, _ in points], "n_periods": n,
            "average": avg, "classification": "", "proposed": ZERO, "notes": [],
        }
        already = ap_current.get(gk)
        if already is not None:
            item["classification"] = "excluded - actual invoice already recorded this period"
            item["notes"].append(
                f"AP register already has {already:,.2f} for this vendor/category "
                f"within {period_start}-{period_end} - no estimate needed")
        elif n < min_history:
            item["classification"] = "exception - insufficient history"
            item["notes"].append(
                f"only {n} prior period(s) available, need at least {min_history} - "
                f"not proposed. Supply more vendor history or add as a manual item "
                f"with an explicit basis.")
        else:
            item["classification"] = "proposed"
            item["proposed"] = avg
            item["notes"].append(
                f"trailing average of {n} period(s) "
                f"({', '.join(str(p) for p, _ in points)}): {avg:,.2f}")
        if gk in prior_not_reversed:
            item["notes"].append(
                "prior-period accrual for this vendor/category was not marked "
                "reversed - check for double counting before booking")
        out.append(item)
    return out


# ------------------------------------------------------------------------- manual

def process_manual(rows) -> list[dict]:
    out = []
    for i, r in enumerate(rows, start=1):
        basis = clean(r.get("basis"))
        item = {
            "row": i, "vendor": clean(r.get("vendor")) or "(no vendor)",
            "description": clean(r.get("description")), "gl_account": clean(r.get("gl_account")),
            "amount": d0(r.get("amount")), "basis": basis,
            "reversing": yn(r.get("reversing"), default=True),
            "reversal_rationale": clean(r.get("reversal_rationale")),
            "classification": "", "notes": [],
        }
        if not basis:
            item["classification"] = "exception - missing basis"
            item["notes"].append(
                "every manual item requires a stated basis - a preparer's assertion "
                "with no basis is not a proposal")
        elif item["amount"] == ZERO:
            item["classification"] = "exception - zero amount"
            item["notes"].append("amount is zero or blank")
        elif not item["reversing"] and not item["reversal_rationale"]:
            item["classification"] = "exception - non-reversing with no rationale"
            item["notes"].append(
                "marked as not reversing but no rationale given - a non-reversing "
                "accrual needs a stated reason, not a default")
        else:
            item["classification"] = "proposed"
        out.append(item)
    return out


# ------------------------------------------------------------------------ batch

def build_batch(po_items, recurring_items, manual_items, credit_account, period_end,
                 reversal_date):
    batch = []
    n = 0
    for x in po_items:
        if x["classification"] != "proposed":
            continue
        n += 1
        batch.append({
            "je": n, "date": period_end, "source": "open PO / GRNI",
            "vendor": x["vendor"], "description": x["description"] or f"PO {x['po_no']}",
            "debit_account": x["gl_account"] or "(unspecified expense/asset)",
            "amount": x["proposed"], "credit_account": credit_account,
            "basis": "; ".join(x["notes"]), "reversing": True,
            "reversal_date": reversal_date, "reference": x["po_no"],
        })
    for x in recurring_items:
        if x["classification"] != "proposed":
            continue
        n += 1
        batch.append({
            "je": n, "date": period_end, "source": "recurring vendor pattern",
            "vendor": x["vendor"], "description": x["category"],
            "debit_account": x["gl_account"] or "(unspecified expense)",
            "amount": x["proposed"], "credit_account": credit_account,
            "basis": "; ".join(x["notes"]), "reversing": True,
            "reversal_date": reversal_date, "reference": x["category"],
        })
    for x in manual_items:
        if x["classification"] != "proposed":
            continue
        n += 1
        batch.append({
            "je": n, "date": period_end, "source": "manual item",
            "vendor": x["vendor"], "description": x["description"],
            "debit_account": x["gl_account"] or "(unspecified expense/asset)",
            "amount": x["amount"], "credit_account": credit_account,
            "basis": x["basis"], "reversing": x["reversing"],
            "reversal_date": reversal_date if x["reversing"] else None,
            "reference": x["vendor"],
        })
    return batch


# ------------------------------------------------------------------------ sheets

def hdr(ws, n, row=1):
    for c in range(1, n + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def widths(ws, w):
    for c, v in w.items():
        ws.column_dimensions[get_column_letter(c)].width = v


def sheet_summary(wb, meta, tests, stats, esc, overall, materiality):
    ws = wb.active
    ws.title = "Summary"
    widths(ws, {1: 4, 2: 52, 3: 16, 4: 18, 5: 62})
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

    line("MONTH-END ACCRUAL PROPOSALS", bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    v = ws.cell(row=r, column=2,
                value="PROPOSAL ONLY - not posted. Every input row is classified; "
                      "nothing is booked without reviewer approval."
                if overall else
                "GATE FAILED - see the failing test(s) below. No batch produced.")
    v.font = OK_FONT if overall else BAD_FONT
    if not overall:
        v.fill = BAD_FILL
    r += 2

    line("PROPOSED ACCRUALS", bold=True, size=12, fill=SUB_FILL)
    line("  From open POs (GRNI)", stats["po_n"], stats["po_value"])
    line("  From recurring vendor patterns", stats["rec_n"], stats["rec_value"])
    line("  From manual items", stats["man_n"], stats["man_value"])
    line("  TOTAL PROPOSED (debit = credit)", None, stats["total"], bold=True)
    if materiality is not None:
        line("  Materiality", None, materiality)
        over = stats["total"] > materiality
        line("  Exceeds materiality?", "YES" if over else "no", bold=True,
             fill=BAD_FILL if over else None)
    r += 1

    line("NOT PROPOSED", bold=True, size=12, fill=SUB_FILL)
    line("  Excluded (already recorded / fully invoiced / future period)",
         stats["excluded_n"], stats["excluded_value"])
    line("  Exceptions (require review before proposing)",
         stats["exception_n"], stats["exception_value"],
         fill=BAD_FILL if stats["exception_n"] else None)
    r += 1

    for t in tests:
        line(f"TEST - {t['name']}", "PASS" if t["passed"] else "FAIL",
             bold=True, fill=None if t["passed"] else BAD_FILL, note=t.get("note", ""))
        c = ws.cell(row=r - 1, column=3)
        c.font = OK_FONT if t["passed"] else BAD_FONT
    r += 1

    if stats["not_reversed_n"]:
        line("PRIOR ACCRUALS NOT MARKED REVERSED", stats["not_reversed_n"], None,
             bold=True, fill=WARN_FILL,
             note="Check these for double counting before booking this period's proposals.")
        r += 1

    if esc:
        line("ESCALATE", bold=True, size=12, fill=BAD_FILL)
        for e in esc:
            line("  " + e)
        r += 1
    r += 1
    line("Prepared by / date:  __________________  ____________", bold=True)
    line("Reviewed / approved by (Controller) / date:  __________________  ____________",
         bold=True)


def sheet_batch(wb, batch, stats):
    ws = wb.create_sheet("Proposed Accrual JE Batch")
    heads = ["JE #", "Date", "Source", "Vendor", "Description", "Debit account",
              "Debit", "Credit account", "Credit", "Reversing?", "Reversal date",
              "Reference", "Basis"]
    ws.append(heads)
    hdr(ws, len(heads))
    for x in batch:
        ws.append([x["je"], x["date"], x["source"], x["vendor"], x["description"] or None,
                   x["debit_account"], float(x["amount"]), None, None,
                   "yes" if x["reversing"] else "no", x["reversal_date"],
                   x["reference"] or None, x["basis"]])
        ws.append([None, None, None, None, "  " + x["credit_account"], None, None,
                   x["credit_account"], float(x["amount"]), None, None, None, None])
    if not batch:
        ws.cell(row=2, column=3, value="No accruals proposed this period.")
    else:
        last = ws.max_row
        r = last + 2
        ws.cell(row=r, column=5, value="TOTAL (must be equal)").font = Font(bold=True)
        for col, letter in ((7, "G"), (9, "I")):
            c = ws.cell(row=r, column=col, value=f"=SUM({letter}2:{letter}{last})")
            c.number_format, c.font, c.border = MONEY, Font(bold=True), TOP
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[1].number_format = DATEF
        row[6].number_format = MONEY
        row[8].number_format = MONEY
        row[10].number_format = DATEF
        row[12].alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    widths(ws, {1: 6, 2: 12, 3: 20, 4: 22, 5: 26, 6: 24, 7: 14, 8: 24, 9: 14,
                10: 11, 11: 13, 12: 16, 13: 60})


def sheet_reversal(wb, batch):
    ws = wb.create_sheet("Auto-Reversal Schedule")
    heads = ["JE #", "Reversal date", "Debit account (was credit)", "Debit",
              "Credit account (was debit)", "Credit", "Reference"]
    ws.append(heads)
    hdr(ws, len(heads))
    rev = [x for x in batch if x["reversing"]]
    for x in rev:
        ws.append([x["je"], x["reversal_date"], x["credit_account"], float(x["amount"]),
                   x["debit_account"], float(x["amount"]), x["reference"] or None])
    if not rev:
        ws.cell(row=2, column=1, value="No reversing entries in this batch.")
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[1].number_format = DATEF
        row[3].number_format = MONEY
        row[5].number_format = MONEY
    non_rev = [x for x in batch if not x["reversing"]]
    if non_rev:
        r = ws.max_row + 2
        ws.cell(row=r, column=1, value="NON-REVERSING (rationale required, see batch tab)"
                ).font = Font(bold=True)
        for x in non_rev:
            r += 1
            ws.cell(row=r, column=1, value=x["je"])
            ws.cell(row=r, column=3, value=f"{x['vendor']} - {x['description']}")
    widths(ws, {1: 6, 2: 14, 3: 30, 4: 14, 5: 30, 6: 14, 7: 16})


def sheet_pos(wb, items):
    ws = wb.create_sheet("Open PO Detail (GRNI)")
    heads = ["PO", "Vendor", "Description", "GL account", "Received", "Invoiced (PO report)",
              "Invoiced (AP register)", "Outstanding", "Receipt date", "Classification",
              "Proposed", "Notes"]
    ws.append(heads)
    hdr(ws, len(heads))
    order = {"proposed": 0}
    for x in sorted(items, key=lambda y: (order.get(y["classification"], 1),
                                          -y["outstanding"])):
        ws.append([x["po_no"], x["vendor"], x["description"] or None, x["gl_account"] or None,
                   float(x["received_amount"]), float(x["invoiced_amount"]),
                   float(x["ap_tie"]) if x["ap_tie"] is not None else None,
                   float(x["outstanding"]), x["receipt_date"], x["classification"],
                   float(x["proposed"]) if x["proposed"] else None, "; ".join(x["notes"])])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in (4, 5, 6, 7, 10):
            row[i].number_format = MONEY
        row[8].number_format = DATEF
        cls = str(row[9].value)
        if cls.startswith("exception"):
            row[9].font, row[9].fill = BAD_FONT, BAD_FILL
        elif cls == "proposed":
            row[9].fill = WARN_FILL
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:L{max(ws.max_row, 2)}"
    widths(ws, {1: 12, 2: 24, 3: 30, 4: 16, 5: 14, 6: 16, 7: 16, 8: 14, 9: 12, 10: 30,
                11: 14, 12: 60})
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value="No open PO report supplied.")


def sheet_recurring(wb, items):
    ws = wb.create_sheet("Recurring Vendor Patterns")
    heads = ["Vendor", "Category", "GL account", "Periods used", "# periods",
              "Trailing average", "Classification", "Proposed", "Notes"]
    ws.append(heads)
    hdr(ws, len(heads))
    for x in sorted(items, key=lambda y: (y["classification"] != "proposed", y["vendor"])):
        ws.append([x["vendor"], x["category"], x["gl_account"] or None,
                   ", ".join(x["periods_used"]), x["n_periods"], float(x["average"]),
                   x["classification"], float(x["proposed"]) if x["proposed"] else None,
                   "; ".join(x["notes"])])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[5].number_format = MONEY
        row[7].number_format = MONEY
        cls = str(row[6].value)
        if cls.startswith("exception"):
            row[6].font, row[6].fill = BAD_FONT, BAD_FILL
        elif cls == "proposed":
            row[6].fill = WARN_FILL
    ws.freeze_panes = "A2"
    widths(ws, {1: 22, 2: 20, 3: 16, 4: 34, 5: 10, 6: 16, 7: 30, 8: 14, 9: 60})
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value="No vendor history supplied.")


def sheet_manual(wb, items):
    ws = wb.create_sheet("Manual Items")
    heads = ["Vendor", "Description", "GL account", "Amount", "Basis", "Reversing?",
              "Rationale (if not reversing)", "Classification"]
    ws.append(heads)
    hdr(ws, len(heads))
    for x in items:
        ws.append([x["vendor"], x["description"] or None, x["gl_account"] or None,
                   float(x["amount"]), x["basis"] or None,
                   "yes" if x["reversing"] else "no", x["reversal_rationale"] or None,
                   x["classification"]])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[3].number_format = MONEY
        cls = str(row[7].value)
        if cls.startswith("exception"):
            row[7].font, row[7].fill = BAD_FONT, BAD_FILL
        elif cls == "proposed":
            row[7].fill = WARN_FILL
    widths(ws, {1: 22, 2: 30, 3: 16, 4: 14, 5: 44, 6: 11, 7: 34, 8: 26})
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value="No manual items supplied.")


def sheet_exceptions(wb, po_items, recurring_items, manual_items):
    ws = wb.create_sheet("Exceptions")
    heads = ["Source", "Vendor / category", "Amount at risk", "Classification", "Notes"]
    ws.append(heads)
    hdr(ws, len(heads))
    n = 0
    for x in po_items:
        if not x["classification"].startswith("exception"):
            continue
        n += 1
        ws.append(["open PO", f"{x['vendor']} ({x['po_no']})", float(x["outstanding"]),
                   x["classification"], "; ".join(x["notes"])])
    for x in recurring_items:
        if not x["classification"].startswith("exception"):
            continue
        n += 1
        ws.append(["recurring pattern", f"{x['vendor']} / {x['category']}",
                   float(x["average"]), x["classification"], "; ".join(x["notes"])])
    for x in manual_items:
        if not x["classification"].startswith("exception"):
            continue
        n += 1
        ws.append(["manual item", x["vendor"], float(x["amount"]), x["classification"],
                   "; ".join(x["notes"])])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[2].number_format = MONEY
        row[3].font, row[3].fill = BAD_FONT, BAD_FILL
    widths(ws, {1: 16, 2: 30, 3: 16, 4: 34, 5: 60})
    if n == 0:
        ws.cell(row=2, column=1, value="None. Every candidate was proposed or excluded "
                                       "with a stated, non-exception reason.")


def sheet_prior(wb, rows, not_reversed_keys):
    ws = wb.create_sheet("Prior Accrual Reversal Check")
    heads = ["Vendor", "Category", "Period", "Amount", "Reversed?", "Reversal date", "Flag"]
    ws.append(heads)
    hdr(ws, len(heads))
    for r in rows:
        vendor, category = clean(r.get("vendor")), clean(r.get("category"))
        reversed_ = yn(r.get("reversed"), default=False)
        flagged = (key(vendor), key(category)) in not_reversed_keys
        ws.append([vendor, category, clean(r.get("period")), float(d0(r.get("amount"))),
                   "yes" if reversed_ else "no", clean(r.get("reversal_date")) or None,
                   "NOT REVERSED - check for double counting" if flagged else ""])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[3].number_format = MONEY
        if row[6].value:
            row[4].font, row[4].fill = BAD_FONT, BAD_FILL
            row[6].font = BAD_FONT
    widths(ws, {1: 22, 2: 20, 3: 12, 4: 14, 5: 11, 6: 14, 7: 38})
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value="No prior accrual file supplied.")


# ------------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                  formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--open-pos")
    ap.add_argument("--ap-register")
    ap.add_argument("--vendor-history")
    ap.add_argument("--prior-accruals")
    ap.add_argument("--manual-items")
    ap.add_argument("--period-end", required=True)
    ap.add_argument("--period-start", help="Defaults to the first day of the "
                     "period-end month; override for a non-calendar-month close.")
    ap.add_argument("--reversal-date", help="Defaults to the first day of the "
                     "month after period end.")
    ap.add_argument("--materiality")
    ap.add_argument("--min-history", type=int, default=3,
                     help="Minimum prior periods of vendor history required before a "
                          "recurring accrual is proposed rather than flagged (default 3).")
    ap.add_argument("--credit-account", default="Accrued Liabilities")
    ap.add_argument("--client", default="")
    ap.add_argument("--out", required=True)
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    period_end = pdate(args.period_end)
    period_start = pdate(args.period_start) if args.period_start else month_start(period_end)
    reversal_date = pdate(args.reversal_date) if args.reversal_date else \
        next_month_start(period_end)
    materiality = dec(args.materiality)

    open_pos = load_simple(Path(args.open_pos) if args.open_pos else None, "open POs")
    ap_reg = load_simple(Path(args.ap_register) if args.ap_register else None, "AP register")
    vendor_hist = load_simple(Path(args.vendor_history) if args.vendor_history else None,
                               "vendor history")
    prior_accr = load_simple(Path(args.prior_accruals) if args.prior_accruals else None,
                              "prior accruals")
    manual = load_simple(Path(args.manual_items) if args.manual_items else None,
                          "manual items")

    ap_by_po = defaultdict(lambda: ZERO)
    ap_current_vc = defaultdict(lambda: ZERO)
    have_po_amt, have_vc_amt = set(), set()
    for r in ap_reg:
        d = pdate(r.get("invoice_date"))
        amt = d0(r.get("amount"))
        po_no = clean(r.get("po_no"))
        if po_no and d and d <= period_end:
            ap_by_po[key(po_no)] += amt
            have_po_amt.add(key(po_no))
        vendor, category = clean(r.get("vendor")), clean(r.get("category"))
        if vendor and category and d and period_start <= d <= period_end:
            gk = (key(vendor), key(category))
            ap_current_vc[gk] += amt
            have_vc_amt.add(gk)
    ap_by_po = {k: v for k, v in ap_by_po.items() if k in have_po_amt}
    ap_current_vc = {k: v for k, v in ap_current_vc.items() if k in have_vc_amt}

    not_reversed_keys = set()
    for r in prior_accr:
        if not yn(r.get("reversed"), default=False):
            not_reversed_keys.add((key(r.get("vendor")), key(r.get("category"))))

    po_items = process_open_pos(open_pos, ap_by_po, period_end)
    recurring_items = process_recurring(vendor_hist, ap_current_vc, not_reversed_keys,
                                        period_start, period_end, args.min_history)
    manual_items = process_manual(manual)

    batch = build_batch(po_items, recurring_items, manual_items, args.credit_account,
                        period_end, reversal_date)

    def bucket(items, key_fn=lambda x: x["classification"]):
        n_p = sum(1 for x in items if key_fn(x) == "proposed")
        n_e = sum(1 for x in items if key_fn(x).startswith("exception"))
        n_x = len(items) - n_p - n_e
        return n_p, n_e, n_x

    po_p, po_e, po_x = bucket(po_items)
    rec_p, rec_e, rec_x = bucket(recurring_items)
    man_p, man_e, man_x = bucket(manual_items)

    po_value = sum((x["proposed"] for x in po_items if x["classification"] == "proposed"), ZERO)
    rec_value = sum((x["proposed"] for x in recurring_items
                     if x["classification"] == "proposed"), ZERO)
    man_value = sum((x["amount"] for x in manual_items
                     if x["classification"] == "proposed"), ZERO)
    exception_value = (sum((x["outstanding"] for x in po_items
                            if x["classification"].startswith("exception")), ZERO) +
                       sum((x["average"] for x in recurring_items
                            if x["classification"].startswith("exception")), ZERO) +
                       sum((x["amount"] for x in manual_items
                            if x["classification"].startswith("exception")), ZERO))
    excluded_n = po_x + rec_x + man_x
    excluded_value = (sum((x["outstanding"] for x in po_items
                           if x["classification"].startswith("excluded")), ZERO) +
                      sum((x["average"] for x in recurring_items
                           if x["classification"].startswith("excluded")), ZERO))

    total = po_value + rec_value + man_value
    debit_total = sum((x["amount"] for x in batch), ZERO)

    stats = {
        "po_n": po_p, "po_value": po_value, "rec_n": rec_p, "rec_value": rec_value,
        "man_n": man_p, "man_value": man_value, "total": total,
        "excluded_n": excluded_n, "excluded_value": excluded_value,
        "exception_n": po_e + rec_e + man_e, "exception_value": exception_value,
        "not_reversed_n": len(not_reversed_keys),
    }

    def classified_fully(items):
        return all(x["classification"] for x in items)

    batch_balances = debit_total == total
    missing_basis = [x for x in batch if not x["basis"]]

    tests = [
        {"name": "Every open PO line classified",
         "passed": classified_fully(po_items),
         "note": "" if classified_fully(po_items) else "at least one PO line has no "
                 "classification - this is a defect in the run, not the data"},
        {"name": "Every recurring vendor/category candidate classified",
         "passed": classified_fully(recurring_items),
         "note": ""},
        {"name": "Every manual item classified",
         "passed": classified_fully(manual_items),
         "note": ""},
        {"name": "Every proposed line has a basis",
         "passed": len(missing_basis) == 0,
         "note": "" if not missing_basis else
                 f"{len(missing_basis)} proposed line(s) have no basis - blocked"},
        {"name": "Proposed batch balances (debits = credits)",
         "passed": batch_balances,
         "note": "" if batch_balances else
                 f"debit total {debit_total:,.2f} vs proposed total {total:,.2f}"},
    ]
    overall = all(t["passed"] for t in tests)

    esc = []
    if stats["exception_n"]:
        esc.append(f"{stats['exception_n']} candidate(s) totalling "
                   f"{stats['exception_value']:,.2f} could not be proposed and need "
                   f"review - see the Exceptions tab.")
    if not_reversed_keys:
        esc.append(f"{len(not_reversed_keys)} prior-period accrual(s) are not marked "
                   f"reversed. Confirm they were not double-booked with this period's "
                   f"actuals or proposals before approving this batch.")
    if materiality is not None and total > materiality:
        esc.append(f"Total proposed accruals of {total:,.2f} exceed materiality of "
                   f"{materiality:,.2f}.")
    mismatched_po = [x for x in po_items
                     if x["classification"] == "exception - PO report does not tie to AP register"]
    if mismatched_po:
        esc.append(f"{len(mismatched_po)} PO(s) show an invoiced amount on the PO report "
                   f"that does not match the AP register - reconcile the PO report before "
                   f"relying on it for other periods.")
    if open_pos and not ap_reg:
        esc.append("No AP register supplied - open-PO candidates could not be "
                   "cross-checked against recorded invoices. Proposed GRNI amounts "
                   "rely entirely on the PO report's own 'invoiced' field.")

    meta = {
        "Client": args.client or "(not stated)",
        "Period start": str(period_start),
        "Period end": str(period_end),
        "Auto-reversal date": str(reversal_date),
        "Minimum history for a recurring accrual": args.min_history,
        "Credit account": args.credit_account,
        "Materiality": float(materiality) if materiality is not None else "(not supplied)",
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    # ---- console
    print("=" * 76)
    print("MONTH-END ACCRUAL PROPOSALS")
    print("=" * 76)
    print(f"Client : {meta['Client']}   Period {period_start} to {period_end}   "
          f"Reversal {reversal_date}")
    print()
    print(f"Proposed from open POs (GRNI)        {po_value:>15,.2f}  ({po_p} items)")
    print(f"Proposed from recurring patterns      {rec_value:>15,.2f}  ({rec_p} items)")
    print(f"Proposed from manual items            {man_value:>15,.2f}  ({man_p} items)")
    print(f"TOTAL PROPOSED                        {total:>15,.2f}")
    if materiality is not None:
        print(f"Materiality                           {materiality:>15,.2f}   "
              f"{'*** EXCEEDS MATERIALITY ***' if total > materiality else 'below'}")
    print()
    print(f"Excluded (already recorded / n/a)     {excluded_n} item(s)")
    print(f"Exceptions (need review)              {stats['exception_n']} item(s), "
          f"{exception_value:,.2f}")
    print(f"Prior accruals not marked reversed    {len(not_reversed_keys)}")
    print()
    for t in tests:
        print(f"Test: {'PASS' if t['passed'] else '*** FAIL ***':<14} {t['name']}")
        if t.get("note"):
            print(f"      {t['note']}")
    if esc:
        print("\nESCALATE:")
        for e in esc:
            print(f"  ! {e}")

    if not overall and not args.force:
        print("\n" + "=" * 76)
        print("WORKBOOK NOT WRITTEN. A failing gate here means the batch is not")
        print("safe to hand to a reviewer as-is.")
        print("=" * 76)
        return 1

    wb = Workbook()
    sheet_summary(wb, meta, tests, stats, esc, overall, materiality)
    sheet_batch(wb, batch, stats)
    sheet_reversal(wb, batch)
    sheet_pos(wb, po_items)
    sheet_recurring(wb, recurring_items)
    sheet_manual(wb, manual_items)
    sheet_exceptions(wb, po_items, recurring_items, manual_items)
    sheet_prior(wb, prior_accr, not_reversed_keys)
    wb.active = 0
    out = Path(args.out)
    if not overall:
        out = out.with_name(out.stem + " [GATE FAILED]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\n{'COMPLETE' if overall else 'GATE FAILED (forced)'}.  Workbook: {out}")
    print("PROPOSAL ONLY - not posted. Controller review and approval required before "
          "booking to the GL.")
    return 0 if overall else 1


if __name__ == "__main__":
    sys.exit(main())
