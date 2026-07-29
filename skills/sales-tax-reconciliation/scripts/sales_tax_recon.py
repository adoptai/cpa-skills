#!/usr/bin/env python3
"""
Sales and use tax reconciliation: filed returns <-> sales ledger <-> GL.

Six tests. Asserts NO tax rate, nexus threshold, transaction count, filing
frequency, or due date. Instead:
  * effective rates are DERIVED per jurisdiction per period as
    tax reported / taxable sales, and compared across periods, so a rate the
    client never updated is detectable without knowing the correct rate
  * sales and transaction counts are reported by jurisdiction as a SCREENING
    schedule with an instruction to confirm each state's current threshold

Marketplace-facilitated sales are tracked separately throughout, because in most
states they count toward the facilitator's obligation rather than the seller's.

Usage:
    python3 sales_tax_recon.py --sales sales.csv --returns returns.csv --gl gl.csv \
        --registrations registrations.csv --marketplace marketplace.csv \
        --reconciling-items recon_items.csv \
        --client "Ridgeline Outdoor Co" --period FY2025 \
        --out "Ridgeline - FY2025 Sales Tax Reconciliation.xlsx"

--sales CSV (per jurisdiction per period, ship-to based):
    jurisdiction, period, gross_sales, exempt_sales, taxable_sales,
    tax_collected, transaction_count, customer_count (optional)

--returns CSV (per jurisdiction per period, as filed):
    jurisdiction, period, gross_sales, exempt_sales, taxable_sales,
    tax_reported, tax_remitted, discount_taken (optional)

--gl CSV: account, description, amount
    Recognised keywords: "sales revenue", "sales tax liability beginning",
    "sales tax liability ending", "sales tax collected", "sales tax remitted"

--registrations CSV: jurisdiction, registered, filing_frequency, nexus_conclusion
--marketplace CSV:   jurisdiction, period, marketplace_sales, marketplace_transactions
--certificates CSV:  jurisdiction, customer, certificate_on_file
--reconciling-items CSV: test, cause, amount, evidence, notes
"""

from __future__ import annotations

import argparse
import csv
import sys
from collections import defaultdict
from datetime import datetime
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
RATE = '0.0000%'

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
TOP = Border(top=Side(style="thin"))
DBL = Border(top=Side(style="thin"), bottom=Side(style="double"))

TOLERANCE = ZERO


def dec(raw):
    if raw is None:
        return None
    s = str(raw).strip()
    if s in ("", "-", "--", "n/a", "N/A", "None"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace(",", "").replace("$", "").replace("%", "").strip()
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


def truthy(s) -> bool:
    return clean(s).lower() in ("yes", "y", "true", "1", "x")


def rows_of(path: Path, label: str):
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No rows in {label} ({path})")
    return rows


# --------------------------------------------------------------------- loading

def load_sales(path: Path) -> list[dict]:
    out = []
    for i, r in enumerate(rows_of(path, "sales"), start=1):
        j = clean(r.get("jurisdiction"))
        if not j:
            raise ValueError(f"Sales row {i}: jurisdiction is required.")
        out.append({
            "jurisdiction": j, "period": clean(r.get("period")) or "(period)",
            "gross_sales": d0(r.get("gross_sales")),
            "exempt_sales": d0(r.get("exempt_sales")),
            "taxable_sales": d0(r.get("taxable_sales")),
            "tax_collected": d0(r.get("tax_collected")),
            "transaction_count": int(d0(r.get("transaction_count"))),
            "customer_count": int(d0(r.get("customer_count"))),
        })
    return out


def load_returns(path: Path) -> list[dict]:
    out = []
    for i, r in enumerate(rows_of(path, "returns"), start=1):
        j = clean(r.get("jurisdiction"))
        if not j:
            raise ValueError(f"Return row {i}: jurisdiction is required.")
        out.append({
            "jurisdiction": j, "period": clean(r.get("period")) or "(period)",
            "gross_sales": d0(r.get("gross_sales")),
            "exempt_sales": d0(r.get("exempt_sales")),
            "taxable_sales": d0(r.get("taxable_sales")),
            "tax_reported": d0(r.get("tax_reported")),
            "tax_remitted": d0(r.get("tax_remitted")),
            "discount_taken": d0(r.get("discount_taken")),
        })
    return out


GL_KEYS = {
    "sales revenue": "sales_revenue",
    "sales tax liability beginning": "liab_begin",
    "sales tax liability ending": "liab_end",
    "sales tax collected": "collected",
    "sales tax remitted": "remitted",
}


def load_gl(path: Path | None) -> dict:
    out = {v: None for v in GL_KEYS.values()}
    if not path or not path.exists():
        return out
    for r in rows_of(path, "GL"):
        desc = clean(r.get("description")).lower()
        amt = dec(r.get("amount"))
        if amt is None:
            continue
        for kw, key in GL_KEYS.items():
            if kw in desc:
                out[key] = (out[key] or ZERO) + amt
                break
    return out


def load_simple(path: Path | None, label: str, key_fields: tuple) -> list[dict]:
    if not path or not path.exists():
        return []
    return [{k: clean(r.get(k)) if k in key_fields else r.get(k)
             for k in r} for r in rows_of(path, label)]


def load_items(path: Path | None) -> dict[str, list[dict]]:
    out: dict[str, list[dict]] = defaultdict(list)
    if not path or not path.exists():
        return out
    for r in rows_of(path, "reconciling items"):
        t = clean(r.get("test"))
        amt = dec(r.get("amount"))
        if not t or amt is None:
            continue
        out[t].append({"cause": clean(r.get("cause")), "amount": amt,
                       "evidence": clean(r.get("evidence")),
                       "notes": clean(r.get("notes"))})
    return out


# ----------------------------------------------------------------------- tests

def mk(num, name, parts, items, *, note=""):
    base = parts[0][1]
    diff = base
    for _, v in parts[1:]:
        diff = diff - v
    explained = sum((i["amount"] for i in items), ZERO)
    residual = diff - explained
    within = residual != ZERO and abs(residual) <= TOLERANCE
    return {"num": str(num), "name": name, "parts": parts, "difference": diff,
            "explained": explained, "residual": residual, "items": items,
            "note": note, "within_tolerance": within,
            "passed": residual == ZERO or within}


def by_jur(sales, returns, marketplace) -> dict:
    out = {}
    for s in sales:
        k = s["jurisdiction"]
        a = out.setdefault(k, blank_jur(k))
        for f in ("gross_sales", "exempt_sales", "taxable_sales", "tax_collected"):
            a["ledger"][f] += s[f]
        a["ledger"]["transaction_count"] += s["transaction_count"]
        a["periods"].add(s["period"])
    for r in returns:
        k = r["jurisdiction"]
        a = out.setdefault(k, blank_jur(k))
        for f in ("gross_sales", "exempt_sales", "taxable_sales", "tax_reported",
                  "tax_remitted", "discount_taken"):
            a["filed"][f] += r[f]
        a["filed_periods"].add(r["period"])
    for m in marketplace:
        k = clean(m.get("jurisdiction"))
        if not k:
            continue
        a = out.setdefault(k, blank_jur(k))
        a["marketplace"]["sales"] += d0(m.get("marketplace_sales"))
        a["marketplace"]["transactions"] += int(d0(m.get("marketplace_transactions")))
    return out


def blank_jur(k):
    return {
        "jurisdiction": k,
        "ledger": {"gross_sales": ZERO, "exempt_sales": ZERO, "taxable_sales": ZERO,
                   "tax_collected": ZERO, "transaction_count": 0},
        "filed": {"gross_sales": ZERO, "exempt_sales": ZERO, "taxable_sales": ZERO,
                  "tax_reported": ZERO, "tax_remitted": ZERO, "discount_taken": ZERO},
        "marketplace": {"sales": ZERO, "transactions": 0},
        "periods": set(), "filed_periods": set(),
    }


def derived_rates(returns) -> list[dict]:
    out = []
    for r in returns:
        rate = (r["tax_reported"] / r["taxable_sales"]) if r["taxable_sales"] else None
        out.append({**r, "derived_rate": rate})
    return out


def rate_drift(rates) -> list[dict]:
    per = defaultdict(list)
    for r in rates:
        if r["derived_rate"] is not None:
            per[r["jurisdiction"]].append(r)
    flags = []
    for j, rs in per.items():
        rs = sorted(rs, key=lambda x: x["period"])
        for a, b in zip(rs, rs[1:]):
            if a["derived_rate"] != b["derived_rate"]:
                flags.append({"jurisdiction": j, "from_period": a["period"],
                              "to_period": b["period"],
                              "from_rate": a["derived_rate"],
                              "to_rate": b["derived_rate"],
                              "change": b["derived_rate"] - a["derived_rate"]})
    return flags


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


def sheet_recon(wb, meta, tests, roll, overall, escalations):
    ws = wb.active
    ws.title = "Reconciliation"
    widths(ws, {1: 4, 2: 56, 3: 18, 4: 18, 5: 62})
    r = 1

    def line(label, v=None, *, bold=False, size=11, fill=None, border=None, note=""):
        nonlocal r
        c = ws.cell(row=r, column=2, value=label)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c.fill = fill
        if v is not None:
            m = ws.cell(row=r, column=3,
                        value=float(v) if isinstance(v, Decimal) else v)
            if isinstance(v, Decimal):
                m.number_format = MONEY
            m.font = Font(bold=bold)
            if border:
                m.border = border
        if note:
            n = ws.cell(row=r, column=5, value=note)
            n.font = Font(italic=True, size=9)
            n.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    line("SALES AND USE TAX RECONCILIATION", bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    v = ws.cell(row=r, column=2,
                value="RECONCILED - returns agree to the ledger and the liability "
                      "rollforward foots"
                if overall else
                "NOT RECONCILED - unexplained differences or unfiled jurisdictions remain")
    v.font = OK_FONT if overall else BAD_FONT
    if not overall:
        v.fill = BAD_FILL
    r += 2

    for t in tests:
        c = ws.cell(row=r, column=2, value=f"TEST {t['num']} - {t['name']}")
        c.font, c.fill = Font(bold=True, size=12), SUB_FILL
        r += 1
        if t.get("skipped"):
            line("  " + t["note"], fill=WARN_FILL)
            r += 1
            continue
        for label, amt in t["parts"]:
            line("  " + label, amt)
        line("  Difference", t["difference"], bold=True, border=TOP)
        if t["items"]:
            for it in t["items"]:
                line(f"    less explained: {it['cause']}", it["amount"],
                     note=f"evidence: {it['evidence'] or '(none cited)'}")
            line("  Unexplained residual", t["residual"], bold=True, border=DBL)
        rc = ws.cell(row=r - 1, column=3)
        rc.font = OK_FONT if t["passed"] else BAD_FONT
        if not t["passed"]:
            rc.fill = BAD_FILL
        elif t.get("within_tolerance"):
            rc.fill = WARN_FILL
            line(f"    absorbed by the stated tolerance of {TOLERANCE:,.2f} - "
                 f"disclosed, not plugged")
            ws.cell(row=r - 1, column=2).font = Font(italic=True, size=9)
        if t["note"]:
            line("  " + t["note"])
            ws.cell(row=r - 1, column=2).font = Font(italic=True, size=9)
        r += 1

    c = ws.cell(row=r, column=2, value="LIABILITY ROLLFORWARD")
    c.font, c.fill = Font(bold=True, size=12), SUB_FILL
    r += 1
    if roll is None:
        line("  Not performed - liability balances not supplied. Without this test, "
             "collected-but-unremitted tax is undetectable.", fill=BAD_FILL)
    else:
        for label, val, bold, border in [
            ("Beginning sales tax liability", roll["begin"], False, None),
            ("Add: tax collected", roll["collected"], False, None),
            ("Less: tax remitted", -roll["remitted"], False, None),
            ("= Computed ending liability", roll["computed"], True, TOP),
            ("Ending liability per GL", roll["end"], False, None),
            ("Difference", roll["difference"], True, DBL),
        ]:
            line("  " + label, val, bold=bold, border=border)
            if label == "Difference":
                cc = ws.cell(row=r - 1, column=3)
                cc.font = OK_FONT if val == ZERO else BAD_FONT
                if val != ZERO:
                    cc.fill = BAD_FILL
        line("  Ending liability in excess of the final period's tax",
             roll["excess"], bold=True,
             fill=BAD_FILL if roll["excess"] > ZERO else None,
             note="ESCALATE. Sales tax collected is trust-fund money. An unremitted "
                  "balance attaches personally to responsible persons in most states."
                  if roll["excess"] > ZERO else "")
    r += 1

    if escalations:
        c = ws.cell(row=r, column=2, value="ESCALATE")
        c.font, c.fill = Font(bold=True, size=12), BAD_FILL
        r += 1
        for e in escalations:
            line("  " + e)
        r += 1

    r += 1
    line("Prepared by / date:  __________________  ____________", bold=True)
    line("Reviewed by / date:  __________________  ____________", bold=True)


def sheet_by_jur(wb, jur, rates):
    ws = wb.create_sheet("By Jurisdiction")
    heads = ["Jurisdiction", "Periods with sales", "Returns filed",
             "Gross sales (ledger)", "Gross sales (returns)", "Gross diff",
             "Exempt (ledger)", "Taxable (ledger)", "Taxable (returns)",
             "Taxable diff", "Tax collected", "Tax reported", "Collected less reported",
             "Tax remitted", "Reported less remitted", "Derived rate"]
    ws.append(heads)
    hdr(ws, len(heads))
    ratemap = defaultdict(list)
    for r in rates:
        if r["derived_rate"] is not None:
            ratemap[r["jurisdiction"]].append(r["derived_rate"])
    for j in sorted(jur.values(), key=lambda x: x["jurisdiction"]):
        L, F = j["ledger"], j["filed"]
        rr = ratemap.get(j["jurisdiction"])
        avg = (sum(rr, Decimal(0)) / len(rr)) if rr else None
        ws.append([j["jurisdiction"], len(j["periods"]), len(j["filed_periods"]),
                   float(L["gross_sales"]), float(F["gross_sales"]),
                   float(L["gross_sales"] - F["gross_sales"]),
                   float(L["exempt_sales"]), float(L["taxable_sales"]),
                   float(F["taxable_sales"]),
                   float(L["taxable_sales"] - F["taxable_sales"]),
                   float(L["tax_collected"]), float(F["tax_reported"]),
                   float(L["tax_collected"] - F["tax_reported"]),
                   float(F["tax_remitted"]),
                   float(F["tax_reported"] - F["tax_remitted"]),
                   float(avg) if avg is not None else None])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in range(3, 15):
            row[i].number_format = MONEY
        row[15].number_format = RATE
        for i in (5, 9, 12, 14):
            if isinstance(row[i].value, (int, float)) and row[i].value:
                row[i].font, row[i].fill = BAD_FONT, BAD_FILL
        if row[2].value == 0 and row[1].value:
            row[2].fill = BAD_FILL
    ws.freeze_panes = "B2"
    widths(ws, {1: 20, 2: 12, 3: 12, 4: 18, 5: 18, 6: 14, 7: 15, 8: 16, 9: 16,
                10: 13, 11: 15, 12: 15, 13: 18, 14: 15, 15: 18, 16: 13})


def sheet_nexus(wb, jur, regs):
    ws = wb.create_sheet("Nexus Screening")
    heads = ["Jurisdiction", "Direct sales", "Marketplace sales", "Total sales",
             "Direct transactions", "Marketplace transactions", "Registered",
             "Returns filed", "Nexus conclusion on file", "Status", "To confirm"]
    ws.append(heads)
    hdr(ws, len(heads))
    reg = {clean(r.get("jurisdiction")): r for r in regs}
    for j in sorted(jur.values(), key=lambda x: -x["ledger"]["gross_sales"]):
        rg = reg.get(j["jurisdiction"], {})
        registered = truthy(rg.get("registered"))
        filed = len(j["filed_periods"]) > 0
        conclusion = clean(rg.get("nexus_conclusion"))
        direct = j["ledger"]["gross_sales"] - j["marketplace"]["sales"]
        if filed:
            status = "filing"
        elif conclusion:
            status = "no return - conclusion documented"
        elif j["ledger"]["gross_sales"] > ZERO:
            status = "ACTIVITY, NO RETURN, NO CONCLUSION"
        else:
            status = "no activity"
        if not filed and not registered and j["ledger"]["gross_sales"] > ZERO:
            confirm = ("Confirm this state's CURRENT economic nexus threshold, its "
                       "measurement period, whether it measures gross or taxable "
                       "sales, and whether marketplace sales count toward the "
                       "seller's threshold. Do not rely on a remembered figure.")
        elif filed and j["ledger"]["gross_sales"] == ZERO:
            confirm = ("Filing with no activity creates a permanent obligation and "
                       "penalty exposure for non-filing of zero returns. Confirm "
                       "whether the registration should be closed.")
        else:
            confirm = "Confirm current rate and filing frequency."
        ws.append([j["jurisdiction"], float(direct),
                   float(j["marketplace"]["sales"]),
                   float(j["ledger"]["gross_sales"]),
                   j["ledger"]["transaction_count"],
                   j["marketplace"]["transactions"],
                   "Y" if registered else "", "Y" if filed else "",
                   conclusion or "", status, confirm])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in (1, 2, 3):
            row[i].number_format = MONEY
        if "NO CONCLUSION" in str(row[9].value):
            row[9].font, row[9].fill = BAD_FONT, BAD_FILL
        elif "permanent obligation" in str(row[10].value):
            row[9].fill = WARN_FILL
    r = ws.max_row + 2
    for note in [
        "SCREENING SCHEDULE, NOT A NEXUS CONCLUSION.",
        "This workpaper asserts no threshold. Thresholds, measurement periods, and",
        "whether marketplace sales count toward the seller's threshold all vary by state",
        "and several have changed since being introduced.",
        "",
        "Marketplace sales are shown separately on purpose: in most states they count",
        "toward the facilitator's obligation, not the seller's. Aggregating them inflates",
        "apparent nexus and produces registrations the client does not need.",
    ]:
        ws.cell(row=r, column=1, value=note).font = Font(italic=True)
        r += 1
    ws.freeze_panes = "B2"
    widths(ws, {1: 20, 2: 16, 3: 18, 4: 16, 5: 17, 6: 21, 7: 11, 8: 12, 9: 26,
                10: 32, 11: 74})


def sheet_rollforward(wb, roll, returns):
    ws = wb.create_sheet("Liability Rollforward")
    widths(ws, {1: 4, 2: 48, 3: 18, 4: 70})
    r = 1
    ws.cell(row=r, column=2,
            value="SALES TAX LIABILITY ROLLFORWARD").font = Font(bold=True, size=14)
    r += 2
    if roll is None:
        ws.cell(row=r, column=2,
                value="Not performed - beginning and ending liability balances were not "
                      "supplied. This is the test that detects collected-but-unremitted "
                      "tax, which is trust-fund money.").fill = BAD_FILL
        return
    for label, val, bold, border in [
        ("Beginning liability", roll["begin"], False, None),
        ("Add: tax collected per GL", roll["collected"], False, None),
        ("Less: tax remitted", -roll["remitted"], False, None),
        ("= Computed ending liability", roll["computed"], True, TOP),
        ("Ending liability per GL", roll["end"], False, None),
        ("Difference", roll["difference"], True, DBL),
        ("Final period tax reported", roll["final_period_tax"], False, None),
        ("Ending liability in excess of final period", roll["excess"], True, TOP),
    ]:
        ws.cell(row=r, column=2, value=label).font = Font(bold=bold)
        c = ws.cell(row=r, column=3, value=float(val))
        c.number_format, c.font = MONEY, Font(bold=bold)
        if border:
            c.border = border
        if label == "Difference":
            c.font = OK_FONT if val == ZERO else BAD_FONT
            if val != ZERO:
                c.fill = BAD_FILL
        if label.startswith("Ending liability in excess"):
            if val > ZERO:
                c.font, c.fill = BAD_FONT, BAD_FILL
                ws.cell(row=r, column=4,
                        value="ESCALATE IMMEDIATELY. Tax collected from customers and not "
                              "remitted is trust-fund money. In most states an unremitted "
                              "balance attaches personally to responsible persons. Days "
                              "matter."
                        ).alignment = Alignment(wrap_text=True, vertical="top")
            else:
                c.font = OK_FONT
        r += 1
    r += 1
    ws.cell(row=r, column=2, value="By period").font = Font(bold=True, size=12)
    r += 1
    for h, col in zip(["Period", "Tax reported", "Tax remitted", "Shortfall"],
                      (2, 3, 4, 5)):
        c = ws.cell(row=r, column=col, value=h)
        c.fill, c.font = HDR_FILL, HDR_FONT
    r += 1
    per = defaultdict(lambda: {"rep": ZERO, "rem": ZERO})
    for x in returns:
        per[x["period"]]["rep"] += x["tax_reported"]
        per[x["period"]]["rem"] += x["tax_remitted"]
    for p, a in sorted(per.items()):
        ws.cell(row=r, column=2, value=p)
        for col, v in ((3, a["rep"]), (4, a["rem"]), (5, a["rep"] - a["rem"])):
            c = ws.cell(row=r, column=col, value=float(v))
            c.number_format = MONEY
        if a["rep"] - a["rem"] != ZERO:
            ws.cell(row=r, column=5).font = BAD_FONT
        r += 1


def sheet_rates(wb, rates, drift):
    ws = wb.create_sheet("Rate Analysis")
    heads = ["Jurisdiction", "Period", "Taxable sales", "Tax reported",
             "Derived effective rate"]
    ws.append(heads)
    hdr(ws, len(heads))
    for x in sorted(rates, key=lambda r: (r["jurisdiction"], r["period"])):
        ws.append([x["jurisdiction"], x["period"], float(x["taxable_sales"]),
                   float(x["tax_reported"]),
                   float(x["derived_rate"]) if x["derived_rate"] is not None else None])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[2].number_format = MONEY
        row[3].number_format = MONEY
        row[4].number_format = RATE
    r = ws.max_row + 2
    if drift:
        c = ws.cell(row=r, column=1, value="RATE MOVEMENT BETWEEN PERIODS")
        c.font = Font(bold=True, size=12)
        r += 1
        for h, col in zip(["Jurisdiction", "From period", "To period", "From rate",
                           "To rate", "Change"], range(1, 7)):
            cc = ws.cell(row=r, column=col, value=h)
            cc.fill, cc.font = HDR_FILL, HDR_FONT
        r += 1
        for d in drift:
            ws.cell(row=r, column=1, value=d["jurisdiction"])
            ws.cell(row=r, column=2, value=d["from_period"])
            ws.cell(row=r, column=3, value=d["to_period"])
            for col, v in ((4, d["from_rate"]), (5, d["to_rate"]), (6, d["change"])):
                cc = ws.cell(row=r, column=col, value=float(v))
                cc.number_format = RATE
            ws.cell(row=r, column=6).fill = WARN_FILL
            r += 1
        r += 1
    else:
        ws.cell(row=r, column=1,
                value="No rate movement between periods in any jurisdiction.").font = OK_FONT
        r += 2
    for note in [
        "Rates here are DERIVED as tax reported / taxable sales from the filed returns.",
        "This workpaper asserts no statutory rate. A rate that moved without a stated",
        "cause is a finding regardless of what the correct rate is - and a rate that did",
        "NOT move in a jurisdiction that changed its rate is the more common error.",
    ]:
        ws.cell(row=r, column=1, value=note).font = Font(italic=True)
        r += 1
    ws.freeze_panes = "C2"
    widths(ws, {1: 20, 2: 14, 3: 18, 4: 16, 5: 20, 6: 14})


def sheet_exempt(wb, jur, certs, rates):
    ws = wb.create_sheet("Exempt Sales")
    heads = ["Jurisdiction", "Exempt sales", "Derived rate", "Exposure if no certificate",
             "Certificates on file", "Status"]
    ws.append(heads)
    hdr(ws, len(heads))
    ratemap = defaultdict(list)
    for r in rates:
        if r["derived_rate"] is not None:
            ratemap[r["jurisdiction"]].append(r["derived_rate"])
    cert_count = defaultdict(int)
    for c in certs:
        if truthy(c.get("certificate_on_file")):
            cert_count[clean(c.get("jurisdiction"))] += 1
    total_exposure = ZERO
    for j in sorted(jur.values(), key=lambda x: -x["ledger"]["exempt_sales"]):
        ex = j["ledger"]["exempt_sales"]
        if ex == ZERO:
            continue
        rr = ratemap.get(j["jurisdiction"])
        rate = (sum(rr, Decimal(0)) / len(rr)) if rr else None
        exposure = (ex * rate) if rate is not None else None
        n = cert_count.get(j["jurisdiction"], 0)
        status = ("certificates on file" if n else
                  "NO CERTIFICATE REGISTER - exempt sales unsupported")
        if exposure is not None and not n:
            total_exposure += exposure
        ws.append([j["jurisdiction"], float(ex),
                   float(rate) if rate is not None else None,
                   float(exposure) if exposure is not None else None,
                   n or "", status])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[1].number_format = MONEY
        row[2].number_format = RATE
        row[3].number_format = MONEY
        if "NO CERTIFICATE" in str(row[5].value):
            row[5].font, row[5].fill = BAD_FONT, BAD_FILL
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value="No exempt sales recorded.")
    else:
        r = ws.max_row + 2
        ws.cell(row=r, column=1, value="TOTAL UNSUPPORTED EXPOSURE").font = Font(bold=True)
        c = ws.cell(row=r, column=4, value=float(total_exposure))
        c.number_format, c.font, c.border = MONEY, Font(bold=True), TOP
        r += 2
        for note in [
            "In an audit, an exempt sale without a valid certificate is a taxable sale.",
            "The tax is assessed against the seller, who can rarely collect it from the",
            "customer years later. Exposure above is exempt sales at the derived rate.",
        ]:
            ws.cell(row=r, column=1, value=note).font = Font(italic=True)
            r += 1
    widths(ws, {1: 20, 2: 18, 3: 13, 4: 24, 5: 20, 6: 44})


def sheet_memo(wb, tests, escalations, meta):
    ws = wb.create_sheet("Exceptions & Memo")
    widths(ws, {1: 4, 2: 108})
    r = 1

    def para(t, bold=False, size=11, fill=None):
        nonlocal r
        c = ws.cell(row=r, column=2, value=t)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c.fill = fill
        r += 1

    para("SALES AND USE TAX MEMO", bold=True, size=14)
    r += 1
    para(f"{meta.get('Client', '')}    Period: {meta.get('Period', '')}", bold=True)
    r += 1

    unexp = [t for t in tests if not t.get("skipped") and not t["passed"]]
    if unexp:
        para("UNEXPLAINED DIFFERENCES", bold=True, size=12, fill=BAD_FILL)
        for t in unexp:
            para(f"  Test {t['num']} - {t['name']}: unexplained {t['residual']:,.2f}")
        r += 1

    if escalations:
        para("ESCALATIONS", bold=True, size=12, fill=BAD_FILL)
        for e in escalations:
            para("  " + e)
        r += 1

    para("FIGURES REQUIRING CONFIRMATION AGAINST CURRENT AUTHORITY", bold=True, size=12)
    for t in [
        "Each jurisdiction's current tax rate, including local and district rates",
        "Each jurisdiction's current economic nexus threshold, its measurement period, "
        "and whether it measures gross or taxable sales",
        "Whether marketplace-facilitated sales count toward the seller's threshold in "
        "each state",
        "Filing frequency assigned in each jurisdiction, and current due dates",
        "Sourcing rules - destination versus origin - for each jurisdiction",
        "Availability and terms of any voluntary disclosure programme, which generally "
        "requires approaching the state BEFORE being contacted",
    ]:
        para("  - " + t)
    r += 1

    para("USE TAX", bold=True, size=12)
    para("  Confirm whether use tax has been self-assessed on purchases where no sales "
         "tax was charged - fixed assets, out-of-state vendors, software, and services. "
         "Clients routinely miss this and auditors look for it first because it is easy "
         "to find.")
    r += 1
    para("CONTROL OBSERVATIONS TO CONSIDER", bold=True, size=12)
    for t in [
        "Whether the exemption certificate register is complete and certificates are "
        "current",
        "Whether new jurisdictions are reviewed for nexus as sales grow",
        "Whether the tax engine or rate table is updated on a defined schedule",
        "Whether the liability account is reconciled and cleared monthly",
        "Whether marketplace and direct sales are separated in the ledger",
    ]:
        para("  - " + t)
    r += 1
    para("Whether to file a voluntary disclosure, amend a return, register, or close a "
         "registration is a decision for the person signing. This workpaper reports "
         "facts and quantifies exposure.", bold=True)


# ------------------------------------------------------------------------ main

def main() -> int:
    global TOLERANCE
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--sales", required=True)
    ap.add_argument("--returns", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--gl")
    ap.add_argument("--registrations")
    ap.add_argument("--marketplace")
    ap.add_argument("--certificates")
    ap.add_argument("--reconciling-items")
    ap.add_argument("--tolerance", default="0.00",
                    help="absolute per-test tolerance (default 0.00). Anything "
                         "absorbed is reported separately.")
    ap.add_argument("--client", default="")
    ap.add_argument("--period", default="")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()
    TOLERANCE = dec(args.tolerance) or ZERO

    sales = load_sales(Path(args.sales))
    returns = load_returns(Path(args.returns))
    gl = load_gl(Path(args.gl) if args.gl else None)
    regs = load_simple(Path(args.registrations) if args.registrations else None,
                       "registrations", ("jurisdiction",))
    mkt = load_simple(Path(args.marketplace) if args.marketplace else None,
                      "marketplace", ("jurisdiction",))
    certs = load_simple(Path(args.certificates) if args.certificates else None,
                        "certificates", ("jurisdiction", "customer"))
    items = load_items(Path(args.reconciling_items) if args.reconciling_items else None)

    jur = by_jur(sales, returns, mkt)
    rates = derived_rates(returns)
    drift = rate_drift(rates)

    led_gross = sum((s["gross_sales"] for s in sales), ZERO)
    ret_gross = sum((r["gross_sales"] for r in returns), ZERO)
    led_taxable = sum((s["taxable_sales"] for s in sales), ZERO)
    ret_taxable = sum((r["taxable_sales"] for r in returns), ZERO)
    led_collected = sum((s["tax_collected"] for s in sales), ZERO)
    ret_reported = sum((r["tax_reported"] for r in returns), ZERO)
    ret_remitted = sum((r["tax_remitted"] for r in returns), ZERO)

    tests = []
    g = lambda n: items.get(str(n), [])

    if gl.get("sales_revenue") is not None:
        tests.append(mk(1, "GL sales revenue bridged to gross sales per returns",
                        [("GL sales revenue", gl["sales_revenue"]),
                         ("Gross sales per all returns", ret_gross)], g(1),
                        note="Differences here are usually basis (accrual GL against "
                             "cash-basis returns), freight, marketplace sales, "
                             "intercompany, or discounts and returns."))
    else:
        tests.append({"num": "1", "name": "GL sales revenue bridge", "skipped": True,
                      "passed": True, "parts": [], "difference": ZERO,
                      "explained": ZERO, "residual": ZERO, "items": [],
                      "note": "NOT PERFORMED - GL sales revenue not supplied"})

    tests.append(mk(2, "Gross less exempt equals taxable, ledger to returns",
                    [("Taxable sales per ledger", led_taxable),
                     ("Taxable sales per returns", ret_taxable)], g(2)))
    tests.append(mk(3, "Tax collected per ledger equals tax reported per returns",
                    [("Tax collected per sales ledger", led_collected),
                     ("Tax reported per returns", ret_reported)], g(3),
                    note="A shortfall means tax was collected and not reported. An "
                         "excess means tax was reported that was never collected - the "
                         "client funded it out of margin."))
    tests.append(mk(4, "Tax reported equals tax remitted",
                    [("Tax reported per returns", ret_reported),
                     ("Tax remitted per returns", ret_remitted)], g(4)))

    unfiled = [j for j in jur.values()
               if j["ledger"]["gross_sales"] > ZERO and not j["filed_periods"]
               and not clean((next((r for r in regs
                                    if clean(r.get("jurisdiction")) == j["jurisdiction"]),
                                   {}) or {}).get("nexus_conclusion"))]
    filed_no_activity = [j for j in jur.values()
                         if j["filed_periods"] and j["ledger"]["gross_sales"] == ZERO]
    tests.append({"num": "5",
                  "name": "Every jurisdiction with activity has a return or a "
                          "documented nexus conclusion",
                  "parts": [("Jurisdictions with sales", Decimal(len(
                      [j for j in jur.values() if j["ledger"]["gross_sales"] > ZERO]))),
                            ("Jurisdictions with activity, no return, no conclusion",
                             Decimal(len(unfiled)))],
                  "difference": Decimal(len(unfiled)), "explained": ZERO,
                  "residual": Decimal(len(unfiled)), "items": [],
                  "passed": not unfiled, "is_count": True,
                  "note": "" if not unfiled else
                          "Where no return was ever filed the lookback is typically open "
                          "indefinitely. This is not a filing error, it is an "
                          "unquantified liability."})

    tests.append({"num": "6", "name": "Derived rate consistency across periods",
                  "parts": [("Jurisdictions with rate movement",
                             Decimal(len({d['jurisdiction'] for d in drift})))],
                  "difference": Decimal(len(drift)), "explained": ZERO,
                  "residual": Decimal(len(drift)), "items": [],
                  "passed": not drift, "is_count": True,
                  "note": "Rate movement is not necessarily an error - rates change. "
                          "Each movement needs a stated cause. A rate that did NOT move "
                          "in a jurisdiction that changed its rate is the more common "
                          "problem and will not appear here."})

    roll = None
    if gl.get("liab_begin") is not None and gl.get("liab_end") is not None:
        collected = gl.get("collected")
        if collected is None:
            collected = led_collected
        remitted = gl.get("remitted")
        if remitted is None:
            remitted = ret_remitted
        computed = gl["liab_begin"] + collected - remitted
        periods = sorted({r["period"] for r in returns})
        final_tax = sum((r["tax_reported"] for r in returns
                         if r["period"] == periods[-1]), ZERO) if periods else ZERO
        roll = {"begin": gl["liab_begin"], "collected": collected,
                "remitted": remitted, "computed": computed, "end": gl["liab_end"],
                "difference": computed - gl["liab_end"],
                "final_period_tax": final_tax,
                "excess": gl["liab_end"] - final_tax}

    escalations = []
    if roll and roll["excess"] > ZERO:
        escalations.append(
            f"Sales tax liability exceeds the final period's tax by "
            f"{roll['excess']:,.2f}. Tax collected from customers and not remitted is "
            f"trust-fund money and attaches personally to responsible persons in most "
            f"states. Escalate immediately.")
    if roll and roll["difference"] != ZERO:
        escalations.append(
            f"Liability rollforward does not foot by {roll['difference']:,.2f} - the "
            f"liability account contains activity that is neither tax collected nor a "
            f"remittance.")
    for j in unfiled:
        escalations.append(
            f"{j['jurisdiction']}: {j['ledger']['gross_sales']:,.2f} of sales, "
            f"{j['ledger']['transaction_count']} transactions, no return filed and no "
            f"nexus conclusion on file. Confirm the current threshold; if a filing "
            f"obligation exists, note that voluntary disclosure generally requires "
            f"approaching the state before being contacted.")
    for j in filed_no_activity:
        escalations.append(
            f"{j['jurisdiction']}: returns filed with no sales activity. A live "
            f"registration creates a permanent filing obligation and penalty exposure "
            f"for non-filing of zero returns.")
    shortfall = ret_reported - ret_remitted
    if shortfall > ZERO:
        escalations.append(
            f"Tax reported exceeds tax remitted by {shortfall:,.2f} across all "
            f"jurisdictions.")
    mkt_in_returns = [j for j in jur.values()
                      if j["marketplace"]["sales"] > ZERO and j["filed_periods"]
                      and j["filed"]["gross_sales"] >= j["ledger"]["gross_sales"]]
    if mkt_in_returns:
        escalations.append(
            f"{len(mkt_in_returns)} jurisdiction(s) have marketplace-facilitated sales "
            f"and gross sales per the return at or above total ledger sales - the "
            f"marketplace sales may have been reported on the seller's return as well "
            f"as by the facilitator, which overstates the liability and may have been "
            f"remitted twice.")
    if not certs:
        ex_total = sum((j["ledger"]["exempt_sales"] for j in jur.values()), ZERO)
        if ex_total > ZERO:
            escalations.append(
                f"{ex_total:,.2f} of exempt sales with no certificate register supplied. "
                f"In an audit an exempt sale without a valid certificate is a taxable "
                f"sale, assessed against the seller.")

    unexplained = [t for t in tests if not t.get("skipped") and not t["passed"]]
    absorbed = [t for t in tests if t.get("within_tolerance")]
    overall = not unexplained and (roll is None or roll["difference"] == ZERO)

    meta = {
        "Client": args.client or "(not stated)",
        "Period": args.period or "(not stated)",
        "Jurisdictions with sales": len([j for j in jur.values()
                                         if j["ledger"]["gross_sales"] > ZERO]),
        "Jurisdictions with returns filed": len([j for j in jur.values()
                                                 if j["filed_periods"]]),
        "Return periods": len({r["period"] for r in returns}),
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    # ---- console
    print("=" * 76)
    print("SALES AND USE TAX RECONCILIATION")
    print("=" * 76)
    print(f"Client : {meta['Client']}    Period: {meta['Period']}")
    print(f"Jurisdictions with sales {meta['Jurisdictions with sales']}   "
          f"with returns filed {meta['Jurisdictions with returns filed']}")
    print()
    for t in tests:
        if t.get("skipped"):
            print(f"Test {t['num']}: NOT PERFORMED - {t['note']}")
            continue
        status = ("PASS" if t["residual"] == ZERO
                  else "PASS (tolerance)" if t.get("within_tolerance")
                  else "*** FAIL ***")
        if t.get("is_count"):
            print(f"Test {t['num']}: {status:<17} {int(t['difference']):>6} item(s)   "
                  f"{t['name']}")
        else:
            print(f"Test {t['num']}: {status:<17} diff {t['difference']:>15,.2f}  "
                  f"unexplained {t['residual']:>14,.2f}")
            print(f"         {t['name']}")
    if roll:
        print()
        print(f"Liability: begin {roll['begin']:>13,.2f} + collected "
              f"{roll['collected']:>13,.2f} - remitted {roll['remitted']:>13,.2f} "
              f"= {roll['computed']:>13,.2f}")
        print(f"           ending per GL {roll['end']:>13,.2f}   difference "
              f"{roll['difference']:>13,.2f}")
        print(f"           excess over final period {roll['excess']:>13,.2f}")
    else:
        print("\nLiability rollforward: NOT PERFORMED - balances not supplied")
    if absorbed:
        print(f"\nABSORBED BY THE STATED TOLERANCE ({TOLERANCE:,.2f}):")
        for t in absorbed:
            print(f"  Test {t['num']}: {t['residual']:,.2f} - disclosed on the workpaper")
    if escalations:
        print("\nESCALATE:")
        for e in escalations:
            print(f"  ! {e}")

    if not overall and not args.force:
        print("\n" + "=" * 76)
        print("WORKBOOK NOT WRITTEN.")
        print("Classify each difference in --reconciling-items, and obtain a documented")
        print("nexus conclusion for every jurisdiction with activity and no return.")
        print("Never plug the sales tax liability account - it holds trust-fund money.")
        print("=" * 76)
        return 1

    wb = Workbook()
    sheet_recon(wb, meta, tests, roll, overall, escalations)
    sheet_by_jur(wb, jur, rates)
    sheet_nexus(wb, jur, regs)
    sheet_rollforward(wb, roll, returns)
    sheet_rates(wb, rates, drift)
    sheet_exempt(wb, jur, certs, rates)
    sheet_memo(wb, tests, escalations, meta)
    wb.active = 0
    out = Path(args.out)
    if not overall:
        out = out.with_name(out.stem + " [NOT RECONCILED]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\n{'RECONCILED' if overall else 'NOT RECONCILED (forced)'}.  Workbook: {out}")
    return 0 if overall else 1


if __name__ == "__main__":
    sys.exit(main())
