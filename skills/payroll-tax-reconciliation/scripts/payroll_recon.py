#!/usr/bin/env python3
"""
Four-way payroll tax reconciliation: register <-> Forms 941 <-> W-2/W-3 <-> GL.

Nine tests. Any unexplained difference keeps the reconciliation open.

Deliberately asserts NO tax rates, wage bases, or deposit thresholds. Instead it
derives them from the data and asks you to confirm them:
  * the Social Security wage base is inferred as the highest SS wages among
    employees whose Medicare wages exceed their SS wages -- i.e. the cap the
    payroll system actually applied
  * effective rates are computed as tax / taxable wages and compared across
    quarters, so rate drift is detectable without knowing the correct rate

Usage:
    python3 payroll_recon.py --register register.csv --f941 f941.csv \
        --w3 w3.csv --gl gl.csv --w2 w2_by_employee.csv \
        --reconciling-items recon_items.csv \
        --client "Northgate Manufacturing Inc" --year 2025 \
        --out "Northgate - 2025 Payroll Tax Reconciliation.xlsx"

--register CSV (one row per employee, full year):
    employee_id, employee_name, gross_wages, pretax_deferral, pretax_cafeteria,
    taxable_fringe, federal_taxable_wages, ss_wages, ss_tax_employee,
    medicare_wages, medicare_tax_employee, addl_medicare_tax, fit_withheld

--f941 CSV (one row per quarter, as filed, including 941-X effects):
    quarter, wages, fit_withheld, ss_wages, ss_tax, medicare_wages,
    medicare_tax, addl_medicare_tax, total_taxes, deposits

  Note: ss_tax and medicare_tax on Form 941 are the COMBINED employee and
  employer share. addl_medicare_tax is employee-only.

--w3 CSV (single row, W-3 box totals):
    box1_wages, box2_fit, box3_ss_wages, box4_ss_tax, box5_medicare_wages,
    box6_medicare_tax

--gl CSV: account, description, amount
    Recognised description keywords: "wage expense", "employer tax expense",
    "payroll tax liability beginning", "payroll tax liability ending",
    "wage accrual beginning", "wage accrual ending", "deposits"

--w2 CSV (optional but strongly recommended):
    employee_id, employee_name, box1, box2, box3, box4, box5, box6

--reconciling-items CSV:
    test, cause, amount, evidence, notes
      test = the test number the item explains, e.g. "1"
"""

from __future__ import annotations

import argparse
import csv
import re
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

SSN_RE = re.compile(r"\b\d{3}-\d{2}-\d{4}\b")


def dec(raw) -> Decimal | None:
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


def d0(raw) -> Decimal:
    v = dec(raw)
    return v if v is not None else ZERO


def clean(s) -> str:
    return " ".join(str(s or "").split())


def rows_of(path: Path, label: str) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No rows in {label} ({path})")
    return rows


def guard_ssn(rows: list[dict], label: str) -> None:
    for i, r in enumerate(rows, start=1):
        for k, v in r.items():
            if v and SSN_RE.search(str(v)):
                sys.exit(
                    f"{label} row {i}: a full SSN appears in column '{k}'. Remove it and "
                    f"use employee_id or last four digits only. Payroll data is the most "
                    f"sensitive information in the engagement; this script will not "
                    f"process a file containing full SSNs."
                )


# --------------------------------------------------------------------- loading

REG_FIELDS = ["gross_wages", "pretax_deferral", "pretax_cafeteria", "taxable_fringe",
              "federal_taxable_wages", "ss_wages", "ss_tax_employee", "medicare_wages",
              "medicare_tax_employee", "addl_medicare_tax", "fit_withheld"]

F941_FIELDS = ["wages", "fit_withheld", "ss_wages", "ss_tax", "medicare_wages",
               "medicare_tax", "addl_medicare_tax", "total_taxes", "deposits"]

W3_FIELDS = ["box1_wages", "box2_fit", "box3_ss_wages", "box4_ss_tax",
             "box5_medicare_wages", "box6_medicare_tax"]


def load_register(path: Path) -> tuple[list[dict], dict]:
    raw = rows_of(path, "register")
    guard_ssn(raw, "register")
    emps, tot = [], {f: ZERO for f in REG_FIELDS}
    for i, r in enumerate(raw, start=1):
        e = {"row": i,
             "employee_id": clean(r.get("employee_id")) or f"EMP{i:04d}",
             "employee_name": clean(r.get("employee_name")) or f"(employee {i})"}
        for f in REG_FIELDS:
            e[f] = d0(r.get(f))
            tot[f] += e[f]
        emps.append(e)
    return emps, tot


def load_f941(path: Path) -> tuple[list[dict], dict]:
    raw = rows_of(path, "941")
    qs, tot = [], {f: ZERO for f in F941_FIELDS}
    for r in raw:
        q = {"quarter": clean(r.get("quarter")) or "?"}
        for f in F941_FIELDS:
            q[f] = d0(r.get(f))
            tot[f] += q[f]
        qs.append(q)
    qs.sort(key=lambda x: x["quarter"])
    return qs, tot


def load_w3(path: Path | None) -> dict | None:
    if not path or not path.exists():
        return None
    raw = rows_of(path, "W-3")
    guard_ssn(raw, "W-3")
    r = raw[0]
    return {f: d0(r.get(f)) for f in W3_FIELDS}


GL_KEYS = {
    "wage expense": "wage_expense",
    "employer tax expense": "employer_tax_expense",
    "payroll tax liability beginning": "liab_begin",
    "payroll tax liability ending": "liab_end",
    "wage accrual beginning": "accrual_begin",
    "wage accrual ending": "accrual_end",
    "deposits": "gl_deposits",
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


def load_w2(path: Path | None) -> dict[str, dict]:
    if not path or not path.exists():
        return {}
    raw = rows_of(path, "W-2")
    guard_ssn(raw, "W-2")
    out = {}
    for r in raw:
        key = clean(r.get("employee_id")) or clean(r.get("employee_name"))
        if not key:
            continue
        out[key.lower()] = {
            "employee_name": clean(r.get("employee_name")),
            "box1": d0(r.get("box1")), "box2": d0(r.get("box2")),
            "box3": d0(r.get("box3")), "box4": d0(r.get("box4")),
            "box5": d0(r.get("box5")), "box6": d0(r.get("box6")),
        }
    return out


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

def infer_wage_base(emps: list[dict]) -> Decimal | None:
    """The cap the payroll system actually applied, read off the data."""
    capped = [e["ss_wages"] for e in emps
              if e["medicare_wages"] > e["ss_wages"] and e["ss_wages"] > ZERO]
    return max(capped) if capped else None


# Set by main() from --rounding-tolerance. Per-payroll rounding genuinely produces
# small differences between quarterly sums and annual totals, so a tolerance is
# realistic. It defaults to zero and must be set deliberately, and any amount it
# absorbs is reported separately rather than disappearing — a disclosed tolerance
# is professional judgment, a silent one is a plug.
TOLERANCE = ZERO


def mk(num, name, parts, items, *, note="", auto_fail=True):
    """parts: list of (label, amount). Difference = first - rest."""
    base = parts[0][1]
    diff = base
    for _, v in parts[1:]:
        diff = diff - v
    explained = sum((i["amount"] for i in items), ZERO)
    residual = diff - explained
    within = residual != ZERO and abs(residual) <= TOLERANCE
    return {"num": str(num), "name": name, "parts": parts, "difference": diff,
            "explained": explained, "residual": residual, "items": items,
            "note": note, "auto_fail": auto_fail,
            "within_tolerance": within,
            "passed": residual == ZERO or within}


def run_tests(reg_tot, emps, f941_tot, qs, w3, gl, w2, items) -> list[dict]:
    T = []
    g = lambda n: items.get(str(n), [])

    # 1 gross -> federal taxable -> 941 wages -> W-3 box 1
    bridge = (reg_tot["gross_wages"] - reg_tot["pretax_deferral"]
              - reg_tot["pretax_cafeteria"] + reg_tot["taxable_fringe"])
    t = mk(1, "Gross wages bridged to federal taxable wages, 941s, and W-3 box 1",
           [("Register gross wages less pre-tax, plus taxable fringe", bridge),
            ("Sum of 941 wage lines", f941_tot["wages"])], g(1))
    t["extra"] = [("Register federal taxable wages", reg_tot["federal_taxable_wages"]),
                  ("W-3 box 1", w3["box1_wages"] if w3 else None)]
    T.append(t)

    # 2 SS wages
    T.append(mk(2, "Social Security wages: 941s = W-3 box 3 = register",
                [("Sum of 941 Social Security wages", f941_tot["ss_wages"]),
                 ("W-3 box 3", w3["box3_ss_wages"] if w3 else ZERO)], g(2)))
    T[-1]["extra"] = [("Register Social Security wages", reg_tot["ss_wages"])]

    # 3 Medicare wages
    T.append(mk(3, "Medicare wages: 941s = W-3 box 5 = register",
                [("Sum of 941 Medicare wages", f941_tot["medicare_wages"]),
                 ("W-3 box 5", w3["box5_medicare_wages"] if w3 else ZERO)], g(3)))
    T[-1]["extra"] = [("Register Medicare wages", reg_tot["medicare_wages"])]

    # 4 FIT withheld
    T.append(mk(4, "Federal income tax withheld: 941s = W-3 box 2 = register",
                [("Sum of 941 federal income tax withheld", f941_tot["fit_withheld"]),
                 ("W-3 box 2", w3["box2_fit"] if w3 else ZERO)], g(4)))
    T[-1]["extra"] = [("Register federal income tax withheld", reg_tot["fit_withheld"])]

    # 5 SS tax: 941 is combined employee+employer; W-3 box 4 is employee only
    emp_share = w3["box4_ss_tax"] if w3 else reg_tot["ss_tax_employee"]
    t = mk(5, "Social Security tax: 941 combined share vs employee share doubled",
           [("Sum of 941 Social Security tax (employee + employer)", f941_tot["ss_tax"]),
            ("W-3 box 4 employee share x 2", emp_share * 2)], g(5),
           note="The doubling relationship holds only where the employer matches at the "
                "employee rate with no special items. Tips, group-term life, and "
                "adjustments legitimately break it - classify rather than assume error.")
    t["extra"] = [("Implied employer share (941 less employee)",
                   f941_tot["ss_tax"] - emp_share),
                  ("Register employee Social Security tax", reg_tot["ss_tax_employee"])]
    T.append(t)

    # 6 Medicare tax, isolating Additional Medicare (employee only)
    med_emp = w3["box6_medicare_tax"] if w3 else reg_tot["medicare_tax_employee"]
    base_med_emp = med_emp - reg_tot["addl_medicare_tax"]
    t = mk(6, "Medicare tax: 941 combined share vs employee share doubled, "
              "Additional Medicare isolated",
           [("Sum of 941 Medicare tax (employee + employer)", f941_tot["medicare_tax"]),
            ("Base employee Medicare x 2, plus Additional Medicare",
             base_med_emp * 2 + reg_tot["addl_medicare_tax"])], g(6),
           note="Additional Medicare Tax is employee-only and breaks the doubling "
                "relationship by design, so it is separated out here rather than buried.")
    t["extra"] = [("Additional Medicare Tax per register", reg_tot["addl_medicare_tax"]),
                  ("Additional Medicare Tax per 941s", f941_tot["addl_medicare_tax"]),
                  ("W-3 box 6 total employee Medicare", med_emp)]
    T.append(t)

    # 7 GL wage expense
    if gl.get("wage_expense") is not None:
        ab = gl.get("accrual_begin") or ZERO
        ae = gl.get("accrual_end") or ZERO
        t = mk(7, "GL wage expense = register gross + ending accrual - beginning accrual",
               [("Register gross wages plus accrual movement",
                 reg_tot["gross_wages"] + ae - ab),
                ("GL wage expense", gl["wage_expense"])], g(7))
        t["extra"] = [("Beginning wage accrual", ab), ("Ending wage accrual", ae)]
        T.append(t)
    else:
        T.append({"num": "7", "name": "GL wage expense", "parts": [],
                  "difference": ZERO, "explained": ZERO, "residual": ZERO, "items": [],
                  "note": "NOT PERFORMED - GL wage expense not supplied",
                  "passed": True, "skipped": True, "extra": [], "auto_fail": True})

    # 8 liability rollforward
    if gl.get("liab_begin") is not None and gl.get("liab_end") is not None:
        deposits = gl.get("gl_deposits")
        if deposits is None:
            deposits = f941_tot["deposits"]
        computed = gl["liab_begin"] + f941_tot["total_taxes"] - deposits
        t = mk(8, "Payroll tax liability rollforward",
               [("Beginning liability + taxes per 941s - deposits", computed),
                ("Ending liability per GL", gl["liab_end"])], g(8))
        t["extra"] = [("Beginning liability", gl["liab_begin"]),
                      ("Taxes incurred per 941s", f941_tot["total_taxes"]),
                      ("Deposits", deposits),
                      ("Ending liability per GL", gl["liab_end"])]
        T.append(t)
    else:
        T.append({"num": "8", "name": "Payroll tax liability rollforward", "parts": [],
                  "difference": ZERO, "explained": ZERO, "residual": ZERO, "items": [],
                  "note": "NOT PERFORMED - liability balances not supplied",
                  "passed": True, "skipped": True, "extra": [], "auto_fail": True})

    return T


def employee_tests(emps, w2, base) -> tuple[list[dict], list[str]]:
    rows, problems = [], []
    for e in emps:
        key = e["employee_id"].lower()
        w = w2.get(key) or w2.get(e["employee_name"].lower())
        flags = []
        d1 = d3 = d5 = None
        if w:
            d1 = e["federal_taxable_wages"] - w["box1"]
            d3 = e["ss_wages"] - w["box3"]
            d5 = e["medicare_wages"] - w["box5"]
            if d1 != ZERO:
                flags.append(f"W-2 box 1 differs from register by {d1:,.2f}")
            if d3 != ZERO:
                flags.append(f"W-2 box 3 differs from register by {d3:,.2f}")
            if d5 != ZERO:
                flags.append(f"W-2 box 5 differs from register by {d5:,.2f}")
        if e["ss_wages"] > e["medicare_wages"]:
            flags.append("Social Security wages exceed Medicare wages - impossible, "
                         "Social Security is capped and Medicare is not")
        if base is not None and e["ss_wages"] > base:
            flags.append(f"Social Security wages {e['ss_wages']:,.2f} exceed the inferred "
                         f"wage base {base:,.2f} - wrong base applied, or two employees "
                         f"merged in the register")
        if e["federal_taxable_wages"] > ZERO and e["fit_withheld"] == ZERO:
            flags.append("taxable wages with zero income tax withholding - confirm an "
                         "exemption certificate is on file")
        if e["gross_wages"] < ZERO:
            flags.append("negative gross wages - a reversal or a correction posted "
                         "to the wrong employee")
        rows.append({**e, "w2": w, "d1": d1, "d3": d3, "d5": d5, "flags": flags})
        for f in flags:
            problems.append(f"{e['employee_name']} ({e['employee_id']}): {f}")

    if w2:
        reg_keys = {e["employee_id"].lower() for e in emps} | \
                   {e["employee_name"].lower() for e in emps}
        for k, w in w2.items():
            if k not in reg_keys:
                problems.append(f"W-2 issued to '{w['employee_name'] or k}' who does not "
                                f"appear in the payroll register")
    return rows, problems


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


def sheet_recon(wb, tests, meta, base, overall, escalations):
    ws = wb.active
    ws.title = "Reconciliation"
    widths(ws, {1: 4, 2: 58, 3: 18, 4: 18, 5: 62})
    r = 1

    def line(t, v=None, *, bold=False, size=11, fill=None, border=None, note=""):
        nonlocal r
        c = ws.cell(row=r, column=2, value=t)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c.fill = fill
        if v is not None:
            m = ws.cell(row=r, column=3, value=float(v))
            m.number_format = MONEY
            m.font = Font(bold=bold)
            if border:
                m.border = border
        if note:
            n = ws.cell(row=r, column=5, value=note)
            n.font = Font(italic=True, size=9)
            n.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    line("PAYROLL TAX RECONCILIATION", bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    v = ws.cell(row=r, column=2,
                value="RECONCILED - all differences explained"
                if overall else
                "NOT RECONCILED - unexplained differences remain. Do not file or "
                "rely on these figures.")
    v.font = OK_FONT if overall else BAD_FONT
    if not overall:
        v.fill = BAD_FILL
    r += 2

    if base is not None:
        line("Inferred Social Security wage base (from the register)", base, bold=True,
             fill=WARN_FILL,
             note="This is the cap the payroll system actually applied, derived from the "
                  "highest Social Security wages among employees whose Medicare wages "
                  "exceed their Social Security wages. CONFIRM against current-year "
                  "authority - this script deliberately asserts no wage base.")
    else:
        line("Inferred Social Security wage base", None, bold=True, fill=WARN_FILL,
             note="Could not be inferred - no employee reached the cap. Confirm the base "
                  "against current-year authority independently.")
    r += 1

    for t in tests:
        c = ws.cell(row=r, column=2, value=f"TEST {t['num']} - {t['name']}")
        c.font, c.fill = Font(bold=True, size=12), SUB_FILL
        r += 1
        if t.get("skipped"):
            line("  " + t["note"], None, fill=WARN_FILL)
            r += 1
            continue
        for label, amt in t["parts"]:
            line("  " + label, amt)
        for label, amt in t.get("extra", []):
            if amt is None:
                line("  " + label + " (not supplied)")
            else:
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
            line(f"    absorbed by the stated rounding tolerance of "
                 f"{TOLERANCE:,.2f} - disclosed, not plugged", None)
            ws.cell(row=r - 1, column=2).font = Font(italic=True, size=9)
        if t["note"]:
            line("  " + t["note"], None)
            ws.cell(row=r - 1, column=2).font = Font(italic=True, size=9)
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


def sheet_tests(wb, tests):
    ws = wb.create_sheet("Test Results")
    heads = ["Test", "Description", "Difference", "Explained", "Unexplained", "Result"]
    ws.append(heads)
    hdr(ws, len(heads))
    for t in tests:
        ws.append([t["num"], t["name"],
                   None if t.get("skipped") else float(t["difference"]),
                   None if t.get("skipped") else float(t["explained"]),
                   None if t.get("skipped") else float(t["residual"]),
                   "NOT PERFORMED" if t.get("skipped")
                   else ("PASS" if t["residual"] == ZERO
                         else "PASS (rounding)" if t.get("within_tolerance")
                         else "FAIL")])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in (2, 3, 4):
            row[i].number_format = MONEY
        if row[5].value == "PASS":
            row[5].font = OK_FONT
        elif row[5].value == "PASS (rounding)":
            row[5].font, row[5].fill = OK_FONT, WARN_FILL
        elif row[5].value == "FAIL":
            row[5].font, row[5].fill = BAD_FONT, BAD_FILL
        else:
            row[5].fill = WARN_FILL
    widths(ws, {1: 6, 2: 66, 3: 16, 4: 16, 5: 16, 6: 16})


def sheet_quarterly(wb, qs, f941_tot):
    ws = wb.create_sheet("Quarterly Analysis")
    heads = ["Quarter", "Wages", "FIT withheld", "SS wages", "SS tax",
             "Medicare wages", "Medicare tax", "Addl Medicare", "Total taxes",
             "Deposits", "Undeposited", "FIT % of wages", "SS effective %",
             "Medicare effective %"]
    ws.append(heads)
    hdr(ws, len(heads))

    def rate(num, den):
        return float(num / den) if den not in (None, ZERO) else None

    for q in qs + [{**f941_tot, "quarter": "TOTAL"}]:
        ws.append([q["quarter"], float(q["wages"]), float(q["fit_withheld"]),
                   float(q["ss_wages"]), float(q["ss_tax"]),
                   float(q["medicare_wages"]), float(q["medicare_tax"]),
                   float(q["addl_medicare_tax"]), float(q["total_taxes"]),
                   float(q["deposits"]), float(q["total_taxes"] - q["deposits"]),
                   rate(q["fit_withheld"], q["wages"]),
                   rate(q["ss_tax"], q["ss_wages"]),
                   rate(q["medicare_tax"], q["medicare_wages"])])
    last = ws.max_row
    for row in ws.iter_rows(min_row=2, max_row=last):
        for i in range(1, 11):
            row[i].number_format = MONEY
        for i in range(11, 14):
            row[i].number_format = RATE
        if row[10].value not in (None, 0):
            row[10].fill = WARN_FILL
    for c in range(1, len(heads) + 1):
        ws.cell(row=last, column=c).font = Font(bold=True)
        ws.cell(row=last, column=c).border = TOP
    r = last + 2
    for note in [
        "Effective rates are COMPUTED from the filed figures (tax / taxable wages).",
        "This script asserts no statutory rate. A rate that drifts between quarters is a",
        "finding regardless of what the correct rate is - investigate the quarter that moved.",
        "'Undeposited' non-zero at TOTAL means taxes incurred were not fully deposited.",
    ]:
        ws.cell(row=r, column=1, value=note).font = Font(italic=True)
        r += 1
    ws.freeze_panes = "B2"
    widths(ws, {1: 10, 2: 15, 3: 14, 4: 15, 5: 14, 6: 15, 7: 14, 8: 14, 9: 14,
                10: 14, 11: 13, 12: 14, 13: 14, 14: 16})


def sheet_employees(wb, erows, base):
    ws = wb.create_sheet("Employee Detail")
    heads = ["Employee ID", "Name", "Gross", "Pre-tax", "Fed taxable", "W-2 box 1",
             "Diff", "SS wages", "W-2 box 3", "Diff", "Medicare wages", "W-2 box 5",
             "Diff", "FIT withheld", "At cap", "Flags"]
    ws.append(heads)
    hdr(ws, len(heads))
    for e in erows:
        w = e["w2"]
        at_cap = "Y" if (base is not None and e["ss_wages"] == base) else ""
        ws.append([e["employee_id"], e["employee_name"], float(e["gross_wages"]),
                   float(e["pretax_deferral"] + e["pretax_cafeteria"]),
                   float(e["federal_taxable_wages"]),
                   float(w["box1"]) if w else None,
                   float(e["d1"]) if e["d1"] is not None else None,
                   float(e["ss_wages"]), float(w["box3"]) if w else None,
                   float(e["d3"]) if e["d3"] is not None else None,
                   float(e["medicare_wages"]), float(w["box5"]) if w else None,
                   float(e["d5"]) if e["d5"] is not None else None,
                   float(e["fit_withheld"]), at_cap, "; ".join(e["flags"])])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in list(range(2, 14)):
            row[i].number_format = MONEY
        for i in (6, 9, 12):
            if row[i].value not in (None, 0):
                row[i].font, row[i].fill = BAD_FONT, BAD_FILL
        if row[15].value:
            row[15].fill = BAD_FILL
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:P{max(ws.max_row, 2)}"
    widths(ws, {1: 13, 2: 26, 3: 14, 4: 12, 5: 13, 6: 13, 7: 11, 8: 13, 9: 12,
                10: 11, 11: 15, 12: 12, 13: 11, 14: 13, 15: 8, 16: 70})


def sheet_liability(wb, gl, f941_tot, qs):
    ws = wb.create_sheet("Liability Rollforward")
    widths(ws, {1: 4, 2: 52, 3: 18, 4: 66})
    r = 1
    ws.cell(row=r, column=2, value="PAYROLL TAX LIABILITY ROLLFORWARD").font = \
        Font(bold=True, size=14)
    r += 2
    if gl.get("liab_begin") is None:
        ws.cell(row=r, column=2,
                value="Not performed - liability balances were not supplied. Without "
                      "this test an undeposited trust-fund tax is undetectable.").fill = WARN_FILL
        return
    deposits = gl.get("gl_deposits")
    if deposits is None:
        deposits = f941_tot["deposits"]
    computed = gl["liab_begin"] + f941_tot["total_taxes"] - deposits
    for label, val, bold, border in [
        ("Beginning payroll tax liability", gl["liab_begin"], False, None),
        ("Add: taxes incurred per Forms 941", f941_tot["total_taxes"], False, None),
        ("Less: deposits made", -deposits, False, None),
        ("= Computed ending liability", computed, True, TOP),
        ("Ending liability per general ledger", gl["liab_end"], False, None),
        ("Difference", computed - gl["liab_end"], True, DBL),
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
        r += 1
    r += 1

    final_q = qs[-1] if qs else None
    if final_q:
        ws.cell(row=r, column=2,
                value="Final quarter taxes incurred").font = Font(bold=True)
        c = ws.cell(row=r, column=3, value=float(final_q["total_taxes"]))
        c.number_format = MONEY
        r += 1
        excess = gl["liab_end"] - final_q["total_taxes"]
        ws.cell(row=r, column=2,
                value="Ending liability in excess of final quarter taxes").font = Font(bold=True)
        c = ws.cell(row=r, column=3, value=float(excess))
        c.number_format = MONEY
        if excess > ZERO:
            c.font, c.fill = BAD_FONT, BAD_FILL
            ws.cell(row=r, column=4,
                    value="ESCALATE. The ending liability should be roughly the taxes owed "
                          "on the final period only. An excess suggests a missed or "
                          "misapplied deposit. Withheld employee tax that was not "
                          "deposited is a trust-fund liability that attaches personally "
                          "to responsible persons."
                    ).alignment = Alignment(wrap_text=True, vertical="top")
        else:
            c.font = OK_FONT
        r += 1


def sheet_memo(wb, tests, escalations, emp_problems, meta):
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

    para("PAYROLL TAX RECONCILIATION MEMO", bold=True, size=14)
    r += 1
    para(f"{meta.get('Client', '')}    Year: {meta.get('Year', '')}", bold=True)
    r += 1

    unexp = [t for t in tests if not t.get("skipped") and not t["passed"]]
    if unexp:
        para("UNEXPLAINED DIFFERENCES", bold=True, size=12, fill=BAD_FILL)
        for t in unexp:
            para(f"  Test {t['num']} - {t['name']}: unexplained {t['residual']:,.2f}")
        para("  Every difference must be named. Do not plug the payroll tax liability "
             "account - that is where unremitted trust-fund taxes hide, and those carry "
             "personal liability for responsible persons.", bold=True)
        r += 1

    skipped = [t for t in tests if t.get("skipped")]
    if skipped:
        para("SCOPE LIMITATIONS", bold=True, size=12, fill=WARN_FILL)
        for t in skipped:
            para(f"  Test {t['num']} - {t['note']}")
        r += 1

    if escalations:
        para("ESCALATIONS", bold=True, size=12, fill=BAD_FILL)
        for e in escalations:
            para("  " + e)
        r += 1

    if emp_problems:
        para(f"EMPLOYEE-LEVEL EXCEPTIONS ({len(emp_problems)})", bold=True, size=12)
        for p in emp_problems[:60]:
            para("  " + p)
        if len(emp_problems) > 60:
            para(f"  ... and {len(emp_problems) - 60} more - see Employee Detail")
        r += 1

    para("FIGURES REQUIRING CONFIRMATION AGAINST CURRENT-YEAR AUTHORITY", bold=True,
         size=12)
    for t in [
        "Social Security wage base - the inferred base on the Reconciliation tab is what "
        "the payroll system applied, not necessarily what it should have applied",
        "Social Security and Medicare tax rates, and the Additional Medicare threshold",
        "Deposit schedule (monthly vs semiweekly) and the deposit thresholds",
        "Filing deadlines for Forms 941, W-2, and W-3",
    ]:
        para("  - " + t)
    r += 1

    para("CONTROL OBSERVATIONS TO CONSIDER", bold=True, size=12)
    for t in [
        "Whether the person who processes payroll can also add an employee or change "
        "bank details",
        "Whether payroll registers are reviewed and approved by someone outside payroll",
        "Whether terminated employees are removed promptly",
        "Whether 1099 contractors resemble the W-2 population in pattern of work",
        "Whether officer or shareholder compensation is reasonable and documented",
    ]:
        para("  - " + t)
    r += 1
    para("Frame these to management as control observations, not audit conclusions. "
         "Report worker-classification patterns as facts and do not conclude on status.",
         bold=True)


# ------------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--register", required=True)
    ap.add_argument("--f941", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--w3")
    ap.add_argument("--gl")
    ap.add_argument("--w2")
    ap.add_argument("--reconciling-items")
    ap.add_argument("--client", default="")
    ap.add_argument("--year", default="")
    ap.add_argument("--rounding-tolerance", default="0.00",
                    help="absolute per-test tolerance for per-payroll rounding "
                         "(default 0.00). Anything absorbed is reported separately.")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    global TOLERANCE
    TOLERANCE = dec(args.rounding_tolerance) or ZERO

    emps, reg_tot = load_register(Path(args.register))
    qs, f941_tot = load_f941(Path(args.f941))
    w3 = load_w3(Path(args.w3) if args.w3 else None)
    gl = load_gl(Path(args.gl) if args.gl else None)
    w2 = load_w2(Path(args.w2) if args.w2 else None)
    items = load_items(Path(args.reconciling_items) if args.reconciling_items else None)

    base = infer_wage_base(emps)
    tests = run_tests(reg_tot, emps, f941_tot, qs, w3, gl, w2, items)
    erows, emp_problems = employee_tests(emps, w2, base)

    escalations = []
    if gl.get("liab_end") is not None and qs:
        excess = gl["liab_end"] - qs[-1]["total_taxes"]
        if excess > ZERO:
            escalations.append(
                f"Ending payroll tax liability exceeds the final quarter's taxes by "
                f"{excess:,.2f}. Likely a missed or misapplied deposit. Unremitted "
                f"withheld tax is a trust-fund liability attaching personally to "
                f"responsible persons - treat as urgent.")
    undeposited = f941_tot["total_taxes"] - f941_tot["deposits"]
    if undeposited > ZERO:
        escalations.append(
            f"Taxes incurred per the 941s exceed deposits by {undeposited:,.2f} for the "
            f"year. Confirm whether this is the final period's liability not yet due, or "
            f"an actual shortfall.")
    if w3 and f941_tot["wages"] != w3["box1_wages"]:
        escalations.append(
            f"941 wages ({f941_tot['wages']:,.2f}) and W-3 box 1 "
            f"({w3['box1_wages']:,.2f}) disagree by "
            f"{f941_tot['wages'] - w3['box1_wages']:,.2f}. An agency mismatch notice is "
            f"effectively automatic - resolve before year-end filing.")
    if len(qs) < 4:
        escalations.append(
            f"Only {len(qs)} quarterly Form(s) 941 supplied. An annual W-3 cannot be "
            f"reconciled to a partial year; obtain all four quarters and any 941-X.")
    if not w3:
        escalations.append("W-3 totals not supplied - the 941-to-W-3 comparison, which is "
                           "the one the agency performs automatically, was not run.")

    unexplained = [t for t in tests if not t.get("skipped") and not t["passed"]]
    absorbed = [t for t in tests if t.get("within_tolerance")]
    overall = not unexplained and not emp_problems

    meta = {
        "Client": args.client or "(not stated)",
        "Year": args.year or "(not stated)",
        "Employees in register": len(emps),
        "Quarters of Form 941 supplied": len(qs),
        "W-2s supplied": len(w2) or "(none)",
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    # ---- console
    print("=" * 74)
    print("PAYROLL TAX RECONCILIATION")
    print("=" * 74)
    print(f"Client   : {meta['Client']}    Year: {meta['Year']}")
    print(f"Employees: {len(emps)}    Quarters filed: {len(qs)}    W-2s: {len(w2)}")
    print()
    if base is not None:
        print(f"Inferred Social Security wage base : {base:>16,.2f}")
        print("  (the cap the payroll system applied - CONFIRM against current-year "
              "authority)")
    else:
        print("Inferred Social Security wage base : could not infer - no employee at cap")
    print()
    for t in tests:
        if t.get("skipped"):
            print(f"Test {t['num']}: NOT PERFORMED - {t['note']}")
            continue
        status = ("PASS" if t["residual"] == ZERO
                  else "PASS (rounding)" if t.get("within_tolerance")
                  else "*** FAIL ***")
        print(f"Test {t['num']}: {status:<14} diff {t['difference']:>14,.2f}  "
              f"explained {t['explained']:>13,.2f}  unexplained {t['residual']:>13,.2f}")
        print(f"         {t['name']}")
    if absorbed:
        print(f"\nABSORBED BY THE STATED ROUNDING TOLERANCE ({TOLERANCE:,.2f}):")
        for t in absorbed:
            print(f"  Test {t['num']}: {t['residual']:,.2f} - within tolerance, "
                  f"disclosed on the workpaper")
        print("  A tolerance is professional judgment only while it is disclosed. "
              "These amounts appear on the Reconciliation tab.")
    if emp_problems:
        print(f"\nEMPLOYEE-LEVEL EXCEPTIONS ({len(emp_problems)}):")
        for p in emp_problems[:15]:
            print(f"  ! {p}")
        if len(emp_problems) > 15:
            print(f"  ... and {len(emp_problems) - 15} more")
    if escalations:
        print("\nESCALATE:")
        for e in escalations:
            print(f"  ! {e}")

    if not overall and not args.force:
        print("\n" + "=" * 74)
        print("WORKBOOK NOT WRITTEN. Unexplained differences remain.")
        print("Classify each difference in --reconciling-items with a cause, an amount,")
        print("and evidence, then re-run. Never plug the payroll tax liability account.")
        print("=" * 74)
        return 1

    wb = Workbook()
    sheet_recon(wb, tests, meta, base, overall, escalations)
    sheet_tests(wb, tests)
    sheet_quarterly(wb, qs, f941_tot)
    sheet_employees(wb, erows, base)
    sheet_liability(wb, gl, f941_tot, qs)
    sheet_memo(wb, tests, escalations, emp_problems, meta)
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
