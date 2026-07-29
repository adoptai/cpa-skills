#!/usr/bin/env python3
"""
Trial balance and chart of accounts integrity.

Run this before anything is built on the trial balance. Debits equalling credits is
necessary and tells you almost nothing - every accounting system enforces it. What
it does not enforce: two accounts for the same thing, an account mapped to no
statement line, or a balance sitting in the wrong direction.

Gates:
  * debits must equal credits
  * every account must map to a financial statement line - an unmapped account is
    reported, never defaulted to a caption
  * statement subtotals derived from the mapping must reproduce the trial balance
  * every account must have a type, or it cannot be tested for sign

Nothing is merged automatically. Which duplicates are real is a judgement about how
the business wants to see its results, and merging changes comparatives.

Usage:
    python3 tb_integrity.py --trial-balance tb.csv --mapping mapping.csv \
        --prior-year tb_prior.csv \
        --client "Halstead Regional Foods" --period 12/31/2025 \
        --out "Halstead - 2025 Trial Balance Integrity.xlsx"

--trial-balance CSV:
    account, name, debit, credit, type, contra (yes/no), activity_count,
    last_activity_date
      type: asset | liability | equity | revenue | expense

--mapping CSV:      account, statement, line
      statement: balance_sheet | income_statement
--prior-year CSV:   account, name, debit, credit
--target CSV:       old_account, new_account, new_name    (optional, for conversions)
"""

from __future__ import annotations

import argparse
import csv
import re
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

TYPES = {"asset", "liability", "equity", "revenue", "expense"}
# natural balance: debit-positive types
DEBIT_TYPES = {"asset", "expense"}

SUSPENSE_WORDS = ("suspense", "clearing", "conversion", "plug", "difference",
                  "to balance", "misc", "miscellaneous", "unallocated", "temporary",
                  "tbd", "other", "unknown", "holding")


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


def truthy(s) -> bool:
    return clean(s).lower() in ("yes", "y", "true", "1", "x")


def norm_name(s: str) -> str:
    """Collapse punctuation, case, spacing, &/and, and a trailing number."""
    t = clean(s).lower().replace("&", "and")
    t = re.sub(r"[^a-z0-9 ]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    t = re.sub(r"\s+\d+$", "", t)
    return t


def pdate(raw):
    s = clean(raw)
    if not s:
        return None
    for f in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d-%b-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s.split(" ")[0], f).date()
        except ValueError:
            continue
    return None


def rows_of(path: Path, label: str):
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No rows in {label} ({path})")
    return rows


def load_tb(path: Path) -> list[dict]:
    out = []
    for i, r in enumerate(rows_of(path, "trial balance"), start=1):
        t = clean(r.get("type")).lower()
        acct = clean(r.get("account")) or f"(no number {i})"
        if t and t not in TYPES:
            sys.exit(f"Trial balance row {i} ({acct}): type '{t}' is not recognised. "
                     f"Valid types are {sorted(TYPES)}.")
        dr, cr = d0(r.get("debit")), d0(r.get("credit"))
        out.append({
            "row": i, "account": acct,
            "name": clean(r.get("name")),
            "debit": dr, "credit": cr, "net": dr - cr,
            "type": t, "contra": truthy(r.get("contra")),
            "activity_count": dec(r.get("activity_count")),
            "last_activity": pdate(r.get("last_activity_date")),
            "statement": "", "line": "", "prior_net": None, "movement": None,
            "flags": [],
        })
    return out


def load_mapping(path: Path | None) -> dict:
    out = {}
    if not path or not path.exists():
        return out
    for r in rows_of(path, "mapping"):
        a = clean(r.get("account"))
        if a:
            out[a] = {"statement": clean(r.get("statement")).lower(),
                      "line": clean(r.get("line"))}
    return out


# ----------------------------------------------------------------------- tests

def analyse(tb, mapping, prior, target) -> dict:
    prior_idx = {clean(p.get("account")): d0(p.get("debit")) - d0(p.get("credit"))
                 for p in prior}

    unmapped, untyped, sign_exc, dormant = [], [], [], []
    dupes = defaultdict(list)
    suspense = []

    for a in tb:
        m = mapping.get(a["account"])
        if m and m["line"]:
            a["statement"], a["line"] = m["statement"], m["line"]
        else:
            a["flags"].append("NOT MAPPED to a financial statement line - it will "
                              "disappear from the statements while the trial balance "
                              "still balances")
            unmapped.append(a)

        if not a["type"]:
            a["flags"].append("no account type - cannot be tested for sign or "
                              "classification")
            untyped.append(a)
        else:
            natural_debit = a["type"] in DEBIT_TYPES
            if a["contra"]:
                natural_debit = not natural_debit
            if a["net"] != ZERO:
                is_debit = a["net"] > ZERO
                if is_debit != natural_debit:
                    a["flags"].append(
                        f"{a['type']}{' (contra)' if a['contra'] else ''} with a "
                        f"{'debit' if is_debit else 'credit'} balance of "
                        f"{abs(a['net']):,.2f} - not automatically wrong, but it needs "
                        f"a reason")
                    sign_exc.append(a)

        if a["account"] in prior_idx:
            a["prior_net"] = prior_idx[a["account"]]
            a["movement"] = a["net"] - a["prior_net"]

        no_activity = (a["activity_count"] is not None and a["activity_count"] == ZERO)
        if a["net"] == ZERO and (no_activity or a["activity_count"] is None):
            a["flags"].append("zero balance"
                              + (" with no activity" if no_activity else
                                 " (activity count not supplied)"))
            dormant.append(a)
        elif no_activity and a["net"] != ZERO:
            a["flags"].append("balance carried with NO activity in the period - a dormant "
                              "account holding a balance is where misposting hides")
            dormant.append(a)

        nn = norm_name(a["name"])
        if nn:
            dupes[nn].append(a)

        low = a["name"].lower()
        if any(w in low for w in SUSPENSE_WORDS) and a["net"] != ZERO:
            a["flags"].append(f"suspense/clearing-style account name with a balance of "
                              f"{a['net']:,.2f} - these are meant to be empty")
            suspense.append(a)

    dup_groups = []
    for nn, group in dupes.items():
        if len(group) > 1:
            with_balance = [g for g in group if g["net"] != ZERO]
            dup_groups.append({"normalised": nn, "accounts": group,
                               "both_with_balance": len(with_balance) > 1,
                               "total": sum((g["net"] for g in group), ZERO)})
            for g in group:
                others = ", ".join(x["account"] for x in group if x is not g)
                g["flags"].append(f"DUPLICATE CANDIDATE with {others} - same normalised "
                                  f"name. Not merged automatically; merging changes "
                                  f"comparatives and is a decision")

    # numbering structure
    numeric = [a for a in tb if re.fullmatch(r"\d+", a["account"])]
    depths = {len(a["account"]) for a in numeric}
    struct = []
    if len(depths) > 1:
        struct.append(f"Inconsistent account number length: {sorted(depths)}. A mixed "
                      f"numbering depth breaks range-based reporting and sorting.")
    by_type_range = defaultdict(list)
    for a in numeric:
        if a["type"]:
            by_type_range[a["type"]].append(int(a["account"]))
    for t, nums in by_type_range.items():
        lo, hi = min(nums), max(nums)
        for other, onums in by_type_range.items():
            if other == t:
                continue
            overlap = [n for n in onums if lo <= n <= hi]
            if overlap:
                struct.append(f"{other} accounts {overlap[:5]} fall inside the {t} range "
                              f"{lo}-{hi}. Overlapping ranges make range-based mapping "
                              f"unreliable.")
                break

    return {"unmapped": unmapped, "untyped": untyped, "sign_exceptions": sign_exc,
            "dormant": dormant, "duplicates": dup_groups, "suspense": suspense,
            "structure": struct}


def totals(tb) -> dict:
    dr = sum((a["debit"] for a in tb), ZERO)
    cr = sum((a["credit"] for a in tb), ZERO)
    by_type = defaultdict(lambda: ZERO)
    for a in tb:
        if a["type"]:
            by_type[a["type"]] += a["net"]
    assets = by_type["asset"]
    liabs = -by_type["liability"]
    equity = -by_type["equity"]
    rev = -by_type["revenue"]
    exp = by_type["expense"]
    ni = rev - exp
    return {"debits": dr, "credits": cr, "difference": dr - cr,
            "assets": assets, "liabilities": liabs, "equity": equity,
            "revenue": rev, "expense": exp, "net_income": ni,
            "equation": assets - (liabs + equity + ni)}


def mapping_proof(tb) -> dict:
    by_line = defaultdict(lambda: ZERO)
    mapped_net = ZERO
    for a in tb:
        if a["line"]:
            by_line[(a["statement"], a["line"])] += a["net"]
            mapped_net += a["net"]
    unmapped_net = sum((a["net"] for a in tb if not a["line"]), ZERO)
    return {"by_line": dict(by_line), "mapped_net": mapped_net,
            "unmapped_net": unmapped_net, "total_net": mapped_net + unmapped_net}


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


def sheet_summary(wb, meta, tests, tot, res, esc, overall):
    ws = wb.active
    ws.title = "Integrity Summary"
    widths(ws, {1: 4, 2: 52, 3: 18, 4: 14, 5: 60})
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

    line("TRIAL BALANCE AND CHART OF ACCOUNTS INTEGRITY", bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    v = ws.cell(row=r, column=2,
                value="SOUND - safe to build on" if overall else
                "NOT SOUND - resolve before building anything on this trial balance. "
                "Errors here propagate into every downstream workpaper.")
    v.font = OK_FONT if overall else BAD_FONT
    if not overall:
        v.fill = BAD_FILL
    r += 2

    line("TOTALS", bold=True, size=12, fill=SUB_FILL)
    line("  Total debits", tot["debits"])
    line("  Total credits", tot["credits"])
    line("  Difference", tot["difference"], bold=True, border=TOP)
    c = ws.cell(row=r - 1, column=3)
    c.font = OK_FONT if tot["difference"] == ZERO else BAD_FONT
    if tot["difference"] != ZERO:
        c.fill = BAD_FILL
    r += 1
    line("  Assets", tot["assets"])
    line("  Liabilities", tot["liabilities"])
    line("  Equity", tot["equity"])
    line("  Net income (revenue less expense)", tot["net_income"])
    line("  Accounting equation check", tot["equation"], bold=True, border=TOP,
         note="Assets less (liabilities + equity + net income). A break here means "
              "accounts are typed wrongly even though the trial balance balances."
              if tot["equation"] != ZERO else "")
    c = ws.cell(row=r - 1, column=3)
    c.font = OK_FONT if tot["equation"] == ZERO else BAD_FONT
    if tot["equation"] != ZERO:
        c.fill = BAD_FILL
    r += 1

    for t in tests:
        line(f"TEST {t['num']} - {t['name']}", None,
             "PASS" if t["passed"] else "FAIL", note=t.get("note", ""))
    r += 1

    line("FINDINGS BY TYPE", bold=True, size=12, fill=SUB_FILL)
    for label, items in [("Unmapped accounts", res["unmapped"]),
                         ("Untyped accounts", res["untyped"]),
                         ("Duplicate candidate groups", res["duplicates"]),
                         ("Sign exceptions", res["sign_exceptions"]),
                         ("Dormant or zero-balance accounts", res["dormant"]),
                         ("Suspense-style accounts with a balance", res["suspense"]),
                         ("Structure observations", res["structure"])]:
        line("  " + label, len(items),
             fill=BAD_FILL if items and label.startswith(("Unmapped", "Untyped"))
             else (WARN_FILL if items else None))
    r += 1

    if esc:
        line("ESCALATE", bold=True, size=12, fill=BAD_FILL)
        for e in esc:
            line("  " + e)
        r += 1
    r += 1
    line("Prepared by / date:  __________________  ____________", bold=True)
    line("Reviewed by / date:  __________________  ____________", bold=True)


def sheet_tb(wb, tb):
    ws = wb.create_sheet("Trial Balance")
    heads = ["Account", "Name", "Type", "Contra", "Debit", "Credit", "Net",
             "Statement", "Line", "Prior year net", "Movement", "Activity",
             "Last activity", "Flags"]
    ws.append(heads)
    hdr(ws, len(heads))
    for a in sorted(tb, key=lambda x: (0 if x["flags"] else 1, x["account"])):
        ws.append([a["account"], a["name"], a["type"] or None,
                   "Y" if a["contra"] else "", float(a["debit"]), float(a["credit"]),
                   float(a["net"]), a["statement"] or None, a["line"] or None,
                   float(a["prior_net"]) if a["prior_net"] is not None else None,
                   float(a["movement"]) if a["movement"] is not None else None,
                   float(a["activity_count"]) if a["activity_count"] is not None else None,
                   a["last_activity"], "; ".join(a["flags"])])
    last = ws.max_row
    ws.append(["TOTAL", "", "", "", f"=SUM(E2:E{last})", f"=SUM(F2:F{last})",
               f"=SUM(G2:G{last})"] + [None] * 7)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in (4, 5, 6, 9, 10):
            row[i].number_format = MONEY
        row[12].number_format = DATEF
        if not row[8].value:
            row[8].fill = BAD_FILL
        if row[13].value:
            row[13].fill = WARN_FILL
    for c in range(1, 8):
        cell = ws.cell(row=ws.max_row, column=c)
        cell.font, cell.border = Font(bold=True), TOP
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:N{last}"
    widths(ws, {1: 12, 2: 36, 3: 12, 4: 8, 5: 15, 6: 15, 7: 15, 8: 16, 9: 28,
                10: 16, 11: 15, 12: 10, 13: 13, 14: 74})


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
        ws.cell(row=2, column=1, value=empty).font = OK_FONT
    elif footer:
        rr = ws.max_row + 2
        c = ws.cell(row=rr, column=1, value=footer)
        c.font = Font(italic=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    widths(ws, {i: 20 for i in range(1, len(header) + 1)} | {2: 36,
                                                            len(header): 60})


def sheet_mapping_proof(wb, mp, tot, tb):
    ws = wb.create_sheet("Mapping Proof")
    widths(ws, {1: 4, 2: 44, 3: 18, 4: 60})
    r = 1
    ws.cell(row=r, column=2, value="MAPPING PROOF").font = Font(bold=True, size=14)
    r += 2
    for label, v, bold, border in [
        ("Net of mapped accounts", mp["mapped_net"], False, None),
        ("Net of unmapped accounts", mp["unmapped_net"], False, None),
        ("= Total net", mp["total_net"], True, TOP),
        ("Trial balance net (debits less credits)", tot["difference"], False, None),
        ("Difference", mp["total_net"] - tot["difference"], True, DBL),
    ]:
        ws.cell(row=r, column=2, value=label).font = Font(bold=bold)
        c = ws.cell(row=r, column=3, value=float(v))
        c.number_format, c.font = MONEY, Font(bold=bold)
        if border:
            c.border = border
        if label == "Difference":
            c.font = OK_FONT if v == ZERO else BAD_FONT
            if v != ZERO:
                c.fill = BAD_FILL
        r += 1
    if mp["unmapped_net"] != ZERO:
        ws.cell(row=r, column=2,
                value=f"{mp['unmapped_net']:,.2f} of net balance sits in unmapped "
                      f"accounts. That amount will not appear in the financial statements "
                      f"while the trial balance still balances - which is exactly the "
                      f"failure this test exists to prevent.").fill = BAD_FILL
        r += 2
    ws.cell(row=r, column=2, value="Mapped statement lines").font = Font(bold=True, size=12)
    r += 1
    for h, col in zip(["Statement", "Line", "Net"], (2, 3, 4)):
        c = ws.cell(row=r, column=col, value=h)
        c.fill, c.font = HDR_FILL, HDR_FONT
    r += 1
    for (stmt, line), v in sorted(mp["by_line"].items()):
        ws.cell(row=r, column=2, value=stmt)
        ws.cell(row=r, column=3, value=line)
        c = ws.cell(row=r, column=4, value=float(v))
        c.number_format = MONEY
        r += 1


# ------------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--trial-balance", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--mapping")
    ap.add_argument("--prior-year")
    ap.add_argument("--target")
    ap.add_argument("--client", default="")
    ap.add_argument("--period", default="")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    tb = load_tb(Path(args.trial_balance))
    mapping = load_mapping(Path(args.mapping) if args.mapping else None)
    prior = (rows_of(Path(args.prior_year), "prior year")
             if args.prior_year and Path(args.prior_year).exists() else [])
    target = (rows_of(Path(args.target), "target")
              if args.target and Path(args.target).exists() else [])

    res = analyse(tb, mapping, prior, target)
    tot = totals(tb)
    mp = mapping_proof(tb)

    tests = [
        {"num": 1, "name": "Debits equal credits", "passed": tot["difference"] == ZERO,
         "note": "" if tot["difference"] == ZERO else
                 f"out of balance by {tot['difference']:,.2f}"},
        {"num": 2, "name": "Every account maps to a financial statement line",
         "passed": not res["unmapped"] and bool(mapping),
         "note": ("no mapping supplied - producing one is the deliverable, and the "
                  "Unmapped Accounts tab lists everything still to be placed"
                  if not mapping else
                  "" if not res["unmapped"] else
                  f"{len(res['unmapped'])} account(s) unmapped, net "
                  f"{mp['unmapped_net']:,.2f}")},
        {"num": 3, "name": "Mapping foots to the trial balance",
         "passed": mp["total_net"] == tot["difference"],
         "note": "" if mp["total_net"] == tot["difference"] else
                 "the mapping double-counts or drops an account"},
        {"num": 4, "name": "Accounting equation holds from account types",
         "passed": tot["equation"] == ZERO,
         "note": "" if tot["equation"] == ZERO else
                 f"off by {tot['equation']:,.2f} - accounts are typed wrongly even "
                 f"though the trial balance balances"},
        {"num": 5, "name": "No unexplained balance-direction exceptions",
         "passed": not res["sign_exceptions"],
         "note": "" if not res["sign_exceptions"] else
                 f"{len(res['sign_exceptions'])} account(s) with a balance opposite to "
                 f"their type - not automatically wrong, but each needs a reason"},
        {"num": 6, "name": "No duplicate or near-duplicate accounts",
         "passed": not res["duplicates"],
         "note": "" if not res["duplicates"] else
                 f"{len(res['duplicates'])} candidate group(s) - nothing merged "
                 f"automatically, since merging changes comparatives"},
        {"num": 7, "name": "No dormant accounts holding balances",
         "passed": not [a for a in res["dormant"] if a["net"] != ZERO],
         "note": f"{len(res['dormant'])} dormant or zero-balance account(s)"},
        {"num": 8, "name": "Account structure and numbering consistent",
         "passed": not res["structure"],
         "note": "; ".join(res["structure"]) if res["structure"] else ""},
    ]

    esc = []
    if tot["difference"] != ZERO:
        esc.append(f"The trial balance does not balance: {tot['difference']:,.2f}. Nothing "
                   f"built on it is reliable.")
    material_unmapped = [a for a in res["unmapped"] if abs(a["net"]) > ZERO]
    if material_unmapped:
        esc.append(f"{len(material_unmapped)} unmapped account(s) carry a net balance of "
                   f"{sum((a['net'] for a in material_unmapped), ZERO):,.2f}. That amount "
                   f"will disappear from the financial statements while the trial balance "
                   f"still balances.")
    both = [g for g in res["duplicates"] if g["both_with_balance"]]
    if both:
        for g in both:
            accts = ", ".join(f"{a['account']} {a['name']} ({a['net']:,.2f})"
                              for a in g["accounts"])
            esc.append(f"Duplicate accounts both carrying balances: {accts}. Every "
                       f"analysis on either one is wrong.")
    for a in res["suspense"]:
        esc.append(f"{a['account']} {a['name']}: {a['net']:,.2f}. Suspense, clearing and "
                   f"conversion accounts are meant to be empty; a balance means something "
                   f"was parked and forgotten.")
    if prior:
        re_acct = [a for a in tb if "retained earning" in a["name"].lower()
                   or "accumulated deficit" in a["name"].lower()]
        for a in re_acct:
            if a["prior_net"] is not None:
                expected = a["prior_net"] - tot["net_income"]
                if a["net"] != expected:
                    esc.append(
                        f"{a['account']} {a['name']}: {-a['net']:,.2f} against prior year "
                        f"{-a['prior_net']:,.2f} plus net income {tot['net_income']:,.2f}. "
                        f"Difference {-(a['net'] - expected):,.2f} - one of the fastest "
                        f"ways to detect an entry posted directly to equity. Confirm "
                        f"distributions before concluding.")
    for a in res["untyped"]:
        esc.append(f"{a['account']} {a['name']}: no account type, so it cannot be tested "
                   f"for sign or classification.")

    overall = all(t["passed"] for t in tests)

    meta = {
        "Client": args.client or "(not stated)",
        "Period": args.period or "(not stated)",
        "Accounts": len(tb),
        "Mapping supplied": "yes" if mapping else "NO",
        "Prior year supplied": "yes" if prior else "no",
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    # ---- console
    print("=" * 76)
    print("TRIAL BALANCE AND CHART OF ACCOUNTS INTEGRITY")
    print("=" * 76)
    print(f"Client : {meta['Client']}   Period: {meta['Period']}   Accounts: {len(tb)}")
    print()
    print(f"  Total debits    {tot['debits']:>18,.2f}")
    print(f"  Total credits   {tot['credits']:>18,.2f}")
    print(f"  Difference      {tot['difference']:>18,.2f}   "
          f"{'BALANCES' if tot['difference'] == ZERO else '*** OUT OF BALANCE ***'}")
    print(f"  Assets          {tot['assets']:>18,.2f}")
    print(f"  Liabilities     {tot['liabilities']:>18,.2f}")
    print(f"  Equity          {tot['equity']:>18,.2f}")
    print(f"  Net income      {tot['net_income']:>18,.2f}")
    print(f"  Equation check  {tot['equation']:>18,.2f}   "
          f"{'OK' if tot['equation'] == ZERO else '*** TYPES ARE WRONG ***'}")
    print()
    for t in tests:
        print(f"Test {t['num']}: {'PASS' if t['passed'] else '*** FAIL ***':<14} {t['name']}")
        if t.get("note"):
            print(f"         {t['note']}")
    if res["duplicates"]:
        print(f"\nDUPLICATE CANDIDATES ({len(res['duplicates'])}):")
        for g in res["duplicates"][:10]:
            accts = "  |  ".join(f"{a['account']} {a['name']} {a['net']:,.2f}"
                                 for a in g["accounts"])
            print(f"  {'BOTH WITH BALANCE  ' if g['both_with_balance'] else ''}{accts}")
    if esc:
        print("\nESCALATE:")
        for e in esc:
            print(f"  ! {e}")

    if not overall and not args.force:
        print("\n" + "=" * 76)
        print("WORKBOOK NOT WRITTEN.")
        print("Resolve before building anything on this trial balance - a duplicate or an")
        print("unmapped account propagates into every downstream workpaper and is far")
        print("harder to find there.")
        print("=" * 76)
        return 1

    wb = Workbook()
    sheet_summary(wb, meta, tests, tot, res, esc, overall)
    sheet_tb(wb, tb)
    sheet_list(wb, "Unmapped Accounts",
               ["Account", "Name", "Type", "Net", "Statement line to assign"],
               [[a["account"], a["name"], a["type"], float(a["net"]), None]
                for a in res["unmapped"]],
               {3: MONEY},
               "None - every account maps to a financial statement line.",
               footer="An unmapped account disappears from the financial statements while "
                      "the trial balance still balances.")
    sheet_mapping_proof(wb, mp, tot, tb)
    sheet_list(wb, "Duplicate Candidates",
               ["Group", "Account", "Name", "Net", "Both with balance?",
                "Merge decision"],
               [[i + 1, a["account"], a["name"], float(a["net"]),
                 "YES" if g["both_with_balance"] else "", None]
                for i, g in enumerate(res["duplicates"]) for a in g["accounts"]],
               {3: MONEY},
               "None - no duplicate or near-duplicate account names.",
               footer="Nothing is merged automatically. Which duplicates are real is a "
                      "judgement about how the business wants to see its results, and "
                      "merging changes comparatives.")
    sheet_list(wb, "Sign Exceptions",
               ["Account", "Name", "Type", "Contra", "Net", "Condition", "Reason"],
               [[a["account"], a["name"], a["type"], "Y" if a["contra"] else "",
                 float(a["net"]),
                 next((f for f in a["flags"] if "balance of" in f), ""), None]
                for a in res["sign_exceptions"]],
               {4: MONEY},
               "None - every balance sits in its natural direction.",
               footer="A balance opposite to its type is not automatically wrong. An "
                      "overdrawn bank account, an accumulated deficit, and a declared "
                      "contra account are all legitimate. Each needs a reason.")
    sheet_list(wb, "Unused and Dormant",
               ["Account", "Name", "Type", "Net", "Activity count", "Last activity",
                "Keep or close"],
               [[a["account"], a["name"], a["type"], float(a["net"]),
                 float(a["activity_count"]) if a["activity_count"] is not None else None,
                 a["last_activity"], None] for a in res["dormant"]],
               {3: MONEY, 5: DATEF},
               "None - every account was used in the period.",
               footer="A dormant account holding a balance is where misposting hides, "
                      "because nobody reviews it.")
    if target:
        sheet_list(wb, "Conversion Mapping",
                   ["Old account", "Old name", "Net", "New account", "New name",
                    "Status"],
                   [[clean(t.get("old_account")), "", None,
                     clean(t.get("new_account")), clean(t.get("new_name")),
                     "mapped" if clean(t.get("new_account")) else "NOT MAPPED"]
                    for t in target],
                   {2: MONEY},
                   "No target structure supplied.",
                   footer="An account with no new-account assignment blocks the "
                          "conversion.")
    wb.active = 0
    out = Path(args.out)
    if not overall:
        out = out.with_name(out.stem + " [NOT SOUND]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\n{'SOUND' if overall else 'NOT SOUND (forced)'}.  Workbook: {out}")
    return 0 if overall else 1


if __name__ == "__main__":
    sys.exit(main())
