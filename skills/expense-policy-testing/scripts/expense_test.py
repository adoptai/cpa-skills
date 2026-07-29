#!/usr/bin/env python3
"""
Expense and T&E policy testing against the client's own written policy.

Gates:
  * the population must tie to the GL expense total
  * every transaction receives a disposition - compliant plus exceptions must equal
    the population in count and in value
  * policy rules are supplied, never assumed

Output is organised by employee and by approver before transaction, because the
findings are behavioural and that is the order in which anything gets fixed. An
approver with an exception rate far above peers explains the other findings.

Usage:
    python3 expense_test.py --expenses expenses.csv --gl-total 1284550.00 \
        --policy policy.csv --roster roster.csv \
        --period FY2025 --client "Brightline Media Group" \
        --out "Brightline - FY2025 Expense Policy Testing.xlsx"

--expenses CSV:
    txn_id, txn_date, employee, employee_id, amount, category, merchant,
    description, approver, receipt (yes/no), report_ref, cost_centre,
    payment_method

--policy CSV (the client's WRITTEN policy - nothing is assumed):
    rule, category, threshold, note
      rule: approval_required_above | receipt_required_above | category_cap
          | prohibited_category | approver_limit | weekend_allowed_categories
      For approver_limit, put the approver name in `category` and their limit in
      `threshold`.

--roster CSV:
    employee, employee_id, manager, termination_date
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

EXC_TYPES = [
    "missing approval", "self-approved", "over approval authority",
    "circular approval", "over limit", "threshold-adjacent", "split transaction",
    "missing receipt", "generic business purpose", "duplicate",
    "prohibited category", "weekend/holiday spend", "round number",
    "claim after termination",
]
RECOVERABLE = {"over limit", "duplicate", "prohibited category", "split transaction",
               "claim after termination"}

GENERIC_PURPOSE = [
    "client meeting", "business", "business meal", "misc", "miscellaneous",
    "per policy", "travel", "expenses", "reimbursement", "meeting", "lunch",
    "dinner", "meal", "n/a", "none", "various", "team", "office",
]


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


def truthy(s) -> bool:
    return clean(s).lower() in ("yes", "y", "true", "1", "x")


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


def load_expenses(path: Path) -> list[dict]:
    out = []
    for i, r in enumerate(rows_of(path, "expenses"), start=1):
        out.append({
            "row": i,
            "txn_id": clean(r.get("txn_id")) or f"TXN{i:05d}",
            "txn_date": pdate(r.get("txn_date")),
            "employee": clean(r.get("employee")) or "(no employee)",
            "employee_id": clean(r.get("employee_id")),
            "amount": d0(r.get("amount")),
            "category": clean(r.get("category")) or "(uncategorised)",
            "merchant": clean(r.get("merchant")),
            "description": clean(r.get("description")),
            "approver": clean(r.get("approver")),
            "receipt": truthy(r.get("receipt")),
            "report_ref": clean(r.get("report_ref")),
            "cost_centre": clean(r.get("cost_centre")),
            "payment_method": clean(r.get("payment_method")),
            "exceptions": [], "notes": [],
        })
    return out


def load_policy(path: Path | None) -> dict:
    p = {"approval_above": None, "receipt_above": None, "category_cap": {},
         "prohibited": set(), "approver_limit": {}, "weekend_ok": set(),
         "supplied": False, "rules": []}
    if not path or not path.exists():
        return p
    p["supplied"] = True
    for r in rows_of(path, "policy"):
        rule = clean(r.get("rule")).lower()
        cat = clean(r.get("category"))
        thr = dec(r.get("threshold"))
        note = clean(r.get("note"))
        p["rules"].append({"rule": rule, "category": cat, "threshold": thr, "note": note})
        if rule == "approval_required_above":
            p["approval_above"] = thr
        elif rule == "receipt_required_above":
            p["receipt_above"] = thr
        elif rule == "category_cap" and cat:
            p["category_cap"][cat.lower()] = thr
        elif rule == "prohibited_category" and cat:
            p["prohibited"].add(cat.lower())
        elif rule == "approver_limit" and cat:
            p["approver_limit"][key(cat)] = thr
        elif rule == "weekend_allowed_categories" and cat:
            p["weekend_ok"].add(cat.lower())
    return p


def load_roster(path: Path | None) -> dict:
    out = {}
    if not path or not path.exists():
        return out
    for r in rows_of(path, "roster"):
        n = clean(r.get("employee"))
        if not n:
            continue
        out[key(n)] = {"employee": n, "employee_id": clean(r.get("employee_id")),
                       "manager": clean(r.get("manager")),
                       "termination_date": pdate(r.get("termination_date"))}
    return out


# ----------------------------------------------------------------------- tests

def is_round(v: Decimal) -> bool:
    a = abs(v)
    if a < Decimal("100"):
        return False
    return a % Decimal("50") == ZERO


def run_tests(rows, pol, roster, holidays, threshold_band: Decimal) -> list[dict]:
    groups = []

    def flag(x, t, detail=""):
        if t not in x["exceptions"]:
            x["exceptions"].append(t)
        if detail:
            x["notes"].append(f"{t}: {detail}")

    thresholds = [t for t in (pol["approval_above"],) if t is not None]
    thresholds += [v for v in pol["category_cap"].values() if v is not None]

    for x in rows:
        amt = x["amount"]

        # authorisation
        if pol["approval_above"] is not None and amt > pol["approval_above"]:
            if not x["approver"]:
                flag(x, "missing approval",
                     f"{amt:,.2f} exceeds the {pol['approval_above']:,.2f} approval "
                     f"threshold with no approver recorded")
        elif not x["approver"] and pol["approval_above"] is None:
            flag(x, "missing approval", "no approver recorded and no threshold in policy")

        if x["approver"] and key(x["approver"]) == key(x["employee"]):
            flag(x, "self-approved",
                 f"approved by the claimant ({x['employee']}) - a control finding at any "
                 f"amount")

        lim = pol["approver_limit"].get(key(x["approver"]))
        if lim is not None and amt > lim:
            flag(x, "over approval authority",
                 f"{amt:,.2f} approved by {x['approver']} whose stated limit is "
                 f"{lim:,.2f}")

        # thresholds
        cap = pol["category_cap"].get(x["category"].lower())
        if cap is not None and amt > cap:
            flag(x, "over limit",
                 f"{amt:,.2f} exceeds the {cap:,.2f} cap for {x['category']}")

        for t in thresholds:
            band = t * threshold_band
            if t - band <= amt < t:
                flag(x, "threshold-adjacent",
                     f"{amt:,.2f} is {t - amt:,.2f} below the {t:,.0f} threshold")
                break

        # documentation
        if pol["receipt_above"] is not None and amt > pol["receipt_above"] \
                and not x["receipt"]:
            flag(x, "missing receipt",
                 f"{amt:,.2f} exceeds the {pol['receipt_above']:,.2f} receipt threshold "
                 f"with no receipt")
        d = x["description"].lower().strip()
        if not d:
            flag(x, "generic business purpose", "no business purpose recorded")
        elif any(d == g or d.startswith(g + " ") and len(d) <= len(g) + 4
                 for g in GENERIC_PURPOSE):
            flag(x, "generic business purpose", f'"{x["description"]}"')

        # category and pattern
        if x["category"].lower() in pol["prohibited"]:
            flag(x, "prohibited category",
                 f"{x['category']} is prohibited by policy")
        if x["txn_date"]:
            weekend = x["txn_date"].weekday() >= 5
            holiday = x["txn_date"] in holidays
            if (weekend or holiday) and pol["weekend_ok"] \
                    and x["category"].lower() not in pol["weekend_ok"]:
                flag(x, "weekend/holiday spend",
                     f"{x['txn_date'].strftime('%A %Y-%m-%d')} on {x['category']}, which "
                     f"is not an allowed weekend category")
        if is_round(amt) and pol["receipt_above"] is not None \
                and pol["receipt_above"] - (pol["receipt_above"] * threshold_band) \
                <= amt <= pol["receipt_above"]:
            flag(x, "round number",
                 f"{amt:,.2f} is a round amount at or just under the receipt threshold")

        # termination
        emp = roster.get(key(x["employee"]))
        if emp and emp["termination_date"] and x["txn_date"] \
                and x["txn_date"] > emp["termination_date"]:
            flag(x, "claim after termination",
                 f"transaction dated {x['txn_date']} after termination "
                 f"{emp['termination_date']}")

    # circular approval
    if roster:
        approves = defaultdict(set)
        for x in rows:
            if x["approver"]:
                approves[key(x["approver"])].add(key(x["employee"]))
        for a, emps in approves.items():
            for e in emps:
                if a in approves.get(e, set()) and a != e:
                    for x in rows:
                        if key(x["approver"]) == a and key(x["employee"]) == e:
                            flag(x, "circular approval",
                                 f"{x['employee']} and {x['approver']} approve each "
                                 f"other's claims")

    # duplicates
    def add_group(pattern, items, kind):
        if len(items) > 1:
            groups.append({"pattern": pattern, "kind": kind, "items": items,
                           "at_risk": sum((i["amount"] for i in items[1:]), ZERO)})
            for i in items:
                flag(i, kind, pattern)

    by_exact = defaultdict(list)
    for x in rows:
        if x["txn_date"]:
            by_exact[(key(x["employee"]), x["amount"], x["txn_date"])].append(x)
    for k, items in by_exact.items():
        if len(items) > 1:
            add_group("same employee, amount and date", items, "duplicate")

    by_amt = defaultdict(list)
    for x in rows:
        if x["txn_date"]:
            by_amt[(key(x["employee"]), x["amount"])].append(x)
    for k, items in by_amt.items():
        if len(items) < 2:
            continue
        items = sorted(items, key=lambda i: i["txn_date"])
        for a in range(len(items)):
            for b in range(a + 1, len(items)):
                gap = (items[b]["txn_date"] - items[a]["txn_date"]).days
                if 0 < gap <= 30 and not any(
                        "same employee, amount and date" in n
                        for n in items[a]["notes"] + items[b]["notes"]):
                    add_group("same employee and amount within 30 days",
                              [items[a], items[b]], "duplicate")

    by_merchant = defaultdict(list)
    for x in rows:
        if x["merchant"] and x["txn_date"]:
            by_merchant[(key(x["merchant"]), x["amount"], x["txn_date"])].append(x)
    for k, items in by_merchant.items():
        emps = {key(i["employee"]) for i in items}
        if len(items) > 1 and len(emps) > 1:
            add_group("same merchant, amount and date claimed by different employees",
                      items, "duplicate")

    # split transactions
    if pol["approval_above"] is not None:
        thr = pol["approval_above"]
        by_day = defaultdict(list)
        for x in rows:
            if x["txn_date"] and x["amount"] < thr:
                by_day[(key(x["employee"]), x["txn_date"],
                        x["category"].lower())].append(x)
        for (emp, d, cat), items in by_day.items():
            total = sum((i["amount"] for i in items), ZERO)
            if len(items) > 1 and total > thr:
                add_group(f"{len(items)} claims on {d} totalling {total:,.2f}, each under "
                          f"the {thr:,.0f} approval threshold",
                          items, "split transaction")

    return groups


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


def rollup(rows, field) -> list[dict]:
    agg = defaultdict(lambda: {"n": 0, "value": ZERO, "exc_n": 0, "exc_value": ZERO,
                               "types": defaultdict(int), "self": 0})
    for x in rows:
        k = x[field] or "(none recorded)"
        a = agg[k]
        a["n"] += 1
        a["value"] += x["amount"]
        if x["exceptions"]:
            a["exc_n"] += 1
            a["exc_value"] += x["amount"]
            for t in x["exceptions"]:
                a["types"][t] += 1
        if "self-approved" in x["exceptions"]:
            a["self"] += 1
    out = []
    for k, a in agg.items():
        dom = max(a["types"].items(), key=lambda kv: kv[1])[0] if a["types"] else ""
        out.append({"name": k, **a, "dominant": dom,
                    "rate": (a["exc_n"] / a["n"] * 100) if a["n"] else 0})
    return sorted(out, key=lambda x: x["exc_value"], reverse=True)


def sheet_summary(wb, meta, stats, proof, emp_roll, appr_roll, obs, pol):
    ws = wb.active
    ws.title = "Summary"
    widths(ws, {1: 4, 2: 46, 3: 14, 4: 18, 5: 60})
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

    line("EXPENSE AND T&E POLICY TESTING", bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    ok = proof["balanced"] and proof["ties"]
    v = ws.cell(row=r, column=2,
                value="POPULATION PROVEN" if ok else
                "POPULATION NOT PROVEN - see the Population Proof tab")
    v.font = OK_FONT if ok else BAD_FONT
    if not ok:
        v.fill = BAD_FILL
    r += 2

    if not pol["supplied"]:
        line("NO WRITTEN POLICY SUPPLIED", bold=True, size=12, fill=BAD_FILL,
             note="This is the headline finding. An expense programme with no written "
                  "policy has no standard to test against and no basis for discipline. "
                  "Testing below used zero thresholds.")
        r += 1

    ws.cell(row=r, column=3, value="Count").font = Font(bold=True)
    ws.cell(row=r, column=4, value="Value").font = Font(bold=True)
    r += 1
    line("Transactions in population", stats["n"], stats["value"], bold=True)
    line("  compliant", stats["compliant_n"], stats["compliant_value"])
    line("  with at least one exception", stats["exc_n"], stats["exc_value"],
         fill=WARN_FILL)
    r += 1
    line("EXCEPTIONS BY TYPE", bold=True, size=12, fill=SUB_FILL)
    for t in EXC_TYPES:
        b = stats["by_type"][t]
        if b["n"]:
            line("  " + t, b["n"], b["value"])
    r += 1
    line("Recovery opportunity", stats["recoverable_n"], stats["recoverable_value"],
         bold=True, fill=WARN_FILL,
         note="Over-limit, duplicate, prohibited, split, and post-termination claims. "
              "Investigate before recovering.")
    r += 1

    if appr_roll:
        line("BY APPROVER - read this first", bold=True, size=12, fill=SUB_FILL)
        for h, col in zip(["Approver", "Approved", "Exception rate", "Exception value",
                           "Self-approved"], (2, 3, 4, 5, 6)):
            c = ws.cell(row=r, column=col, value=h)
            c.fill, c.font = HDR_FILL, HDR_FONT
        r += 1
        for a in appr_roll[:8]:
            ws.cell(row=r, column=2, value=a["name"])
            ws.cell(row=r, column=3, value=a["n"])
            cc = ws.cell(row=r, column=4, value=a["rate"])
            cc.number_format = PCT
            if a["rate"] >= 40:
                cc.fill = BAD_FILL
            cv = ws.cell(row=r, column=5, value=float(a["exc_value"]))
            cv.number_format = MONEY
            cs = ws.cell(row=r, column=6, value=a["self"] or "")
            if a["self"]:
                cs.fill = BAD_FILL
            r += 1
        r += 1
        line("An approver whose exception rate is far above peers is approving without "
             "reading. That explains the other findings and is more valuable than any "
             "individual claim.", bold=True)
        r += 1

    if emp_roll:
        line("BY EMPLOYEE", bold=True, size=12, fill=SUB_FILL)
        for h, col in zip(["Employee", "Claims", "Exceptions", "Exception value",
                           "Dominant type"], (2, 3, 4, 5, 6)):
            c = ws.cell(row=r, column=col, value=h)
            c.fill, c.font = HDR_FILL, HDR_FONT
        r += 1
        for e in emp_roll[:10]:
            ws.cell(row=r, column=2, value=e["name"])
            ws.cell(row=r, column=3, value=e["n"])
            ws.cell(row=r, column=4, value=e["exc_n"])
            cv = ws.cell(row=r, column=5, value=float(e["exc_value"]))
            cv.number_format = MONEY
            ws.cell(row=r, column=6, value=e["dominant"])
            r += 1
        top3 = sum((e["exc_value"] for e in emp_roll[:3]), ZERO)
        share = (top3 / stats["exc_value"] * 100) if stats["exc_value"] else ZERO
        r += 1
        line("Top three employees as a share of exception value", None, None,
             bold=True)
        cc = ws.cell(row=r - 1, column=3, value=float(share))
        cc.number_format = PCT
        ws.cell(row=r - 1, column=5,
                value="If a few people drive most of the value, the programme does not "
                      "have a policy problem - it has a small number of specific "
                      "problems.").font = Font(italic=True, size=9)
        r += 1

    if obs:
        line("ESCALATE / CONTROL OBSERVATIONS", bold=True, size=12, fill=BAD_FILL)
        for o in obs:
            line("  " + o)
        r += 1
    r += 1
    line("Report facts and amounts. Do not characterise intent - an expense finding "
         "lands on an individual, and management and HR handle it through their own "
         "process.", bold=True)
    r += 2
    line("Prepared by / date:  __________________  ____________", bold=True)
    line("Reviewed by / date:  __________________  ____________", bold=True)


def sheet_proof(wb, proof, stats, pol):
    ws = wb.create_sheet("Population Proof")
    widths(ws, {1: 4, 2: 48, 3: 14, 4: 18, 5: 58})
    r = 1
    ws.cell(row=r, column=2, value="POPULATION PROOF").font = Font(bold=True, size=14)
    r += 2
    ws.cell(row=r, column=3, value="Count").font = Font(bold=True)
    ws.cell(row=r, column=4, value="Value").font = Font(bold=True)
    r += 1
    for label, n, v, bold, border in [
        ("Compliant", proof["compliant_n"], proof["compliant_value"], False, None),
        ("With exceptions", proof["exc_n"], proof["exc_value"], False, None),
        ("= Total accounted for", proof["acc_n"], proof["acc_value"], True, TOP),
        ("Population", stats["n"], stats["value"], False, None),
        ("Difference", proof["n_diff"], proof["v_diff"], True, TOP),
    ]:
        ws.cell(row=r, column=2, value=label).font = Font(bold=bold)
        c1 = ws.cell(row=r, column=3, value=n)
        c2 = ws.cell(row=r, column=4, value=float(v))
        c2.number_format = MONEY
        for c in (c1, c2):
            c.font = Font(bold=bold)
            if border:
                c.border = border
        if label == "Difference":
            for c in (c1, c2):
                c.font = OK_FONT if proof["balanced"] else BAD_FONT
                if not proof["balanced"]:
                    c.fill = BAD_FILL
        r += 1
    r += 1
    if proof["gl_total"] is None:
        ws.cell(row=r, column=2,
                value="GL total NOT SUPPLIED - the completeness of the population is "
                      "unproven. A filtered card extract proves nothing about the "
                      "expense line.").fill = BAD_FILL
        r += 1
    else:
        for label, v in [("Population value", stats["value"]),
                         ("GL expense total", proof["gl_total"]),
                         ("Difference", stats["value"] - proof["gl_total"])]:
            ws.cell(row=r, column=2, value=label).font = Font(bold=label == "Difference")
            c = ws.cell(row=r, column=4, value=float(v))
            c.number_format = MONEY
            if label == "Difference":
                c.font = OK_FONT if v == ZERO else BAD_FONT
                if v != ZERO:
                    c.fill = BAD_FILL
            r += 1
    r += 2
    ws.cell(row=r, column=2, value="POLICY RULES APPLIED").font = Font(bold=True, size=12)
    r += 1
    if not pol["rules"]:
        ws.cell(row=r, column=2,
                value="None - no written policy was supplied. Zero thresholds were used "
                      "and the absence of a policy is reported as the headline finding."
                ).fill = BAD_FILL
        return
    for h, col in zip(["Rule", "Category / approver", "Threshold", "Note"],
                      (2, 3, 4, 5)):
        c = ws.cell(row=r, column=col, value=h)
        c.fill, c.font = HDR_FILL, HDR_FONT
    r += 1
    for rule in pol["rules"]:
        ws.cell(row=r, column=2, value=rule["rule"])
        ws.cell(row=r, column=3, value=rule["category"] or None)
        if rule["threshold"] is not None:
            c = ws.cell(row=r, column=4, value=float(rule["threshold"]))
            c.number_format = MONEY
        ws.cell(row=r, column=5, value=rule["note"] or None)
        r += 1


def sheet_rollup(wb, title, roll, label):
    ws = wb.create_sheet(title)
    heads = [label, "Claims", "Value", "Exceptions", "Exception rate",
             "Exception value", "Self-approved", "Dominant type", "Likely cause"]
    ws.append(heads)
    hdr(ws, len(heads))
    causes = {
        "missing receipt": "training or process gap - low severity, high volume",
        "split transaction": "implies knowledge of the approval threshold - escalate",
        "self-approved": "segregation of duties failure in the workflow configuration",
        "duplicate": "submission or intake control gap",
        "generic business purpose": "no substantiation - a documentation discipline issue",
        "over limit": "policy breach - recoverable",
        "missing approval": "workflow not enforced",
        "prohibited category": "policy breach - recoverable",
    }
    for a in roll:
        ws.append([a["name"], a["n"], float(a["value"]), a["exc_n"], a["rate"],
                   float(a["exc_value"]), a["self"] or "", a["dominant"],
                   causes.get(a["dominant"], "")])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[2].number_format = MONEY
        row[4].number_format = PCT
        row[5].number_format = MONEY
        if isinstance(row[4].value, (int, float)) and row[4].value >= 40:
            row[4].fill = BAD_FILL
        if row[6].value:
            row[6].fill = BAD_FILL
    ws.freeze_panes = "B2"
    widths(ws, {1: 28, 2: 9, 3: 15, 4: 12, 5: 14, 6: 16, 7: 14, 8: 24, 9: 56})


def sheet_detail(wb, rows, want_exceptions: bool, title: str):
    ws = wb.create_sheet(title)
    heads = ["Txn", "Date", "Employee", "Amount", "Category", "Merchant",
             "Description", "Approver", "Receipt", "Exception type(s)", "Detail",
             "Investigation", "Disposition"]
    ws.append(heads)
    hdr(ws, len(heads))
    for x in rows:
        has = bool(x["exceptions"])
        if has != want_exceptions:
            continue
        ws.append([x["txn_id"], x["txn_date"], x["employee"], float(x["amount"]),
                   x["category"], x["merchant"] or None, x["description"] or None,
                   x["approver"] or None, "Y" if x["receipt"] else "",
                   ", ".join(x["exceptions"]), "; ".join(x["notes"]), None, None])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[1].number_format = DATEF
        row[3].number_format = MONEY
        if row[9].value:
            row[9].fill = BAD_FILL if any(
                t in str(row[9].value) for t in
                ("split", "self-approved", "duplicate", "termination")) else WARN_FILL
        if not row[8].value:
            row[8].fill = WARN_FILL
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value="None.")
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:M{max(ws.max_row, 2)}"
    widths(ws, {1: 12, 2: 11, 3: 24, 4: 13, 5: 20, 6: 24, 7: 34, 8: 20, 9: 9,
                10: 34, 11: 66, 12: 24, 13: 20})


def sheet_groups(wb, groups):
    ws = wb.create_sheet("Splits and Duplicates")
    heads = ["Group", "Kind", "Pattern", "Txn", "Date", "Employee", "Amount",
             "Merchant", "At risk", "Disposition"]
    ws.append(heads)
    hdr(ws, len(heads))
    for i, g in enumerate(sorted(groups, key=lambda x: -x["at_risk"]), start=1):
        for j, x in enumerate(g["items"]):
            ws.append([i if j == 0 else None, g["kind"] if j == 0 else None,
                       g["pattern"] if j == 0 else None,
                       x["txn_id"], x["txn_date"], x["employee"], float(x["amount"]),
                       x["merchant"] or None,
                       float(g["at_risk"]) if j == 0 else None, None])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[4].number_format = DATEF
        row[6].number_format = MONEY
        row[8].number_format = MONEY
        if row[8].value:
            row[8].font = BAD_FONT
        if row[1].value == "split transaction":
            row[1].fill = BAD_FILL
    if ws.max_row == 1:
        ws.cell(row=2, column=3, value="None detected.")
    else:
        r = ws.max_row + 2
        ws.cell(row=r, column=3, value="TOTAL AT RISK").font = Font(bold=True)
        c = ws.cell(row=r, column=9,
                    value=float(sum((g["at_risk"] for g in groups), ZERO)))
        c.number_format, c.font, c.border = MONEY, Font(bold=True), TOP
        r += 2
        ws.cell(row=r, column=3,
                value="Split transactions are the highest-signal finding here because "
                      "they imply knowledge of the approval threshold. Report the "
                      "pattern, not the single instance - one is a coincidence, a habit "
                      "is not.").alignment = Alignment(wrap_text=True)
    ws.freeze_panes = "D2"
    widths(ws, {1: 7, 2: 20, 3: 52, 4: 12, 5: 11, 6: 24, 7: 13, 8: 24, 9: 14, 10: 20})


# ------------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--expenses", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--gl-total")
    ap.add_argument("--policy")
    ap.add_argument("--roster")
    ap.add_argument("--holidays", help="CSV with a 'date' column")
    ap.add_argument("--threshold-band", default="0.10",
                    help="fraction below a threshold treated as threshold-adjacent "
                         "(default 0.10)")
    ap.add_argument("--client", default="")
    ap.add_argument("--period", default="")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    rows = load_expenses(Path(args.expenses))
    pol = load_policy(Path(args.policy) if args.policy else None)
    roster = load_roster(Path(args.roster) if args.roster else None)
    holidays = set()
    if args.holidays and Path(args.holidays).exists():
        holidays = {pdate(r.get("date")) for r in rows_of(Path(args.holidays),
                                                         "holidays")} - {None}
    band = dec(args.threshold_band) or Decimal("0.10")

    groups = run_tests(rows, pol, roster, holidays, band)

    total_value = sum((x["amount"] for x in rows), ZERO)
    compliant = [x for x in rows if not x["exceptions"]]
    exceptions = [x for x in rows if x["exceptions"]]
    recoverable = [x for x in rows if set(x["exceptions"]) & RECOVERABLE]
    by_type = {t: {"n": 0, "value": ZERO} for t in EXC_TYPES}
    for x in rows:
        for t in x["exceptions"]:
            by_type[t]["n"] += 1
            by_type[t]["value"] += x["amount"]

    gl = dec(args.gl_total)
    proof = {
        "compliant_n": len(compliant),
        "compliant_value": sum((x["amount"] for x in compliant), ZERO),
        "exc_n": len(exceptions),
        "exc_value": sum((x["amount"] for x in exceptions), ZERO),
        "gl_total": gl,
    }
    proof["acc_n"] = proof["compliant_n"] + proof["exc_n"]
    proof["acc_value"] = proof["compliant_value"] + proof["exc_value"]
    proof["n_diff"] = proof["acc_n"] - len(rows)
    proof["v_diff"] = proof["acc_value"] - total_value
    proof["balanced"] = proof["n_diff"] == 0 and proof["v_diff"] == ZERO
    proof["ties"] = gl is None or total_value == gl

    stats = {"n": len(rows), "value": total_value,
             "compliant_n": len(compliant), "compliant_value": proof["compliant_value"],
             "exc_n": len(exceptions), "exc_value": proof["exc_value"],
             "by_type": by_type,
             "recoverable_n": len(recoverable),
             "recoverable_value": sum((x["amount"] for x in recoverable), ZERO)}

    emp_roll = rollup(rows, "employee")
    appr_roll = [a for a in rollup(rows, "approver") if a["name"] != "(none recorded)"]

    obs = []
    if not pol["supplied"]:
        obs.append("No written expense policy was supplied. This is the headline finding: "
                   "a programme with no written policy has no standard to test against "
                   "and no basis for discipline.")
    splits = [g for g in groups if g["kind"] == "split transaction"]
    if splits:
        by_emp = defaultdict(int)
        for g in splits:
            by_emp[g["items"][0]["employee"]] += 1
        repeat = {k: v for k, v in by_emp.items() if v > 1}
        obs.append(f"{len(splits)} split-transaction group(s) totalling "
                   f"{sum((g['at_risk'] for g in splits), ZERO):,.2f} at risk."
                   + (f" Repeated by: {', '.join(f'{k} ({v})' for k, v in repeat.items())}"
                      f" - a habit, not a coincidence." if repeat else ""))
    selfs = [x for x in rows if "self-approved" in x["exceptions"]]
    if selfs:
        obs.append(f"{len(selfs)} claim(s) approved by the claimant, totalling "
                   f"{sum((x['amount'] for x in selfs), ZERO):,.2f}. A segregation of "
                   f"duties failure in the workflow configuration.")
    for a in appr_roll:
        if a["rate"] >= 40 and a["n"] >= 5:
            obs.append(f"Approver {a['name']}: {a['rate']:.0f}% of {a['n']} claims "
                       f"approved carry an exception. Approving without reading - this "
                       f"explains other findings.")
    terms = [x for x in rows if "claim after termination" in x["exceptions"]]
    if terms:
        obs.append(f"{len(terms)} claim(s) dated after the claimant's termination date, "
                   f"totalling {sum((x['amount'] for x in terms), ZERO):,.2f}.")
    cross = [g for g in groups
             if "different employees" in g["pattern"]]
    if cross:
        obs.append(f"{len(cross)} instance(s) of the same merchant, amount and date "
                   f"claimed by different employees - the same receipt claimed twice.")
    if not proof["balanced"]:
        obs.append("POPULATION DOES NOT BALANCE - compliant plus exceptions does not "
                   "equal the population.")

    # ---- console
    print("=" * 76)
    print("EXPENSE AND T&E POLICY TESTING")
    print("=" * 76)
    print(f"Client : {args.client or '(not stated)'}   Period: {args.period or '(not stated)'}")
    print(f"Transactions {len(rows):,}   value {total_value:>16,.2f}")
    if gl is not None:
        d = total_value - gl
        print(f"GL total     {gl:>16,.2f}   difference {d:>14,.2f}  "
              f"{'TIES' if d == ZERO else '*** DOES NOT TIE ***'}")
    else:
        print("GL total     NOT SUPPLIED")
    if not pol["supplied"]:
        print("\n*** NO WRITTEN POLICY SUPPLIED - this is the headline finding ***")
    print()
    print(f"Compliant {len(compliant):,}   with exceptions {len(exceptions):,}")
    for t in EXC_TYPES:
        b = by_type[t]
        if b["n"]:
            print(f"  {t:<26} {b['n']:>5}   {b['value']:>15,.2f}")
    print(f"\nPopulation proof: {proof['acc_n']} of {len(rows)}  "
          f"{'BALANCED' if proof['balanced'] else '*** DOES NOT BALANCE ***'}")
    print(f"Recovery opportunity: {stats['recoverable_value']:,.2f} "
          f"({stats['recoverable_n']} claims)")
    if appr_roll:
        print("\nBy approver (read first):")
        for a in appr_roll[:6]:
            print(f"  {a['name'][:26]:<26} {a['n']:>4} approved  "
                  f"{a['rate']:>5.0f}% exceptions  {a['exc_value']:>13,.2f}"
                  f"{'  SELF-APPROVED ' + str(a['self']) if a['self'] else ''}")
    if emp_roll:
        print("\nBy employee:")
        for e in emp_roll[:6]:
            print(f"  {e['name'][:26]:<26} {e['exc_n']:>4} exc  "
                  f"{e['exc_value']:>13,.2f}  mostly {e['dominant']}")
    if obs:
        print("\nESCALATE / OBSERVATIONS:")
        for o in obs:
            print(f"  ! {o}")

    if not proof["balanced"] and not args.force:
        print("\n" + "=" * 76)
        print("WORKBOOK NOT WRITTEN. Every transaction must be either compliant or an")
        print("exception. A transaction accounted for by neither has fallen out of the")
        print("test.")
        print("=" * 76)
        return 1

    meta = {
        "Client": args.client or "(not stated)",
        "Period": args.period or "(not stated)",
        "Transactions": len(rows),
        "Written policy supplied": "yes" if pol["supplied"] else "NO",
        "Roster supplied": "yes" if roster else "no",
        "Threshold-adjacent band": f"{band * 100:.0f}% below a threshold",
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    wb = Workbook()
    sheet_summary(wb, meta, stats, proof, emp_roll, appr_roll, obs, pol)
    sheet_proof(wb, proof, stats, pol)
    sheet_rollup(wb, "By Approver", appr_roll, "Approver")
    sheet_rollup(wb, "By Employee", emp_roll, "Employee")
    sheet_detail(wb, rows, True, "Exception Detail")
    sheet_groups(wb, groups)
    sheet_detail(wb, rows, False, "Compliant Detail")
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
