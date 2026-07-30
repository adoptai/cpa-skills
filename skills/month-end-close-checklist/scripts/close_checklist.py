#!/usr/bin/env python3
"""
Month-end close checklist DERIVED from the trial balance, not from a template.

A generic checklist is a list of tasks somebody else's business needed. It says
nothing about the accounts on THIS trial balance that moved this period and nobody
has looked at. The usual close failure is not a forgotten task - it is an account
nobody considered, because it was not on the template.

Gates:
  * every account with activity is reconciled, waived with a reason, or open -
    and reconciled + waived + open + blocked must equal that population
  * nothing above materiality is waived without a reason
  * every open item has an owner and a due date
  * every reconciled account names its support
  * the trial balance must balance - if it does not, the close has not started

Also assigns the right PROCEDURE per account based on what the account actually is,
and names the skill in this repository that performs it.

Usage:
    python3 close_checklist.py --trial-balance tb.csv --materiality 25000 \
        --prior-checklist prior_close.csv \
        --period 2025-11 --target-close-date 2025-12-05 \
        --client "Halstead Regional Foods" \
        --out "Halstead - 2025-11 Close Checklist.xlsx"

--trial-balance CSV (same schema trial-balance-integrity uses):
    account, name, debit, credit, type, contra, prior_debit, prior_credit

--prior-checklist CSV:
    account, status, owner, reason, periods_open, periods_waived
--status CSV (this period's statuses, as they are worked):
    account, status, owner, due_date, support_reference, reason, blocker, notes
      status: reconciled | waived | open | blocked
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

STATUSES = ("reconciled", "waived", "open", "blocked")

# Procedure assignment: (keywords, procedure, skill in this repository)
PROCEDURES = [
    (("cash", "bank", "checking", "money market", "operating account", "petty"),
     "Reconcile to the bank statement; propose journal entries for unrecorded bank items",
     "bank-rec-to-gl"),
    (("receivable", "a/r", "unbilled", "contract asset"),
     "Tie the aging to the control account; recompute every ageing bucket from the "
     "invoice date",
     "ar-aging-tie-out"),
    (("allowance", "doubtful", "bad debt"),
     "Assess the allowance against the aging profile and subsequent collections",
     "ar-aging-tie-out"),
    (("payable", "a/p", "accrued", "accrual"),
     "Search for unrecorded liabilities; classify on the service date",
     "cutoff-and-unrecorded-liabilities"),
    (("fixed asset", "machinery", "equipment", "property", "plant", "leasehold",
      "vehicle", "furniture", "accumulated depreciation", "cip",
      "construction in progress"),
     "Roll forward cost and accumulated depreciation; agree beginning to the prior "
     "period as filed",
     "depreciation-tie-out"),
    (("depreciation expense", "amortization expense", "amortisation expense"),
     "Recompute depreciation and agree to the fixed asset register",
     "depreciation-tie-out"),
    (("wage", "salar", "payroll", "bonus", "commission", "fica", "medicare",
      "withholding", "941"),
     "Reconcile the payroll register to the filings and the general ledger",
     "payroll-tax-reconciliation"),
    (("sales tax", "use tax", "vat", "gst"),
     "Reconcile filed returns to the ledger; roll forward the liability",
     "sales-tax-reconciliation"),
    (("revenue", "sales", "fees earned", "service income"),
     "Test cutoff and occurrence; trace a sample to invoice, contract and cash",
     "revenue-trace-to-source"),
    (("inventor", "cost of goods", "cogs"),
     "Agree to the physical count or roll forward; test cutoff on receipts",
     "cutoff-and-unrecorded-liabilities"),
    (("travel", "entertainment", "meals", "card", "expense reimbursement"),
     "Test the population against written expense policy",
     "expense-policy-testing"),
    (("suspense", "clearing", "conversion", "unallocated", "temporary", "holding"),
     "MUST BE CLEARED - a suspense or clearing account with a balance at close means "
     "something was parked and the close moved on without it",
     "trial-balance-integrity"),
    (("debt", "loan", "note payable", "line of credit", "interest payable"),
     "Agree to the lender statement and the amortisation schedule; recompute interest",
     None),
    (("prepaid", "deferred charge"),
     "Roll forward and confirm amortisation of the prepayment",
     None),
    (("deferred revenue", "unearned"),
     "Roll forward; confirm recognition matches performance",
     "revenue-trace-to-source"),
    (("equity", "capital", "retained earnings", "distribution", "dividend",
      "common stock", "member"),
     "Roll forward; agree movement to net income and distributions",
     "trial-balance-integrity"),
    (("intercompany", "due from", "due to", "affiliate", "related party"),
     "Agree to the counterparty balance and eliminate",
     None),
]

DEFAULT_PROCEDURE = ("Reconcile the balance to supporting detail, or waive with a "
                     "documented reason")


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


def assign_procedure(name: str, acct_type: str) -> tuple[str, str | None]:
    low = name.lower()
    for words, proc, skill in PROCEDURES:
        if any(w in low for w in words):
            return proc, skill
    return DEFAULT_PROCEDURE, None


def load_tb(path: Path) -> list[dict]:
    out = []
    for i, r in enumerate(rows_of(path, "trial balance"), start=1):
        dr, cr = d0(r.get("debit")), d0(r.get("credit"))
        pdr, pcr = dec(r.get("prior_debit")), dec(r.get("prior_credit"))
        prior = ((pdr or ZERO) - (pcr or ZERO)) if (pdr is not None or pcr is not None) \
            else None
        net = dr - cr
        name = clean(r.get("name"))
        proc, skill = assign_procedure(name, clean(r.get("type")))
        out.append({
            "row": i,
            "account": clean(r.get("account")) or f"(no number {i})",
            "name": name,
            "type": clean(r.get("type")).lower(),
            "debit": dr, "credit": cr, "net": net,
            "prior_net": prior,
            "movement": (net - prior) if prior is not None else None,
            "procedure": proc, "skill": skill,
            "status": "", "owner": "", "due_date": None,
            "support": "", "reason": "", "blocker": "", "notes": "",
            "periods_open": 0, "periods_waived": 0, "flags": [],
        })
    return out


def load_status(path: Path | None) -> dict:
    out = {}
    if not path or not path.exists():
        return out
    for r in rows_of(path, "status"):
        a = clean(r.get("account"))
        if not a:
            continue
        st = clean(r.get("status")).lower()
        if st and st not in STATUSES:
            sys.exit(f"Status row for account {a}: '{st}' is not a recognised status. "
                     f"Valid statuses are {list(STATUSES)}.")
        out[a] = {"status": st, "owner": clean(r.get("owner")),
                  "due_date": pdate(r.get("due_date")),
                  "support": clean(r.get("support_reference")),
                  "reason": clean(r.get("reason")),
                  "blocker": clean(r.get("blocker")),
                  "notes": clean(r.get("notes"))}
    return out


def load_prior(path: Path | None) -> dict:
    out = {}
    if not path or not path.exists():
        return out
    for r in rows_of(path, "prior checklist"):
        a = clean(r.get("account"))
        if not a:
            continue
        out[a] = {"status": clean(r.get("status")).lower(),
                  "owner": clean(r.get("owner")),
                  "periods_open": int(d0(r.get("periods_open"))),
                  "periods_waived": int(d0(r.get("periods_waived")))}
    return out


# ----------------------------------------------------------------------- logic

def apply_status(tb, statuses, prior, materiality: Decimal | None, target: date | None,
                 today: date) -> None:
    for a in tb:
        s = statuses.get(a["account"], {})
        p = prior.get(a["account"], {})
        # An account absent from the status file entirely is NOT considered - it is not
        # silently defaulted to "open". That distinction is the whole point: the usual
        # close failure is an account nobody considered, and defaulting it to a real
        # status would hide exactly what this is meant to surface.
        a["considered"] = a["account"] in statuses
        a["status"] = s.get("status", "") if a["considered"] else ""
        a["owner"] = s.get("owner") or p.get("owner", "")
        a["due_date"] = s.get("due_date")
        a["support"] = s.get("support", "")
        a["reason"] = s.get("reason", "")
        a["blocker"] = s.get("blocker", "")
        a["notes"] = s.get("notes", "")
        a["periods_open"] = p.get("periods_open", 0) + (1 if a["status"] == "open" else 0)
        a["periods_waived"] = p.get("periods_waived", 0) + (1 if a["status"] == "waived" else 0)

        material = materiality is None or abs(a["net"]) >= materiality

        if a["status"] == "reconciled" and not a["support"]:
            a["flags"].append("marked reconciled with no support reference - that is an "
                              "assertion, not a reconciliation")
        if a["status"] == "waived":
            if not a["reason"]:
                a["flags"].append("WAIVED WITH NO REASON - a waiver is a documented "
                                  "judgement. A blank reason is an omission, not a waiver")
            if material and materiality is not None:
                a["flags"].append(f"waived at {abs(a['net']):,.2f}, at or above "
                                  f"materiality of {materiality:,.2f}")
            if a["periods_waived"] >= 3:
                a["flags"].append(f"waived in {a['periods_waived']} consecutive periods - "
                                  f"this is how an account disappears from the close "
                                  f"permanently")
        if a["status"] == "open":
            if not a["owner"]:
                a["flags"].append("OPEN WITH NO OWNER - it does not close")
            if not a["due_date"]:
                a["flags"].append("open with no due date")
            elif target and a["due_date"] > target:
                a["flags"].append(f"due {a['due_date']}, after the target close date "
                                  f"{target}")
            if a["periods_open"] >= 3:
                a["flags"].append(f"open at {a['periods_open']} consecutive closes - not a "
                                  f"busy month, a process gap")
        if a["status"] == "blocked" and not a["blocker"]:
            a["flags"].append("BLOCKED with no named blocker - that is an open item "
                              "wearing a disguise")
        if not a["status"]:
            if a["considered"]:
                a["flags"].append("appears on the status file with a blank status - decide "
                                  "reconciled, waived, open or blocked")
            else:
                a["flags"].append("NOT CONSIDERED - this account moved and does not appear "
                                  "on the close checklist at all. This is the usual close "
                                  "failure: not a forgotten task, an account nobody knew "
                                  "to look at")
        if a["prior_net"] is None:
            a["flags"].append("no prior-period balance supplied, so movement could not be "
                              "identified for this account")
        if a["account"] not in prior and prior and a["net"] != ZERO:
            a["flags"].append("did not appear on the prior close checklist - a new account, "
                              "or one that appeared without anyone noticing. This is where "
                              "misposting lands")
        low = a["name"].lower()
        if any(w in low for w in ("suspense", "clearing", "conversion", "unallocated",
                                  "temporary", "holding")) and a["net"] != ZERO:
            a["flags"].append(f"suspense/clearing account with a balance of {a['net']:,.2f} "
                              f"at close - something was parked and the close moved on")


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


def sheet_status(wb, meta, tests, stats, active, esc, overall, target, today, owners):
    ws = wb.active
    ws.title = "Close Status"
    widths(ws, {1: 4, 2: 48, 3: 14, 4: 18, 5: 62})
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

    line("MONTH-END CLOSE STATUS", bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    v = ws.cell(row=r, column=2,
                value="CLOSE COMPLETE - every account with activity is reconciled or "
                      "waived with a reason"
                if overall else
                "CLOSE NOT COMPLETE - items remain open. Issuing statements now means "
                "issuing them on an incomplete basis; make that a decision rather than "
                "an accident.")
    v.font = OK_FONT if overall else BAD_FONT
    if not overall:
        v.fill = BAD_FILL
    r += 2

    if target:
        days = (today - target).days
        line("Target close date", str(target), bold=True)
        line("Days against target", days, bold=True,
             fill=BAD_FILL if days > 0 else None,
             note=f"{days} day(s) past target" if days > 0 else "on or ahead of target")
    r += 1

    ws.cell(row=r, column=3, value="Count").font = Font(bold=True)
    ws.cell(row=r, column=4, value="Value").font = Font(bold=True)
    r += 1
    line("Accounts with activity", stats["active_n"], stats["active_value"], bold=True)
    for st in STATUSES:
        b = stats["by_status"][st]
        line("  " + st, b["n"], b["value"],
             fill=BAD_FILL if st in ("open", "blocked") and b["n"] else None)
    b = stats["by_status"][""]
    if b["n"]:
        line("  NO STATUS AT ALL", b["n"], b["value"], fill=BAD_FILL,
             note="These accounts moved and nobody has considered them. This is the usual "
                  "close failure - not a forgotten task, an account that was never on the "
                  "template.")
    r += 1

    for t in tests:
        line(f"TEST {t['num']} - {t['name']}", None, None,
             fill=None if t["passed"] else BAD_FILL, note=t.get("note", ""))
        sc = ws.cell(row=r - 1, column=3, value="PASS" if t["passed"] else "FAIL")
        sc.font = OK_FONT if t["passed"] else BAD_FONT
    r += 1

    if owners:
        line("OPEN ITEMS BY OWNER", bold=True, size=12, fill=SUB_FILL)
        for o, d in owners:
            line("  " + (o or "(NO OWNER)"), d["n"], d["value"],
                 fill=BAD_FILL if not o else None)
        top = owners[0] if owners else None
        if top and stats["by_status"]["open"]["n"] and \
                top[1]["n"] / max(stats["by_status"]["open"]["n"], 1) > 0.5:
            line(f"  {top[0] or '(NO OWNER)'} owns more than half the open items - the "
                 f"close is single-threaded and the date is at risk regardless of effort.",
                 bold=True)
        r += 1

    if esc:
        line("ESCALATE", bold=True, size=12, fill=BAD_FILL)
        for e in esc:
            line("  " + e)
        r += 1
    r += 1
    line("Prepared by / date:  __________________  ____________", bold=True)
    line("Reviewed by / date:  __________________  ____________", bold=True)


def sheet_checklist(wb, active):
    ws = wb.create_sheet("Checklist")
    heads = ["Account", "Name", "Balance", "Prior balance", "Movement", "Status",
             "Procedure required", "Skill to run", "Support reference", "Owner",
             "Due date", "Periods open", "Periods waived", "Reason / blocker", "Flags"]
    ws.append(heads)
    hdr(ws, len(heads))
    order = {"": 0, "blocked": 1, "open": 2, "waived": 3, "reconciled": 4}
    for a in sorted(active, key=lambda x: (order.get(x["status"], 5), -abs(x["net"]))):
        ws.append([a["account"], a["name"], float(a["net"]),
                   float(a["prior_net"]) if a["prior_net"] is not None else None,
                   float(a["movement"]) if a["movement"] is not None else None,
                   a["status"] or "NO STATUS", a["procedure"], a["skill"] or "",
                   a["support"] or None, a["owner"] or None, a["due_date"],
                   a["periods_open"] or None, a["periods_waived"] or None,
                   a["reason"] or a["blocker"] or None, "; ".join(a["flags"])])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in (2, 3, 4):
            row[i].number_format = MONEY
        row[10].number_format = DATEF
        st = row[5].value
        if st == "NO STATUS":
            row[5].font, row[5].fill = BAD_FONT, BAD_FILL
        elif st in ("open", "blocked"):
            row[5].fill = WARN_FILL
        elif st == "reconciled":
            row[5].font = OK_FONT
        if row[14].value:
            row[14].fill = WARN_FILL
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:O{max(ws.max_row, 2)}"
    widths(ws, {1: 12, 2: 34, 3: 15, 4: 15, 5: 15, 6: 12, 7: 56, 8: 30, 9: 24,
                10: 18, 11: 12, 12: 12, 13: 13, 14: 34, 15: 66})


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
    widths(ws, {i: 20 for i in range(1, len(header) + 1)} | {2: 34,
                                                            len(header): 62})


def sheet_coverage(wb, stats, active):
    ws = wb.create_sheet("Coverage Proof")
    widths(ws, {1: 4, 2: 44, 3: 14, 4: 18, 5: 58})
    r = 1
    ws.cell(row=r, column=2, value="COVERAGE PROOF").font = Font(bold=True, size=14)
    r += 2
    ws.cell(row=r, column=3, value="Count").font = Font(bold=True)
    ws.cell(row=r, column=4, value="Value").font = Font(bold=True)
    r += 1
    tot_n = tot_v = 0
    for st in list(STATUSES) + [""]:
        b = stats["by_status"][st]
        if not b["n"] and st == "":
            continue
        ws.cell(row=r, column=2, value="  " + (st or "NO STATUS"))
        ws.cell(row=r, column=3, value=b["n"])
        c = ws.cell(row=r, column=4, value=float(b["value"]))
        c.number_format = MONEY
        tot_n += b["n"]
        tot_v += b["value"]
        r += 1
    ws.cell(row=r, column=2, value="= Total accounted for").font = Font(bold=True)
    c1 = ws.cell(row=r, column=3, value=tot_n)
    c2 = ws.cell(row=r, column=4, value=float(tot_v))
    c2.number_format = MONEY
    for c in (c1, c2):
        c.font, c.border = Font(bold=True), TOP
    r += 1
    ws.cell(row=r, column=2, value="Accounts with activity")
    ws.cell(row=r, column=3, value=stats["active_n"])
    c = ws.cell(row=r, column=4, value=float(stats["active_value"]))
    c.number_format = MONEY
    r += 1
    ws.cell(row=r, column=2, value="Difference").font = Font(bold=True)
    ok = tot_n == stats["active_n"]
    d = ws.cell(row=r, column=3, value=tot_n - stats["active_n"])
    d.font = OK_FONT if ok else BAD_FONT
    if not ok:
        d.fill = BAD_FILL
    r += 2
    ws.cell(row=r, column=2,
            value="This is a completeness test, not a to-do list. Every account that moved "
                  "must be in one of the statuses above. An account in none of them is one "
                  "nobody considered - which is the usual close failure."
            ).alignment = Alignment(wrap_text=True)


# ------------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--trial-balance", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--status")
    ap.add_argument("--prior-checklist")
    ap.add_argument("--materiality")
    ap.add_argument("--period", default="")
    ap.add_argument("--target-close-date")
    ap.add_argument("--as-of", help="today's date for lateness, defaults to the system date")
    ap.add_argument("--client", default="")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    tb = load_tb(Path(args.trial_balance))
    statuses = load_status(Path(args.status) if args.status else None)
    prior = load_prior(Path(args.prior_checklist) if args.prior_checklist else None)
    materiality = dec(args.materiality)
    target = pdate(args.target_close_date) if args.target_close_date else None
    today = pdate(args.as_of) if args.as_of else date.today()

    dr = sum((a["debit"] for a in tb), ZERO)
    cr = sum((a["credit"] for a in tb), ZERO)
    if dr != cr:
        print("=" * 76)
        print("THE TRIAL BALANCE DOES NOT BALANCE")
        print("=" * 76)
        print(f"  debits {dr:,.2f}   credits {cr:,.2f}   difference {dr - cr:,.2f}")
        print("\nThe close has not started. Run trial-balance-integrity first - a close")
        print("checklist built on an unbalanced trial balance is meaningless.")
        print("=" * 76)
        if not args.force:
            return 1

    apply_status(tb, statuses, prior, materiality, target, today)

    active = [a for a in tb
              if a["net"] != ZERO or (a["movement"] not in (None, ZERO))]
    by_status = {st: {"n": 0, "value": ZERO} for st in list(STATUSES) + [""]}
    for a in active:
        b = by_status[a["status"]]
        b["n"] += 1
        b["value"] += abs(a["net"])

    stats = {"active_n": len(active),
             "active_value": sum((abs(a["net"]) for a in active), ZERO),
             "by_status": by_status}

    open_items = [a for a in active if a["status"] in ("open", "blocked", "")]
    no_owner = [a for a in open_items if not a["owner"]]
    waived_no_reason = [a for a in active
                        if a["status"] == "waived" and not a["reason"]]
    waived_material = [a for a in active if a["status"] == "waived"
                       and materiality is not None and abs(a["net"]) >= materiality]
    recon_no_support = [a for a in active
                        if a["status"] == "reconciled" and not a["support"]]
    no_status = [a for a in active if not a["status"]]
    recurring_open = [a for a in active if a["periods_open"] >= 3]
    recurring_waived = [a for a in active if a["periods_waived"] >= 3]
    suspense = [a for a in active
                if any(w in a["name"].lower() for w in
                       ("suspense", "clearing", "conversion", "unallocated",
                        "temporary", "holding")) and a["net"] != ZERO]
    new_accounts = [a for a in active if prior and a["account"] not in prior
                    and a["net"] != ZERO]

    owners_agg = defaultdict(lambda: {"n": 0, "value": ZERO})
    for a in open_items:
        o = owners_agg[a["owner"]]
        o["n"] += 1
        o["value"] += abs(a["net"])
    owners = sorted(owners_agg.items(), key=lambda kv: -kv[1]["n"])

    tests = [
        {"num": 1, "name": "Every account with activity is accounted for",
         "passed": not no_status,
         "note": "" if not no_status else
                 f"{len(no_status)} account(s) moved with no status - nobody has "
                 f"considered them"},
        {"num": 2, "name": "Nothing above materiality waived without a reason",
         "passed": not waived_no_reason and not waived_material,
         "note": "" if not (waived_no_reason or waived_material) else
                 f"{len(waived_no_reason)} waived with no reason, "
                 f"{len(waived_material)} waived at or above materiality"},
        {"num": 3, "name": "Every open item has an owner and a due date",
         "passed": not no_owner,
         "note": "" if not no_owner else
                 f"{len(no_owner)} open item(s) with no owner - they do not close"},
        {"num": 4, "name": "Every reconciled account names its support",
         "passed": not recon_no_support,
         "note": "" if not recon_no_support else
                 f"{len(recon_no_support)} marked reconciled with no support reference"},
        {"num": 5, "name": "Nothing left open",
         "passed": not open_items,
         "note": "" if not open_items else
                 f"{len(open_items)} item(s) open or blocked, "
                 f"{sum((abs(a['net']) for a in open_items), ZERO):,.2f}"},
    ]

    esc = []
    for a in no_status:
        esc.append(f"{a['account']} {a['name']} ({a['net']:,.2f}): moved and is NOT ON THE "
                   f"CLOSE CHECKLIST at all. Nobody knew to look at it - assign an owner "
                   f"and a procedure. Suggested: {a['procedure'][:70]}")
    for a in suspense:
        esc.append(f"{a['account']} {a['name']}: {a['net']:,.2f} at close. Suspense and "
                   f"clearing accounts are meant to be empty - something was parked and "
                   f"the close moved on without it.")
    for a in recurring_open:
        esc.append(f"{a['account']} {a['name']}: open at {a['periods_open']} consecutive "
                   f"closes. Not a busy month, a process gap.")
    for a in recurring_waived:
        esc.append(f"{a['account']} {a['name']}: waived in {a['periods_waived']} "
                   f"consecutive periods. This is how an account disappears from the "
                   f"close permanently.")
    for a in waived_material:
        esc.append(f"{a['account']} {a['name']}: waived at {abs(a['net']):,.2f}, at or "
                   f"above materiality.")
    for a in no_owner:
        esc.append(f"{a['account']} {a['name']}: open with no owner.")
    for a in new_accounts:
        esc.append(f"{a['account']} {a['name']}: was not on the prior close checklist. A "
                   f"new account, or one that appeared without anyone noticing - this is "
                   f"where misposting lands.")
    if target and (today - target).days > 0 and open_items:
        esc.append(f"The close is {(today - target).days} day(s) past the target date with "
                   f"{len(open_items)} item(s) still open.")

    overall = all(t["passed"] for t in tests)

    meta = {
        "Client": args.client or "(not stated)",
        "Period": args.period or "(not stated)",
        "Accounts in trial balance": len(tb),
        "Accounts with activity": len(active),
        "Materiality": float(materiality) if materiality is not None else "(not supplied)",
        "Target close date": str(target) if target else "(not supplied)",
        "As of": str(today),
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    # ---- console
    print("=" * 76)
    print("MONTH-END CLOSE CHECKLIST")
    print("=" * 76)
    print(f"Client : {meta['Client']}   Period: {meta['Period']}")
    print(f"Accounts with activity: {len(active)} of {len(tb)}")
    if target:
        d = (today - today).days if not target else (today - target).days
        print(f"Target close {target}   as of {today}   "
              f"{'ON TRACK' if d <= 0 else f'*** {d} DAY(S) LATE ***'}")
    print()
    for st in list(STATUSES) + [""]:
        b = by_status[st]
        if b["n"]:
            print(f"  {(st or 'NO STATUS'):<14} {b['n']:>5}   {b['value']:>15,.2f}")
    print()
    for t in tests:
        print(f"Test {t['num']}: {'PASS' if t['passed'] else '*** FAIL ***':<14} {t['name']}")
        if t.get("note"):
            print(f"         {t['note']}")
    print()
    print("Procedures assigned from the accounts themselves:")
    by_skill = defaultdict(int)
    for a in active:
        if a["skill"]:
            by_skill[a["skill"]] += 1
    for sk, n in sorted(by_skill.items(), key=lambda kv: -kv[1]):
        print(f"  {sk:<36} {n:>3} account(s)")
    unassigned = len([a for a in active if not a["skill"]])
    if unassigned:
        print(f"  {'(no skill - reconcile manually)':<36} {unassigned:>3} account(s)")
    if owners:
        print("\nOpen items by owner:")
        for o, d in owners[:8]:
            print(f"  {(o or '(NO OWNER)'):<26} {d['n']:>4} item(s)  {d['value']:>14,.2f}")
    if esc:
        print(f"\nESCALATE ({len(esc)}):")
        for e in esc[:15]:
            print(f"  ! {e}")
        if len(esc) > 15:
            print(f"  ... and {len(esc) - 15} more")

    if not overall and not args.force:
        print("\n" + "=" * 76)
        print("CLOSE NOT COMPLETE.")
        print("This is a completeness test, not a to-do list. Every account that moved must")
        print("be reconciled or waived with a reason. Issuing statements with items open")
        print("means issuing them on an incomplete basis - make that a decision, not an")
        print("accident. Re-run with --status once the items are worked.")
        print("=" * 76)

    wb = Workbook()
    sheet_status(wb, meta, tests, stats, active, esc, overall, target, today, owners)
    sheet_checklist(wb, active)
    sheet_list(wb, "Open Items",
               ["Account", "Name", "Balance", "Status", "Owner", "Due date",
                "Periods open", "Procedure required", "Flags"],
               [[a["account"], a["name"], float(a["net"]), a["status"] or "NO STATUS",
                 a["owner"] or "(NO OWNER)", a["due_date"], a["periods_open"] or None,
                 a["procedure"], "; ".join(a["flags"])]
                for a in sorted(open_items, key=lambda x: -abs(x["net"]))],
               {2: MONEY, 5: DATEF},
               "None - nothing is open.",
               footer="Sorted by materiality. An open item with no owner does not close.")
    sheet_list(wb, "Waivers",
               ["Account", "Name", "Balance", "Reason", "Consecutive periods waived",
                "Flags"],
               [[a["account"], a["name"], float(a["net"]),
                 a["reason"] or "(NO REASON GIVEN)", a["periods_waived"] or None,
                 "; ".join(a["flags"])]
                for a in active if a["status"] == "waived"],
               {2: MONEY},
               "No accounts waived this period.",
               footer="A waiver is a documented judgement that the account does not warrant "
                      "a reconciliation this period. Immaterial and dormant accounts "
                      "genuinely do not. A waiver without a reason is how an account "
                      "disappears from the close permanently.")
    sheet_coverage(wb, stats, active)
    sheet_list(wb, "Recurring Issues",
               ["Account", "Name", "Balance", "Status", "Periods open", "Periods waived",
                "Note"],
               [[a["account"], a["name"], float(a["net"]), a["status"],
                 a["periods_open"] or None, a["periods_waived"] or None,
                 "; ".join(a["flags"])]
                for a in active if a["periods_open"] >= 2 or a["periods_waived"] >= 2],
               {2: MONEY},
               "None - no account has been open or waived across consecutive periods.",
               footer="An item recurring across closes is a process gap, not a busy month.")
    sheet_list(wb, "Next Period Carry-Forward",
               ["account", "status", "owner", "reason", "periods_open", "periods_waived"],
               [[a["account"], "", a["owner"], a["reason"],
                 a["periods_open"] if a["status"] == "open" else 0,
                 a["periods_waived"] if a["status"] == "waived" else 0]
                for a in active],
               {},
               "No accounts to carry forward.",
               footer="Save as CSV and pass to --prior-checklist next period. Owners are "
                      "retained and recurring items stay visible, so the next close starts "
                      "from evidence rather than a blank template.")
    wb.active = 0
    out = Path(args.out)
    if not overall:
        out = out.with_name(out.stem + " [OPEN ITEMS]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\n{'CLOSE COMPLETE' if overall else 'OPEN ITEMS REMAIN'}.  Workbook: {out}")
    return 0 if overall else 1


if __name__ == "__main__":
    sys.exit(main())
