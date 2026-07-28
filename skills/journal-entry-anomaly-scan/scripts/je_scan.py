#!/usr/bin/env python3
"""
Full-population journal entry anomaly scan and audit workpaper builder.

Proves the population first (every entry balances; entry-number continuity;
field population), then runs the anomaly tests and ranks entries by accumulated
risk score. Converging flags are the signal; a single flag rarely is.

Outputs attributes and facts only. It does not conclude on intent.

Usage:
    python3 je_scan.py --entries gl.csv --validate-only

    python3 je_scan.py --entries gl.csv \
        --period-start 2025-01-01 --period-end 2025-12-31 \
        --materiality 75000 --approval-thresholds "10000,50000,250000" \
        --business-hours 7-19 --client "Acme Holdings LLC" \
        --out "Acme - FY2025 JE Scan Workpaper.xlsx"

Input CSV - one row per journal entry LINE:
    entry_id (req), line_no, effective_date (req), posted_timestamp,
    account (req), account_name, debit, credit, description,
    posted_by, approved_by, source, reversal_of
"""

from __future__ import annotations

import argparse
import csv
import math
import re
import sys
from collections import Counter, defaultdict
from datetime import date, datetime, time
from decimal import Decimal, InvalidOperation
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is required.  pip install openpyxl")

ZERO = Decimal("0.00")
MONEY = '#,##0.00;[Red](#,##0.00)'
PCT = '0.0"%"'
DATEF = "yyyy-mm-dd"
DTF = "yyyy-mm-dd hh:mm"

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
INFO_FILL = PatternFill("solid", fgColor="DDEBF7")

# --- test catalogue: code -> (weight, label)
TESTS = {
    "MANUAL":        (1, "Manually posted (not a subsystem interface)"),
    "WEEKEND":       (3, "Posted on a weekend"),
    "HOLIDAY":       (3, "Posted on a holiday"),
    "AFTERHOURS":    (2, "Posted outside business hours"),
    "PERIODEND":     (2, "Effective in the last 5 days of the period"),
    "POSTLAG":       (3, "Posted >5 days after its effective date"),
    "POSTAFTERCLOSE":(4, "Posted after period end but dated within the period"),
    "ROUND":         (2, "Round-dollar amount relative to its magnitude"),
    "ROUNDCOMPUTED": (3, "Round-dollar amount on an account that should be computed"),
    "THRESHOLD":     (5, "Amount falls just below an approval threshold"),
    "SPLIT":         (4, "Possible split: multiple sub-threshold entries, same day/account/user"),
    "SELFAPPROVED":  (4, "Preparer is also the approver"),
    "UNAPPROVED":    (3, "No approver recorded"),
    "UNAUTHUSER":    (5, "Posted by a user not on the authorized list"),
    "DUPEXACT":      (4, "Exact duplicate: same account, amount and effective date"),
    "DUPNEAR":       (2, "Near duplicate: same account and amount within the date window"),
    "BLANKDESC":     (3, "Blank description on a manual entry"),
    "GENERICDESC":   (2, "Generic or self-explanatory description"),
    "RAREACCOUNT":   (2, "Account used very rarely in the population"),
    "SUSPENSE":      (3, "Suspense / clearing / other account used as counter-account"),
    "REVENUEMANUAL": (3, "Manual entry crediting revenue"),
    "RESERVEMANUAL": (3, "Manual entry adjusting a reserve, allowance or accrual"),
    "CASHREVENUE":   (4, "Entry touches cash and revenue with no receivable leg"),
    "UNREVERSED":    (3, "Marked as reversing but no reversal found in the population"),
    "REVAMOUNT":     (3, "Reversal amount differs from the original entry"),
    "MATERIAL":      (2, "Entry equals or exceeds materiality"),
}

GENERIC_DESC = [
    "adjustment", "adjusting entry", "adj", "reclass", "reclassification", "plug",
    "to balance", "balance", "per client", "per discussion", "as discussed",
    "correction", "correcting entry", "true up", "true-up", "trueup", "misc",
    "miscellaneous", "see attached", "do not reverse", "dnr", "per management",
    "to agree", "to tie", "catch up", "catch-up", "cleanup", "clean up", "entry",
    "journal entry", "je", "monthly", "month end", "month-end", "accrual", "reverse",
]

REVENUE_HINTS = ("revenue", "sales", "income", "fees earned", "service revenue")
RESERVE_HINTS = ("reserve", "allowance", "accrual", "accrued", "provision",
                 "impairment", "valuation")
COMPUTED_HINTS = RESERVE_HINTS + ("depreciation", "amortization", "depletion",
                                  "interest expense", "bad debt")
SUSPENSE_HINTS = ("suspense", "clearing", "other", "unallocated", "unapplied",
                  "misc", "temporary", "plug", "to be determined", "tbd")
CASH_HINTS = ("cash", "bank", "operating account", "checking", "money market")
AR_HINTS = ("receivable", "a/r", "ar ", "unbilled", "contract asset")

ROUND_MODULI = (Decimal("10000000"), Decimal("1000000"), Decimal("500000"),
                Decimal("100000"), Decimal("50000"), Decimal("25000"),
                Decimal("10000"), Decimal("5000"), Decimal("1000"))
ROUND_MIN = Decimal("1000")
ROUND_SIGNIFICANCE = Decimal("20")


# ------------------------------------------------------------------ parsing

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


def pdate(raw) -> date | None:
    s = str(raw or "").strip()
    if not s:
        return None
    s = s.split("T")[0].split(" ")[0]
    for f in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d-%b-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            continue
    raise ValueError(f"Unrecognized date {raw!r}. Prefer ISO YYYY-MM-DD.")


def pdt(raw) -> datetime | None:
    s = str(raw or "").strip()
    if not s:
        return None
    s = s.replace("T", " ").split(".")[0]
    for f in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%m/%d/%Y %H:%M:%S",
              "%m/%d/%Y %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, f)
        except ValueError:
            continue
    d = pdate(s)
    return datetime.combine(d, time(0, 0)) if d else None


def is_round(v: Decimal) -> str:
    a = abs(v)
    if a < ROUND_MIN:
        return ""
    need = a / ROUND_SIGNIFICANCE
    for m in ROUND_MODULI:
        if m >= need and a >= m and a % m == ZERO:
            return f"multiple of {m:,.0f}"
    return ""


def hinted(text: str, hints) -> bool:
    t = (text or "").lower()
    return any(h in t for h in hints)


def load(path: Path) -> tuple[list[dict], dict[str, dict]]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        raw = list(csv.DictReader(fh))
    if not raw:
        sys.exit(f"No rows in {path}")

    lines: list[dict] = []
    for i, r in enumerate(raw, start=1):
        eid = " ".join(str(r.get("entry_id", "") or "").split())
        if not eid:
            raise ValueError(f"Row {i}: entry_id is required - it is what groups lines "
                             f"into entries, and every test depends on it.")
        d, c = dec(r.get("debit")), dec(r.get("credit"))
        if d is not None and c is not None and d != ZERO and c != ZERO:
            raise ValueError(f"Row {i} ({eid}): both debit and credit populated.")
        eff = pdate(r.get("effective_date"))
        if eff is None:
            raise ValueError(f"Row {i} ({eid}): effective_date is required.")
        lines.append({
            "row": i, "entry_id": eid,
            "line_no": " ".join(str(r.get("line_no", "") or "").split()),
            "effective_date": eff,
            "posted_timestamp": pdt(r.get("posted_timestamp")),
            "account": " ".join(str(r.get("account", "") or "").split()),
            "account_name": " ".join(str(r.get("account_name", "") or "").split()),
            "debit": d or ZERO, "credit": c or ZERO,
            "description": " ".join(str(r.get("description", "") or "").split()),
            "posted_by": " ".join(str(r.get("posted_by", "") or "").split()),
            "approved_by": " ".join(str(r.get("approved_by", "") or "").split()),
            "source": " ".join(str(r.get("source", "") or "").split()).upper(),
            "reversal_of": " ".join(str(r.get("reversal_of", "") or "").split()),
        })

    entries: dict[str, dict] = {}
    for ln in lines:
        e = entries.setdefault(ln["entry_id"], {
            "entry_id": ln["entry_id"], "lines": [], "debits": ZERO, "credits": ZERO,
        })
        e["lines"].append(ln)
        e["debits"] += ln["debit"]
        e["credits"] += ln["credit"]

    for e in entries.values():
        L = e["lines"]
        e["effective_date"] = min(x["effective_date"] for x in L)
        stamps = [x["posted_timestamp"] for x in L if x["posted_timestamp"]]
        e["posted_timestamp"] = min(stamps) if stamps else None
        e["amount"] = max(e["debits"], e["credits"])
        e["balanced"] = e["debits"] == e["credits"]
        e["out_of_balance"] = e["debits"] - e["credits"]
        e["accounts"] = [x["account"] for x in L]
        e["account_names"] = " | ".join(sorted({x["account_name"] for x in L if x["account_name"]}))
        e["description"] = next((x["description"] for x in L if x["description"]), "")
        e["posted_by"] = next((x["posted_by"] for x in L if x["posted_by"]), "")
        e["approved_by"] = next((x["approved_by"] for x in L if x["approved_by"]), "")
        e["source"] = next((x["source"] for x in L if x["source"]), "")
        e["reversal_of"] = next((x["reversal_of"] for x in L if x["reversal_of"]), "")
        e["line_count"] = len(L)
        e["lag_days"] = ((e["posted_timestamp"].date() - e["effective_date"]).days
                         if e["posted_timestamp"] else None)
        e["flags"] = []
        e["score"] = 0

    return lines, entries


# ------------------------------------------------------------------ integrity

def integrity(lines, entries, holidays: set) -> dict:
    unbalanced = [e for e in entries.values() if not e["balanced"]]

    # numeric-suffix continuity, if the ids look sequential
    nums, prefix = [], None
    ok_seq = True
    for eid in entries:
        m = re.match(r"^(.*?)(\d+)$", eid)
        if not m:
            ok_seq = False
            break
        p, n = m.group(1), int(m.group(2))
        if prefix is None:
            prefix = p
        elif p != prefix:
            ok_seq = False
            break
        nums.append(n)
    gaps = []
    if ok_seq and len(nums) > 1:
        s = sorted(set(nums))
        for a, b in zip(s, s[1:]):
            if b - a > 1:
                gaps.append((f"{prefix}{a}", f"{prefix}{b}", b - a - 1))

    fields = ["posted_timestamp", "posted_by", "approved_by", "source",
              "description", "account_name", "reversal_of"]
    pop = {}
    n = len(lines)
    for f in fields:
        filled = sum(1 for x in lines if x[f])
        pop[f] = (filled, n, (filled / n * 100) if n else 0.0)

    disabled = []
    if pop["posted_timestamp"][0] == 0:
        disabled.append("posted_timestamp empty - weekend, holiday, after-hours, "
                        "posting-lag and post-close tests are DISABLED")
    if pop["posted_by"][0] == 0:
        disabled.append("posted_by empty - user, self-approval and authorized-user "
                        "tests are DISABLED")
    if pop["approved_by"][0] == 0:
        disabled.append("approved_by empty - approval tests are DISABLED")
    if pop["source"][0] == 0:
        disabled.append("source empty - manual-vs-interface classification is DISABLED; "
                        "every entry is treated as manual, which inflates the exception "
                        "population")

    by_acct = defaultdict(lambda: {"dr": ZERO, "cr": ZERO, "n": 0})
    for ln in lines:
        a = by_acct[f"{ln['account']} {ln['account_name']}".strip()]
        a["dr"] += ln["debit"]
        a["cr"] += ln["credit"]
        a["n"] += 1

    return {
        "line_count": n, "entry_count": len(entries),
        "total_debits": sum((x["debit"] for x in lines), ZERO),
        "total_credits": sum((x["credit"] for x in lines), ZERO),
        "unbalanced": unbalanced, "seq_checked": ok_seq, "gaps": gaps,
        "population": pop, "disabled": disabled, "by_account": dict(by_acct),
        "date_min": min(x["effective_date"] for x in lines),
        "date_max": max(x["effective_date"] for x in lines),
    }


# ------------------------------------------------------------------ tests

def flag(e: dict, code: str, detail: str = "") -> None:
    w, label = TESTS[code]
    e["flags"].append({"code": code, "weight": w, "label": label, "detail": detail})
    e["score"] += w


def run_tests(entries: dict, args, holidays: set, authorized: set,
              thresholds: list[Decimal], bh: tuple[int, int],
              materiality: Decimal | None, period_end: date | None) -> None:
    acct_use = Counter()
    for e in entries.values():
        for a in set(e["accounts"]):
            acct_use[a] += 1
    rare_cut = max(2, int(len(entries) * 0.005))

    dup_exact = defaultdict(list)
    dup_near = defaultdict(list)
    split_groups = defaultdict(list)
    reversed_ids = {e["reversal_of"] for e in entries.values() if e["reversal_of"]}

    for e in entries.values():
        src, amt = e["source"], e["amount"]
        manual = (not src) or src in ("MANUAL", "MAN", "JE", "GJ", "TOPSIDE",
                                      "ADJ", "CONSOL", "USER")
        e["manual"] = manual
        if manual:
            flag(e, "MANUAL", src or "no source field")

        if materiality is not None and amt >= materiality:
            flag(e, "MATERIAL", f"{amt:,.2f} >= {materiality:,.2f}")

        ts = e["posted_timestamp"]
        if ts:
            if ts.weekday() >= 5:
                flag(e, "WEEKEND", ts.strftime("%A %Y-%m-%d %H:%M"))
            if ts.date() in holidays:
                flag(e, "HOLIDAY", ts.date().isoformat())
            if not (bh[0] <= ts.hour < bh[1]):
                flag(e, "AFTERHOURS",
                     f"{ts.strftime('%H:%M')} (business hours {bh[0]:02d}:00-{bh[1]:02d}:00)")
            if e["lag_days"] is not None and e["lag_days"] > 5:
                flag(e, "POSTLAG", f"{e['lag_days']} days after effective date")
            if period_end and ts.date() > period_end and e["effective_date"] <= period_end:
                flag(e, "POSTAFTERCLOSE",
                     f"posted {ts.date()} for effective {e['effective_date']}")
        if period_end and 0 <= (period_end - e["effective_date"]).days < 5:
            flag(e, "PERIODEND", e["effective_date"].isoformat())

        rnd = is_round(amt)
        if rnd:
            computed = hinted(e["account_names"], COMPUTED_HINTS)
            if computed:
                flag(e, "ROUNDCOMPUTED", f"{amt:,.2f} is a {rnd} on a computed account")
            else:
                flag(e, "ROUND", f"{amt:,.2f} is a {rnd}")

        for t in thresholds:
            band = t * Decimal("0.05")
            if t - band <= amt < t:
                flag(e, "THRESHOLD",
                     f"{amt:,.2f} is {t - amt:,.2f} below the {t:,.0f} approval limit")
                if manual:
                    split_groups[(e["effective_date"], e["posted_by"],
                                  tuple(sorted(set(e["accounts"]))), t)].append(e)
                break

        if e["posted_by"] and e["approved_by"] and e["posted_by"] == e["approved_by"]:
            flag(e, "SELFAPPROVED", f"user {e['posted_by']}")
        if e["posted_by"] and not e["approved_by"]:
            flag(e, "UNAPPROVED", f"posted by {e['posted_by']}, no approver")
        if authorized and e["posted_by"] and e["posted_by"].lower() not in authorized:
            flag(e, "UNAUTHUSER", f"user {e['posted_by']} not on the authorized list")

        d = e["description"].lower().strip()
        if manual and not d:
            flag(e, "BLANKDESC")
        elif manual and d:
            # The test is whether the description says essentially NOTHING beyond a
            # generic term. "Monthly depreciation per FA schedule 3" starts with a
            # generic word but is a perfectly good description, so a startswith rule
            # would bury the real findings in false positives. Require the generic
            # term to be substantially the whole description.
            for g in GENERIC_DESC:
                if re.fullmatch(rf"{re.escape(g)}[\s\W]*\d{{0,6}}[\s\W]*", d):
                    flag(e, "GENERICDESC", f'"{e["description"]}"')
                    break
            else:
                if re.fullmatch(r"per [a-z.\s]{1,18}", d) or len(d) <= 3:
                    flag(e, "GENERICDESC", f'"{e["description"]}"')

        for a in set(e["accounts"]):
            if acct_use[a] <= rare_cut:
                flag(e, "RAREACCOUNT",
                     f"account {a} appears in {acct_use[a]} of {len(entries)} entries")
                break

        names = [x["account_name"] for x in e["lines"]]
        if manual and any(hinted(n, SUSPENSE_HINTS) for n in names):
            flag(e, "SUSPENSE", next(n for n in names if hinted(n, SUSPENSE_HINTS)))
        if manual and any(hinted(x["account_name"], REVENUE_HINTS) and x["credit"] > ZERO
                          for x in e["lines"]):
            flag(e, "REVENUEMANUAL")
        if manual and any(hinted(x["account_name"], RESERVE_HINTS) for x in e["lines"]):
            flag(e, "RESERVEMANUAL")
        has_cash = any(hinted(n, CASH_HINTS) for n in names)
        has_rev = any(hinted(n, REVENUE_HINTS) for n in names)
        has_ar = any(hinted(n, AR_HINTS) for n in names)
        if has_cash and has_rev and not has_ar:
            flag(e, "CASHREVENUE")

        if e["reversal_of"] and e["reversal_of"] in entries:
            orig = entries[e["reversal_of"]]
            if orig["amount"] != e["amount"]:
                flag(e, "REVAMOUNT",
                     f"reverses {orig['entry_id']} of {orig['amount']:,.2f} "
                     f"with {e['amount']:,.2f}")

        for ln in e["lines"]:
            amt_key = ln["debit"] if ln["debit"] else ln["credit"]
            dup_exact[(ln["account"], amt_key, ln["effective_date"])].append(e)
            dup_near[(ln["account"], amt_key)].append((ln["effective_date"], e))

    # duplicates
    for (acct, amt, d), es in dup_exact.items():
        uniq = {x["entry_id"]: x for x in es}
        if len(uniq) > 1 and amt != ZERO:
            for e in uniq.values():
                # one DUPEXACT flag per entry, however many lines collided
                if any(f["code"] == "DUPEXACT" for f in e["flags"]):
                    continue
                flag(e, "DUPEXACT",
                     f"account {acct} {amt:,.2f} on {d} also in "
                     f"{', '.join(k for k in uniq if k != e['entry_id'])}")
    win = args.duplicate_window
    for (acct, amt), items in dup_near.items():
        if amt == ZERO or len(items) < 2:
            continue
        items.sort(key=lambda t: t[0])
        for i in range(len(items)):
            for j in range(i + 1, len(items)):
                d1, e1 = items[i]
                d2, e2 = items[j]
                gap = (d2 - d1).days
                if gap == 0 or gap > win or e1["entry_id"] == e2["entry_id"]:
                    continue
                if any(f["code"] == "DUPEXACT" for f in e1["flags"]):
                    continue
                for e in (e1, e2):
                    if not any(f["code"] == "DUPNEAR" for f in e["flags"]):
                        flag(e, "DUPNEAR",
                             f"account {acct} {amt:,.2f} appears in "
                             f"{e1['entry_id']} ({d1}) and {e2['entry_id']} ({d2})")

    for key, es in split_groups.items():
        uniq = {x["entry_id"]: x for x in es}
        if len(uniq) > 1:
            d, user, accts, t = key
            for e in uniq.values():
                flag(e, "SPLIT",
                     f"{len(uniq)} sub-{t:,.0f} entries on {d} by {user or 'unknown'} "
                     f"to the same account(s)")

    # accruals never reversed
    for e in entries.values():
        if not e["manual"]:
            continue
        if hinted(e["description"], ("accrual", "accrue", "accrued", "reverse next")) \
                and e["entry_id"] not in reversed_ids:
            flag(e, "UNREVERSED",
                 "description indicates an accrual but no reversing entry references it")


def benford(entries: dict) -> list[dict]:
    counts = Counter()
    total = 0
    for e in entries.values():
        a = abs(e["amount"])
        if a < 10:
            continue
        m = re.search(r"[1-9]", str(a).replace(".", "").replace("-", ""))
        if not m:
            continue
        counts[int(m.group())] += 1
        total += 1
    out = []
    for d in range(1, 10):
        exp = math.log10(1 + 1 / d) * 100
        act = (counts[d] / total * 100) if total else 0.0
        out.append({"digit": d, "count": counts[d], "actual": act,
                    "expected": exp, "delta": act - exp})
    return out


# ------------------------------------------------------------------ workbook

def hdr(ws, n, row=1):
    for c in range(1, n + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def widths(ws, w):
    for c, v in w.items():
        ws.column_dimensions[get_column_letter(c)].width = v


EX_COLS = [
    ("score", "Risk score", 11), ("entry_id", "Entry ID", 16),
    ("effective_date", "Effective", 12), ("posted_timestamp", "Posted", 17),
    ("lag_days", "Lag (d)", 9), ("amount", "Amount", 16),
    ("source", "Source", 12), ("posted_by", "Posted by", 14),
    ("approved_by", "Approved by", 14),
    ("account_names", "Accounts", 40), ("description", "Description", 44),
    ("flag_codes", "Flags", 34), ("flag_detail", "Flag detail", 70),
    ("selected", "Selected for testing", 18),
    ("support", "Support obtained", 22),
    ("disposition", "Disposition", 26),
]


def sheet_exceptions(wb, rows, title):
    ws = wb.create_sheet(title)
    ws.append([l for _, l, _ in EX_COLS])
    hdr(ws, len(EX_COLS))
    for e in rows:
        ws.append([
            e["score"], e["entry_id"], e["effective_date"], e["posted_timestamp"],
            e["lag_days"], float(e["amount"]), e["source"] or None,
            e["posted_by"] or None, e["approved_by"] or None,
            e["account_names"] or None, e["description"] or None,
            ", ".join(f["code"] for f in e["flags"]),
            " | ".join(f"{f['code']}: {f['detail']}" for f in e["flags"] if f["detail"]),
            None, None, None,
        ])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[2].number_format = DATEF
        row[3].number_format = DTF
        row[5].number_format = MONEY
        s = row[0].value or 0
        row[0].font = Font(bold=True)
        if s >= 12:
            row[0].fill = BAD_FILL
            row[0].font = BAD_FONT
        elif s >= 7:
            row[0].fill = WARN_FILL
        elif s >= 4:
            row[0].fill = INFO_FILL
    if ws.max_row == 1:
        ws.cell(row=2, column=2, value="None.")
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(EX_COLS))}{max(ws.max_row, 2)}"
    widths(ws, {i: w for i, (_, _, w) in enumerate(EX_COLS, start=1)})


def sheet_summary(wb, integ, entries, flagged, meta, criteria, bench):
    ws = wb.active
    ws.title = "Workpaper Summary"
    widths(ws, {1: 4, 2: 50, 3: 22, 4: 18, 5: 60})
    r = 1

    def line(t, *, bold=False, size=11, fill=None, col=2):
        nonlocal r
        c = ws.cell(row=r, column=col, value=t)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c.fill = fill
        r += 1

    line("JOURNAL ENTRY SCAN - AUDIT WORKPAPER", bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    line("OBJECTIVE", bold=True, size=12)
    line("Scan the complete population of journal entries for characteristics "
         "associated with error or override of controls, rank entries by accumulated "
         "risk attributes, and identify entries for further testing. This workpaper "
         "records attributes and facts; it does not conclude on intent.")
    r += 1

    line("POPULATION AND COMPLETENESS", bold=True, size=12)
    for label, val, fmt in [
        ("Journal entries", integ["entry_count"], None),
        ("Journal entry lines", integ["line_count"], None),
        ("Total debits", float(integ["total_debits"]), MONEY),
        ("Total credits", float(integ["total_credits"]), MONEY),
        ("Effective date range",
         f"{integ['date_min']} to {integ['date_max']}", None),
    ]:
        ws.cell(row=r, column=2, value="  " + label)
        c = ws.cell(row=r, column=3, value=val)
        if fmt:
            c.number_format = fmt
        r += 1

    ub = integ["unbalanced"]
    c = ws.cell(row=r, column=2, value="  Entries where debits = credits")
    v = ws.cell(row=r, column=3,
                value="ALL BALANCE" if not ub else f"{len(ub)} DO NOT BALANCE")
    v.font = OK_FONT if not ub else BAD_FONT
    if ub:
        v.fill = BAD_FILL
        ws.cell(row=r, column=5,
                value="No ERP posts an unbalanced entry. This means the extract is "
                      "truncated - no anomaly result is reliable until it is fixed."
                ).font = Font(italic=True, size=9)
    r += 1

    c = ws.cell(row=r, column=2, value="  Entry number continuity")
    if not integ["seq_checked"]:
        ws.cell(row=r, column=3, value="not testable (non-sequential IDs)")
    elif integ["gaps"]:
        g = ws.cell(row=r, column=3, value=f"{len(integ['gaps'])} gap(s)")
        g.font, g.fill = BAD_FONT, WARN_FILL
        ws.cell(row=r, column=5,
                value="Gaps mean an incomplete extract or deleted entries. "
                      "Both matter and need different follow-ups."
                ).font = Font(italic=True, size=9)
    else:
        g = ws.cell(row=r, column=3, value="no gaps")
        g.font = OK_FONT
    r += 1
    line("  Tie to trial balance: activity by account is on the Population "
         "Integrity tab. Agree it to the TB before relying on this workpaper.")
    r += 1

    if integ["disabled"]:
        line("SCOPE LIMITATIONS", bold=True, size=12, fill=WARN_FILL)
        for d in integ["disabled"]:
            line("  ! " + d)
        line("  A system that does not capture posting user or timestamp is itself a "
             "control finding worth reporting.", bold=True)
        r += 1

    line("CRITERIA AND THRESHOLDS USED", bold=True, size=12)
    for k, v in criteria.items():
        ws.cell(row=r, column=2, value="  " + k)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    line("RESULTS BY TEST", bold=True, size=12)
    for h, col in zip(["Test", "Entries", "Dollars", "Weight", "Description"],
                      (2, 3, 4, 5, 6)):
        hc = ws.cell(row=r, column=col, value=h)
        hc.fill, hc.font = HDR_FILL, HDR_FONT
    r += 1
    per = defaultdict(lambda: {"n": 0, "amt": ZERO})
    for e in entries.values():
        for f in e["flags"]:
            per[f["code"]]["n"] += 1
            per[f["code"]]["amt"] += e["amount"]
    for code, (w, label) in sorted(TESTS.items(), key=lambda kv: -kv[1][0]):
        p = per.get(code, {"n": 0, "amt": ZERO})
        ws.cell(row=r, column=2, value=code)
        ws.cell(row=r, column=3, value=p["n"])
        mc = ws.cell(row=r, column=4, value=float(p["amt"]))
        mc.number_format = MONEY
        ws.cell(row=r, column=5, value=w)
        ws.cell(row=r, column=6, value=label)
        if p["n"]:
            ws.cell(row=r, column=3).font = Font(bold=True)
        r += 1
    r += 1

    line("RANKING", bold=True, size=12)
    line(f"  Entries with at least one flag: {len(flagged)} of {integ['entry_count']}")
    for lo, hi, lbl in ((12, 10**9, "score 12+ (converging attributes)"),
                        (7, 12, "score 7-11"), (4, 7, "score 4-6"), (1, 4, "score 1-3")):
        n = sum(1 for e in flagged if lo <= e["score"] < hi)
        ws.cell(row=r, column=2, value="  " + lbl)
        ws.cell(row=r, column=3, value=n)
        r += 1
    r += 1
    line("A single flag is weak evidence. Read the top of the ranking first - the "
         "entry appearing on four tests at once is the one that matters.", bold=True)
    r += 1

    top = flagged[:15]
    if top:
        line("HIGHEST-RANKED ENTRIES", bold=True, size=12)
        for h, col in zip(["Score", "Entry", "Date", "Amount", "User", "Flags"],
                          (2, 3, 4, 5, 6, 7)):
            hc = ws.cell(row=r, column=col, value=h)
            hc.fill, hc.font = HDR_FILL, HDR_FONT
        r += 1
        for e in top:
            ws.cell(row=r, column=2, value=e["score"]).font = Font(bold=True)
            ws.cell(row=r, column=3, value=e["entry_id"])
            dc = ws.cell(row=r, column=4, value=e["effective_date"])
            dc.number_format = DATEF
            mc = ws.cell(row=r, column=5, value=float(e["amount"]))
            mc.number_format = MONEY
            ws.cell(row=r, column=6, value=e["posted_by"] or None)
            ws.cell(row=r, column=7, value=", ".join(f["code"] for f in e["flags"]))
            r += 1
        r += 1

    line("SELECTION BASIS", bold=True, size=12)
    line("  State the basis explicitly - an unstated selection basis is the defect "
         "that makes journal entry testing indefensible. Recommended: all entries "
         "above a score threshold, all material manual entries, all self-approved "
         "entries above a floor, plus a random sample from the unflagged remainder "
         "so that population is not left untested. Record what was not selected and "
         "why on the Not Selected tab.")
    r += 2
    line("CONCLUSION", bold=True, size=12)
    line("  [Complete after investigation. State whether the procedure identified "
         "misstatements, control observations, or unresolved matters, and how each "
         "was dispositioned.]")
    r += 2
    line("Prepared by / date:  __________________  ____________", bold=True)
    line("Reviewed by / date:  __________________  ____________", bold=True)


def sheet_integrity(wb, integ):
    ws = wb.create_sheet("Population Integrity")
    widths(ws, {1: 4, 2: 44, 3: 18, 4: 18, 5: 14, 6: 50})
    r = 1
    ws.cell(row=r, column=2, value="POPULATION INTEGRITY").font = Font(bold=True, size=14)
    r += 2

    ws.cell(row=r, column=2, value="Unbalanced entries").font = Font(bold=True, size=12)
    r += 1
    if not integ["unbalanced"]:
        c = ws.cell(row=r, column=2, value="None - every entry has debits equal to credits.")
        c.font = OK_FONT
        r += 1
    else:
        for h, col in zip(["Entry ID", "Debits", "Credits", "Out of balance"], (2, 3, 4, 5)):
            hc = ws.cell(row=r, column=col, value=h)
            hc.fill, hc.font = HDR_FILL, HDR_FONT
        r += 1
        for e in integ["unbalanced"]:
            ws.cell(row=r, column=2, value=e["entry_id"])
            for col, k in ((3, "debits"), (4, "credits"), (5, "out_of_balance")):
                c = ws.cell(row=r, column=col, value=float(e[k]))
                c.number_format = MONEY
            ws.cell(row=r, column=5).font = BAD_FONT
            r += 1
    r += 1

    ws.cell(row=r, column=2, value="Entry number continuity").font = Font(bold=True, size=12)
    r += 1
    if not integ["seq_checked"]:
        ws.cell(row=r, column=2,
                value="Entry IDs are not a single sequential series - continuity not "
                      "testable this way. Consider testing continuity per document type.")
        r += 1
    elif not integ["gaps"]:
        ws.cell(row=r, column=2, value="No gaps in the sequence.").font = OK_FONT
        r += 1
    else:
        for h, col in zip(["After", "Before", "Missing count"], (2, 3, 4)):
            hc = ws.cell(row=r, column=col, value=h)
            hc.fill, hc.font = HDR_FILL, HDR_FONT
        r += 1
        for a, b, n in integ["gaps"]:
            ws.cell(row=r, column=2, value=a)
            ws.cell(row=r, column=3, value=b)
            ws.cell(row=r, column=4, value=n).font = BAD_FONT
            r += 1
    r += 1

    ws.cell(row=r, column=2, value="Field population").font = Font(bold=True, size=12)
    r += 1
    for h, col in zip(["Field", "Populated", "Of lines", "%", "Effect if empty"],
                      (2, 3, 4, 5, 6)):
        hc = ws.cell(row=r, column=col, value=h)
        hc.fill, hc.font = HDR_FILL, HDR_FONT
    r += 1
    notes = {
        "posted_timestamp": "disables weekend, holiday, after-hours, lag, post-close",
        "posted_by": "disables user, self-approval, authorized-user tests",
        "approved_by": "disables approval tests",
        "source": "disables manual-vs-interface split; all entries treated as manual",
        "description": "disables description-quality tests",
        "account_name": "weakens account-relationship tests (name matching)",
        "reversal_of": "disables reversal tests",
    }
    for f, (filled, n, pctv) in integ["population"].items():
        ws.cell(row=r, column=2, value=f)
        ws.cell(row=r, column=3, value=filled)
        ws.cell(row=r, column=4, value=n)
        pc = ws.cell(row=r, column=5, value=pctv)
        pc.number_format = PCT
        if filled == 0:
            pc.fill, pc.font = BAD_FILL, BAD_FONT
        elif pctv < 90:
            pc.fill = WARN_FILL
        ws.cell(row=r, column=6, value=notes.get(f, ""))
        r += 1
    r += 2

    ws.cell(row=r, column=2, value="Activity by account - agree to the trial balance"
            ).font = Font(bold=True, size=12)
    r += 1
    for h, col in zip(["Account", "Lines", "Debits", "Credits", "Net"], (2, 3, 4, 5, 6)):
        hc = ws.cell(row=r, column=col, value=h)
        hc.fill, hc.font = HDR_FILL, HDR_FONT
    r += 1
    for acct, a in sorted(integ["by_account"].items()):
        ws.cell(row=r, column=2, value=acct)
        ws.cell(row=r, column=3, value=a["n"])
        for col, v in ((4, a["dr"]), (5, a["cr"]), (6, a["dr"] - a["cr"])):
            c = ws.cell(row=r, column=col, value=float(v))
            c.number_format = MONEY
        r += 1


def sheet_users(wb, entries):
    ws = wb.create_sheet("User Activity Profile")
    heads = ["User", "Entries", "Dollars", "Manual entries", "Manual %",
             "Self-approved", "Unapproved", "Weekend/after-hours",
             "First activity", "Last activity", "Months active"]
    ws.append(heads)
    hdr(ws, len(heads))
    agg = defaultdict(lambda: {"n": 0, "amt": ZERO, "man": 0, "self": 0,
                               "unap": 0, "odd": 0, "months": set(),
                               "first": None, "last": None})
    for e in entries.values():
        u = e["posted_by"] or "(no user recorded)"
        a = agg[u]
        a["n"] += 1
        a["amt"] += e["amount"]
        if e["manual"]:
            a["man"] += 1
        codes = {f["code"] for f in e["flags"]}
        if "SELFAPPROVED" in codes:
            a["self"] += 1
        if "UNAPPROVED" in codes:
            a["unap"] += 1
        if codes & {"WEEKEND", "AFTERHOURS", "HOLIDAY"}:
            a["odd"] += 1
        d = e["effective_date"]
        a["months"].add((d.year, d.month))
        a["first"] = d if a["first"] is None else min(a["first"], d)
        a["last"] = d if a["last"] is None else max(a["last"], d)
    for u, a in sorted(agg.items(), key=lambda kv: -kv[1]["n"]):
        ws.append([u, a["n"], float(a["amt"]), a["man"],
                   (a["man"] / a["n"] * 100) if a["n"] else 0,
                   a["self"], a["unap"], a["odd"], a["first"], a["last"],
                   len(a["months"])])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[2].number_format = MONEY
        row[4].number_format = PCT
        row[8].number_format = DATEF
        row[9].number_format = DATEF
        if row[5].value:
            row[5].fill = BAD_FILL
        if row[7].value:
            row[7].fill = WARN_FILL
    ws.freeze_panes = "B2"
    widths(ws, {1: 24, 2: 10, 3: 16, 4: 15, 5: 11, 6: 14, 7: 12, 8: 20, 9: 14,
                10: 14, 11: 14})


def sheet_benford(wb, bench):
    ws = wb.create_sheet("Benford")
    ws.append(["Leading digit", "Count", "Actual %", "Expected %", "Delta (pp)"])
    hdr(ws, 5)
    for b in bench:
        ws.append([b["digit"], b["count"], b["actual"], b["expected"], b["delta"]])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in (2, 3, 4):
            row[i].number_format = PCT
        if abs(row[4].value or 0) > 5:
            row[4].fill = WARN_FILL
    r = ws.max_row + 2
    for note in [
        "DIRECTIONAL ONLY. Benford tells you where to look, not what is wrong.",
        "Unreliable on small populations and on amounts with natural bounds - anything",
        "priced, capped, contractual, or clustered around approval limits will deviate",
        "for entirely innocent reasons. Do not draw a conclusion from this tab alone.",
    ]:
        ws.cell(row=r, column=1, value=note).font = Font(italic=True)
        r += 1
    widths(ws, {1: 14, 2: 10, 3: 12, 4: 12, 5: 13})


def sheet_not_selected(wb, flagged):
    ws = wb.create_sheet("Not Selected")
    heads = ["Entry ID", "Score", "Amount", "Flags", "Reason not selected"]
    ws.append(heads)
    hdr(ws, len(heads))
    for e in flagged:
        ws.append([e["entry_id"], e["score"], float(e["amount"]),
                   ", ".join(f["code"] for f in e["flags"]), None])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[2].number_format = MONEY
    r = ws.max_row + 2
    ws.cell(row=r, column=1,
            value="Delete the rows that were selected and give a reason for each row "
                  "that remains. The absence of this tab is what makes a selection "
                  "look arbitrary.").font = Font(italic=True)
    ws.freeze_panes = "A2"
    widths(ws, {1: 16, 2: 9, 3: 16, 4: 40, 5: 60})


# ------------------------------------------------------------------ main

def load_set(path: str | None, col: str) -> set:
    if not path or not Path(path).exists():
        return set()
    out = set()
    with Path(path).open(newline="", encoding="utf-8-sig") as fh:
        rdr = csv.DictReader(fh)
        for r in rdr:
            v = (r.get(col) or "").strip()
            if v:
                out.add(v.lower() if col != "date" else v)
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--entries", required=True)
    ap.add_argument("--out")
    ap.add_argument("--validate-only", action="store_true")
    ap.add_argument("--period-start")
    ap.add_argument("--period-end")
    ap.add_argument("--materiality")
    ap.add_argument("--approval-thresholds", default="",
                    help="comma-separated, e.g. 10000,50000,250000")
    ap.add_argument("--business-hours", default="7-19")
    ap.add_argument("--holidays", help="CSV with a 'date' column")
    ap.add_argument("--authorized-users", help="CSV with a 'user' column")
    ap.add_argument("--duplicate-window", type=int, default=7)
    ap.add_argument("--client", default="")
    ap.add_argument("--force", action="store_true",
                    help="run the scan even though the population failed integrity")
    args = ap.parse_args()

    lines, entries = load(Path(args.entries))
    holidays = {pdate(d) for d in load_set(args.holidays, "date")} - {None}
    authorized = load_set(args.authorized_users, "user")
    integ = integrity(lines, entries, holidays)

    print("=" * 72)
    print("POPULATION INTEGRITY")
    print("=" * 72)
    print(f"Entries {integ['entry_count']:,}   lines {integ['line_count']:,}")
    print(f"Total debits  {integ['total_debits']:>18,.2f}")
    print(f"Total credits {integ['total_credits']:>18,.2f}")
    print(f"Date range    {integ['date_min']} to {integ['date_max']}")
    ub = integ["unbalanced"]
    if ub:
        print(f"\n*** {len(ub)} ENTRY(S) DO NOT BALANCE ***")
        for e in ub[:10]:
            print(f"  {e['entry_id']}: dr {e['debits']:,.2f} cr {e['credits']:,.2f} "
                  f"out {e['out_of_balance']:,.2f}")
        print("No ERP posts an unbalanced entry - the extract is truncated.")
    else:
        print("Balancing test: PASS - every entry balances")
    if integ["seq_checked"]:
        gap_txt = "no gaps" if not integ["gaps"] else f"{len(integ['gaps'])} GAP(S)"
        print(f"Entry continuity: {gap_txt}")
        for a, b, n in integ["gaps"][:10]:
            print(f"  gap: {a} -> {b} ({n} missing)")
    else:
        print("Entry continuity: not testable (non-sequential IDs)")
    for f, (filled, n, p) in integ["population"].items():
        if p < 100:
            print(f"  field {f}: {p:.0f}% populated")
    if integ["disabled"]:
        print("\nSCOPE LIMITATIONS:")
        for d in integ["disabled"]:
            print(f"  ! {d}")
    print("\nTie activity by account to the trial balance before relying on results.")

    if args.validate_only:
        return 0 if not ub else 1
    if ub and not args.force:
        print("\n" + "=" * 72)
        print("SCAN NOT RUN. Fix the extract first - anomaly results computed on a")
        print("truncated population create false comfort. Use --force to override.")
        print("=" * 72)
        return 1
    if not args.out:
        sys.exit("--out is required unless --validate-only is used.")

    m = re.match(r"^(\d{1,2})\s*-\s*(\d{1,2})$", args.business_hours.strip())
    if not m:
        sys.exit("--business-hours must look like 7-19")
    bh = (int(m.group(1)), int(m.group(2)))
    thresholds = sorted(
        [t for t in (dec(x) for x in args.approval_thresholds.split(",") if x.strip())
         if t is not None]
    )
    materiality = dec(args.materiality) if args.materiality else None
    period_end = pdate(args.period_end) if args.period_end else None

    run_tests(entries, args, holidays, authorized, thresholds, bh,
              materiality, period_end)

    flagged = sorted([e for e in entries.values() if e["flags"]],
                     key=lambda e: (-e["score"], -e["amount"]))
    bench = benford(entries)

    meta = {
        "Client": args.client or "(not stated)",
        "Period": f"{args.period_start or '?'} to {args.period_end or '?'}",
        "Population source file": Path(args.entries).name,
        "Entries in population": integ["entry_count"],
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    criteria = {
        "Materiality": float(materiality) if materiality else "(not supplied)",
        "Approval thresholds": ", ".join(f"{t:,.0f}" for t in thresholds) or "(not supplied - THRESHOLD and SPLIT tests inactive)",
        "Business hours": f"{bh[0]:02d}:00-{bh[1]:02d}:00 (timestamp time zone: confirm with client)",
        "Holidays supplied": len(holidays) or "(none - HOLIDAY test inactive)",
        "Authorized user list": len(authorized) or "(none - UNAUTHUSER test inactive)",
        "Duplicate window (days)": args.duplicate_window,
    }

    wb = Workbook()
    sheet_summary(wb, integ, entries, flagged, meta, criteria, bench)
    sheet_integrity(wb, integ)
    sheet_exceptions(wb, flagged, "Ranked Exceptions")
    for code in TESTS:
        subset = [e for e in flagged if any(f["code"] == code for f in e["flags"])]
        if subset:
            sheet_exceptions(wb, subset, f"T-{code}"[:31])
    sheet_users(wb, entries)
    sheet_benford(wb, bench)
    sheet_not_selected(wb, flagged)
    wb.active = 0

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))

    print("\n" + "=" * 72)
    print("SCAN RESULTS")
    print("=" * 72)
    print(f"Entries with at least one flag: {len(flagged):,} of {integ['entry_count']:,}")
    per = Counter()
    for e in entries.values():
        for f in e["flags"]:
            per[f["code"]] += 1
    for code, (w, label) in sorted(TESTS.items(), key=lambda kv: -kv[1][0]):
        if per[code]:
            print(f"  [w{w}] {code:<15} {per[code]:>6,}   {label}")
    print("\nHighest-ranked entries (read these first):")
    for e in flagged[:12]:
        print(f"  score {e['score']:>3}  {e['entry_id']:<14} {e['effective_date']}  "
              f"{e['amount']:>14,.2f}  {e['posted_by'] or '-':<10} "
              f"{', '.join(f['code'] for f in e['flags'])[:58]}")
    worst = max((abs(b["delta"]) for b in bench), default=0)
    print(f"\nBenford: largest leading-digit deviation {worst:.1f}pp "
          f"(directional only - see the Benford tab caveats)")
    print(f"\nWorkpaper: {out}")
    print("Next: state the selection basis, complete the Not Selected tab, obtain "
          "support, and disposition each selection.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
