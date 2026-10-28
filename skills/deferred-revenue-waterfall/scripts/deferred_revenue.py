#!/usr/bin/env python3
"""
Build the deferred revenue waterfall from a contract/invoice population and tie
the ending balance to the GL.

Gates:
  * every contract needs recognition_method, service_start, service_end, with
    service_end >= service_start - never defaulted to straight_line or to a
    guessed service period
  * milestone/usage contracts are never mechanically recognized - they need a
    user-supplied stated_ending_deferred balance to enter the tie-out, or they
    are a visible exception
  * the sum of ending deferred balances (mechanical + stated) must tie to the
    supplied GL ending deferred revenue balance exactly, or produce a named,
    dollar-quantified unreconciled difference - never plugged
  * no duplicate invoice numbers within a customer
  * recognized revenue for a contract can never exceed what was available to
    recognize (beginning deferred + additions) for that contract this period

Usage - emit a blank worksheet:
    python3 deferred_revenue.py --contracts contracts.csv \
        --period-start 2025-10-01 --period-end 2025-12-31 \
        --template --out contracts_worksheet.csv

Usage - run:
    python3 deferred_revenue.py --contracts contracts_worksheet.csv \
        --period-start 2025-10-01 --period-end 2025-12-31 \
        --gl-beginning 1842500.00 --gl-ending 2015300.00 \
        --client "Northwind Systems Inc" \
        --out "Northwind - Q4 2025 Deferred Revenue Waterfall.xlsx"

--contracts CSV:
    contract_id, customer, invoice_no, invoice_date, amount,
    service_start, service_end, recognition_method, stated_ending_deferred

  recognition_method: straight_line | milestone | usage
  stated_ending_deferred: required only for milestone/usage rows - the
    client's own determination of that contract's ending deferred balance,
    since this skill does not mechanically recognize those contracts.
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
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is required.  pip install openpyxl")

ZERO = Decimal("0.00")
TOL = Decimal("0.01")
MONEY = '#,##0.00;[Red](#,##0.00)'
DATEF = "yyyy-mm-dd"

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")

METHODS = {"straight_line", "milestone", "usage"}

COLS = ["contract_id", "customer", "invoice_no", "invoice_date", "amount",
        "service_start", "service_end", "recognition_method",
        "stated_ending_deferred"]


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


def load(path: Path) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No contracts in {path}")
    out = []
    for i, r in enumerate(rows, start=1):
        x = {k: clean(r.get(k)) for k in COLS}
        x["row"] = i
        out.append(x)
    return out


def write_template(rows: list[dict], out: Path) -> None:
    with out.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(COLS)
        for r in rows:
            w.writerow([r.get(c, "") for c in COLS])
    print(f"Wrote worksheet with {len(rows)} row(s) and required columns: {out}")


def overlap_days(a_start: date, a_end: date, b_start: date, b_end: date) -> int:
    start = max(a_start, b_start)
    end = min(a_end, b_end)
    if end < start:
        return 0
    return (end - start).days + 1


def evaluate(rows, period_start, period_end, gl_beginning, gl_ending):
    errors, exceptions, judgment = [], [], []
    detail = []
    seen_invoices = defaultdict(set)

    for x in rows:
        rid = x["contract_id"] or f"row{x['row']}"
        cust = x["customer"]
        method = x["recognition_method"]
        amount = dec(x["amount"])
        s_start = None
        s_end = None
        bad = False

        if x["invoice_no"]:
            if x["invoice_no"] in seen_invoices[cust]:
                errors.append(f"{rid}: duplicate invoice {x['invoice_no']!r} for "
                              f"customer {cust!r}")
            seen_invoices[cust].add(x["invoice_no"])

        if method not in METHODS:
            exceptions.append(f"{rid}: recognition_method {method!r} is not one of "
                               f"{sorted(METHODS)} - excluded from the waterfall")
            bad = True
        try:
            s_start = pdate(x["service_start"])
            s_end = pdate(x["service_end"])
        except ValueError as e:
            exceptions.append(f"{rid}: {e}")
            bad = True
        if not bad and (s_start is None or s_end is None):
            exceptions.append(f"{rid}: service_start/service_end missing - "
                               f"never defaulted")
            bad = True
        if not bad and s_end < s_start:
            exceptions.append(f"{rid}: service_end {s_end} is before "
                               f"service_start {s_start}")
            bad = True
        if amount is None:
            exceptions.append(f"{rid}: amount is missing or unparseable")
            bad = True
        inv_date = None
        if x["invoice_date"]:
            try:
                inv_date = pdate(x["invoice_date"])
            except ValueError as e:
                exceptions.append(f"{rid}: {e}")
                bad = True

        if bad:
            detail.append({**x, "contract_id": rid, "status": "excluded",
                            "beginning_deferred": None, "additions": None,
                            "recognized": None, "ending_deferred": None})
            continue

        if method in ("milestone", "usage"):
            stated = dec(x["stated_ending_deferred"])
            if stated is None:
                errors.append(f"{rid}: {method} contract has no "
                              f"stated_ending_deferred - cannot enter the tie-out "
                              f"without the client's own determination")
            judgment.append({"contract_id": rid, "customer": cust, "method": method,
                              "amount": amount, "stated_ending_deferred": stated})
            detail.append({**x, "contract_id": rid, "status": f"judgment ({method})",
                            "beginning_deferred": None, "additions": None,
                            "recognized": None, "ending_deferred": stated})
            continue

        # straight_line mechanical recognition
        total_days = (s_end - s_start).days + 1
        daily_rate = amount / total_days

        if inv_date and inv_date > period_end:
            detail.append({**x, "contract_id": rid, "status": "not yet on books",
                            "beginning_deferred": None, "additions": None,
                            "recognized": None, "ending_deferred": None})
            continue

        booked_at_period_start = bool(inv_date) and inv_date < period_start
        if booked_at_period_start:
            days_before = overlap_days(s_start, s_end, s_start,
                                        period_start - timedelta(days=1))
            beginning_deferred = amount - daily_rate * days_before
            beginning_deferred = max(min(beginning_deferred, amount), ZERO)
        else:
            beginning_deferred = ZERO

        additions = amount if (inv_date and period_start <= inv_date <= period_end) else ZERO

        days_recognized_in_period = overlap_days(s_start, s_end, period_start, period_end)
        recognized = daily_rate * days_recognized_in_period
        if not (inv_date and inv_date <= period_end):
            recognized = ZERO

        ending_deferred = beginning_deferred + additions - recognized

        available = beginning_deferred + additions
        if recognized - available > TOL:
            errors.append(f"{rid}: recognized {recognized:,.2f} exceeds the "
                           f"{available:,.2f} available to recognize this period")

        detail.append({**x, "contract_id": rid, "status": "recognized",
                        "beginning_deferred": beginning_deferred.quantize(TOL),
                        "additions": additions.quantize(TOL),
                        "recognized": recognized.quantize(TOL),
                        "ending_deferred": ending_deferred.quantize(TOL)})

    mech_ending = sum((d["ending_deferred"] for d in detail
                        if d["status"] == "recognized"), ZERO)
    stated_ending = sum((j["stated_ending_deferred"] for j in judgment
                          if j["stated_ending_deferred"] is not None), ZERO)
    computed_ending = mech_ending + stated_ending

    tie_diff = None
    if gl_ending is not None:
        tie_diff = (computed_ending - gl_ending).quantize(TOL)
        if abs(tie_diff) > TOL:
            errors.append(f"Computed ending deferred revenue {computed_ending:,.2f} "
                           f"does not tie to GL ending balance {gl_ending:,.2f} - "
                           f"unreconciled difference {tie_diff:,.2f}")
    else:
        errors.append("--gl-ending not supplied - the waterfall cannot be tied "
                       "to the GL without it")

    return detail, errors, exceptions, judgment, mech_ending, stated_ending, computed_ending, tie_diff


def style_header(ws, headers):
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"


def autosize(ws):
    for i, col in enumerate(ws.columns, start=1):
        width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
        ws.column_dimensions[get_column_letter(i)].width = min(width + 2, 46)


def sheet_summary(wb, meta, errors, mech_ending, stated_ending, computed_ending,
                   gl_beginning, gl_ending, tie_diff, valid):
    ws = wb.active
    ws.title = "Workpaper Summary"
    ws.append(["Deferred Revenue Waterfall"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])
    for k, v in meta.items():
        ws.append([k, v])
    ws.append([])
    ws.append(["GL beginning deferred revenue", gl_beginning])
    ws.append(["Mechanically recognized (straight-line) ending", mech_ending])
    ws.append(["Judgment items (milestone/usage) stated ending", stated_ending])
    ws.append(["Computed ending deferred revenue", computed_ending])
    ws.append(["GL ending deferred revenue", gl_ending])
    ws.append(["Unreconciled difference", tie_diff])
    for r in range(ws.max_row - 5, ws.max_row + 1):
        for c in (2,):
            cell = ws.cell(row=r, column=c)
            if isinstance(cell.value, Decimal):
                cell.number_format = MONEY
    ws.append([])
    ws.append(["CONCLUSION", "TIES TO GL - CLEAN" if valid else "GATE NOT MET"])
    ws[f"B{ws.max_row}"].font = OK_FONT if valid else BAD_FONT
    if errors:
        ws.append([])
        ws.append(["Gate failures"])
        ws["A" + str(ws.max_row)].font = BAD_FONT
        for e in errors:
            ws.append([e])
    autosize(ws)


def sheet_waterfall(wb, detail):
    ws = wb.create_sheet("Waterfall Schedule")
    headers = ["Contract", "Customer", "Invoice No", "Invoice Date", "Amount",
               "Service Start", "Service End", "Status",
               "Beginning Deferred", "Additions", "Recognized", "Ending Deferred"]
    style_header(ws, headers)
    for d in detail:
        row = [d["contract_id"], d["customer"], d["invoice_no"], d["invoice_date"],
               dec(d["amount"]), d["service_start"], d["service_end"], d["status"],
               d["beginning_deferred"], d["additions"], d["recognized"],
               d["ending_deferred"]]
        ws.append(row)
        r = ws.max_row
        for col in (5, 9, 10, 11, 12):
            cell = ws.cell(row=r, column=col)
            if isinstance(cell.value, Decimal):
                cell.number_format = MONEY
        if d["status"] not in ("recognized",) and "judgment" not in d["status"]:
            for col in range(1, len(headers) + 1):
                ws.cell(row=r, column=col).fill = WARN_FILL
    totals_row = ["TOTAL", "", "", "", "",
                  "", "", "",
                  sum((d["beginning_deferred"] for d in detail if d["beginning_deferred"] is not None), ZERO),
                  sum((d["additions"] for d in detail if d["additions"] is not None), ZERO),
                  sum((d["recognized"] for d in detail if d["recognized"] is not None), ZERO),
                  sum((d["ending_deferred"] for d in detail if d["ending_deferred"] is not None), ZERO)]
    ws.append(totals_row)
    r = ws.max_row
    for c in range(1, len(headers) + 1):
        ws.cell(row=r, column=c).font = Font(bold=True)
    for col in (9, 10, 11, 12):
        ws.cell(row=r, column=col).number_format = MONEY
    autosize(ws)


def sheet_je(wb, detail):
    ws = wb.create_sheet("Recognition JEs")
    headers = ["Contract", "Customer", "Dr Deferred Revenue", "Cr Revenue", "Memo"]
    style_header(ws, headers)
    for d in detail:
        if d["status"] == "recognized" and d["recognized"] and d["recognized"] > ZERO:
            ws.append([d["contract_id"], d["customer"], d["recognized"], d["recognized"],
                       "Recognize earned revenue for the period - straight-line"])
            r = ws.max_row
            ws.cell(row=r, column=3).number_format = MONEY
            ws.cell(row=r, column=4).number_format = MONEY
    autosize(ws)


def sheet_tie(wb, gl_beginning, mech_ending, stated_ending, computed_ending, gl_ending, tie_diff):
    ws = wb.create_sheet("GL Tie-Out")
    style_header(ws, ["Line", "Amount"])
    rows = [
        ("GL beginning deferred revenue", gl_beginning),
        ("+ Mechanically recognized (straight-line) ending deferred", mech_ending),
        ("+ Judgment items (milestone/usage) stated ending deferred", stated_ending),
        ("= Computed ending deferred revenue", computed_ending),
        ("GL ending deferred revenue (per client)", gl_ending),
        ("Unreconciled difference", tie_diff),
    ]
    for label, val in rows:
        ws.append([label, val])
        cell = ws.cell(row=ws.max_row, column=2)
        if isinstance(val, Decimal):
            cell.number_format = MONEY
    last = ws.cell(row=ws.max_row, column=2)
    if tie_diff is not None:
        last.font = OK_FONT if abs(tie_diff) <= TOL else BAD_FONT
        if abs(tie_diff) > TOL:
            ws.cell(row=ws.max_row, column=1).fill = BAD_FILL
            last.fill = BAD_FILL
    autosize(ws)


def sheet_exceptions(wb, errors, exceptions):
    ws = wb.create_sheet("Exceptions")
    style_header(ws, ["Type", "Detail"])
    for e in errors:
        ws.append(["Gate failure", e])
        ws.cell(row=ws.max_row, column=1).font = BAD_FONT
    for e in exceptions:
        ws.append(["Data exception", e])
        ws.cell(row=ws.max_row, column=1).font = BAD_FONT
    autosize(ws)


def sheet_judgment(wb, judgment):
    ws = wb.create_sheet("Judgment Items")
    style_header(ws, ["Contract", "Customer", "Method", "Amount",
                       "Stated Ending Deferred", "Note"])
    for j in judgment:
        ws.append([j["contract_id"], j["customer"], j["method"], j["amount"],
                   j["stated_ending_deferred"],
                   "Not mechanically recognized - recognition timing is a "
                   "professional judgment call, not computed by this skill"])
        r = ws.max_row
        ws.cell(row=r, column=4).number_format = MONEY
        if j["stated_ending_deferred"] is not None:
            ws.cell(row=r, column=5).number_format = MONEY
        else:
            ws.cell(row=r, column=5).fill = BAD_FILL
    autosize(ws)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--contracts", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--period-start", required=True)
    ap.add_argument("--period-end", required=True)
    ap.add_argument("--gl-beginning")
    ap.add_argument("--gl-ending")
    ap.add_argument("--client", default="")
    ap.add_argument("--template", action="store_true",
                     help="emit a blank contract worksheet instead of running")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    rows = load(Path(args.contracts))

    if args.template:
        write_template(rows, Path(args.out))
        return 0

    period_start = pdate(args.period_start)
    period_end = pdate(args.period_end)
    if period_end < period_start:
        sys.exit("--period-end is before --period-start")
    gl_beginning = dec(args.gl_beginning)
    gl_ending = dec(args.gl_ending)

    (detail, errors, exceptions, judgment, mech_ending, stated_ending,
     computed_ending, tie_diff) = evaluate(rows, period_start, period_end,
                                            gl_beginning, gl_ending)

    valid = not errors

    meta = {
        "Client": args.client or "(not stated)",
        "Period": f"{period_start} to {period_end}",
        "Contracts in population": len(rows),
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    print("=" * 76)
    print("DEFERRED REVENUE WATERFALL")
    print("=" * 76)
    print(f"Client  : {meta['Client']}    Period: {period_start} to {period_end}")
    print(f"Contracts: {len(rows)}   mechanical ending {mech_ending:,.2f}   "
          f"judgment stated ending {stated_ending:,.2f}")
    print(f"Computed ending {computed_ending:,.2f}   GL ending "
          f"{gl_ending:,.2f}" if gl_ending is not None else
          f"Computed ending {computed_ending:,.2f}   GL ending NOT SUPPLIED")
    if tie_diff is not None:
        print(f"Unreconciled difference: {tie_diff:,.2f}")
    if exceptions:
        print(f"\nDATA EXCEPTIONS ({len(exceptions)}):")
        for e in exceptions[:15]:
            print(f"  ! {e}")
    if errors:
        print(f"\nGATE FAILURES ({len(errors)}):")
        for e in errors[:15]:
            print(f"  ! {e}")
    if judgment:
        print(f"\nJUDGMENT ITEMS - not mechanically recognized ({len(judgment)}):")
        for j in judgment[:10]:
            print(f"  ? {j['contract_id']} ({j['customer'][:22]}): {j['method']}")

    if not valid and not args.force:
        print("\n" + "=" * 76)
        print("WORKBOOK NOT WRITTEN.")
        print("The waterfall does not tie to the GL, or a contract failed the data")
        print("standard. Resolve the exceptions above, or re-run with --force to see")
        print("a workbook marked [GATE NOT MET].")
        print("=" * 76)
        return 1

    wb = Workbook()
    sheet_summary(wb, meta, errors, mech_ending, stated_ending, computed_ending,
                  gl_beginning, gl_ending, tie_diff, valid)
    sheet_waterfall(wb, detail)
    sheet_je(wb, detail)
    sheet_tie(wb, gl_beginning, mech_ending, stated_ending, computed_ending,
              gl_ending, tie_diff)
    sheet_exceptions(wb, errors, exceptions)
    sheet_judgment(wb, judgment)
    wb.active = 0

    out = Path(args.out)
    if not valid:
        out = out.with_name(out.stem + " [GATE NOT MET]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\n{'TIES TO GL' if valid else 'NOT MET (forced)'}.  Workbook: {out}")
    return 0 if valid else 1


if __name__ == "__main__":
    sys.exit(main())
