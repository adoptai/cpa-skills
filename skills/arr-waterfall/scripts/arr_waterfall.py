#!/usr/bin/env python3
"""
Build the ARR/MRR movement waterfall (new, expansion, contraction, churn) from
a subscription or invoice export and the prior period's ending state, and tie
ending ARR to recognized revenue.

Gates:
  * beginning ARR + new + expansion - contraction - churn must equal ending
    ARR exactly - a closed arithmetic identity, checked, not assumed
  * every customer appearing in either file lands in exactly one movement
    bucket, never double-counted or silently dropped (a customer present last
    period but absent this period is still picked up as churn)
  * ending ARR must tie to the supplied GL recognized revenue figure within a
    user-supplied tolerance - no tolerance supplied means an exact tie or a
    reported difference, never a silently-accepted gap

Usage - emit a blank worksheet:
    python3 arr_waterfall.py --current current.csv --prior prior.csv \
        --template --out current_worksheet.csv

Usage - run:
    python3 arr_waterfall.py --current current_worksheet.csv --prior prior.csv \
        --gl-revenue 512300.00 --tolerance 250.00 \
        --client "Northwind Systems Inc" \
        --out "Northwind - Q4 2025 ARR Waterfall.xlsx"

--current CSV (this period's subscription/invoice export):
    customer, plan, mrr_amount, effective_date

--prior CSV (prior period's ending ARR/MRR by customer):
    customer, prior_mrr
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
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is required.  pip install openpyxl")

ZERO = Decimal("0.00")
TOL_DEFAULT = Decimal("0.00")
MONEY = '#,##0.00;[Red](#,##0.00)'

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")

CUR_COLS = ["customer", "plan", "mrr_amount", "effective_date"]
PRIOR_COLS = ["customer", "prior_mrr"]


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


def load(path: Path, cols: list[str]) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No rows in {path}")
    out = []
    for i, r in enumerate(rows, start=1):
        x = {k: clean(r.get(k)) for k in cols}
        x["row"] = i
        out.append(x)
    return out


def write_template(rows: list[dict], out: Path) -> None:
    with out.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(CUR_COLS)
        for r in rows:
            w.writerow([r.get(c, "") for c in CUR_COLS])
    print(f"Wrote worksheet with {len(rows)} row(s) and required columns: {out}")


def evaluate(current_rows, prior_rows, gl_revenue, tolerance):
    errors, exceptions = [], []
    current_mrr = defaultdict(Decimal)
    plan_by_customer = {}
    seen_current = set()

    for x in current_rows:
        cust = x["customer"]
        if not cust:
            exceptions.append(f"row {x['row']}: current extract row has no customer")
            continue
        if cust in seen_current:
            exceptions.append(f"{cust}: appears more than once in the current "
                               f"extract - amounts summed, verify this is intended "
                               f"(multiple subscriptions) and not a duplicate row")
        seen_current.add(cust)
        amt = dec(x["mrr_amount"])
        if amt is None:
            exceptions.append(f"{cust}: mrr_amount is missing or unparseable - "
                               f"treated as excluded, not zero")
            continue
        current_mrr[cust] += amt
        plan_by_customer[cust] = x["plan"]

    prior_mrr = {}
    for x in prior_rows:
        cust = x["customer"]
        if not cust:
            exceptions.append(f"row {x['row']}: prior extract row has no customer")
            continue
        amt = dec(x["prior_mrr"])
        if amt is None:
            exceptions.append(f"{cust}: prior_mrr is missing or unparseable - "
                               f"treated as excluded, not zero")
            continue
        if cust in prior_mrr:
            prior_mrr[cust] += amt
        else:
            prior_mrr[cust] = amt

    all_customers = sorted(set(current_mrr) | set(prior_mrr))
    detail = []
    new_t = exp_t = con_t = chu_t = ZERO
    beg_arr = sum((v for v in prior_mrr.values()), ZERO) * 12
    end_arr = sum((v for v in current_mrr.values()), ZERO) * 12

    for cust in all_customers:
        prior = prior_mrr.get(cust, ZERO)
        cur = current_mrr.get(cust, ZERO)
        delta = cur - prior
        if prior == ZERO and cur > ZERO:
            bucket = "New"
            new_t += delta
        elif prior > ZERO and cur == ZERO:
            bucket = "Churn"
            chu_t += delta
        elif cur > prior > ZERO:
            bucket = "Expansion"
            exp_t += delta
        elif ZERO < cur < prior:
            bucket = "Contraction"
            con_t += delta
        elif cur == prior:
            bucket = "No Change"
        else:
            bucket = "No Change"
        detail.append({
            "customer": cust,
            "plan": plan_by_customer.get(cust, ""),
            "prior_mrr": prior,
            "current_mrr": cur,
            "delta_mrr": delta,
            "bucket": bucket,
        })

    # movement totals above are in MRR terms (delta per customer); annualize once
    new_arr, exp_arr, con_arr, chu_arr = (new_t * 12, exp_t * 12, con_t * 12, chu_t * 12)
    rebuilt_end = beg_arr + new_arr + exp_arr + con_arr + chu_arr
    identity_diff = (rebuilt_end - end_arr).quantize(Decimal("0.01"))
    if abs(identity_diff) > Decimal("0.01"):
        errors.append(f"Waterfall identity fails: beginning ARR {beg_arr:,.2f} + "
                       f"movements does not equal ending ARR {end_arr:,.2f} - "
                       f"off by {identity_diff:,.2f}")

    tie_diff = None
    if gl_revenue is not None:
        tie_diff = (end_arr - gl_revenue).quantize(Decimal("0.01"))
        tol = tolerance if tolerance is not None else TOL_DEFAULT
        if abs(tie_diff) > tol:
            errors.append(f"Ending ARR {end_arr:,.2f} does not tie to GL recognized "
                           f"revenue {gl_revenue:,.2f} within tolerance {tol:,.2f} - "
                           f"difference {tie_diff:,.2f}")
    else:
        errors.append("--gl-revenue not supplied - ending ARR cannot be tied to "
                       "the GL without it")

    judgment = [d for d in detail if d["bucket"] in ("Contraction",) and
                ZERO < d["current_mrr"] < d["prior_mrr"] * Decimal("0.3")]

    return (detail, errors, exceptions, beg_arr, end_arr, new_arr, exp_arr,
            con_arr, chu_arr, tie_diff, judgment)


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


BUCKET_COLOR = {"Churn": BAD_FILL, "Contraction": WARN_FILL}


def sheet_summary(wb, meta, beg_arr, new_arr, exp_arr, con_arr, chu_arr, end_arr,
                   gl_revenue, tie_diff, errors, valid):
    ws = wb.active
    ws.title = "Workpaper Summary"
    ws.append(["ARR Waterfall"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])
    for k, v in meta.items():
        ws.append([k, v])
    ws.append([])
    for label, val in [
        ("Beginning ARR", beg_arr), ("New", new_arr), ("Expansion", exp_arr),
        ("Contraction", con_arr), ("Churn", chu_arr), ("Ending ARR", end_arr),
        ("GL recognized revenue (annualized basis supplied)", gl_revenue),
        ("Tie difference", tie_diff),
    ]:
        ws.append([label, val])
        cell = ws.cell(row=ws.max_row, column=2)
        if isinstance(val, Decimal):
            cell.number_format = MONEY
    ws.append([])
    ws.append(["CONCLUSION", "TIES - CLEAN" if valid else "GATE NOT MET"])
    ws[f"B{ws.max_row}"].font = OK_FONT if valid else BAD_FONT
    if errors:
        ws.append([])
        ws.append(["Gate failures"])
        ws["A" + str(ws.max_row)].font = BAD_FONT
        for e in errors:
            ws.append([e])
    autosize(ws)


def sheet_waterfall(wb, detail, beg_arr, new_arr, exp_arr, con_arr, chu_arr, end_arr):
    ws = wb.create_sheet("Waterfall")
    headers = ["Customer", "Plan", "Prior MRR", "Current MRR", "Delta MRR", "Movement"]
    style_header(ws, headers)
    for d in detail:
        ws.append([d["customer"], d["plan"], d["prior_mrr"], d["current_mrr"],
                   d["delta_mrr"], d["bucket"]])
        r = ws.max_row
        for col in (3, 4, 5):
            ws.cell(row=r, column=col).number_format = MONEY
        fill = BUCKET_COLOR.get(d["bucket"])
        if fill:
            for col in range(1, len(headers) + 1):
                ws.cell(row=r, column=col).fill = fill
    ws.append([])
    ws.append(["", "", "", "ARR total", "", ""])
    for label, val in [("Beginning ARR", beg_arr), ("New", new_arr),
                        ("Expansion", exp_arr), ("Contraction", con_arr),
                        ("Churn", chu_arr), ("Ending ARR", end_arr)]:
        ws.append([label, "", "", val, "", ""])
        ws.cell(row=ws.max_row, column=4).number_format = MONEY
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
    autosize(ws)


def sheet_tie(wb, end_arr, gl_revenue, tie_diff):
    ws = wb.create_sheet("Revenue Tie")
    style_header(ws, ["Line", "Amount"])
    for label, val in [("Ending ARR (annualized)", end_arr),
                        ("GL recognized revenue", gl_revenue),
                        ("Difference", tie_diff)]:
        ws.append([label, val])
        cell = ws.cell(row=ws.max_row, column=2)
        if isinstance(val, Decimal):
            cell.number_format = MONEY
    if tie_diff is not None:
        last = ws.cell(row=ws.max_row, column=2)
        last.font = OK_FONT if abs(tie_diff) <= Decimal("0.01") else BAD_FONT
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
    style_header(ws, ["Customer", "Plan", "Prior MRR", "Current MRR", "Note"])
    for j in judgment:
        ws.append([j["customer"], j["plan"], j["prior_mrr"], j["current_mrr"],
                   "Large contraction - verify this is a downsell and not a "
                   "churn event recorded as a reduced renewal"])
        for col in (3, 4):
            ws.cell(row=ws.max_row, column=col).number_format = MONEY
    autosize(ws)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--current", required=True)
    ap.add_argument("--prior", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--gl-revenue")
    ap.add_argument("--tolerance")
    ap.add_argument("--client", default="")
    ap.add_argument("--template", action="store_true",
                     help="emit a blank current-period worksheet instead of running")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    current_rows = load(Path(args.current), CUR_COLS)

    if args.template:
        write_template(current_rows, Path(args.out))
        return 0

    prior_rows = load(Path(args.prior), PRIOR_COLS)
    gl_revenue = dec(args.gl_revenue)
    tolerance = dec(args.tolerance)

    (detail, errors, exceptions, beg_arr, end_arr, new_arr, exp_arr, con_arr,
     chu_arr, tie_diff, judgment) = evaluate(current_rows, prior_rows, gl_revenue,
                                              tolerance)
    valid = not errors

    meta = {
        "Client": args.client or "(not stated)",
        "Customers in current extract": len(current_rows),
        "Customers in prior extract": len(prior_rows),
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    print("=" * 76)
    print("ARR WATERFALL")
    print("=" * 76)
    print(f"Client : {meta['Client']}")
    print(f"Beginning ARR {beg_arr:,.2f}   New {new_arr:,.2f}   "
          f"Expansion {exp_arr:,.2f}   Contraction {con_arr:,.2f}   "
          f"Churn {chu_arr:,.2f}")
    print(f"Ending ARR {end_arr:,.2f}" +
          (f"   GL revenue {gl_revenue:,.2f}" if gl_revenue is not None else
           "   GL revenue NOT SUPPLIED"))
    if tie_diff is not None:
        print(f"Tie difference: {tie_diff:,.2f}")
    if exceptions:
        print(f"\nDATA EXCEPTIONS ({len(exceptions)}):")
        for e in exceptions[:15]:
            print(f"  ! {e}")
    if errors:
        print(f"\nGATE FAILURES ({len(errors)}):")
        for e in errors[:15]:
            print(f"  ! {e}")
    if judgment:
        print(f"\nJUDGMENT ITEMS ({len(judgment)}):")
        for j in judgment[:10]:
            print(f"  ? {j['customer']}: prior {j['prior_mrr']:,.2f} -> "
                  f"current {j['current_mrr']:,.2f}")

    if not valid and not args.force:
        print("\n" + "=" * 76)
        print("WORKBOOK NOT WRITTEN.")
        print("The waterfall identity does not close, or ending ARR does not tie to")
        print("the GL. Resolve the exceptions above, or re-run with --force to see a")
        print("workbook marked [GATE NOT MET].")
        print("=" * 76)
        return 1

    wb = Workbook()
    sheet_summary(wb, meta, beg_arr, new_arr, exp_arr, con_arr, chu_arr, end_arr,
                  gl_revenue, tie_diff, errors, valid)
    sheet_waterfall(wb, detail, beg_arr, new_arr, exp_arr, con_arr, chu_arr, end_arr)
    sheet_tie(wb, end_arr, gl_revenue, tie_diff)
    sheet_exceptions(wb, errors, exceptions)
    sheet_judgment(wb, judgment)
    wb.active = 0

    out = Path(args.out)
    if not valid:
        out = out.with_name(out.stem + " [GATE NOT MET]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\n{'TIES' if valid else 'NOT MET (forced)'}.  Workbook: {out}")
    return 0 if valid else 1


if __name__ == "__main__":
    sys.exit(main())
