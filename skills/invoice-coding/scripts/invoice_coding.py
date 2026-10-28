#!/usr/bin/env python3
"""
Propose GL coding for vendor invoices from historical coding frequency, and
route everything that cannot be proposed safely to an exception queue.

Gates:
  * a vendor with no historical coding rows is an exception - never guessed
    from the line description or any category heuristic
  * a proposed account_code must exist in the chart of accounts AND be active,
    or it is an exception
  * every invoice lands in exactly one bucket: coded-with-confidence, or
    exception - never left blank, never double-counted
  * an unparseable amount or date is a visible parse exception, never coerced
    or dropped silently
  * if --min-confidence is supplied, anything below it is routed to the
    exception queue instead of being auto-coded

Usage - emit a blank invoice intake template:
    python3 invoice_coding.py --template --out invoices_template.csv

Usage - propose coding:
    python3 invoice_coding.py --invoices invoices.csv --coa chart_of_accounts.csv \
        --historical historical_coding.csv --min-confidence 60 \
        --client "Northwind Systems Inc" \
        --out "Northwind - Invoice Coding Register.xlsx"

--invoices CSV:
    invoice_id, vendor, amount, invoice_date, line_description

--coa CSV (chart of accounts):
    account_code, account_name, active        (active: y/n)

--historical CSV (historical coding frequency):
    vendor, account_code, invoice_count
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
MONEY = '#,##0.00;[Red](#,##0.00)'
PCT = '0.0"%"'
DATEF = "yyyy-mm-dd"

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")

INVOICE_COLS = ["invoice_id", "vendor", "amount", "invoice_date", "line_description"]
COA_COLS = ["account_code", "account_name", "active"]
HIST_COLS = ["vendor", "account_code", "invoice_count"]


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


def norm_vendor(s: str) -> str:
    return clean(s).lower()


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


def read_csv(path: Path, cols: list[str]) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        missing = [c for c in cols if c not in (reader.fieldnames or [])]
        if missing:
            sys.exit(f"{path}: missing required column(s) {missing}")
        rows = list(reader)
    if not rows:
        sys.exit(f"No rows in {path}")
    return [{k: clean(r.get(k)) for k in cols} for r in rows]


def write_template(path: Path) -> None:
    with path.open("w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=INVOICE_COLS)
        w.writeheader()
        w.writerow({
            "invoice_id": "INV-1001", "vendor": "Acme Supply Co",
            "amount": "1250.00", "invoice_date": "2026-01-15",
            "line_description": "office supplies - January",
        })
    print(f"Template written: {path}")


def load_coa(path: Path) -> dict:
    coa = {}
    for r in read_csv(path, COA_COLS):
        code = r["account_code"]
        if not code:
            continue
        coa[code] = {
            "account_name": r["account_name"],
            "active": r["active"].strip().lower() in ("y", "yes", "true", "1"),
        }
    return coa


def load_historical(path: Path) -> dict:
    hist = defaultdict(dict)
    errors = []
    for i, r in enumerate(read_csv(path, HIST_COLS), start=1):
        vendor = norm_vendor(r["vendor"])
        if not vendor or not r["account_code"]:
            continue
        try:
            count = int(Decimal(r["invoice_count"]))
        except (InvalidOperation, ValueError):
            errors.append(f"historical row {i}: unparseable invoice_count {r['invoice_count']!r}")
            continue
        hist[vendor][r["account_code"]] = hist[vendor].get(r["account_code"], 0) + count
    return hist, errors


def load_invoices(path: Path) -> tuple[list[dict], list[str]]:
    errors = []
    out = []
    seen_ids = {}
    with path.open(newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        missing = [c for c in INVOICE_COLS if c not in (reader.fieldnames or [])]
        if missing:
            sys.exit(f"{path}: missing required column(s) {missing}")
        for i, raw in enumerate(reader, start=2):  # row 1 is header
            r = {k: clean(raw.get(k)) for k in INVOICE_COLS}
            r["row"] = i
            if not r["invoice_id"]:
                errors.append(f"row {i}: blank invoice_id - cannot be scanned or deduped")
                continue
            if r["invoice_id"] in seen_ids:
                errors.append(
                    f"row {i}: invoice_id {r['invoice_id']!r} duplicates row "
                    f"{seen_ids[r['invoice_id']]} - population is ambiguous"
                )
                continue
            seen_ids[r["invoice_id"]] = i
            try:
                r["amount_dec"] = dec(r["amount"])
                if r["amount_dec"] is None:
                    raise ValueError("blank amount")
            except ValueError as e:
                errors.append(f"row {i} ({r['invoice_id']}): {e}")
                continue
            try:
                r["invoice_date_dec"] = pdate(r["invoice_date"])
            except ValueError as e:
                errors.append(f"row {i} ({r['invoice_id']}): {e}")
                continue
            out.append(r)
    if not out and not errors:
        sys.exit(f"No invoices in {path}")
    return out, errors


def propose(invoices: list[dict], coa: dict, hist: dict, min_conf) -> tuple[list[dict], list[dict]]:
    coded, exceptions = [], []
    for inv in invoices:
        vkey = norm_vendor(inv["vendor"])
        vhist = hist.get(vkey)
        if not vhist:
            exceptions.append({**inv, "reason": "no historical coding for this vendor - new vendor, route for manual coding"})
            continue
        total = sum(vhist.values())
        ranked = sorted(vhist.items(), key=lambda kv: (-kv[1], kv[0]))
        account_code, count = ranked[0]
        confidence = (Decimal(count) / Decimal(total) * 100).quantize(Decimal("0.1"))
        acct = coa.get(account_code)
        if acct is None:
            exceptions.append({**inv, "reason": f"proposed account {account_code} not found in chart of accounts"})
            continue
        if not acct["active"]:
            exceptions.append({**inv, "reason": f"proposed account {account_code} ({acct['account_name']}) is inactive"})
            continue
        if min_conf is not None and confidence < min_conf:
            exceptions.append({
                **inv,
                "reason": f"confidence {confidence}% below floor {min_conf}% "
                          f"(would propose {account_code} on {count}/{total} history)",
            })
            continue
        coded.append({
            **inv,
            "account_code": account_code,
            "account_name": acct["account_name"],
            "confidence": confidence,
            "basis": f"{count}/{total} historical invoices to {account_code}",
        })
    return coded, exceptions


def style_header(ws, cols):
    for c, name in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=c, value=name)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"


def autosize(ws, cols, rows):
    for c, name in enumerate(cols, start=1):
        width = max(len(name), *(len(str(r.get(name, ""))) for r in rows)) if rows else len(name)
        ws.column_dimensions[get_column_letter(c)].width = min(max(width + 2, 10), 48)


def sheet_summary(wb, meta, n_coded, n_exc, n_parse_err, valid):
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = "INVOICE EXTRACTION AND GL CODING"
    ws["A1"].font = Font(bold=True, size=14)
    row = 3
    for k, v in meta.items():
        ws.cell(row=row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=row, column=2, value=v)
        row += 1
    row += 1
    ws.cell(row=row, column=1, value="Coded with confidence").font = Font(bold=True)
    ws.cell(row=row, column=2, value=n_coded)
    row += 1
    ws.cell(row=row, column=1, value="Exceptions").font = Font(bold=True)
    ws.cell(row=row, column=2, value=n_exc)
    row += 1
    ws.cell(row=row, column=1, value="Parse errors (gate)").font = Font(bold=True)
    ws.cell(row=row, column=2, value=n_parse_err)
    row += 2
    status = ws.cell(row=row, column=1, value="GATE STATUS: " + ("CLEAN" if valid else "NOT MET"))
    status.font = OK_FONT if valid else BAD_FONT
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 50


def sheet_coded(wb, coded):
    ws = wb.create_sheet("Coded Invoice Register")
    cols = ["invoice_id", "vendor", "invoice_date", "amount", "line_description",
            "account_code", "account_name", "confidence", "basis"]
    style_header(ws, cols)
    for r, inv in enumerate(coded, start=2):
        ws.cell(row=r, column=1, value=inv["invoice_id"])
        ws.cell(row=r, column=2, value=inv["vendor"])
        dc = ws.cell(row=r, column=3, value=inv["invoice_date_dec"]); dc.number_format = DATEF
        ac = ws.cell(row=r, column=4, value=float(inv["amount_dec"])); ac.number_format = MONEY
        ws.cell(row=r, column=5, value=inv["line_description"])
        ws.cell(row=r, column=6, value=inv["account_code"])
        ws.cell(row=r, column=7, value=inv["account_name"])
        pc = ws.cell(row=r, column=8, value=float(inv["confidence"])); pc.number_format = PCT
        ws.cell(row=r, column=9, value=inv["basis"])
    autosize(ws, cols, [])


def sheet_exceptions(wb, exceptions, parse_errors):
    ws = wb.create_sheet("Exception Queue")
    cols = ["invoice_id", "vendor", "invoice_date", "amount", "reason"]
    style_header(ws, cols)
    r = 2
    for inv in exceptions:
        ws.cell(row=r, column=1, value=inv["invoice_id"])
        ws.cell(row=r, column=2, value=inv["vendor"])
        ws.cell(row=r, column=3, value=str(inv.get("invoice_date_dec", inv.get("invoice_date", ""))))
        amt = inv.get("amount_dec")
        ws.cell(row=r, column=4, value=float(amt) if amt is not None else inv.get("amount", ""))
        rc = ws.cell(row=r, column=5, value=inv["reason"])
        for c in range(1, 6):
            ws.cell(row=r, column=c).fill = WARN_FILL
        r += 1
    if parse_errors:
        ws.cell(row=r, column=1, value="-- parse errors (gate failures) --").font = BAD_FONT
        r += 1
        for e in parse_errors:
            cell = ws.cell(row=r, column=1, value=e)
            cell.fill = BAD_FILL
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
            r += 1
    autosize(ws, cols, exceptions)


def sheet_history(wb, hist):
    ws = wb.create_sheet("Vendor Coding History")
    cols = ["vendor", "account_code", "invoice_count", "share_of_vendor"]
    style_header(ws, cols)
    r = 2
    for vendor in sorted(hist):
        total = sum(hist[vendor].values())
        for code, count in sorted(hist[vendor].items(), key=lambda kv: -kv[1]):
            ws.cell(row=r, column=1, value=vendor)
            ws.cell(row=r, column=2, value=code)
            ws.cell(row=r, column=3, value=count)
            share = ws.cell(row=r, column=4, value=float((Decimal(count) / Decimal(total) * 100).quantize(Decimal("0.1"))))
            share.number_format = PCT
            r += 1
    autosize(ws, cols, [])


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--invoices")
    ap.add_argument("--coa")
    ap.add_argument("--historical")
    ap.add_argument("--out", required=True)
    ap.add_argument("--min-confidence", type=float, default=None,
                    help="invoices proposed below this confidence become exceptions. "
                         "If omitted, no floor is applied.")
    ap.add_argument("--client", default="")
    ap.add_argument("--template", action="store_true",
                    help="emit a blank invoice intake template instead of coding")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    if args.template:
        write_template(Path(args.out))
        return 0

    if not (args.invoices and args.coa and args.historical):
        sys.exit("--invoices, --coa, and --historical are all required (unless --template)")

    coa = load_coa(Path(args.coa))
    hist, hist_errors = load_historical(Path(args.historical))
    invoices, parse_errors = load_invoices(Path(args.invoices))

    coded, exceptions = propose(invoices, coa, hist, args.min_confidence)

    total_in = len(invoices) + len(parse_errors)
    accounted = len(coded) + len(exceptions) + len(parse_errors)
    errors = list(hist_errors) + list(parse_errors)
    if accounted != total_in:
        errors.append(
            f"internal reconciliation failure: {accounted} rows accounted for vs "
            f"{total_in} rows read - every invoice must land in exactly one bucket"
        )
    valid = not errors

    meta = {
        "Client": args.client or "(not stated)",
        "Invoices read": total_in,
        "Confidence floor": f"{args.min_confidence}%" if args.min_confidence is not None else "none applied",
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    print("=" * 76)
    print("INVOICE EXTRACTION AND GL CODING")
    print("=" * 76)
    print(f"Client            : {meta['Client']}")
    print(f"Invoices read     : {total_in}")
    print(f"Coded w/ confidence: {len(coded)}")
    print(f"Exceptions        : {len(exceptions)}")
    print(f"Confidence floor  : {meta['Confidence floor']}")
    if errors:
        print(f"\nGATE FAILURES ({len(errors)}):")
        for e in errors[:25]:
            print(f"  ! {e}")
        if len(errors) > 25:
            print(f"  ... and {len(errors) - 25} more")

    if not valid and not args.force:
        print("\n" + "=" * 76)
        print("WORKBOOK NOT WRITTEN.")
        print("Every invoice must be accounted for with a parseable amount and date, and")
        print("the historical/invoice population must reconcile row-for-row. Fix the")
        print("flagged rows and re-run.")
        print("=" * 76)
        return 1

    wb = Workbook()
    sheet_summary(wb, meta, len(coded), len(exceptions), len(parse_errors), valid)
    sheet_coded(wb, coded)
    sheet_exceptions(wb, exceptions, parse_errors)
    sheet_history(wb, hist)
    wb.active = 0
    out = Path(args.out)
    if not valid:
        out = out.with_name(out.stem + " [GATE NOT MET]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\n{'GATE CLEAN' if valid else 'NOT MET (forced)'}.  Workbook: {out}")
    return 0 if valid else 1


if __name__ == "__main__":
    sys.exit(main())
