---
name: receipts-to-expense
description: Turn scanned or photographed receipts into an expense spreadsheet with vendor, date, amount, tax, tip and category — proving the extraction by requiring each receipt's line items to sum to its own printed total, every receipt image to be accounted for, and the population total to tie to a control total such as the card statement. Output uses the exact schema that expense-policy-testing consumes. Use this whenever the user mentions receipts, expense receipts, turning receipts into a spreadsheet, scanning or photographing receipts, extracting receipt data, a receipt envelope or shoebox, or building an expense report from receipts. Also trigger on "digitise these receipts," "get these receipts into Excel," "extract the receipt data," "build an expense report from these," or when a folder of receipt images is provided. Runs fully local, including OCR — no receipt image leaves the machine.
---

# Receipts → Expense Spreadsheet

Receipt extraction is harder than statement extraction and it is worth being honest about why. A
bank statement is machine-generated on a consistent template with a text layer. A receipt is
thermal paper that has been folded, sat in a wallet, and photographed at an angle under bad
lighting. **Expect a materially higher exception rate than the bank statement skill produces**, and
plan for the exceptions to be worked rather than assumed away.

What makes the extraction provable anyway is that **most receipts carry their own internal proof**:
subtotal plus tax plus tip equals the printed total. That check costs nothing, is independent of
the OCR, and catches the specific failure that matters — a misread digit in an amount.

## The gate

No clean spreadsheet unless:

1. **Every receipt image is accounted for.** One row per receipt, and the count of rows plus the
   count of exceptions must equal the number of images supplied. A receipt that quietly fails to
   process is the one that goes unclaimed or unreviewed.
2. **Each receipt's components sum to its own printed total.** Where subtotal, tax, and tip are
   legible, `subtotal + tax + tip = total`. A break means a digit was misread, and the script names
   the receipt.
3. **The population total ties to a control total** — the corporate card statement total, or the
   reimbursement claim total. Without it, completeness is unproven.
4. **Nothing is inferred.** An illegible amount, date, or vendor becomes a visible exception with
   the image filename. It is never back-solved from the total, and never guessed from context.

## Inputs

1. **The receipt images or PDFs** — one receipt per file, or a multi-page PDF where each page is a
   receipt. Say which, because it changes the accounting.
2. **A control total** — the card statement total for the period, or the claim total. Ask for it.
3. **The employee** the receipts belong to, and the period.
4. **The category list** the client uses, so categories match their chart of accounts rather than
   being invented.
5. **The card transaction listing**, if available. Matching receipts against card transactions is
   the strongest completeness test there is — it finds both a charge with no receipt and a receipt
   with no charge, and the second is the more interesting one.

## Step 1 — Extract

```bash
python3 scripts/extract_receipts.py --input receipts/ --out work/
# add --ocr for photographs and scans, which is most of them
```

Local OCR only. If `ocrmypdf`/`tesseract` is not installed the script says so and stops rather
than falling back to a cloud service — these are employee expense records.

**OCR vigilance specific to receipts.** Beyond the usual digit confusions, watch for:

- **A decimal point lost on thermal paper**, turning `12.50` into `1250`. The internal sum check
  catches this, which is most of why it exists.
- **The tip written by hand** and the printed total not updated, so the card was charged more than
  the receipt shows. Capture both and let the difference be visible.
- **Two receipts photographed together**, which produces one row where there should be two.
- **The merchant name in a stylised logo** rather than text — often not OCR-readable at all.
- **A duplicate photograph of the same receipt**, which is different from a duplicate claim and
  should not be reported as one.

## Step 2 — Normalise and prove

Build one row per receipt with these fields, then validate:

```bash
python3 scripts/build_expense_sheet.py \
  --receipts receipt_lines.csv \
  --control-total 6284.19 \
  --card-transactions card.csv \
  --employee "A. Reyes" --employee-id E001 \
  --period "2025-11" --client "Brightline Media Group" \
  --out "Brightline - A Reyes - 2025-11 Expenses.xlsx"
```

Five tests:

- **Test 1 — Image accounting.** Rows plus exceptions equals images supplied.
- **Test 2 — Internal sum.** Subtotal + tax + tip = printed total, per receipt.
- **Test 3 — Control total.** Population total agrees to the card statement or claim total.
- **Test 4 — Card matching**, where a transaction listing is supplied: every receipt matched to a
  charge and every charge matched to a receipt, both directions reported.
- **Test 5 — Field completeness.** Vendor, date, amount and category present on every row, with
  anything missing listed as an exception rather than defaulted.

## Step 3 — The output is another skill's input

The spreadsheet's `Expenses` tab uses **exactly the schema `expense-policy-testing` reads**:

`txn_id, txn_date, employee, employee_id, amount, category, merchant, description, approver,
receipt, report_ref, cost_centre, payment_method`

So the chain runs straight through:

```bash
python3 ../receipts-to-expense/scripts/build_expense_sheet.py ... --csv-out expenses.csv
python3 ../expense-policy-testing/scripts/expense_test.py \
    --expenses expenses.csv --policy policy.csv --gl-total 6284.19 --out testing.xlsx
```

`receipt` is populated `yes` on every extracted row, because a receipt image is what produced it —
which means the downstream missing-receipt test is testing the *card charges with no receipt*, not
these. That is the right way round, and worth stating on the workpaper so nobody misreads a clean
missing-receipt result.

`approver`, `report_ref`, and `cost_centre` are left blank deliberately. They are not on the
receipt, and inventing them would defeat the downstream approval tests.

## Step 4 — Deliver

**Workbook tabs:**

1. **Summary** — the five tests, image accounting, control total agreement, category totals, and
   the exception count. The signable page.
2. **Expenses** — the normalised rows in `expense-policy-testing` schema, ready to export.
3. **Receipt Detail** — every receipt with subtotal, tax, tip, total, the internal sum check, and
   the source image filename. A reviewer must be able to open the image behind any row.
4. **Exceptions** — illegible fields, failed sum checks, unmatched receipts, with the image name.
5. **Card Matching** — where a listing was supplied: matched, receipt with no charge, charge with
   no receipt.
6. **Category Summary** — totals by category for coding.

**Then, in chat:** receipts processed, total, whether it ties, the exception count, and anything
that needs the employee to be asked. Be candid about the exception rate — on real receipts it will
not be zero, and pretending otherwise wastes the reviewer's time.

## What to escalate

- **A charge on the card with no receipt.** The finding that matters most, and only visible with the
  transaction listing.
- **A receipt with no matching charge** — reimbursement claimed for something not paid on the card,
  or a personal receipt submitted.
- **A handwritten tip that makes the charge exceed the printed total** by more than the client's
  policy allows.
- **The same receipt image submitted twice**, distinguished from two genuine identical purchases by
  comparing image content rather than only amounts.
- **A receipt dated outside the claim period.**
- **Alcohol, personal, or prohibited categories** where the policy restricts them — flag for the
  downstream policy test rather than concluding here.
- **A receipt that is illegible in the amount field.** Do not estimate it. An estimated expense
  claim is a misstatement, however small.

## Security posture

Fully local, OCR included. No image, no receipt data, and no employee information leaves the
machine. Original images are opened read-only. Intermediate OCR output lands in `work/` and can be
deleted after delivery — say so, because receipt images can carry card numbers and the client may
have a retention policy about them.

## Dependencies

```bash
pip install openpyxl pdfplumber pillow
# local OCR, needed for essentially all photographed receipts:
#   macOS:  brew install ocrmypdf
#   Debian: sudo apt install ocrmypdf tesseract-ocr
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*

---

# Appendix — bundled files

If you installed the `.skill` package these files are already in place and you can ignore this appendix. If you copied the skill as text, create the files below at the paths shown, alongside your `SKILL.md`. The skill will not run without them.

## `scripts/build_expense_sheet.py`

```python
#!/usr/bin/env python3
"""
Validate extracted receipt data and build the expense spreadsheet.

Output uses EXACTLY the schema that expense-policy-testing reads, so the two chain
together without translation.

Gates:
  * every receipt image accounted for - rows + exceptions = images supplied
  * each receipt's components must sum to its own printed total
    (subtotal + tax + tip = total). This check is independent of the OCR and catches
    the failure that matters: a misread digit, or a decimal point lost on thermal paper
  * the population total must tie to a control total
  * nothing inferred - an illegible field becomes a visible exception, never a guess

Usage:
    python3 build_expense_sheet.py --receipts receipt_lines.csv \
        --control-total 6284.19 --card-transactions card.csv \
        --employee "A. Reyes" --employee-id E001 --period 2025-11 \
        --client "Brightline Media Group" \
        --out "Brightline - A Reyes - 2025-11 Expenses.xlsx" \
        --csv-out expenses.csv

--receipts CSV (one row per receipt):
    image_file, receipt_date, merchant, subtotal, tax, tip, total, category,
    description, payment_method, illegible_fields
      illegible_fields: comma-separated field names that could not be read.
      Leave an amount BLANK if illegible - never estimate it.

--card-transactions CSV:
    txn_date, merchant, amount, reference
"""

from __future__ import annotations

import argparse
import csv
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

# The schema expense-policy-testing reads. Do not change without changing that skill.
EXPENSE_SCHEMA = ["txn_id", "txn_date", "employee", "employee_id", "amount", "category",
                  "merchant", "description", "approver", "receipt", "report_ref",
                  "cost_centre", "payment_method"]


def dec(raw):
    if raw is None:
        return None
    s = str(raw).strip()
    if s in ("", "-", "--", "n/a", "N/A", "None", "illegible"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace(",", "").replace("$", "").strip()
    if s.endswith("-"):
        neg, s = True, s[:-1]
    try:
        v = Decimal(s)
    except InvalidOperation:
        raise ValueError(f"Unparseable amount: {raw!r}. Leave it blank if illegible - "
                         f"never estimate it.")
    return -v if neg else v


def clean(s) -> str:
    return " ".join(str(s or "").split())


def key(s) -> str:
    return "".join(ch for ch in str(s or "").lower() if ch.isalnum())


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


def load_receipts(path: Path) -> list[dict]:
    out = []
    for i, r in enumerate(rows_of(path, "receipts"), start=1):
        out.append({
            "row": i,
            "image_file": clean(r.get("image_file")) or f"(no image {i})",
            "receipt_date": pdate(r.get("receipt_date")),
            "merchant": clean(r.get("merchant")),
            "subtotal": dec(r.get("subtotal")),
            "tax": dec(r.get("tax")),
            "tip": dec(r.get("tip")),
            "total": dec(r.get("total")),
            "category": clean(r.get("category")),
            "description": clean(r.get("description")),
            "payment_method": clean(r.get("payment_method")),
            "illegible": [clean(x) for x in
                          clean(r.get("illegible_fields")).split(",") if clean(x)],
            "sum_check": "", "flags": [], "matched_txn": None,
        })
    return out


# ----------------------------------------------------------------------- tests

def validate(receipts, images_supplied: int | None) -> list[str]:
    problems = []
    seen_img = defaultdict(list)

    for x in receipts:
        # internal sum - independent of the OCR, catches a misread digit
        parts = [p for p in (x["subtotal"], x["tax"], x["tip"]) if p is not None]
        if x["total"] is None:
            x["sum_check"] = "total illegible"
            x["flags"].append("TOTAL ILLEGIBLE - do not estimate it. An estimated expense "
                              "claim is a misstatement, however small.")
            problems.append(f"{x['image_file']}: total illegible")
        elif x["subtotal"] is None:
            x["sum_check"] = "subtotal illegible - cannot verify"
            x["flags"].append("subtotal illegible, so the internal sum check could not "
                              "run on this receipt")
        else:
            computed = sum(parts, ZERO)
            diff = computed - x["total"]
            if diff == ZERO:
                x["sum_check"] = "ties"
            else:
                x["sum_check"] = f"OFF BY {diff:,.2f}"
                x["flags"].append(
                    f"subtotal + tax + tip = {computed:,.2f} against a printed total of "
                    f"{x['total']:,.2f}. A digit was misread, or a decimal point was lost "
                    f"on thermal paper.")
                problems.append(f"{x['image_file']}: internal sum off by {diff:,.2f}")

        # tip exceeding what the total shows
        if x["tip"] is not None and x["subtotal"] is not None and x["total"] is not None:
            if x["tip"] > ZERO and x["total"] < sum(
                    [p for p in (x["subtotal"], x["tax"]) if p is not None], ZERO) + x["tip"]:
                x["flags"].append("a handwritten tip may not be included in the printed "
                                  "total - confirm what the card was actually charged")

        for f in ("merchant", "category"):
            if not x[f]:
                x["flags"].append(f"{f} missing")
                problems.append(f"{x['image_file']}: {f} missing")
        if x["receipt_date"] is None:
            x["flags"].append("date illegible or missing")
            problems.append(f"{x['image_file']}: date illegible or missing")
        for f in x["illegible"]:
            x["flags"].append(f"reported illegible by the extractor: {f}")

        seen_img[key(x["image_file"])].append(x["row"])

    for img, rowlist in seen_img.items():
        if len(rowlist) > 1:
            problems.append(f"image {img} produced {len(rowlist)} rows (rows {rowlist}) - "
                            f"either two receipts were photographed together, or the same "
                            f"image was processed twice")
            for x in receipts:
                if key(x["image_file"]) == img:
                    x["flags"].append(f"this image produced {len(rowlist)} rows")

    return problems


def match_card(receipts, card, tol_days: int) -> dict:
    if not card:
        return {"matched": [], "receipt_no_charge": [], "charge_no_receipt": [],
                "applicable": False}
    pool = list(card)
    matched, r_only = [], []
    for x in receipts:
        if x["total"] is None:
            r_only.append(x)
            continue
        hit = None
        for c in pool:
            if c["amount"] != x["total"]:
                continue
            if x["receipt_date"] and c["txn_date"]:
                if abs((c["txn_date"] - x["receipt_date"]).days) > tol_days:
                    continue
            hit = c
            break
        if hit:
            pool.remove(hit)
            x["matched_txn"] = hit
            matched.append((x, hit))
        else:
            r_only.append(x)
            x["flags"].append("no matching card charge - reimbursement claimed for "
                              "something not paid on the card, or a personal receipt")
    return {"matched": matched, "receipt_no_charge": r_only,
            "charge_no_receipt": pool, "applicable": True}


def to_expense_rows(receipts, employee, employee_id) -> list[dict]:
    out = []
    for i, x in enumerate(receipts, start=1):
        if x["total"] is None:
            continue
        out.append({
            "txn_id": f"RCPT{i:05d}",
            "txn_date": x["receipt_date"].isoformat() if x["receipt_date"] else "",
            "employee": employee,
            "employee_id": employee_id,
            "amount": f"{x['total']:.2f}",
            "category": x["category"],
            "merchant": x["merchant"],
            "description": x["description"],
            "approver": "",          # not on the receipt - inventing it would defeat
            "receipt": "yes",        # a receipt image is what produced this row
            "report_ref": "",        # the downstream approval tests
            "cost_centre": "",
            "payment_method": x["payment_method"],
        })
    return out


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


def sheet_summary(wb, meta, tests, stats, cats, esc, overall, control):
    ws = wb.active
    ws.title = "Summary"
    widths(ws, {1: 4, 2: 48, 3: 16, 4: 14, 5: 62})
    r = 1

    def line(label, v=None, s=None, *, bold=False, size=11, fill=None, note=""):
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

    line("RECEIPTS TO EXPENSE SPREADSHEET", bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    v = ws.cell(row=r, column=2,
                value="PROVEN - every receipt accounted for and every total ties"
                if overall else
                "NOT PROVEN - see the failing test(s). On real receipts the exception "
                "rate will not be zero; work the exceptions rather than assuming them away.")
    v.font = OK_FONT if overall else BAD_FONT
    if not overall:
        v.fill = BAD_FILL
    r += 2

    line("POPULATION", bold=True, size=12, fill=SUB_FILL)
    line("  Receipts processed", stats["n"])
    line("  Rows produced", stats["rows"])
    line("  Exceptions", stats["exceptions"],
         fill=WARN_FILL if stats["exceptions"] else None)
    line("  Total extracted", stats["total"], bold=True)
    if control is not None:
        line("  Control total", control)
        line("  Difference", stats["total"] - control, bold=True,
             fill=None if stats["total"] == control else BAD_FILL)
        c = ws.cell(row=r - 1, column=3)
        c.font = OK_FONT if stats["total"] == control else BAD_FONT
    else:
        line("  Control total", "NOT SUPPLIED", fill=BAD_FILL,
             note="Without a control total, completeness is unproven.")
    r += 1

    for t in tests:
        line(f"TEST {t['num']} - {t['name']}", None,
             "PASS" if t["passed"] else "FAIL", note=t.get("note", ""))
    r += 1

    if cats:
        line("BY CATEGORY", bold=True, size=12, fill=SUB_FILL)
        for cat, amt in sorted(cats.items(), key=lambda kv: -kv[1]):
            line("  " + (cat or "(uncategorised)"), amt)
        r += 1

    line("DOWNSTREAM", bold=True, size=12, fill=SUB_FILL)
    line("  The Expenses tab uses exactly the schema expense-policy-testing reads. "
         "Export it and run that skill next.")
    line("  Every row carries receipt = yes, because a receipt image produced it. The "
         "downstream missing-receipt test therefore tests CARD CHARGES with no receipt, "
         "not these rows - state that on the workpaper so a clean result is not "
         "misread.", bold=True)
    r += 1

    if esc:
        line("ESCALATE", bold=True, size=12, fill=BAD_FILL)
        for e in esc:
            line("  " + e)
        r += 1
    r += 1
    line("Prepared by / date:  __________________  ____________", bold=True)


def sheet_expenses(wb, rows):
    ws = wb.create_sheet("Expenses")
    ws.append(EXPENSE_SCHEMA)
    hdr(ws, len(EXPENSE_SCHEMA))
    for x in rows:
        ws.append([x[k] for k in EXPENSE_SCHEMA])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[4].number_format = MONEY
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value="No rows produced.")
    else:
        r = ws.max_row + 2
        ws.cell(row=r, column=1,
                value="This is expense-policy-testing's input schema exactly. approver, "
                      "report_ref and cost_centre are blank deliberately - they are not on "
                      "a receipt, and inventing them would defeat the downstream approval "
                      "tests.").font = Font(italic=True)
    ws.freeze_panes = "B2"
    widths(ws, {1: 12, 2: 12, 3: 22, 4: 12, 5: 13, 6: 20, 7: 26, 8: 34, 9: 12,
                10: 9, 11: 12, 12: 12, 13: 18})


def sheet_detail(wb, receipts):
    ws = wb.create_sheet("Receipt Detail")
    heads = ["Image file", "Date", "Merchant", "Subtotal", "Tax", "Tip", "Total",
             "Internal sum check", "Category", "Description", "Card match", "Flags"]
    ws.append(heads)
    hdr(ws, len(heads))
    for x in sorted(receipts, key=lambda y: (0 if y["flags"] else 1, y["image_file"])):
        ws.append([x["image_file"], x["receipt_date"], x["merchant"],
                   float(x["subtotal"]) if x["subtotal"] is not None else None,
                   float(x["tax"]) if x["tax"] is not None else None,
                   float(x["tip"]) if x["tip"] is not None else None,
                   float(x["total"]) if x["total"] is not None else None,
                   x["sum_check"], x["category"], x["description"],
                   x["matched_txn"]["reference"] if x["matched_txn"] else "",
                   "; ".join(x["flags"])])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[1].number_format = DATEF
        for i in (3, 4, 5, 6):
            row[i].number_format = MONEY
        if row[7].value == "ties":
            row[7].font = OK_FONT
        elif row[7].value:
            row[7].font, row[7].fill = BAD_FONT, BAD_FILL
        if row[11].value:
            row[11].fill = WARN_FILL
    r = ws.max_row + 2
    ws.cell(row=r, column=1,
            value="The image filename is on every row so a reviewer can open the receipt "
                  "behind any figure. The internal sum check is independent of the OCR - it "
                  "is what catches a misread digit or a lost decimal point."
            ).font = Font(italic=True)
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:L{max(ws.max_row, 2)}"
    widths(ws, {1: 28, 2: 12, 3: 26, 4: 13, 5: 11, 6: 11, 7: 13, 8: 20, 9: 20,
                10: 30, 11: 16, 12: 70})


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
    widths(ws, {i: 20 for i in range(1, len(header) + 1)} | {1: 28,
                                                            len(header): 66})


# ------------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--receipts", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--csv-out", help="also write the Expenses tab as CSV, ready for "
                                      "expense-policy-testing")
    ap.add_argument("--control-total")
    ap.add_argument("--images-supplied", type=int,
                    help="number of receipt images given to the extractor")
    ap.add_argument("--card-transactions")
    ap.add_argument("--date-tolerance", type=int, default=3)
    ap.add_argument("--employee", default="")
    ap.add_argument("--employee-id", default="")
    ap.add_argument("--period", default="")
    ap.add_argument("--client", default="")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    receipts = load_receipts(Path(args.receipts))
    card = []
    if args.card_transactions and Path(args.card_transactions).exists():
        for r in rows_of(Path(args.card_transactions), "card transactions"):
            card.append({"txn_date": pdate(r.get("txn_date")),
                         "merchant": clean(r.get("merchant")),
                         "amount": dec(r.get("amount")) or ZERO,
                         "reference": clean(r.get("reference"))})

    problems = validate(receipts, args.images_supplied)
    cm = match_card(receipts, card, args.date_tolerance)
    exp_rows = to_expense_rows(receipts, args.employee, args.employee_id)

    control = dec(args.control_total)
    total = sum((x["total"] for x in receipts if x["total"] is not None), ZERO)
    exceptions = [x for x in receipts if x["flags"]]
    images = args.images_supplied if args.images_supplied is not None else len(receipts)

    cats = defaultdict(lambda: ZERO)
    for x in receipts:
        if x["total"] is not None:
            cats[x["category"]] += x["total"]

    sum_breaks = [x for x in receipts if x["sum_check"].startswith("OFF BY")]
    missing_fields = [x for x in receipts
                      if any(f.endswith("missing") or "illegible" in f
                             for f in x["flags"])]

    stats = {"n": len(receipts), "rows": len(exp_rows),
             "exceptions": len(exceptions), "total": total}

    tests = [
        {"num": 1, "name": "Every receipt image accounted for",
         "passed": len(receipts) == images,
         "note": "" if len(receipts) == images else
                 f"{len(receipts)} row(s) from {images} image(s) supplied - a receipt "
                 f"that quietly fails to process is the one that goes unclaimed"},
        {"num": 2, "name": "Each receipt's components sum to its printed total",
         "passed": not sum_breaks,
         "note": "" if not sum_breaks else
                 f"{len(sum_breaks)} receipt(s) fail the internal sum - a digit was "
                 f"misread, or a decimal point was lost"},
        {"num": 3, "name": "Population total ties to the control total",
         "passed": control is not None and total == control,
         "note": ("no control total supplied - completeness is unproven"
                  if control is None else
                  "" if total == control else
                  f"difference {total - control:,.2f}")},
        {"num": 4, "name": "Receipts matched to card transactions",
         "passed": cm["applicable"] and not cm["receipt_no_charge"]
                   and not cm["charge_no_receipt"],
         "note": ("no card transaction listing supplied - this is the strongest "
                  "completeness test available and finds a charge with no receipt"
                  if not cm["applicable"] else
                  f"{len(cm['receipt_no_charge'])} receipt(s) with no charge, "
                  f"{len(cm['charge_no_receipt'])} charge(s) with no receipt")},
        {"num": 5, "name": "Every row has vendor, date, amount and category",
         "passed": not missing_fields,
         "note": "" if not missing_fields else
                 f"{len(missing_fields)} receipt(s) with a missing or illegible field - "
                 f"listed as exceptions rather than defaulted"},
    ]

    esc = []
    for c in cm["charge_no_receipt"]:
        esc.append(f"Card charge {c['txn_date']} {c['merchant']} {c['amount']:,.2f} has no "
                   f"receipt. This is the finding that matters most, and it is only "
                   f"visible with the transaction listing.")
    for x in cm["receipt_no_charge"]:
        if x["total"] is not None:
            esc.append(f"{x['image_file']}: receipt for {x['total']:,.2f} with no matching "
                       f"card charge - reimbursement claimed for something not paid on the "
                       f"card, or a personal receipt.")
    for x in sum_breaks:
        esc.append(f"{x['image_file']}: internal sum does not tie. Re-read the receipt "
                   f"before using the amount.")
    for x in receipts:
        if x["total"] is None:
            esc.append(f"{x['image_file']}: amount illegible. Do not estimate it - an "
                       f"estimated expense claim is a misstatement, however small.")
    if control is not None and total != control:
        esc.append(f"Extracted total {total:,.2f} does not agree to the control total "
                   f"{control:,.2f} (difference {total - control:,.2f}).")

    overall = all(t["passed"] for t in tests)

    meta = {
        "Client": args.client or "(not stated)",
        "Employee": args.employee or "(not stated)",
        "Employee ID": args.employee_id or "(not stated)",
        "Period": args.period or "(not stated)",
        "Images supplied": images,
        "Card listing supplied": "yes" if card else "no",
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    # ---- console
    print("=" * 76)
    print("RECEIPTS TO EXPENSE SPREADSHEET")
    print("=" * 76)
    print(f"Client   : {meta['Client']}   Employee: {meta['Employee']}   "
          f"Period: {meta['Period']}")
    print(f"Receipts {len(receipts)} from {images} image(s)   rows produced "
          f"{len(exp_rows)}   exceptions {len(exceptions)}")
    print(f"Total extracted {total:>16,.2f}")
    if control is not None:
        d = total - control
        print(f"Control total   {control:>16,.2f}   difference {d:>14,.2f}  "
              f"{'TIES' if d == ZERO else '*** DOES NOT TIE ***'}")
    else:
        print("Control total   NOT SUPPLIED")
    print()
    for t in tests:
        print(f"Test {t['num']}: {'PASS' if t['passed'] else '*** FAIL ***':<14} {t['name']}")
        if t.get("note"):
            print(f"         {t['note']}")
    if cats:
        print("\nBy category:")
        for cat, amt in sorted(cats.items(), key=lambda kv: -kv[1]):
            print(f"  {(cat or '(uncategorised)'):<24} {amt:>13,.2f}")
    if problems:
        print(f"\nEXTRACTION PROBLEMS ({len(problems)}):")
        for p in problems[:15]:
            print(f"  ! {p}")
    if esc:
        print(f"\nESCALATE ({len(esc)}):")
        for e in esc[:12]:
            print(f"  ! {e}")

    if not overall and not args.force:
        print("\n" + "=" * 76)
        print("WORKBOOK NOT WRITTEN.")
        print("On real receipts the exception rate will not be zero - work the exceptions")
        print("rather than assuming them away. Never estimate an illegible amount.")
        print("=" * 76)
        return 1

    wb = Workbook()
    sheet_summary(wb, meta, tests, stats, dict(cats), esc, overall, control)
    sheet_expenses(wb, exp_rows)
    sheet_detail(wb, receipts)
    sheet_list(wb, "Exceptions", ["Image file", "Issue"],
               [[x["image_file"], f] for x in receipts for f in x["flags"]], {},
               "None - every receipt extracted cleanly.",
               footer="Every exception carries its image filename so the receipt can be "
                      "re-read.")
    sheet_list(wb, "Card Matching",
               ["Status", "Date", "Merchant", "Amount", "Reference / image"],
               ([["matched", x["receipt_date"], x["merchant"], float(x["total"]),
                  c["reference"]] for x, c in cm["matched"]]
                + [["RECEIPT WITH NO CHARGE", x["receipt_date"], x["merchant"],
                    float(x["total"]) if x["total"] is not None else None,
                    x["image_file"]] for x in cm["receipt_no_charge"]]
                + [["CHARGE WITH NO RECEIPT", c["txn_date"], c["merchant"],
                    float(c["amount"]), c["reference"]]
                   for c in cm["charge_no_receipt"]]),
               {1: DATEF, 3: MONEY},
               "No card transaction listing supplied - matching not performed.",
               footer="A charge with no receipt is the finding that matters most. A receipt "
                      "with no charge is the more interesting of the two directions.")
    sheet_list(wb, "Category Summary", ["Category", "Total"],
               [[k or "(uncategorised)", float(v)]
                for k, v in sorted(cats.items(), key=lambda kv: -kv[1])],
               {1: MONEY}, "No categorised receipts.")
    wb.active = 0
    out = Path(args.out)
    if not overall:
        out = out.with_name(out.stem + " [NOT PROVEN]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\n{'PROVEN' if overall else 'NOT PROVEN (forced)'}.  Workbook: {out}")

    if args.csv_out:
        p = Path(args.csv_out)
        with p.open("w", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=EXPENSE_SCHEMA)
            w.writeheader()
            w.writerows(exp_rows)
        print(f"Expenses CSV: {p}")
        print("  Ready for expense-policy-testing:")
        print(f"    python3 ../expense-policy-testing/scripts/expense_test.py \\")
        print(f"        --expenses {p} --policy policy.csv --gl-total {total:.2f} \\")
        print(f"        --out testing.xlsx")
    return 0 if overall else 1


if __name__ == "__main__":
    sys.exit(main())
```

## `scripts/extract_receipts.py`

```python
#!/usr/bin/env python3
"""
Extract the text layer and word geometry from receipt images and PDFs, fully locally.

Bundled with this skill so a single-skill install works standalone. Receipts are
almost always photographs or scans with no text layer, so --ocr is the normal path
rather than the exception.

Writes, per input PDF:
    <stem>.text.txt    page-delimited text
    <stem>.words.jsonl one JSON object per word with coordinates
    <stem>.meta.json   page count, per-page char counts, scanned flag

Usage:
    python3 extract_text.py --input statements/ --out work/
    python3 extract_text.py --input jan.pdf --out work/ --ocr

No network access. Originals are opened read-only and never modified.
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    sys.exit("pdfplumber is required.  pip install pdfplumber")

# A page with fewer than this many extractable characters is treated as a scan.
SCAN_CHAR_THRESHOLD = 60


def ocr_pdf(src: Path) -> Path:
    """Run local OCR, returning a path to a new searchable PDF. Never touches src."""
    ocrmypdf = shutil.which("ocrmypdf")
    if not ocrmypdf:
        sys.exit(
            "This PDF has no text layer and 'ocrmypdf' is not installed.\n"
            "Install it locally, then re-run with --ocr:\n"
            "  macOS:  brew install ocrmypdf\n"
            "  Debian: sudo apt install ocrmypdf tesseract-ocr\n"
            "Do not upload employee receipt images to a cloud OCR service - they can carry card numbers."
        )
    # Resolve to an absolute input path and confirm it is a real PDF before handing
    # it to another process. No shell is involved (list argv, shell=False), and the
    # binary is the absolute path resolved above rather than a PATH lookup at spawn.
    src = src.resolve(strict=True)
    if not src.is_file() or src.suffix.lower() != ".pdf":
        sys.exit(f"Not a PDF file: {src}")

    out = Path(tempfile.mkdtemp(prefix="ocr_")) / f"{src.stem}.ocr.pdf"
    print(f"  OCR (local): {src.name} -> {out.name}", flush=True)
    subprocess.run(
        [
            ocrmypdf,
            "--force-ocr",
            "--deskew",
            "--optimize", "0",
            "--quiet",
            str(src),
            str(out),
        ],
        check=True,
    )
    return out


def extract(pdf_path: Path, out_dir: Path, use_ocr: bool) -> dict:
    stem = pdf_path.stem
    source = pdf_path

    if use_ocr:
        source = ocr_pdf(pdf_path)

    text_lines: list[str] = []
    words_out: list[str] = []
    page_meta: list[dict] = []

    with pdfplumber.open(str(source)) as pdf:
        for pageno, page in enumerate(pdf.pages, start=1):
            txt = page.extract_text() or ""
            text_lines.append(f"\n===== PAGE {pageno} =====\n{txt}")

            for w in page.extract_words(
                use_text_flow=False,
                keep_blank_chars=False,
                extra_attrs=["size"],
            ):
                words_out.append(
                    json.dumps(
                        {
                            "page": pageno,
                            "text": w["text"],
                            "x0": round(float(w["x0"]), 2),
                            "x1": round(float(w["x1"]), 2),
                            "top": round(float(w["top"]), 2),
                            "bottom": round(float(w["bottom"]), 2),
                            "size": round(float(w.get("size", 0)), 1),
                        },
                        ensure_ascii=False,
                    )
                )

            page_meta.append(
                {
                    "page": pageno,
                    "chars": len(txt),
                    "scanned": len(txt.strip()) < SCAN_CHAR_THRESHOLD,
                    "width": round(float(page.width), 1),
                    "height": round(float(page.height), 1),
                }
            )

    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / f"{stem}.text.txt").write_text("\n".join(text_lines), encoding="utf-8")
    (out_dir / f"{stem}.words.jsonl").write_text(
        "\n".join(words_out) + "\n", encoding="utf-8"
    )

    scanned_pages = [p["page"] for p in page_meta if p["scanned"]]
    meta = {
        "source_file": pdf_path.name,
        "ocr_applied": use_ocr,
        "page_count": len(page_meta),
        "scanned_pages": scanned_pages,
        "scanned": bool(scanned_pages),
        "pages": page_meta,
    }
    (out_dir / f"{stem}.meta.json").write_text(
        json.dumps(meta, indent=2), encoding="utf-8"
    )
    return meta


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input", required=True, help="PDF file or folder of PDFs")
    ap.add_argument("--out", default="work", help="output directory (default: work)")
    ap.add_argument(
        "--ocr",
        action="store_true",
        help="run local OCR first (required for scanned statements)",
    )
    args = ap.parse_args()

    src = Path(args.input).expanduser()
    out_dir = Path(args.out).expanduser()

    if src.is_dir():
        pdfs = sorted(src.glob("*.pdf")) + sorted(src.glob("*.PDF"))
    elif src.is_file():
        pdfs = [src]
    else:
        sys.exit(f"Not found: {src}")

    if not pdfs:
        sys.exit(f"No PDFs found in {src}")

    print(f"Extracting {len(pdfs)} PDF(s) -> {out_dir}/")
    needs_ocr: list[str] = []

    for p in pdfs:
        print(f"- {p.name}")
        meta = extract(p, out_dir, args.ocr)
        flag = ""
        if meta["scanned"] and not args.ocr:
            needs_ocr.append(p.name)
            flag = f"  [SCANNED pages {meta['scanned_pages']} - re-run with --ocr]"
        print(f"    {meta['page_count']} page(s){flag}")

    if needs_ocr:
        print(
            "\nThese files have pages with no usable text layer and MUST be re-run "
            "with --ocr before you read any figures from them:"
        )
        for n in needs_ocr:
            print(f"  - {n}")
        print(
            "Reading amounts off a scanned page without OCR produces silent blanks, "
            "not errors."
        )
        return 2

    print("\nDone. Text layer looks usable on every page.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

