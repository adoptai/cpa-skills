---
name: bank-statement-to-excel
description: Convert PDF bank statements (or credit card, brokerage, and merchant-processor statements) into a clean Excel workbook with dates, descriptions, debits, credits, and running balance — and prove the extraction is complete by tying to the statement's own opening and closing balances. Use this whenever the user mentions converting, extracting, digitizing, or "getting into Excel" a bank statement, bank PDF, credit card statement, or transaction history; asks to turn statements into a spreadsheet or CSV; needs statement data for a reconciliation, cash-flow analysis, or lender package; or drops PDF statements in a folder and asks for the transactions. Also trigger on "parse this statement," "OCR these statements," "statement to spreadsheet," "I have 12 months of PDFs," or any request where transaction-level detail is locked inside a PDF. Runs fully local — no statement data leaves the machine.
---

# Bank Statement → Excel

You are extracting the transaction detail from a statement PDF into a workbook that a
preparer, reviewer, or auditor can rely on without re-reading the PDF.

The standard is not "did you get most of the rows." The standard is: **the extraction
mathematically ties to the statement's own opening and closing balance.** If it doesn't
tie, the file is not delivered — it's flagged. A silently incomplete statement extract is
worse than no extract at all, because it will be relied on.

## Non-negotiables

1. **Never invent, estimate, or interpolate a number.** If a figure is illegible, missing,
   or ambiguous, it goes in the Exceptions tab with the page number and the reason. You do
   not fill the gap by back-solving from the balance column.
2. **Every output row carries its source page.** A reviewer must be able to jump from any
   row in Excel to the exact page of the PDF.
3. **Debits and credits get separate columns.** Never a single signed amount column — sign
   conventions get misread downstream and reverse the reconciliation.
4. **Money is decimal, never float.** Use Python's `Decimal`. Do not round anything; carry
   two decimal places as presented.
5. **The proof runs before delivery, not after.** See Step 4.

## Workflow

### Step 1 — Inventory the source

List every PDF in scope. For each, capture: file name, page count, account number (last 4
only in filenames and tab names — never write a full account number into an output
filename), statement period, and whether the page has a real text layer or is a scan.

```bash
python3 scripts/extract_text.py --input <folder-or-file> --out work/
```

This writes, per PDF:
- `<name>.text.txt` — page-delimited text layer
- `<name>.words.jsonl` — every word with x0/x1/top/bottom coordinates (needed when columns
  collide or descriptions wrap)
- `<name>.meta.json` — page count, per-page character count, `scanned: true|false`

A page with near-zero extractable characters is a scan. Route those through Step 2. Do not
attempt to read numbers off a scan from the text layer — you will get silent blanks.

### Step 2 — OCR only if needed (optional, local)

If `scanned: true`, and `ocrmypdf`/`tesseract` are available locally:

```bash
python3 scripts/extract_text.py --input <file.pdf> --out work/ --ocr
```

If OCR tooling is not installed, say so plainly and stop — offer the install command
(`brew install ocrmypdf` / `apt install ocrmypdf tesseract-ocr`). Do not send the document
to a cloud OCR service. These are client bank records.

**OCR-specific vigilance.** On OCR'd output, treat these as suspect until confirmed against
the balance column: `0/O`, `1/l/I`, `5/S`, `8/B`, `6/G`, comma-vs-period decimal separators,
and dropped leading digits (`1,240.00` read as `240.00`). The balance proof in Step 4 is
what catches these — which is exactly why it is mandatory on OCR'd statements.

### Step 3 — Normalize the transactions

Read the text and word-coordinate output and build a normalized table. You are doing the
layout reasoning here; the scripts do the arithmetic and the file writing.

Columns:

| Field | Rule |
|---|---|
| `row_id` | Sequential from 1, in statement order. Never re-sorted. |
| `source_file` | PDF filename |
| `source_page` | 1-based page number |
| `txn_date` | ISO `YYYY-MM-DD`. See year-assignment rule below. |
| `post_date` | ISO, if the statement shows a separate posting date; else blank |
| `description` | Full description, wrapped lines joined with a single space |
| `check_no` | If present |
| `reference` | Bank reference / trace / confirmation number if present |
| `debit` | Withdrawals/charges. Blank if not a debit. |
| `credit` | Deposits/payments. Blank if not a credit. |
| `balance` | Running balance **as printed on the statement**. Blank if the statement doesn't print one per line. |
| `section` | e.g. `Deposits`, `Checks Paid`, `Electronic Withdrawals`, `Fees` |

**Year assignment.** Most statements print `MM/DD`. Derive the year from the statement
period, and handle the December/January rollover explicitly: on a statement period of
12/16/2025–01/15/2026, a `12/28` row is 2025 and a `01/03` row is 2026. Getting this wrong
silently misstates a fiscal year cutoff. If a date cannot be resolved to a year with
certainty, it is an exception — not a guess.

**Lines that are not transactions.** Exclude, and do not count toward totals: `Balance
Forward`, `Beginning Balance`, `Ending Balance`, `Total Deposits`, `Total Withdrawals`,
`Continued on next page`, subtotals, column headers repeated on continuation pages, daily
balance summary tables, interest-rate disclosures, and the check-image grid. These are the
single most common source of double-counting. If you include a subtotal as a transaction,
your proof will fail by exactly that subtotal — which is a useful tell.

**Wrapped descriptions.** A description continuing on the next physical line has no date and
no amount. Join it to the transaction above. Use the word coordinates to confirm it starts
in the description column, not the date column.

**Multi-account statements.** Some PDFs contain several accounts. Split by account into
separate normalized tables and separate worksheet tabs. Never commingle — each account gets
its own independent balance proof.

Write the normalized table to `work/<name>.transactions.csv`.

### Step 4 — Prove it (mandatory gate)

```bash
python3 scripts/build_workbook.py \
  --transactions work/<name>.transactions.csv \
  --opening <opening balance from statement> \
  --closing <closing balance from statement> \
  --out "<Client> - <Bank> x<last4> - <Period> - Transactions.xlsx"
```

Three independent tests run, all of which must pass:

- **Test A — Balance roll-forward.** `opening + Σcredits − Σdebits = closing`, using the
  opening and closing balances printed on the statement. Any difference is reported to the
  cent.
- **Test B — Statement control totals.** Where the statement prints "Total Deposits" and
  "Total Withdrawals," extracted sums must agree to those figures.
- **Test C — Line-balance continuity.** Where the statement prints a per-line running
  balance: for every row, `prior balance ± amount = printed balance`. This isolates *which*
  row is wrong rather than just telling you the total is off — it is the fastest debugging
  tool you have and it catches OCR digit errors precisely.

If any test fails, the script exits non-zero and writes a break report. **Do not hand the
workbook over.** Go find the row. In practice the cause is one of five things, in
descending order of frequency: a subtotal captured as a transaction, a wrapped description
line captured as a zero-amount transaction, a transaction on a continuation page missed
entirely, a debit classified as a credit, or an OCR digit misread. Test C tells you which.

### Step 5 — Deliver

The workbook has four tabs, in this order:

1. **Transactions** — the normalized rows. Frozen header, autofilter, dates as real dates,
   amounts as accounting-format numbers (not text), `source_page` visible. A `Recomputed
   Balance` column sits beside the printed `balance` so a reviewer can see they agree.
2. **Proof** — the three tests, showing opening balance, total credits, total debits,
   computed closing, statement closing, and the difference (which reads `0.00`). This is
   the tab a reviewer looks at first. It is what makes the file signable.
3. **Exceptions** — every illegible figure, unresolved date, ambiguous classification, and
   judgment call, with page reference. An empty Exceptions tab is a valid and good outcome;
   a *missing* Exceptions tab is not.
4. **Source Map** — one row per PDF page: file, page, transactions extracted from that page,
   page subtotal. Lets a reviewer sample-check any page in seconds.

Then report in chat, briefly: statements processed, period covered, transaction count,
total debits and credits, proof result, and the exception count. If anything failed, lead
with that.

## Multi-statement engagements

For a 12-month or multi-account set, run each statement independently — each gets its own
proof — then add a **Continuity** tab: each statement's closing balance must equal the next
statement's opening balance, in date order. A gap here means a missing statement month, and
that is worth flagging loudly. It is the failure that most often makes it all the way to a
lender or an auditor unnoticed.

## Security posture

Everything runs locally: `pdfplumber` for extraction, optional local `ocrmypdf`/`tesseract`,
`openpyxl` for output. No network calls, no cloud OCR, no uploads, no telemetry. Original
PDFs are opened read-only and never modified. Intermediate files land in `work/` and can be
deleted after delivery — tell the user they exist so they can shred them under their
document-retention policy.

## Dependencies

```bash
pip install pdfplumber openpyxl
# optional, for scanned statements:
#   macOS:  brew install ocrmypdf
#   Debian: sudo apt install ocrmypdf tesseract-ocr
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*

---

# Appendix — bundled files

If you installed the `.skill` package these files are already in place and you can ignore this appendix. If you copied the skill as text, create the files below at the paths shown, alongside your `SKILL.md`. The skill will not run without them.

## `scripts/build_workbook.py`

```python
#!/usr/bin/env python3
"""
Prove a normalized statement extract and write the deliverable workbook.

Runs three independent completeness tests and REFUSES to produce a clean
workbook if any of them fail:

  Test A  Balance roll-forward:  opening + credits - debits == closing
  Test B  Statement control totals (if supplied)
  Test C  Line-balance continuity, where the statement prints a running balance

Usage:
    python3 build_workbook.py \
        --transactions work/jan.transactions.csv \
        --opening 42150.22 --closing 38904.71 \
        --out "Acme - First National x4471 - 2026-01 - Transactions.xlsx"

Optional:
    --stmt-deposits 18200.00 --stmt-withdrawals 21445.51   (enables Test B)
    --account-label "First National x4471"
    --period "2026-01-01 to 2026-01-31"
    --force        write the workbook even if a test fails (still marked FAILED)

Input CSV columns (blank where not applicable):
    row_id, source_file, source_page, txn_date, post_date, description,
    check_no, reference, debit, credit, balance, section

Money is handled as Decimal throughout. Nothing is rounded.
"""

from __future__ import annotations

import argparse
import csv
import sys
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is required.  pip install openpyxl")

ZERO = Decimal("0.00")
MONEY_FMT = '#,##0.00;[Red](#,##0.00)'
DATE_FMT = "yyyy-mm-dd"

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
PASS_FONT = Font(bold=True, color="006100")
FAIL_FONT = Font(bold=True, color="9C0006")
FAIL_FILL = PatternFill("solid", fgColor="FFC7CE")

COLUMNS = [
    "row_id", "source_file", "source_page", "txn_date", "post_date",
    "description", "check_no", "reference", "section",
    "debit", "credit", "balance",
]


def dec(raw) -> Decimal | None:
    """Parse a money string to Decimal. Blank -> None. Never guesses."""
    if raw is None:
        return None
    s = str(raw).strip()
    if s in ("", "-", "--", "n/a", "N/A", "None"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace(",", "").replace("$", "").strip()
    if s.endswith("-"):          # trailing-minus convention
        neg, s = True, s[:-1]
    try:
        v = Decimal(s)
    except InvalidOperation:
        raise ValueError(f"Unparseable amount: {raw!r}")
    return -v if neg else v


def parse_date(raw) -> date | None:
    s = str(raw or "").strip()
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        raise ValueError(f"Dates must be ISO YYYY-MM-DD, got {raw!r}")


def load_rows(path: Path) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No rows in {path}")

    out = []
    for i, r in enumerate(rows, start=1):
        d, c = dec(r.get("debit")), dec(r.get("credit"))
        if d is not None and c is not None:
            raise ValueError(
                f"Row {i}: both debit and credit populated. A transaction is one "
                f"or the other. ({r.get('description','')!r})"
            )
        if d is None and c is None:
            raise ValueError(
                f"Row {i}: neither debit nor credit populated. This is usually a "
                f"wrapped description line or a subtotal that should not be a "
                f"transaction row. ({r.get('description','')!r})"
            )
        out.append({
            "row_id": int(r.get("row_id") or i),
            "source_file": (r.get("source_file") or "").strip(),
            "source_page": int(r["source_page"]) if str(r.get("source_page") or "").strip() else None,
            "txn_date": parse_date(r.get("txn_date")),
            "post_date": parse_date(r.get("post_date")),
            "description": " ".join((r.get("description") or "").split()),
            "check_no": (r.get("check_no") or "").strip(),
            "reference": (r.get("reference") or "").strip(),
            "section": (r.get("section") or "").strip(),
            "debit": d, "credit": c,
            "balance": dec(r.get("balance")),
        })
    return out


# --------------------------------------------------------------------------- tests

def test_a(rows, opening: Decimal, closing: Decimal) -> dict:
    deb = sum((r["debit"] for r in rows if r["debit"] is not None), ZERO)
    cre = sum((r["credit"] for r in rows if r["credit"] is not None), ZERO)
    computed = opening + cre - deb
    diff = computed - closing
    return {
        "name": "Test A - Balance roll-forward",
        "detail": "opening + total credits - total debits = statement closing balance",
        "debits": deb, "credits": cre,
        "computed_closing": computed, "stmt_closing": closing,
        "difference": diff, "passed": diff == ZERO,
    }


def test_b(rows, stmt_dep, stmt_wd) -> dict | None:
    if stmt_dep is None and stmt_wd is None:
        return None
    deb = sum((r["debit"] for r in rows if r["debit"] is not None), ZERO)
    cre = sum((r["credit"] for r in rows if r["credit"] is not None), ZERO)
    parts, ok = [], True
    if stmt_dep is not None:
        d = cre - stmt_dep
        ok &= d == ZERO
        parts.append(("Total deposits / credits", cre, stmt_dep, d))
    if stmt_wd is not None:
        d = deb - stmt_wd
        ok &= d == ZERO
        parts.append(("Total withdrawals / debits", deb, stmt_wd, d))
    return {
        "name": "Test B - Statement control totals",
        "detail": "extracted sums agree to the totals printed on the statement",
        "parts": parts, "passed": ok,
    }


def test_c(rows, opening: Decimal) -> dict:
    """Walk the printed running balance. Isolates the exact offending row."""
    with_bal = [r for r in rows if r["balance"] is not None]
    if not with_bal:
        return {
            "name": "Test C - Line-balance continuity",
            "detail": "statement does not print a per-line running balance - not applicable",
            "breaks": [], "passed": True, "applicable": False,
        }
    breaks, prior = [], opening
    for r in rows:
        amt = (r["credit"] or ZERO) - (r["debit"] or ZERO)
        expected = prior + amt
        if r["balance"] is not None:
            if r["balance"] != expected:
                breaks.append({
                    "row_id": r["row_id"], "page": r["source_page"],
                    "date": r["txn_date"], "description": r["description"][:70],
                    "prior": prior, "amount": amt,
                    "expected": expected, "printed": r["balance"],
                    "difference": r["balance"] - expected,
                })
            prior = r["balance"]      # resync so one bad row doesn't cascade
        else:
            prior = expected
    return {
        "name": "Test C - Line-balance continuity",
        "detail": "prior balance +/- amount = printed balance, for every row",
        "breaks": breaks, "passed": not breaks, "applicable": True,
    }


def recomputed_balances(rows, opening: Decimal) -> list[Decimal]:
    out, run = [], opening
    for r in rows:
        run += (r["credit"] or ZERO) - (r["debit"] or ZERO)
        out.append(run)
    return out


# --------------------------------------------------------------------------- excel

def style_header(ws, ncols: int, row: int = 1) -> None:
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 28


def autosize(ws, widths: dict[int, int]) -> None:
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def write_transactions(wb, rows, opening) -> None:
    ws = wb.active
    ws.title = "Transactions"
    headers = [
        "Row", "Source File", "Pg", "Txn Date", "Post Date", "Description",
        "Check No", "Reference", "Section", "Debit", "Credit",
        "Balance (printed)", "Balance (recomputed)", "Tie",
    ]
    ws.append(headers)
    style_header(ws, len(headers))

    recomp = recomputed_balances(rows, opening)
    for r, rb in zip(rows, recomp):
        tie = "" if r["balance"] is None else ("OK" if r["balance"] == rb else "BREAK")
        ws.append([
            r["row_id"], r["source_file"], r["source_page"], r["txn_date"],
            r["post_date"], r["description"], r["check_no"], r["reference"],
            r["section"],
            float(r["debit"]) if r["debit"] is not None else None,
            float(r["credit"]) if r["credit"] is not None else None,
            float(r["balance"]) if r["balance"] is not None else None,
            float(rb), tie,
        ])

    last = ws.max_row
    for row in ws.iter_rows(min_row=2, max_row=last):
        for idx in (4, 5):
            row[idx - 1].number_format = DATE_FMT
        for idx in (10, 11, 12, 13):
            row[idx - 1].number_format = MONEY_FMT
        if row[13].value == "BREAK":
            row[13].font, row[13].fill = FAIL_FONT, FAIL_FILL

    # totals strip
    ws.append([])
    t = ws.max_row + 1
    ws.cell(row=t, column=9, value="TOTALS").font = Font(bold=True)
    ws.cell(row=t, column=10, value=f"=SUBTOTAL(109,J2:J{last})").number_format = MONEY_FMT
    ws.cell(row=t, column=11, value=f"=SUBTOTAL(109,K2:K{last})").number_format = MONEY_FMT
    for c in (10, 11):
        ws.cell(row=t, column=c).font = Font(bold=True)

    ws.freeze_panes = "D2"
    ws.auto_filter.ref = f"A1:N{last}"
    autosize(ws, {1: 6, 2: 26, 3: 5, 4: 12, 5: 12, 6: 54, 7: 10, 8: 20,
                  9: 20, 10: 14, 11: 14, 12: 17, 13: 19, 14: 8})


def write_proof(wb, results, meta, overall_pass: bool) -> None:
    ws = wb.create_sheet("Proof")
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 46

    r = 1
    ws.cell(row=r, column=1, value="COMPLETENESS PROOF").font = Font(bold=True, size=14)
    r += 1
    verdict = ws.cell(
        row=r, column=1,
        value="RESULT: PASS - extraction ties to the statement"
        if overall_pass else
        "RESULT: FAILED - DO NOT RELY ON THIS FILE",
    )
    verdict.font = PASS_FONT if overall_pass else FAIL_FONT
    if not overall_pass:
        verdict.fill = FAIL_FILL
    r += 2

    for k, v in meta.items():
        ws.cell(row=r, column=1, value=k).font = Font(bold=True)
        ws.cell(row=r, column=2, value=v)
        r += 1
    r += 1

    a = results["a"]
    ws.cell(row=r, column=1, value=a["name"]).font = Font(bold=True, size=12)
    ws.cell(row=r, column=5, value=a["detail"]).font = Font(italic=True, size=9)
    r += 1
    for label, val in [
        ("Opening balance per statement", meta["Opening balance"]),
        ("Add: total credits / deposits", a["credits"]),
        ("Less: total debits / withdrawals", -a["debits"]),
        ("= Computed closing balance", a["computed_closing"]),
        ("Closing balance per statement", a["stmt_closing"]),
        ("Difference", a["difference"]),
    ]:
        ws.cell(row=r, column=1, value=label)
        cell = ws.cell(row=r, column=2, value=float(val))
        cell.number_format = MONEY_FMT
        if label.startswith(("=", "Difference")):
            cell.font = Font(bold=True)
        if label == "Difference":
            cell.font = PASS_FONT if a["passed"] else FAIL_FONT
            if not a["passed"]:
                cell.fill = FAIL_FILL
        r += 1
    ws.cell(row=r, column=1, value="Test A verdict").font = Font(bold=True)
    vc = ws.cell(row=r, column=2, value="PASS" if a["passed"] else "FAIL")
    vc.font = PASS_FONT if a["passed"] else FAIL_FONT
    r += 2

    b = results.get("b")
    if b:
        ws.cell(row=r, column=1, value=b["name"]).font = Font(bold=True, size=12)
        ws.cell(row=r, column=5, value=b["detail"]).font = Font(italic=True, size=9)
        r += 1
        for h, col in zip(["Item", "Extracted", "Per statement", "Difference"], range(1, 5)):
            hc = ws.cell(row=r, column=col, value=h)
            hc.fill, hc.font = HDR_FILL, HDR_FONT
        r += 1
        for label, extracted, stmt, diff in b["parts"]:
            ws.cell(row=r, column=1, value=label)
            for col, val in ((2, extracted), (3, stmt), (4, diff)):
                c = ws.cell(row=r, column=col, value=float(val))
                c.number_format = MONEY_FMT
            dc = ws.cell(row=r, column=4)
            dc.font = PASS_FONT if diff == ZERO else FAIL_FONT
            if diff != ZERO:
                dc.fill = FAIL_FILL
            r += 1
        ws.cell(row=r, column=1, value="Test B verdict").font = Font(bold=True)
        vc = ws.cell(row=r, column=2, value="PASS" if b["passed"] else "FAIL")
        vc.font = PASS_FONT if b["passed"] else FAIL_FONT
        r += 2
    else:
        ws.cell(row=r, column=1, value="Test B - Statement control totals").font = Font(bold=True, size=12)
        ws.cell(row=r, column=2, value="NOT PERFORMED - statement totals not supplied")
        r += 2

    c = results["c"]
    ws.cell(row=r, column=1, value=c["name"]).font = Font(bold=True, size=12)
    ws.cell(row=r, column=5, value=c["detail"]).font = Font(italic=True, size=9)
    r += 1
    if not c["applicable"]:
        ws.cell(row=r, column=2, value="NOT APPLICABLE - no per-line balance printed")
        r += 1
    elif c["passed"]:
        vc = ws.cell(row=r, column=2, value="PASS - every printed balance agrees")
        vc.font = PASS_FONT
        r += 1
    else:
        ws.cell(row=r, column=1, value=f"{len(c['breaks'])} break(s) - each row below is where "
                                       f"the extract diverges from the statement").font = FAIL_FONT
        r += 1
        heads = ["Row", "Pg", "Date", "Description", "Prior bal", "Amount",
                 "Expected", "Printed", "Difference"]
        for col, h in enumerate(heads, start=1):
            hc = ws.cell(row=r, column=col, value=h)
            hc.fill, hc.font = HDR_FILL, HDR_FONT
        r += 1
        for brk in c["breaks"]:
            ws.cell(row=r, column=1, value=brk["row_id"])
            ws.cell(row=r, column=2, value=brk["page"])
            dc = ws.cell(row=r, column=3, value=brk["date"])
            dc.number_format = DATE_FMT
            ws.cell(row=r, column=4, value=brk["description"])
            for col, key in ((5, "prior"), (6, "amount"), (7, "expected"),
                             (8, "printed"), (9, "difference")):
                mc = ws.cell(row=r, column=col, value=float(brk[key]))
                mc.number_format = MONEY_FMT
            ws.cell(row=r, column=9).font = FAIL_FONT
            r += 1

    r += 1
    ws.cell(row=r, column=1, value="Most likely causes when a test fails, in order:").font = Font(bold=True)
    for line in [
        "1. A subtotal or 'Total Deposits' line captured as a transaction",
        "2. A wrapped description line captured as its own row",
        "3. A transaction on a continuation page missed entirely",
        "4. A debit classified as a credit (or vice versa)",
        "5. An OCR digit misread (0/O, 1/l, 5/S, 8/B) or a dropped leading digit",
        "Test C above identifies the specific row - start there, not with the totals.",
    ]:
        r += 1
        ws.cell(row=r, column=1, value=line)


def write_exceptions(wb, exceptions_csv: Path | None) -> None:
    ws = wb.create_sheet("Exceptions")
    headers = ["#", "Source File", "Pg", "Type", "Description of issue",
               "As presented", "Treatment applied", "Needs preparer decision"]
    ws.append(headers)
    style_header(ws, len(headers))
    n = 0
    if exceptions_csv and exceptions_csv.exists():
        with exceptions_csv.open(newline="", encoding="utf-8-sig") as fh:
            for i, row in enumerate(csv.DictReader(fh), start=1):
                ws.append([i] + [row.get(k, "") for k in
                                 ["source_file", "source_page", "type", "issue",
                                  "as_presented", "treatment", "needs_decision"]])
                n = i
    if n == 0:
        ws.cell(row=2, column=1, value="None")
        ws.cell(row=2, column=5,
                value="No illegible figures, unresolved dates, or ambiguous "
                      "classifications were encountered.")
    autosize(ws, {1: 5, 2: 26, 3: 5, 4: 20, 5: 56, 6: 18, 7: 34, 8: 22})


def write_source_map(wb, rows) -> None:
    ws = wb.create_sheet("Source Map")
    headers = ["Source File", "Pg", "Transactions", "Debits on page",
               "Credits on page", "Net"]
    ws.append(headers)
    style_header(ws, len(headers))

    agg: dict[tuple, dict] = {}
    for r in rows:
        key = (r["source_file"], r["source_page"])
        a = agg.setdefault(key, {"n": 0, "d": ZERO, "c": ZERO})
        a["n"] += 1
        a["d"] += r["debit"] or ZERO
        a["c"] += r["credit"] or ZERO

    for (f, p), a in sorted(agg.items(), key=lambda kv: (kv[0][0], kv[0][1] or 0)):
        ws.append([f, p, a["n"], float(a["d"]), float(a["c"]), float(a["c"] - a["d"])])

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for idx in (4, 5, 6):
            row[idx - 1].number_format = MONEY_FMT
    ws.freeze_panes = "A2"
    autosize(ws, {1: 30, 2: 6, 3: 14, 4: 16, 5: 16, 6: 16})


# --------------------------------------------------------------------------- main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--transactions", required=True)
    ap.add_argument("--opening", required=True)
    ap.add_argument("--closing", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--stmt-deposits")
    ap.add_argument("--stmt-withdrawals")
    ap.add_argument("--exceptions", help="optional exceptions CSV")
    ap.add_argument("--account-label", default="")
    ap.add_argument("--period", default="")
    ap.add_argument("--force", action="store_true",
                    help="write the workbook even if a test fails")
    args = ap.parse_args()

    rows = load_rows(Path(args.transactions))
    opening, closing = dec(args.opening), dec(args.closing)
    if opening is None or closing is None:
        sys.exit("--opening and --closing are required and must be numeric.")

    results = {
        "a": test_a(rows, opening, closing),
        "b": test_b(rows, dec(args.stmt_deposits), dec(args.stmt_withdrawals)),
        "c": test_c(rows, opening),
    }
    overall = results["a"]["passed"] and results["c"]["passed"] and (
        results["b"]["passed"] if results["b"] else True
    )

    meta = {
        "Account": args.account_label or "(not supplied)",
        "Statement period": args.period or "(not supplied)",
        "Source file(s)": ", ".join(sorted({r["source_file"] for r in rows if r["source_file"]})) or "(not supplied)",
        "Transactions extracted": len(rows),
        "Opening balance": opening,
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    # ---- console report
    print("=" * 68)
    print("COMPLETENESS PROOF")
    print("=" * 68)
    a = results["a"]
    print(f"Transactions extracted : {len(rows)}")
    print(f"Opening balance        : {opening:>14,.2f}")
    print(f"Total credits          : {a['credits']:>14,.2f}")
    print(f"Total debits           : {a['debits']:>14,.2f}")
    print(f"Computed closing       : {a['computed_closing']:>14,.2f}")
    print(f"Statement closing      : {a['stmt_closing']:>14,.2f}")
    print(f"Difference             : {a['difference']:>14,.2f}   "
          f"{'PASS' if a['passed'] else '*** FAIL ***'}")
    if results["b"]:
        print(f"\nTest B (control totals): "
              f"{'PASS' if results['b']['passed'] else '*** FAIL ***'}")
        for label, ex, st, df in results["b"]["parts"]:
            print(f"  {label:<28} extracted {ex:>12,.2f}  stmt {st:>12,.2f}  diff {df:>10,.2f}")
    c = results["c"]
    if not c["applicable"]:
        print("\nTest C (line continuity): N/A - no per-line balance printed")
    elif c["passed"]:
        print("\nTest C (line continuity): PASS")
    else:
        print(f"\nTest C (line continuity): *** FAIL - {len(c['breaks'])} break(s) ***")
        for brk in c["breaks"][:15]:
            print(f"  row {brk['row_id']:>4} pg {brk['page']}  {brk['date']}  "
                  f"{brk['description'][:40]:<40} expected {brk['expected']:>12,.2f} "
                  f"printed {brk['printed']:>12,.2f}  diff {brk['difference']:>10,.2f}")
        if len(c["breaks"]) > 15:
            print(f"  ... and {len(c['breaks']) - 15} more")

    if not overall and not args.force:
        print("\n" + "=" * 68)
        print("WORKBOOK NOT WRITTEN. The extract does not tie to the statement.")
        print("Fix the extract and re-run. Do not deliver an unproven file.")
        print("Check, in order: subtotals captured as transactions; wrapped")
        print("description lines; missed continuation pages; debit/credit flipped;")
        print("OCR digit misreads. Test C above names the offending row.")
        print("Use --force only to produce a marked-FAILED file for debugging.")
        print("=" * 68)
        return 1

    wb = Workbook()
    write_transactions(wb, rows, opening)
    write_proof(wb, results, meta, overall)
    write_exceptions(wb, Path(args.exceptions) if args.exceptions else None)
    write_source_map(wb, rows)
    wb.active = 0

    out = Path(args.out)
    if not overall:
        out = out.with_name(out.stem + " [FAILED PROOF]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))

    print("\n" + ("PROOF PASSED. " if overall else "PROOF FAILED (forced). "))
    print(f"Workbook: {out}")
    return 0 if overall else 1


if __name__ == "__main__":
    sys.exit(main())
```

## `scripts/extract_text.py`

```python
#!/usr/bin/env python3
"""
Extract the text layer and word geometry from statement PDFs, fully locally.

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
    if not shutil.which("ocrmypdf"):
        sys.exit(
            "This PDF has no text layer and 'ocrmypdf' is not installed.\n"
            "Install it locally, then re-run with --ocr:\n"
            "  macOS:  brew install ocrmypdf\n"
            "  Debian: sudo apt install ocrmypdf tesseract-ocr\n"
            "Do not upload client bank records to a cloud OCR service."
        )
    out = Path(tempfile.mkdtemp(prefix="ocr_")) / f"{src.stem}.ocr.pdf"
    print(f"  OCR (local): {src.name} -> {out.name}", flush=True)
    subprocess.run(
        [
            "ocrmypdf",
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

