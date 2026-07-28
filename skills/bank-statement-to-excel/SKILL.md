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
