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
