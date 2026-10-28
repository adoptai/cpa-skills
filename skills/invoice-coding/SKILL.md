---
name: invoice-coding
description: Proposes GL account coding for vendor invoices from historical coding frequency and the chart of accounts, and routes every invoice that cannot be coded safely — new vendors, decommissioned accounts, low-confidence matches — to an exception queue instead of guessing. Use this whenever the user mentions invoice coding, GL coding of invoices, coding AP invoices, auto-coding vendor invoices, proposing account codes, invoice extraction and coding, or building a coded invoice register. Also trigger on "code these invoices," "propose GL coding," "what account does this vendor usually hit," "route invoices for approval," or when an invoice extract, chart of accounts, and historical coding file are provided together. This is the AP-side twin of receipts-to-expense (which handles employee receipts, not vendor invoices) and feeds three-way-match downstream once invoices are coded. Runs fully local — no vendor or invoice data leaves the machine.
---

# Invoice Extraction and GL Coding

This skill assumes invoice text has already been extracted into structured rows — it proposes
coding from that structured data, it does not perform PDF or inbox parsing itself. Its job is
the coding decision: which account, on what basis, and at what confidence — never a guess
dressed up as a model score.

## The gate

A proposal is not "coded" until it is earned by history and a usable account. The script
refuses to produce a clean workbook when:

1. **An invoice is accounted for in neither bucket.** Every invoice lands in exactly one of
   coded-with-confidence or exception — a reconciliation check enforces this and blocks the
   workbook if it ever fails.
2. **An amount or date does not parse.** Reported as a named exception on the row it came from,
   never dropped or defaulted.
3. **A duplicate `invoice_id`** makes the population ambiguous and is rejected outright.

A vendor with no historical coding, a proposed account that is missing or inactive in the chart
of accounts, and anything below `--min-confidence` are not gate failures — they are the
exception queue working as intended, and the workbook is written with them clearly flagged.

## Inputs

1. **Invoice extract** — `invoice_id, vendor, amount, invoice_date, line_description`.
2. **Chart of accounts** — `account_code, account_name, active` (`active`: y/n).
3. **Historical coding** — `vendor, account_code, invoice_count`, a frequency table of how often
   each vendor was coded to each account historically.

## Step 1 — Run the proposal

```bash
python3 scripts/invoice_coding.py --invoices invoices.csv --coa chart_of_accounts.csv \
  --historical historical_coding.csv --min-confidence 60 \
  --client "Northwind Systems Inc" \
  --out "Northwind - Invoice Coding Register.xlsx"
```

Omit `--min-confidence` only if you deliberately want no confidence floor — every
historically-coded vendor with an active account then gets auto-coded, and only brand-new
vendors or inactive/missing accounts land in the exception queue. Stating a floor explicitly is
the safer default for anything going straight to posting.

## Step 2 — How confidence is computed

Confidence is a real, auditable statistic, not an invented score: it is the share of a vendor's
historical invoice count that went to the account being proposed. A vendor coded to one account
100% of the time in history gets 100% confidence; a vendor split 60/40 between two accounts gets
60% confidence on the majority account and the minority 40% is simply the alternative not taken.
The **basis** column on the register always shows the underlying count (`8/10 historical
invoices to 6100`), so a reviewer can second-guess the proposal without re-deriving it.

## Deliver

**Workbook tabs:**

1. **Summary** — population counts, coded vs. exception split, gate status.
2. **Coded Invoice Register** — every auto-coded invoice with proposed account, confidence, and
   basis.
3. **Exception Queue** — new vendors, inactive/missing accounts, below-floor confidence, and any
   parse failures, each with a stated reason.
4. **Vendor Coding History** — the frequency table as supplied, for reference during review.

**Then, in chat:** the coded/exception split, the confidence floor applied, and anything that
blocked the gate. Name the largest-dollar exceptions first — that's where review time belongs.

## What to escalate

- **A vendor whose historical coding has shifted materially** (e.g., recently split differently
  than history suggests) — the script cannot detect intent behind a shift, only report the
  historical pattern.
- **An invoice coded to an expense account that is unusually large** relative to that vendor's
  historical invoice amounts — a dollar-size outlier even at high coding confidence.
- **A chart of accounts with recently deactivated accounts still carrying heavy vendor history**
  — may signal a chart cleanup that hasn't been communicated to AP.
- **A new vendor whose name is a close variant of an existing vendor** — possible duplicate
  vendor master record, not a coding question but worth flagging while the exception is in
  front of a reviewer.

Report each as a fact with the record reference. Do not characterize intent — describe what the
data shows and let the reviewer decide.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.
Vendor and invoice data never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
