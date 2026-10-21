---
name: cash-application
description: Match bank credits and remittance advices to open AR invoices, apply configurable matching rules (remittance detail, embedded invoice numbers, exact and lump-sum amount matching), and flag unapplied cash, short-pays, and overpayments in an exception queue for human review rather than guessing which invoice they belong to. Use whenever the user mentions cash application, applying cash, matching deposits to invoices, lockbox processing, unapplied cash, cash receipts, deposit-to-invoice matching, remittance matching, short-pays, or "why is this payment sitting unapplied." Also trigger on "apply this week's deposits," "match the lockbox file to AR," "clear the unapplied cash bucket," "reconcile customer payments to invoices," or when a bank statement/feed export and an open AR invoice list are given together. For the AR subledger-to-GL tie use balance-sheet-recon-pack; for raw PDF-to-spreadsheet statement extraction use bank-statement-to-excel. Runs fully local - no bank or customer data leaves the machine.
---

# Cash Application

Applying cash is the act of saying which invoice a payment settles. The mechanical part -
matching a deposit amount to an open invoice balance, or to several that sum to it - is
something a script should do, every time, the same way. The part that is not mechanical is
deciding why a payment is $42 short, or who a payment with no recognizable name belongs to.
That distinction is the whole design of this skill: apply what ties exactly, and refuse to
guess at anything else.

## The gate

A deposit is `APPLIED` or `RESOLVED`, and the batch is not `COMPLETE`, unless:

1. **Every deposit dollar is accounted for.** Applied + resolved (classified, even if not tied
   to an invoice) + exception must equal the deposit amount, for every deposit. Nothing falls
   out of the middle of a matching routine silently.
2. **Nothing is force-applied to the nearest invoice.** A deposit that doesn't tie exactly - or
   in an exact combination - to open invoice balances for an identified customer goes to the
   Exception Queue with candidates shown, never auto-applied to "the closest one."
3. **No invoice is applied past its open balance.** An over-application is a data error and is
   reported as one, not silently capped.
4. **Every exception needs a confirmed classification before it clears.** `unidentified_payer`
   is the one classification that never clears the gate - naming a payer "unidentified" states
   the finding, it does not resolve it. See `AGENTS.md` rule 4.
5. **If a control total is supplied, the deposits population must tie to it.** A filtered bank
   export (e.g., one lockbox out of three) will manufacture false unapplied cash and bury the
   real exceptions.

## Scope: what this skill is not

- It does not decide whether a short-pay is an authorized discount, a disputed deduction, or an
  error, or whether an overpayment becomes a credit balance or a refund. The script proposes a
  hypothesis in the Exception Queue's suggested-cause column; the preparer confirms the real
  disposition and feeds it back with `--manual-applications`.
- It does not identify a payer from a fuzzy name guess. Customer identity comes from an explicit
  remittance, an exact (normalized) match to the AR customer name, an `aliases.csv` entry, or an
  invoice number embedded in the bank reference - never from "this name looks close enough." Two
  customers sharing a display name are never resolved from the name alone; they require an
  `aliases.csv` entry.
- It does not OCR or parse remittance PDFs/emails itself; see Step 1. It applies cash once the
  remittance detail is in a CSV.
- It does not post cash receipts to the GL or close the AR subledger. The output is the
  application file an AR clerk or controller reviews and posts from.

## Inputs

1. **Bank credits** - a bank statement or feed export limited to deposits/credits for the
   period: `deposit_id, date, amount, reference, payer_name, description`. If you only have a
   PDF statement, run it through `bank-statement-to-excel` first and filter to credits.
2. **Open AR invoices** - the subledger as of the as-of date: `customer_id, customer_name,
   invoice_number, invoice_date, due_date, original_amount, open_balance`.
3. **Remittance detail**, optional but the single highest-value input - which invoices a
   payment covers, transcribed from remittance advices or emails (Step 1):
   `deposit_id, deposit_reference, customer_id, invoice_number, amount_applied, deduction_code,
   deduction_reason`.
4. **Customer aliases**, optional - how a payer's name appears on the bank side when it differs
   from the AR customer master (common with lockbox remitters): `payer_alias, customer_id`.
5. **A control total**, strongly recommended - the period's total bank credits per the
   statement or feed, to prove the deposits file is the full population.
6. **Manual applications**, fed back after reviewing the Exception Queue (Step 3).

Confirm before running: the invoice list is open balances **as of the same date** the deposits
cover (a payment for an invoice already cleared in a prior batch will otherwise look unmatched),
and the deposits file is credits only, not the full statement.

## Step 1 - Get remittance detail into a CSV

Remittance advices and emails are not machine-uniform, so this is a judgment step, not a script.

For scanned or text-layer remittance PDFs, reuse the extractor from `bank-statement-to-excel`:

```bash
python3 ../bank-statement-to-excel/scripts/extract_text.py --input remittances/ --out work/
```

For email remittances, read the `.eml`/`.txt` body directly. Either way, build
`remittances.csv` by hand from what you read: one row per invoice line, with the amount applied
and any deduction code/reason stated on the advice. **Never infer an invoice number or amount
that isn't printed on the remittance** - a line you can't read becomes a row with the invoice
number blank and a note, which the matching script will surface as an exception rather than
silently drop.

If no remittance was received for a payment, skip this step for that deposit - it is not an
error, just a payment the matcher has to work out from the bank reference and amount alone.

## Step 2 - Match and apply

```bash
python3 scripts/cash_apply.py \
  --deposits deposits.csv --invoices invoices.csv \
  --remittances remittances.csv --aliases aliases.csv \
  --control-total 184250.00 --as-of 2026-10-20 \
  --client "Acme Inc" --period "2026-10" \
  --out "Acme Inc - 2026-10 Cash Application.xlsx"
```

For each deposit, the script runs this cascade, strongest evidence first, and an invoice leaves
the pool once fully consumed so a weaker rule can never steal it from a stronger one:

| Pass | Rule | Notes |
|---|---|---|
| 1 | Remittance detail | Applies exactly what the advice says; any leftover after remittance lines carries forward to the next passes. |
| 2 | Invoice number embedded in bank reference/description | Only when the number uniquely identifies one open invoice for the resolved customer and the amount ties exactly. |
| 3 | Exact amount match | Deposit (or leftover) equals exactly one open invoice's remaining balance. More than one tie is ambiguous and is not guessed. |
| 4 | Lump sum (subset match) | Several open invoices for the same customer sum exactly to the deposit - common with batched customer payment runs. |

Customer identity for passes 2-4 comes only from remittance detail, `aliases.csv`, or a unique
invoice-number match in the reference text - never from approximate name matching.

Anything that doesn't resolve through this cascade becomes a row in the Exception Queue with
candidate invoices and a suggested (not confirmed) cause.

## Step 3 - Classify every exception

Open `cash_app_exceptions.csv` (same rows as the workbook's Exception Queue tab) and classify
each one:

- **`short_pay_discount`** - an authorized early-payment or trade discount; legitimate and
  expected to recur with this customer.
- **`short_pay_deduction`** - a deduction for damages, returns, or a chargeback. Usually needs
  a credit memo.
- **`short_pay_dispute`** - the customer disputes the invoice and the shortfall is unresolved.
  Needs follow-up, not a write-off.
- **`overpayment_credit_balance`** - hold the excess as a credit on the customer's account.
- **`overpayment_refund_due`** - the customer is owed a refund.
- **`unidentified_payer`** - the one you must never bury. Escalate regardless of amount; stating
  it is not resolving it.
- **`data_error`** - wrong invoice, wrong amount, or a duplicate remittance line.
- **`misc_resolved`** - resolved by preparer judgment not covered above; state why in notes.

Feed the completed file back in:

```bash
python3 scripts/cash_apply.py ... --manual-applications cash_app_manual.csv --out "..."
```

`manual-applications.csv` can both apply cash to a specific invoice (fill `invoice_number` and
`amount_applied`) and/or record a classification for the deposit's remaining leftover - a pure
refund or an unidentified payer has no invoice to apply to, only a classification.

## Step 4 - Read the batch

- **A deposit with no remittance and no reference match is not unusual on day one** of a cash
  application cycle - it becomes unusual if the same payer shows up unmatched two periods
  running, which means the alias table or the customer's remittance habits need attention.
- **A short-pay pattern concentrated on one customer** is almost always a standing discount or
  deduction arrangement the alias/classification table hasn't caught up to yet - worth asking
  about once, then encoding so it stops generating exceptions.
- **An `unidentified_payer` item aging past the period** belongs on the month-end close
  checklist, not carried silently - unapplied cash overstates true AR collectability if the
  underlying invoice is actually settled.

## Step 5 - Deliver

**Workbook tabs:**

1. **Application Summary** - totals (applied, resolved, exception), the applied breakdown by
   matching rule, and the gate status. The page a controller reads first.
2. **Population Proof** - every deposit, its applied/resolved/exception split, status, and the
   tie to the control total. The gate, shown.
3. **Applied Detail** - every invoice-level application with the rule that produced it, for
   re-performance.
4. **Exception Queue** - every deposit not fully applied, with candidate invoices, suggested
   cause, confirmed classification, and notes. The working document.
5. **Remaining Open AR** - every invoice with its balance before and after this run, so the next
   cycle starts from a clean position.
6. **Exceptions & Memo** - completeness flags (orphaned remittance lines, control-total tie,
   over-applications) and the never-plug reminder.

**Then, in chat:** total deposits and whether they tie to the control total, applied vs.
resolved vs. exception dollars, the exception count, and anything that needs the AR team to
chase before next period (lead with any `unidentified_payer` amount).

## What to escalate rather than classify and move on

- **Any `unidentified_payer` item**, regardless of size.
- **A short-pay pattern that recurs for one customer without a matching discount or deduction
  arrangement on file** - it is either an undocumented policy or a customer taking an
  unauthorized deduction.
- **A remittance that cannot be linked to any deposit** - a payment advice with no matching
  bank credit is either a timing difference worth tracking or a payment that never arrived.
- **A deposit whose payer name or reference changes from a customer's established pattern**
  shortly before a large payment - worth a second look before applying.

## Security posture

Fully local: standard library plus `openpyxl` (and `pdfplumber`, only if reusing the PDF
extractor from `bank-statement-to-excel` in Step 1). No network calls, no uploads, no telemetry.
Bank account details and customer data never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
