---
name: ar-aging-tie-out
description: Tie the accounts receivable aging to the general ledger control account and test it properly — recomputing every ageing bucket from the invoice date rather than trusting the report, grossing up netted credit balances, and flagging stale balances, unapplied cash, and cutoff errors. Use this whenever the user mentions AR aging, accounts receivable aging, tying the aging to the GL, an AR subledger that does not agree to the control account, stale or uncollectible receivables, the allowance for doubtful accounts, unapplied cash or credits, or receivable confirmations. Also trigger on "does our AR aging tie," "reconcile accounts receivable," "which receivables are old," "check the allowance," "why is AR off," "review the aging," or when an AR aging report and a trial balance are provided together. Runs fully local — no customer or receivable data leaves the machine.
---

# AR Aging Tie-Out

Almost everyone ties the aging total to the general ledger and stops. That test is necessary and
nearly worthless on its own, because **the total can agree while the aging itself is wrong** — and
the aging is what drives the allowance, the collectibility discussion, and the confirmation
sample.

Four things go wrong beneath a correct total:

1. **Mis-bucketed invoices.** Aging reports commonly bucket from the wrong date, or fail to
   re-age after a partial payment or a credit application. So this skill **recomputes every
   bucket from the invoice date and compares** rather than trusting the column the report printed.
2. **Netted credit balances.** A customer with a $90,000 open invoice and a $85,000 unapplied
   credit shows as $5,000 current. Both numbers matter and the netting hides them.
3. **Unapplied cash sitting in the aging** as a negative current balance, which understates real
   exposure and often means cash was received against an invoice nobody closed.
4. **Invoices dated after period end** included in the aging — a cutoff error that overstates both
   revenue and receivables.

## The gate

No clean workpaper unless:

1. **The aging total equals the GL control account balance.** You supply the GL figure.
2. **The buckets foot.** Bucket amounts sum to each customer's total, and customer totals sum to
   the report total.
3. **Every bucket agrees to a recomputation from the invoice date** using the stated bucket
   boundaries and as-of date. Any invoice in the wrong bucket is reported individually.
4. **Debits and credits are shown gross**, not netted, at customer level.

## Inputs

1. **AR aging detail at invoice level** — customer, invoice number, invoice date, due date, original
   amount, open balance, and the bucket the report assigned. Invoice-level detail is essential; a
   customer-level summary cannot be re-aged.
2. **GL control account balance** for accounts receivable at the same date.
3. **As-of date and bucket boundaries** — the client's own, whatever they are. Do not assume
   30/60/90.
4. **Payment terms**, by customer if they vary. Aging past terms is the meaningful measure, not
   aging past an arbitrary 30 days.
5. **Allowance for doubtful accounts** balance and how it was computed, if testing the allowance.
6. **Subsequent cash receipts**, if available — the strongest evidence of collectibility and the
   best test of existence short of a confirmation.

## Step 1 — Tie and re-age

```bash
python3 scripts/ar_aging.py \
  --aging aging_detail.csv \
  --gl-balance 2841655.40 \
  --as-of 2025-12-31 \
  --buckets 30,60,90,120 \
  --allowance 148000 \
  --subsequent-receipts receipts.csv \
  --client "Fairhaven Distributors" \
  --out "Fairhaven - 2025 AR Aging Tie-Out.xlsx"
```

Seven tests:

- **Test 1 — Aging total to GL control account.** The one everybody runs.
- **Test 2 — Bucket footing.** Buckets sum to customer totals, customers sum to the report total.
- **Test 3 — Bucket recomputation.** Every invoice re-aged from its invoice date against the
  stated boundaries. Mis-bucketed invoices are listed individually with both buckets shown. This
  is the test that finds real problems.
- **Test 4 — Credit balances grossed up.** Total debits and total credits shown separately, with
  customers whose net balance masks a material credit identified.
- **Test 5 — Cutoff.** Invoices dated after the as-of date, and invoices with a due date before
  their invoice date.
- **Test 6 — Internal integrity.** Open balance exceeding the original amount, zero-balance rows
  carried in the aging, duplicate invoice numbers, missing dates, and invoices with no customer.
- **Test 7 — Allowance reasonableness**, where an allowance balance is supplied. Reports the
  allowance as a percentage of total AR and of each aged bucket, and the implied coverage of
  balances past terms.

## Step 2 — Read the collectibility picture

The workbook aggregates by customer and by age band, but the useful views are these:

- **Aging past terms, not past 30 days.** A customer on 60-day terms at 45 days is current. A
  customer on 10-day terms at 45 days is a problem. Where terms are supplied the schedule shows
  days past terms alongside days past invoice, and they tell different stories.
- **Concentration.** The largest few customers as a percentage of total AR. A receivable book
  where one customer is 40% of the balance carries a different risk than the aging profile alone
  suggests, and it belongs in the allowance discussion.
- **Balances with no subsequent receipt.** Where subsequent receipts are supplied, the schedule
  flags aged balances with no payment activity after period end. An old balance that has since
  been paid is a collection-timing issue; an old balance with no subsequent activity is a
  valuation issue, and possibly an existence one.
- **Round-number balances** on aged accounts, which sometimes indicate an estimated or
  placeholder invoice rather than a real one.
- **Credit balances aged into old buckets** — an unapplied credit sitting for months usually means
  a customer overpaid or was double-billed, and it may be a refund liability rather than a
  receivable offset.

## Step 3 — Deliver

**Workbook tabs:**

1. **Tie-Out Summary** — the seven tests, the GL agreement, aging profile, concentration, and the
   allowance analysis. The signable page.
2. **Re-aged Detail** — every invoice with the bucket per the report, the recomputed bucket, days
   outstanding, days past terms, and any flag. Mis-bucketed invoices sort to the top.
3. **By Customer** — gross debits, gross credits, net balance, bucket distribution, days past
   terms on the oldest item, and share of total AR.
4. **Exceptions** — every finding from Tests 3 through 6 with the invoice and dollar effect.
5. **Credit Balances** — all credit balances with age, separated from the debit aging entirely.
6. **Allowance Analysis** — allowance as a percentage of total and of each bucket, coverage of
   balances past terms, and the customers driving the exposure.
7. **Subsequent Receipts** — aged balances with and without post-period collection activity.

**Then, in chat:** whether the aging ties to the GL, how many invoices were mis-bucketed and the
dollars involved, the aging profile past terms, concentration, and the allowance conclusion. Lead
with the mis-bucketing if there is any — it means the aging the client has been relying on for
credit decisions and the allowance is wrong.

## What to escalate

- **The aging not tying to the control account.** Until it does, nothing built on the aging is
  reliable. The usual causes are unposted cash, a journal entry made directly to the control
  account, and a subledger cut on a different date.
- **Journal entries posted directly to the AR control account.** These bypass the subledger by
  definition, so they are invisible in the aging and are a common route for both error and
  concealment. Ask for the control account detail and read every manual entry.
- **Material credit balances aged beyond a few months** — potentially a refund liability, and
  possibly unclaimed property depending on the jurisdiction and dormancy rules. Flag the question;
  confirm the applicable rules rather than assuming them.
- **Aged balances with no subsequent receipts and no collection activity**, particularly where the
  customer continues to be invoiced.
- **Receivables from related parties or employees** mixed into trade AR — a classification and
  disclosure matter, and they are rarely subject to the same credit discipline.
- **An allowance that has not moved** while the aging profile deteriorated. A static allowance
  against a worsening book is a judgment that needs documenting, not a rollforward.
- **Invoices dated after period end** in the aging, which overstates revenue and receivables in
  the same stroke.
- **Write-offs late in the period** clearing balances that were about to become conspicuously
  aged. Report the pattern as a fact.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.
Customer names and balances never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
