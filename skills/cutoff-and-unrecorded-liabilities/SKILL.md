---
name: cutoff-and-unrecorded-liabilities
description: Test whether transactions are recorded in the right period and search for liabilities that were never recorded at all — examining subsequent-period payments, invoices received after period end, open purchase orders with goods received, unprocessed invoices, credit memos, and accrual completeness, with the search population proven and every item classified as belonging to the period or not. Use this whenever the user mentions cutoff testing, search for unrecorded liabilities, completeness of accounts payable, unrecorded expenses, subsequent disbursements review, accrual completeness, transactions in the wrong period, or period-end cutoff errors. Also trigger on "search for unrecorded liabilities," "test cutoff," "did we accrue everything," "check subsequent payments," "are expenses in the right period," "AP completeness," or when a subsequent-period cash disbursement listing is provided. Runs fully local — no client data leaves the machine.
---

# Cutoff and Search for Unrecorded Liabilities

Every other test in an audit works on what was recorded. This one works on **what was not**, which
makes it structurally different and is why it is the hardest procedure to fake. You cannot sample
from a population of things that do not exist. You have to find them somewhere else and then ask
whether they should have been recorded.

The three places they hide:

1. **Paid after period end for goods or services received before it.** The subsequent-period cash
   disbursement listing is the single best source of unrecorded liabilities, and it is the first
   thing to request.
2. **Received but never entered.** Invoices sitting in a drawer, an inbox, or an approval queue at
   period end. Goods received with no invoice yet (the GRNI population).
3. **Known but not accrued.** Legal fees, bonuses, commissions, utilities, professional fees,
   vacation, warranty, rebates. All predictable, all frequently missed.

Understatement of liabilities is also the direction management is motivated toward, which is why
this procedure carries weight disproportionate to its cost.

## The gate

No clean workpaper unless:

1. **The search population is proven.** Subsequent disbursements must tie to a stated total from
   the cash account for the search period — a partial listing produces false comfort in the
   direction that matters.
2. **Every item in the search population is classified** — belongs to the period under audit, or
   does not, with a reason. Nothing may be left unexamined, and unclassified items are reported.
3. **Every "belongs to the period" item is traced** to either a recorded liability or a proposed
   adjustment. Identifying an unrecorded liability and not quantifying it is not a test.
4. **The accrual completeness checklist is answered**, item by item, including the ones answered
   "not applicable" — with a reason.

## Inputs

1. **Subsequent-period cash disbursements** — every payment for a defined window after period end
   (commonly the period between year end and fieldwork), with date, payee, amount, and the invoice
   or reference it paid. Plus **the total disbursements for that window from the cash account**, so
   the population can be proven.
2. **Recorded accounts payable and accrual detail** at period end.
3. **Unprocessed invoice file** — invoices received but not entered at period end, if the client
   maintains one.
4. **Open purchase orders with goods received** — the GRNI population. The `three-way-match` skill
   produces this.
5. **Subsequent-period invoice register**, if available.
6. **Subsequent-period credit memos** — these test the other direction, revenue and receivable
   overstatement.
7. **Period end date and the search window end date.**

## Step 1 — Prove the search population, then classify

```bash
python3 scripts/cutoff_test.py \
  --disbursements subsequent_payments.csv \
  --disbursement-total 4182005.44 \
  --recorded-ap recorded_ap.csv \
  --accruals accruals.csv \
  --grni grni.csv \
  --credit-memos credit_memos.csv \
  --period-end 2025-12-31 --search-end 2026-02-28 \
  --materiality 25000 \
  --client "Ardenway Manufacturing" \
  --out "Ardenway - FY2025 Unrecorded Liabilities Search.xlsx"
```

Seven tests:

- **Test 1 — Search population proven.** Disbursements listed tie to the total from the cash
  account for the window.
- **Test 2 — Every disbursement classified.** Each payment is assigned a service or delivery date
  and classified: pre-period-end (should have been a liability), post-period-end (correctly
  excluded), or unclassified. **Unclassified is a finding, not a default.**
- **Test 3 — Pre-period-end payments traced.** Each payment for goods or services received before
  period end must appear in recorded AP or accruals, or become a proposed adjustment.
- **Test 4 — GRNI recorded.** Goods received before period end with no invoice must be accrued.
- **Test 5 — Unprocessed invoices.** Invoices dated on or before period end, received before the
  books closed, and not recorded.
- **Test 6 — Accrual completeness checklist.** Each recurring accrual category answered.
- **Test 7 — Credit memos after period end** reversing period revenue, which tests the receivable
  and revenue direction.

## Step 2 — Get the date right

The classification hinges on **when the service was performed or the goods transferred**, not on
the invoice date and not on the payment date. This is where the test is usually done wrong.

- An invoice dated 15 January for December consulting is a **December liability**. The invoice
  date is irrelevant.
- An invoice dated 20 December for a January insurance renewal is a **January expense**, and at
  period end it is a prepayment, not a liability.
- A payment on 10 February for goods received 28 December is a **December liability**.
- Rent paid in advance on 28 December for January is a **prepayment**.

The script requires a `service_date` on each disbursement and treats a missing one as
unclassified rather than guessing from the invoice or payment date. That is deliberate: guessing
here produces a wrong answer that looks like a right one.

## Step 3 — Work the accrual checklist

Recurring accruals that are missed most often, each of which the script requires an answer for:

Legal and professional fees · audit and tax fees for the period being audited · bonuses and
incentive compensation · commissions · payroll and payroll taxes for the stub period · vacation
and paid time off · utilities · rent and common area charges · property and other non-income taxes
· interest · insurance · warranty and product returns · customer rebates and volume discounts ·
freight in transit · repairs performed but not invoiced · severance and restructuring · self-insured
health claims incurred but not reported

"Not applicable" is a valid answer with a reason. A blank is not.

Audit and tax fees for the period under audit deserve their own mention: they are incurred, the
amount is known to the firm performing the work, and they are omitted with remarkable frequency.

## Step 4 — Deliver

**Workbook tabs:**

1. **Summary** — population proof, classification counts, total unrecorded liabilities identified,
   proposed adjustments, and the accrual checklist status. The signable page.
2. **Population Proof** — disbursements listed against the cash account total, and the
   classification balance: pre + post + unclassified = population.
3. **Subsequent Disbursements** — every payment with service date, classification, whether traced
   to recorded AP or an accrual, and the proposed adjustment where not.
4. **Proposed Adjustments** — debit and credit ready to post, with the supporting reference,
   individually and in total, compared to materiality.
5. **GRNI and Unprocessed** — goods received not invoiced and invoices not entered, with the
   accrual computed.
6. **Accrual Checklist** — each category, answer, amount, and basis.
7. **Credit Memos** — subsequent credits against period revenue.

**Then, in chat:** whether the population ties, how many items were unclassified, the total
unrecorded liability identified, and the proposed adjustment against materiality. Lead with the
total identified — that is the number that changes the financial statements.

## What to escalate

- **A subsequent disbursement listing that does not tie.** In this test specifically, an
  incomplete population biases the result toward finding nothing, which is the direction that
  matters.
- **Items that cannot be classified** because no service date is available. Report the count and
  value; they are a scope limitation.
- **A pattern of invoices held and entered after period end**, particularly if concentrated in the
  first days of the new period.
- **Accruals that were reversed at period end and not re-established.**
- **A recurring accrual present in the prior year and absent this year** with no explanation.
- **Round-number accruals** unchanged from the prior year, which usually means nobody recomputed
  them.
- **Audit or tax fees for the period under audit not accrued.**
- **Disputed invoices excluded from AP** on the basis of the dispute. A disputed liability is
  usually still a liability, and the dispute affects measurement rather than existence.
- **Debit balances in accounts payable**, which often means a payment was recorded against nothing
  or a credit was misapplied.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
