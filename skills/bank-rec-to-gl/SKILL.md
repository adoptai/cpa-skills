---
name: bank-rec-to-gl
description: Reconcile a bank statement to the general ledger cash account — match transactions, identify deposits in transit, outstanding checks, and unrecorded bank items, then produce a signable four-column reconciliation workpaper that proves adjusted bank balance equals adjusted book balance. Use this whenever the user mentions bank reconciliation, reconciling cash, "tie the bank to the GL," unreconciled items, outstanding checks, deposits in transit, an unexplained cash difference, a cash account that won't balance, month-end close of cash, or asks why the bank and the books disagree. Also trigger on "reconcile these," "match the bank to QuickBooks/NetSuite/Sage," "our cash is off by," "clear the old outstanding checks," or when a bank statement and a GL detail export are provided together. Runs fully local — no client cash data leaves the machine.
---

# Bank Reconciliation → General Ledger

A reconciliation is not a matching exercise. It is a **proof that two independently
maintained records of the same cash, which legitimately differ, differ only for reasons you
can name.** Matching is the mechanical part. The reconciliation is the explanation.

The output is a workpaper someone signs. That means every reconciling item has a cause, a
dollar amount, and an expected resolution date — and the unexplained difference line reads
zero. Not "close to zero." Zero.

## The one rule that matters most

**Never plug.** If a difference cannot be explained, it stays visible as an unexplained
difference and the reconciliation is reported as incomplete. A plugged reconciliation is how
misappropriation survives twelve consecutive months of close. The most common form of a plug
is a "reconciling item" with a vague label — `Timing`, `Bank adj`, `Misc difference`,
`Unidentified deposit` — that never resolves in a subsequent month. Treat any such item
inherited from a prior period as an open finding, not as an accepted balance.

## Inputs

You need three things. Ask for whatever is missing before starting.

1. **Bank side** — statement transactions with date, description, amount, and direction.
   If only a PDF statement is available, extract it first (the `bank-statement-to-excel`
   skill does this with a completeness proof; do not reconcile off an unproven extract).
   You also need the statement's opening and closing balance.
2. **Book side** — GL detail for the cash account for the same period: date, memo/payee,
   reference/check number, debit, credit, and the account's opening and closing balance per
   the trial balance.
3. **Prior-period reconciliation**, if one exists — specifically the list of items that were
   outstanding at the prior close. Without it you cannot tell a genuinely new outstanding
   check from one that has been outstanding for nine months.

Confirm before matching: same account, same period, and the GL closing balance agrees to the
trial balance. Reconciling to a GL balance that doesn't agree to the TB wastes the entire
exercise.

## Step 1 — Match

```bash
python3 scripts/reconcile.py \
  --bank bank.csv --book book.csv \
  --opening-bank 42150.22 --closing-bank 38904.71 \
  --opening-book 41980.22 --closing-book 39310.71 \
  --prior-outstanding prior_items.csv \
  --date-tolerance 5 \
  --out "Acme - Cash x4471 - 2026-01 Bank Rec.xlsx"
```

Matching runs as a cascade, most reliable first. Once a transaction is matched it is removed
from the pool, so a weaker rule can never steal a transaction from a stronger one. Every
match records **which rule matched it** — a reviewer needs to know a $4,000 item was matched
on exact amount and date, not on amount alone within a five-day window.

| Pass | Rule | Notes |
|---|---|---|
| 1 | Check number + amount | Strongest evidence available. Run before anything date-based. |
| 2 | Exact amount + exact date | |
| 3 | Exact amount within date tolerance | Default ±5 days for float/posting lag. Ask before widening. |
| 4 | Amount + description similarity | Requires a token overlap threshold, not a loose fuzzy score. |
| 5 | Many-to-one | Several GL receipts deposited as one bank deposit — very common. Subset-sum, capped to keep it honest. |
| 6 | One-to-many | One GL entry settling as several bank items. |

Everything unmatched after pass 6 is a reconciling item and must be classified by hand in
Step 2. **The script does not classify.** Classification is judgment and it is where the
value of a CPA actually sits.

## Step 2 — Classify every unmatched item

This is the part that cannot be automated, and the part a reviewer reads.

**Unmatched on the book side** (in the GL, not at the bank):

- *Deposit in transit* — receipt recorded before period end, deposited/cleared after. Legitimate
  timing. Verify it actually cleared in the subsequent period; note the clearing date. A
  deposit in transit that has not cleared by the time you're reconciling is not timing — it is
  an exception, and it is one of the classic fictitious-revenue and lapping patterns.
- *Outstanding check* — issued and recorded, not yet presented. Legitimate timing. Record the
  issue date, because age is the whole story: see stale-check handling below.
- *Book error* — wrong amount, wrong period, duplicate posting, wrong account. Requires a
  correcting journal entry, not a reconciling item. Draft the JE.

**Unmatched on the bank side** (at the bank, not in the GL):

- *Unrecorded bank item* — service charges, analysis fees, wire fees, interest earned, NSF
  returns and their fees, credit-card processor holdbacks, sweep transfers, direct debits set
  up outside AP. These are real economic events the books simply do not know about yet. They
  require a journal entry, and that JE is a deliverable of this engagement.
- *Bank error* — rare and worth naming loudly when found. Notify the bank; most agreements
  carry a 30–60 day notification deadline, so timing is substantive.
- *Unidentified* — the one you must never bury. An unidentified deposit or, far more
  seriously, an unidentified withdrawal is escalated in the memo regardless of size. Amount
  is not the test for escalation here; the absence of an explanation is.

## Step 3 — Build the reconciliation in proper form

Both sides converge, and the two adjusted balances must be equal:

```
  Balance per bank statement                              38,904.71
  Add:  Deposits in transit                                4,200.00
  Less: Outstanding checks                               ( 3,794.00)
  Add/(less): Bank errors                                      0.00
  = ADJUSTED BANK BALANCE                                 39,310.71

  Balance per general ledger                              39,310.71
  Add:  Unrecorded credits (interest, bank credits)            0.00
  Less: Unrecorded debits (service charges, NSF, fees)         0.00
  Add/(less): Book errors                                      0.00
  = ADJUSTED BOOK BALANCE                                 39,310.71

  UNEXPLAINED DIFFERENCE                                       0.00   <-- must be 0.00
```

The script computes both columns and the difference. If the difference is not zero it exits
non-zero and does not produce a clean workpaper. Note the reconciliation is *to the adjusted
book balance* — the adjusted book balance is what the GL should say once the required journal
entries are posted, which is why those JEs are part of the deliverable and not an
afterthought.

## Step 4 — Age the outstanding items

Aging is where a reconciliation stops being clerical.

- **Over 90 days outstanding** — investigate. Was it delivered? Is the payee waiting? Is it
  lost?
- **Over 6 months** — presumptively stale. Most bank agreements make a check stale-dated at
  six months. Consider void-and-reissue, and record the escheatment question.
- **Unclaimed property** — uncashed checks are reportable to the state after a dormancy
  period (commonly 3–5 years, varying by state and property type, generally the payee's
  state of last known address). Flag candidates. Do not assert the specific dormancy period
  or filing deadline for a given state from memory — name the item as a reportable candidate
  and tell the user to confirm against the applicable state's current rules. Being wrong here
  creates real penalty exposure.
- **Recurring identical reconciling items** — the same amount appearing as a reconciling item
  in consecutive months is almost never timing. It is a systemic posting error or a duplicate
  that has been rolling forward unexamined. Call it out by name.

## Step 5 — Deliver

**Workbook tabs:**

1. **Reconciliation** — the four-column form above. Signature and date block for preparer
   and reviewer. This is the workpaper.
2. **Reconciling Items** — every item: side, date, description, amount, classification,
   age in days, expected resolution, whether a JE is required, and who owns it.
3. **Proposed Journal Entries** — debit/credit ready to post, for the unrecorded bank items
   and book errors. Each with account, amount, and a one-line explanation. This is usually
   the single most-used tab.
4. **Matched Detail** — every matched pair with the rule that matched it and the day
   variance. Lets a reviewer sample and re-perform.
5. **Aging** — outstanding checks and deposits in transit bucketed 0–30 / 31–60 / 61–90 /
   91–180 / 180+, with stale and unclaimed-property candidates flagged.
6. **Exceptions & Memo** — unidentified items, bank errors, suspected duplicates, recurring
   items, control observations, and anything a reviewer must see before signing.

**Then write a short memo in chat**: period and account, bank and book closing balances,
count and total of reconciling items by classification, the required journal entries, the
unexplained difference (zero, or the reconciliation is incomplete), and any escalation. Lead
with escalations if there are any. Three sentences before the detail.

## Control observations worth raising unprompted

You will see these while reconciling, and they are worth more to the client than the
reconciliation itself:

- Checks clearing out of sequence, or gaps in the check sequence
- Round-dollar disbursements to unfamiliar payees
- Transfers between company accounts that appear on only one side (kiting pattern)
- Bank items with no corresponding AP record — payments set up outside the AP process
- The same person preparing the reconciliation and having disbursement authority
- Reconciliations prepared more than 30 days after period end, or unsigned by a reviewer

Raise them in the memo as observations, framed as control matters for management, not as
audit conclusions.

## Security posture

Fully local: `pandas`/`openpyxl` only. No network calls, no uploads, no telemetry. Inputs are
read-only. Intermediate files stay in the working directory and can be shredded after
delivery.

## Dependencies

```bash
pip install pandas openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
