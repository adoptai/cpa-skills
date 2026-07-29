---
name: bank-rec-review
description: Audit a bank reconciliation that someone else prepared — independently recompute it, verify every reconciling item against the bank statement and the subsequent-period statement, confirm each item actually cleared, and detect plugs, stale items rolled forward, and reconciling items that exist only on the reconciliation. Use this whenever the user mentions reviewing or testing a bank reconciliation, auditing a client's reconciliation, verifying reconciling items, whether a reconciliation is supported, an outstanding check that never cleared, a suspicious reconciling item, or cash audit procedures. Also trigger on "review the client's bank rec," "test this reconciliation," "did these items clear," "is this reconciliation supported," "audit cash," or when a prepared reconciliation is provided alongside a bank statement. Distinct from preparing a reconciliation — for that, see bank-rec-to-gl. Runs fully local — no client cash data leaves the machine.
---

# Bank Reconciliation Review

Preparing a reconciliation and reviewing one are different procedures. The preparer starts from
two records and works toward agreement. **The reviewer starts from a document that already claims
agreement, and has to establish whether that claim is true.** Those are not the same task, and the
review is the easier one to do badly, because a reconciliation that foots looks finished.

A reconciliation can foot perfectly and still be worthless. The three ways:

1. **A fabricated reconciling item.** An "outstanding check" or "deposit in transit" that exists
   only on the reconciliation, inserted to make the two sides agree. It has no counterpart in any
   bank statement, before or after.
2. **A stale item rolled forward.** A genuine item from years ago that has never cleared and never
   been investigated, carried each month because removing it would break the reconciliation.
3. **A plug with a plausible label.** `Timing`, `Bank adjustment`, `Unidentified deposit`,
   `Items in transit`. The label is doing the work, not the item.

All three survive a review that only checks the arithmetic. **The test that catches all three is
subsequent clearance**: a real reconciling item appears in the next period's bank statement. A
fabricated one never does.

## The gate

No clean workpaper unless:

1. **The reconciliation is independently recomputed** and the recomputation agrees to what the
   preparer presented. Do not accept the preparer's subtotals.
2. **The bank balance is agreed to the bank statement** and the book balance to the general
   ledger — as two independent sources. Agreeing the reconciliation to itself proves nothing.
3. **Every reconciling item is supported** — traced to the bank statement, the subsequent-period
   statement, or a document. An item with no support is reported, never accepted.
4. **Subsequent clearance is tested** where a subsequent statement is available, and every item
   that did not clear is reported individually.

## Inputs

1. **The reconciliation as prepared**, item by item — not just the totals.
2. **The bank statement** for the period, with its closing balance.
3. **The general ledger** cash balance at the same date, ideally with the account detail.
4. **The subsequent-period bank statement.** This is the most important document in the review and
   it is frequently not requested. Without it, clearance cannot be tested and the review is
   limited in scope.
5. **Prior-period reconciliations**, to detect items rolled forward.
6. **Who prepared and who reviewed it, and when.** A reconciliation prepared two months after
   period end by the person with disbursement authority, unreviewed, is a control finding
   independent of its arithmetic.

## Step 1 — Recompute and test

```bash
python3 scripts/rec_review.py \
  --reconciliation client_rec.csv \
  --bank-balance 1284551.09 \
  --gl-balance 1291204.77 \
  --subsequent-statement feb_statement.csv \
  --prior-reconciliation prior_rec.csv \
  --period-end 2025-12-31 \
  --prepared-by "J. Mensah" --prepared-date 2026-02-18 \
  --reviewed-by "" \
  --client "Fairmont Logistics" \
  --out "Fairmont - 2025 Bank Rec Review.xlsx"
```

Eight tests:

- **Test 1 — Independent recomputation.** Bank balance ± reconciling items = adjusted bank; GL
  balance ± unrecorded items = adjusted book; the two must be equal. Computed from the item detail,
  not from the preparer's subtotals.
- **Test 2 — Bank balance agreed to the statement.** An independent source.
- **Test 3 — Book balance agreed to the general ledger.** The other independent source.
- **Test 4 — Every item supported.** Each reconciling item traced to a document.
- **Test 5 — Subsequent clearance.** Each item matched to the subsequent-period statement.
  Anything that did not clear is listed.
- **Test 6 — Stale items.** Items appearing on the prior reconciliation with the same amount,
  aged, with the number of periods they have been carried.
- **Test 7 — Plug detection.** Items whose description is generic, whose amount is suspiciously
  round, or that appear only on the reconciliation with no counterpart anywhere.
- **Test 8 — Preparation and review controls.** Timeliness, whether it was reviewed, and whether
  the preparer also has disbursement authority.

## Step 2 — Read subsequent clearance carefully

Clearance is the strongest evidence available here, but read it correctly:

- **An outstanding check that cleared in the subsequent period** for the same amount: supported.
  Good.
- **Cleared for a different amount**: the reconciliation was wrong, or the check was altered.
  Worth pursuing either way.
- **Did not clear at all** and is more than a couple of months old: investigate. Was it ever
  issued? Was it delivered?
- **A deposit in transit that did not clear within days**: a deposit in transit is by definition a
  deposit already made. One that never arrives at the bank is not timing — it is an exception, and
  it is the classic lapping and fictitious-receipts pattern.
- **An item that cleared before period end**: it was not outstanding at all and should never have
  been on the reconciliation.

The script computes days-to-clear for each item and flags the ones that do not fit their claimed
category.

## Step 3 — Deliver

**Workbook tabs:**

1. **Review Summary** — independent recomputation against the preparer's figures, the eight tests,
   the control observations, and the conclusion. The signable page.
2. **Independent Recomputation** — the four-column form rebuilt from item detail, alongside what
   the preparer presented, with differences.
3. **Item Testing** — every reconciling item: amount, category, support obtained, whether it
   cleared, when, days to clear, and any flag. Unsupported and uncleared items sort to the top.
4. **Uncleared Items** — everything that did not clear, with age and the follow-up required.
5. **Stale Items** — items carried from prior reconciliations, with the number of periods.
6. **Plug Indicators** — generic descriptions, round amounts, and items with no counterpart.
7. **Control Observations** — preparation timeliness, review evidence, segregation of duties.

**Then, in chat:** whether the recomputation agrees, whether both balances agreed to independent
sources, how many items lacked support, how many did not clear, and the control observations. Lead
with anything unsupported or uncleared — those are the findings; the arithmetic rarely is.

## What to escalate

- **A reconciling item with no counterpart in any bank statement, before or after.** This is the
  fabricated-item pattern and it is escalated regardless of amount.
- **A deposit in transit that never cleared.**
- **An item carried on three or more consecutive reconciliations.** It is not timing.
- **A reconciliation that only foots because of an item labelled generically.** Report the label
  and the amount together — the juxtaposition is the finding.
- **An unreviewed reconciliation**, or one prepared by someone with disbursement authority.
- **A reconciliation prepared long after period end.** The longer the delay, the more opportunity
  to construct rather than reconcile.
- **Adjusting entries made to the cash account rather than to the reconciliation** — these bypass
  the reconciliation entirely.
- **Prior-period reconciling items that disappeared without clearing.** Something was removed
  rather than resolved, and the reconciliation still foots, which means something else absorbed it.

State each as a fact with the item reference and amount. Do not characterise intent.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
