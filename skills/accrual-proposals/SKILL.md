---
name: accrual-proposals
description: Propose month-end accruals from open POs received but not yet invoiced (GRNI), recurring vendor spend patterns, and preparer-supplied one-time items, cross-checked against the AP register so nothing is double-booked, with every line carrying a stated basis and an auto-reversal flag for reviewer approval. Never books anything - output is a proposed JE batch for a Controller to approve before it posts to the GL. Use whenever the user mentions month-end accruals, accrual proposals, what should we accrue this month, GRNI, goods received not invoiced, a recurring accrual estimate, or the accrual JE batch for close. Also trigger on "propose accruals," "draft the accrual entries," "what are we missing on the accrual side," or when an open PO report, AP register, or vendor spend history is supplied together at month end. Inverts cutoff-and-unrecorded-liabilities: that proves a search for what was missed is complete; this proposes what to book. Runs fully local - no client data leaves the machine.
---

# Month-End Accrual Proposals

`cutoff-and-unrecorded-liabilities` proves a search is complete and reports what it found.
This skill runs the other direction: it proposes what should be booked, before a reviewer has
asked for it. Same underlying question — what happened before period end that isn't on the
books yet — but the deliverable is a draft journal entry batch instead of a search memo.

Three sources, in order of reliability:

1. **Open purchase orders with goods or services received but not yet invoiced (GRNI).** The
   most defensible accrual, because it is backed by a receipt, not an estimate.
2. **Recurring vendor spend** — utilities, rent, subscriptions — accrued at the trailing
   average of prior periods, but only when there is enough history and only when the actual
   invoice for the current period isn't already recorded.
3. **Manual one-time items** the preparer knows about — legal fees, bonuses, a settlement —
   each of which must carry its own stated basis. This skill will not accept an amount with no
   reason attached.

Nothing here gets posted. Every proposed line is a **recommendation with a basis and a
reversal flag**, built for a Controller to approve or reject line by line.

## The gate

No clean workpaper unless:

1. **Every open-PO line, every recurring vendor/category candidate, and every manual item is
   classified** — proposed, excluded with a reason, or an exception with a reason. Nothing is
   silently dropped, and nothing is silently proposed either.
2. **Every proposed line has a non-blank basis and an explicit reversal flag.** A non-reversing
   item requires a stated rationale — "no" is never the default.
3. **Open-PO candidates are cross-checked against the AP register.** If the PO report's
   "invoiced" figure doesn't match what the AP register actually shows, that PO becomes an
   exception, not a proposal — a stale PO report is the single most common way this kind of
   tool double-books an accrual.
4. **Recurring accruals require a minimum amount of history** (default three prior periods,
   configurable) and are dropped to an exception, never guessed, when there isn't enough.
5. **The proposed batch balances** — total debits equal total credits — before it is handed to
   anyone.

## Inputs

1. **Open PO report** — purchase orders with `received_amount` and `invoiced_amount` in
   dollars, plus a `receipt_date`. The difference is the candidate accrual; a missing receipt
   date makes the line an exception rather than a guess about when it was received.
2. **AP invoice register** — invoices recorded through period end, with `po_no` so it can be
   tied back to the open PO report, and `vendor`/`category` so a recurring estimate can be
   skipped when the actual invoice already showed up.
3. **Vendor spend history** — prior-period actuals by vendor and category, for the recurring
   accrual average. More periods produce a more defensible average; too few produce an
   exception instead of an estimate.
4. **Prior-period accruals**, with a `reversed` flag — used to flag anything that should have
   reversed and didn't, which is a double-counting risk this period.
5. **Manual items**, if the preparer has known one-time accruals to add — each needs its own
   basis. This file is optional; nothing here fabricates a manual item on its own.
6. **Period end date** and, only if the close period isn't a calendar month, a period start
   override. **Materiality**, if you want the total measured against it.

## Step 1 — Run the proposal

```bash
python3 scripts/accrual_proposal.py \
  --open-pos open_pos.csv --ap-register ap_register.csv \
  --vendor-history vendor_history.csv --prior-accruals prior_accruals.csv \
  --manual-items manual_items.csv \
  --period-end 2026-08-31 --materiality 25000 \
  --client "Ardenway Manufacturing" \
  --out "Ardenway - August 2026 Accrual Proposals.xlsx"
```

Every input file except `--open-pos`, `--period-end`, and `--out` is optional — supply what
you have. Fewer sources means fewer proposals and more exceptions, not guessed numbers filling
the gap. `--min-history` (default 3) sets how many prior periods a recurring accrual needs
before it's proposed rather than flagged; `--reversal-date` defaults to the first day of the
month after period end.

## Step 2 — Get the date and the cross-check right

The GRNI accrual depends on **when the receipt happened**, not the invoice date — a PO
receipted 28 August with no invoice yet is an August liability regardless of when the invoice
eventually arrives.

The AP register cross-check exists because a PO report's own "invoiced" column goes stale
quickly — an invoice keyed against the PO after the report was pulled makes the PO report say
"not invoiced" for something that already is. Supplying the AP register lets the script catch
that; without it, GRNI candidates are proposed on the PO report's word alone, which the tool
says explicitly in its output.

## Step 3 — Work the exceptions before the batch goes to the Controller

The Exceptions tab is not noise — it's the population the tool refused to guess on. Common
reasons: no receipt date, invoiced-exceeds-received (usually a data error in the PO system),
a PO report that doesn't tie to the AP register, insufficient vendor history for a recurring
estimate, a manual item with no basis, and a non-reversing manual item with no rationale. Each
is worth five minutes before the batch is finalized — most convert into either a proposal or a
documented reason to skip it, not a lingering unknown.

## Step 4 — Deliver

**Workbook tabs:**

1. **Summary** — proposal totals by source, materiality comparison, exception count and value,
   prior-accrual-not-reversed count, the gate results, and sign-off lines for preparer and
   Controller.
2. **Proposed Accrual JE Batch** — the deliverable: debit and credit accounts and amounts,
   basis, reversing flag, and reversal date, one pair of lines per proposed entry.
3. **Auto-Reversal Schedule** — the mirrored entries as they'll look on the reversal date, so
   the reviewer can see the whole lifecycle of each accrual, not just the booking.
4. **Open PO Detail (GRNI)** — every PO line with its classification and notes.
5. **Recurring Vendor Patterns** — the trailing-average calculation, periods used, and
   classification for each vendor/category.
6. **Manual Items** — the preparer's one-time items, each with its basis and classification.
7. **Exceptions** — every item that could not be proposed, consolidated across sources, with
   the reason.
8. **Prior Accrual Reversal Check** — prior-period accruals not marked reversed, which is a
   double-counting risk against anything proposed this period.

**Then, in chat:** the total proposed, how many exceptions and their value, whether the total
exceeds materiality, and a plain statement that this is a proposal only — nothing has been
posted, and Controller approval is required before it is booked.

## What to escalate

- **Any PO where the PO report's invoiced amount doesn't match the AP register.** This means
  the PO report can't be trusted for other periods either until it's reconciled.
- **No AP register supplied.** GRNI candidates are then proposed on the PO report's own word,
  with no independent check — say so, don't imply the cross-check ran.
- **Prior-period accruals not marked reversed.** Booking this period's proposal on top of one
  that never reversed is a straightforward double-count.
- **Recurring categories present in the prior year's pattern and absent from this period's
  vendor history** with no explanation — usually means data wasn't pulled, not that the
  expense stopped.
- **A pattern of PO receipts recorded just after period end** that would have been material if
  dated a day earlier — worth a look at whether the receipt date itself is being entered
  accurately.
- **Non-reversing manual items.** These bypass the normal check-and-clear the following month,
  so the rationale needs to actually hold up, not just be present.
- **Total proposed accruals against materiality**, and total exceptions against materiality —
  an exception population large enough to matter is itself a finding.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
