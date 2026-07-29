---
name: three-way-match
description: Perform a three-way match between purchase orders, vendor invoices, and receiving records — testing quantity and price against tolerance, detecting duplicate invoices and over-billing, and proving that every invoice in the population is either matched or reported as an exception. Use this whenever the user mentions three-way match, PO to invoice matching, matching invoices to receipts, accounts payable testing, duplicate payment review, over-billing, price variance on invoices, unmatched invoices, goods-received-not-invoiced or invoiced-not-received, or AP internal control testing. Also trigger on "match these invoices to POs," "did we pay twice," "check AP against receiving," "find duplicate payments," "test our purchasing controls," "GRNI review," or when purchase order, invoice, and receiving files are provided together. Runs fully local — no vendor or purchasing data leaves the machine.
---

# Three-Way Match

A three-way match asks whether the company **ordered** it, **received** it, and was **billed
correctly** for it. The control is old and well understood; what makes testing it valuable is
that the failures are almost never random.

Duplicate payments cluster around the same handful of vendors, because the cause is a process
defect — an invoice submitted twice under slightly different numbers, a credit memo never
applied, a PO closed and reopened. Price variances cluster too, because someone is invoicing off
a stale price list. **So the pattern matters more than the individual exception**, and the output
is organized to make patterns visible rather than to produce a long flat list.

## The gate

The population is proven before it is analysed:

1. **Every invoice is accounted for.** Matched invoices plus exceptions must equal the full
   invoice population, in count and in value. An invoice that quietly falls out of the middle of
   a matching routine is the one worth finding, so the script balances the population and
   refuses to report until it does.
2. **The invoice population ties to a stated control total** — the AP subledger or the GL
   purchases figure. Testing a filtered extract proves nothing about the account.
3. **No invoice is matched twice.** A PO line consumed by one invoice is not available to
   another; otherwise a duplicate looks like a match.
4. **Tolerances are stated, not assumed.** Quantity and price tolerance come from the client's
   own policy and are recorded on the workpaper.

## Inputs

Ask for all three, plus the control total:

1. **Purchase orders** — PO number, line, item, quantity ordered, unit price, vendor, date,
   status.
2. **Vendor invoices** — invoice number, vendor, date, PO reference, line, item, quantity
   billed, unit price, extended amount, and the amount actually paid if available.
3. **Receiving records** — receipt number, PO reference, line, item, quantity received, date.
4. **Tolerance policy** — the client's approved quantity and price variance thresholds, and any
   de minimis amount below which variances are accepted without approval.
5. **Control total** — AP subledger total or GL purchases for the period.

Where the client has no written tolerance policy, that is itself a finding. Use zero tolerance,
say that you did, and note the absence of a policy as a control observation.

## Step 1 — Prove the population

```bash
python3 scripts/three_way_match.py \
  --invoices inv.csv --pos po.csv --receipts rec.csv \
  --control-total 8412990.55 --validate-only
```

Reports counts and totals for each file, invoices with no PO reference, POs with no invoices,
receipts with no PO, and whether the invoice population ties to the control total. Resolve
differences first — in particular, ask what the export excluded. A receiving file limited to
one warehouse will manufacture hundreds of false "invoiced not received" exceptions and bury the
real ones.

## Step 2 — Match

```bash
python3 scripts/three_way_match.py \
  --invoices inv.csv --pos po.csv --receipts rec.csv \
  --control-total 8412990.55 \
  --qty-tolerance-pct 2 --price-tolerance-pct 1 --de-minimis 50 \
  --client "Cascade Industrial Supply" --period "FY2025" \
  --out "Cascade - FY2025 Three-Way Match.xlsx"
```

Matching runs on PO number and line where available, falling back to PO plus item. Each invoice
line receives one of these outcomes, and every outcome is reported:

| Outcome | Meaning |
|---|---|
| `matched` | Quantity and price within tolerance against both PO and receipt |
| `price variance` | Billed unit price outside tolerance of the PO price |
| `quantity variance` | Billed quantity outside tolerance of quantity received |
| `over-PO` | Billed quantity exceeds quantity ordered |
| `over-received` | Billed quantity exceeds quantity received — **billed for goods not received** |
| `not received` | Invoice matched to a PO with no receiving record at all |
| `no PO` | Invoice with no purchase order — the control was bypassed |
| `PO not found` | Invoice references a PO that does not exist in the file |
| `duplicate` | See below |

**Duplicate detection runs on four patterns**, because duplicates rarely repeat exactly:

- Same vendor and same invoice number
- Same vendor, same amount, same invoice date
- Same vendor, same amount, dates within a short window
- Same vendor and same PO line billed more than once

The third and fourth patterns find the ones that matter. An invoice resubmitted as `INV-4471`
and `INV4471`, or a PO line billed on two invoices a month apart, will pass an exact-match check
and has to be caught structurally.

## Step 3 — Read the exceptions as patterns

The workbook aggregates by vendor before it lists by invoice, because that is the order in which
the findings are actionable.

- **Recurring price variance for one vendor** — almost always a stale price list or an unrecorded
  price increase, not fraud. Real money, and recoverable.
- **Recurring over-billing for one vendor** — short shipments billed in full. Also recoverable,
  and worth a conversation about receiving discipline.
- **Duplicates concentrated in one vendor** — a submission process problem at their end or an AP
  intake problem at yours.
- **"No PO" concentrated in one requester or department** — the control is being bypassed
  routinely, which is a more serious finding than any single dollar amount in the listing.
- **Invoices consistently just below an approval threshold** — worth escalating. See below.
- **Goods received not invoiced (GRNI)** — receipts with no matching invoice represent an
  unrecorded liability at period end. This is a completeness finding for the financial
  statements, and it is on its own tab because it is the exception a payables test most often
  misses: nobody complains about an invoice that never arrived.

## Step 4 — Deliver

**Workbook tabs:**

1. **Summary** — population proof, tolerances used, outcome counts and values, recovery
   opportunity quantified, and the control observations. The page a controller reads.
2. **Population Proof** — invoices matched plus exceptions equals the population, in count and
   value, with the tie to the control total. The gate, shown.
3. **Exceptions by Vendor** — vendor rollup with exception counts, values, and the dominant
   exception type. Read this before the detail.
4. **Exception Detail** — every exception line with PO, receipt, invoice figures, variance in
   units and dollars, and columns for investigation and disposition.
5. **Duplicates** — grouped, with which pattern caught each and the amount at risk.
6. **GRNI** — received not invoiced, with age, for the accrual.
7. **Matched Detail** — clean matches with the rule that matched them, for re-performance.

**Then, in chat:** population and whether it ties, the three or four vendor-level patterns that
matter, quantified recovery opportunity, the GRNI accrual figure, and control observations. Lead
with recoverable dollars — that is what funds the next engagement.

## What to escalate rather than list

- **Invoices repeatedly just below an approval threshold** from the same vendor or requester.
  Splitting to avoid approval implies knowledge of the control.
- **A vendor with no PO discipline at all** and material spend.
- **Payment exceeding the invoice**, or payment where no invoice exists.
- **A vendor bank detail change** shortly before a payment, if that data is available — the
  classic payment-diversion pattern.
- **A vendor whose address, bank details, or contact match an employee record.** Report the match
  as a fact; do not characterize it.
- **Receiving records created after the invoice date**, particularly near period end — receiving
  documentation generated to clear a match rather than to record a delivery.

State these as attributes with the supporting documents. Do not conclude on intent.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry. Vendor
master data and bank details, if present in the inputs, are never transmitted and are excluded
from output beyond what the exception requires.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
