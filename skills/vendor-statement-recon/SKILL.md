---
name: vendor-statement-recon
description: Reconciles vendor statements to the AP subledger, per vendor — classifying every statement line and every open subledger line as matched, missing from the subledger, disputed, or carried in the subledger but absent from the statement — and ties the reconciled total to the subledger open balance or reports a quantified, named difference. Use this whenever the user mentions vendor statement reconciliation, reconciling AP to vendor statements, unapplied credits, missing invoices on a vendor statement, or AP subledger tie-out by vendor. Also trigger on "reconcile this vendor statement," "does our AP agree to what the vendor says we owe," "find missing invoices," or when vendor statement PDFs/extracts and an AP subledger export are provided together. High-volume, offshore-displacing grunt work at scale — the AP analog of bank-rec-review. Runs fully local — no vendor or AP data leaves the machine.
---

# Vendor Statement Reconciliation

A vendor statement and an AP subledger will almost never agree line-for-line by coincidence —
timing differences, unapplied credits, and genuinely missing invoices are normal. The point of
this skill is to classify every difference explicitly rather than let a vendor's claimed balance
and the subledger's balance sit unreconciled with no record of why.

## The gate

The script refuses to produce a clean workbook when:

1. **An amount or date does not parse.** A visible, blocking exception on the row it came from —
   never coerced or dropped.
2. **A `(vendor, invoice_no)` pair is duplicated** in the subledger. Matching against an
   ambiguous target is unreliable, so the workbook is not written until the subledger is
   deduplicated or corrected.

A reconciling difference after classification is never a gate failure — it is the normal,
expected output of a recon with real discrepancies, reported as a quantified amount on the
Per-Vendor Reconciliation tab, never forced to zero or absorbed into a plug.

## Inputs

1. **Vendor statement extract**, all vendors together — `vendor, statement_date, ref, amount,
   line_type` (`line_type`: `invoice`, `credit`, or `payment`).
2. **AP subledger export** — `vendor, invoice_no, amount, invoice_date, status` (`status`: `open`
   or `paid`).

## Step 1 — Run the reconciliation

```bash
python3 scripts/vendor_statement_recon.py --statement vendor_statements.csv \
  --subledger ap_subledger.csv --client "Northwind Systems Inc" \
  --out "Northwind - Vendor Statement Recon.xlsx"
```

## Step 2 — How a statement line is classified

| Classification | Meaning |
|---|---|
| **matched** | Statement line's `ref` found in the subledger for that vendor, amount agrees. |
| **disputed** | `ref` found, but the statement and subledger amounts disagree. |
| **missing-from-subledger** | Statement shows an invoice or credit with no matching `ref` in our subledger — a possible unrecorded liability (invoice) or an uncaptured credit. |
| **in-subledger-not-on-statement** | An **open** subledger invoice whose number never appears on the statement — could mean the vendor's statement is stale, or our payable is. |

The reconciled balance starts from the subledger's open total and applies each classified item's
dollar effect; the result is compared to the statement's own implied ending balance. Agreement
within a cent is a tie. Anything else is reported as a named difference, never plugged.

## Step 3 — Proposed actions are suggestions, not determinations

Every discrepancy carries a proposed next step (request an invoice copy, apply a credit, confirm
with the vendor) — explicitly labeled `SUGGESTED (not a determination)`. The skill never
concludes a discrepancy is resolved; it stops at classifying and quantifying, and the engagement
decides the actual resolution.

## Deliver

**Workbook tabs:**

1. **Per-Vendor Reconciliation** — statement balance, subledger open balance, reconciled balance,
   difference, and tie/no-tie per vendor.
2. **Discrepancy List** — every classified item with its suggested (not determined) action.
3. **Unreconciled-Exceptions** — vendors whose difference did not tie, plus any gate failures if
   the workbook was forced.

**Then, in chat:** which vendors tied and which didn't, the largest unreconciled dollar amounts
first, and a count of missing-from-subledger items — those are the ones with the clearest
unrecorded-liability exposure.

## What to escalate

- **Missing-from-subledger invoices concentrated near period end** — unrecorded-liability risk
  for the close, not just an AP housekeeping item.
- **A vendor with a persistent, unexplained difference across multiple statement periods** — a
  recurring tie failure usually means a systemic posting issue, not a one-off.
- **Unapplied credits sitting on a vendor statement for several periods** — cash or discount
  value not yet realized.
- **A disputed amount where the subledger figure is larger than the statement figure** — worth
  confirming we haven't accrued or posted more than the vendor believes we owe.

Report each as a fact with the vendor and reference. Do not characterize intent or conclude a
dispute's resolution — that decision belongs to whoever is working the vendor relationship.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.
Vendor and AP data never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
