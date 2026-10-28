---
name: duplicate-payment-scan
description: Scans the full AP invoice and payment population — never a sample — for duplicate-payment candidates across four risk-ranked tiers (exact duplicate invoice numbers, same vendor/amount near in time, vendor-name variants, and reused payment references), with every finding traceable to the source rows it's based on. Use this whenever the user mentions duplicate payments, double payments, duplicate invoices, paying a vendor twice, AP duplicate scan, finding overpayments, or recovering cash from AP. Also trigger on "scan AP for duplicates," "did we pay this vendor twice," "find duplicate invoices," "duplicate payment risk," or when 12+ months of AP invoice and payment registers are provided for review. Reuses the full-population scan pattern from journal-entry-anomaly-scan, pointed at AP instead of the GL. Runs fully local — no vendor or payment data leaves the machine.
---

# Duplicate Payment Scan

Findings here are self-evidently valuable — every confirmed duplicate is recoverable cash. The
discipline is in proving each finding with the actual matched rows rather than asserting
"duplicate" and moving on.

## The gate

The scan refuses to produce a clean workbook when:

1. **Coverage cannot be proven.** Every row read must end up either scanned or in the
   unscanned/unparsed list with a stated reason. Scanned count plus unscanned count must equal
   total rows read — checked explicitly, not assumed.
2. **A `record_id` is duplicated** in the source file. Row identity must be unambiguous before
   any finding can reliably point back to a specific transaction.

A row with an unparseable amount or date is not itself a gate failure — it is excluded from
matching and listed, with its reason, on the Unscanned/Unparsed Rows tab, so the scope limitation
is visible rather than silently absorbed into a clean-looking result.

## Inputs

**AP invoice and payment register (12+ months)** — one file, one row per invoice or payment
event: `record_id, vendor, amount, invoice_date, invoice_no, payment_date, payment_ref`.
`invoice_no`, `payment_date`, and `payment_ref` may be blank on any given row.

## Step 1 — Run the scan

```bash
python3 scripts/duplicate_payment_scan.py --register ap_register.csv --window-days 5 \
  --client "Northwind Systems Inc" \
  --out "Northwind - Duplicate Payment Scan.xlsx"
```

`--window-days` has no built-in default — state the proximity window you actually want checked
(a tighter window for weekly-run AP, a wider one for monthly-batch AP) rather than relying on a
number this skill would otherwise have to invent.

## Step 2 — What each tier catches, and why it's ranked there

| Tier | Match | Why it's ranked there |
|---|---|---|
| **1** | Exact vendor + amount + `invoice_no` duplicated | The same invoice number paid twice for the same vendor and amount is close to unambiguous — the highest-confidence finding in the set. |
| **2** | Exact vendor + amount + date within the window | Same vendor, same amount, close in time, different invoice numbers — classic re-submission or re-entry duplicate. |
| **3** | Same amount + date within the window, vendor name a normalized variant of another (e.g. "Acme Supply Co" vs. "Acme Supply") | Catches duplicates hidden by vendor-master inconsistency — often the hardest tier to find by eye and the one worth the normalization log. |
| **4** | Same `payment_ref` reused across different invoice numbers | A check or EFT reference applied to more than one invoice is either a batching convention or a posting error — always worth a look. |

Each finding carries the matched records' row references and fields, never a bare assertion.

## Deliver

**Workbook tabs:**

1. **Risk Summary** — population rows read, scanned vs. unscanned, population dollar total, and
   finding counts and dollars at risk by tier.
2. **Findings** — every match with tier, basis, vendor(s), amount, matched record references, and
   dollars at risk.
3. **Unscanned-Unparsed Rows** — every row excluded from scanning, with the reason.
4. **Vendor Normalization Log** — every raw vendor string mapped to its normalized form, so a
   Tier 3 match can be audited back to the exact spelling variants involved.

**Then, in chat:** the population scanned, total dollars at risk by tier, and the highest-tier
findings first. State plainly if the unscanned count is material — a scan with a large unparsed
population proves less than it appears to.

## What to escalate

- **Any Tier 1 finding** — confirm with the bank or payment processor whether both payments
  actually cleared before assuming it's a false positive.
- **A payment reference reused across invoices to different vendors**, if it occurs — a
  reference collision this specific is unusual and worth confirming isn't a data-entry artifact
  versus a real misapplied payment.
- **A vendor-name variant cluster (Tier 3) spanning many distinct normalized groups** — may
  indicate the vendor master itself needs cleanup, a separate project from the duplicates found
  this pass.
- **Concentration of findings in one AP clerk's or approver's batch**, if that data is
  available separately — the scan does not have approver data and cannot make this link itself,
  but a reviewer with access to it should check.

Report each as a fact with the record references. Do not characterize intent — the scan reports
pattern and evidence, never motive.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.
Vendor and payment data never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
