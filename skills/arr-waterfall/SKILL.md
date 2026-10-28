---
name: arr-waterfall
description: Builds the ARR/MRR movement waterfall (new, expansion, contraction, churn) from a subscription export and the prior period's ending ARR, closes the arithmetic identity between beginning and ending ARR, and ties ending ARR to recognized revenue — refusing to deliver when the waterfall doesn't close or ending ARR doesn't tie to the GL. Use this whenever the user mentions the ARR waterfall, MRR movement, new/expansion/contraction/churn classification, the SaaS board metric, or ARR rollforward. Also trigger on "build the ARR waterfall," "what's our net revenue retention," "classify these customers as new/expansion/churn," "tie ARR to revenue," or when current and prior-period subscription exports are supplied together. Commonly paired with `deferred-revenue-waterfall` (customer ARR movement vs. the GL deferred balance) and `revenue-trace-to-source`. Runs fully local.
---

# ARR Waterfall

## The gate

The script refuses to produce a clean waterfall when:

1. **The waterfall doesn't arithmetically close.** Beginning ARR plus new, expansion,
   contraction, and churn must equal ending ARR exactly. This is a closed identity — if it
   doesn't hold, a customer was miscounted or double-counted somewhere in the movement logic,
   and the script treats that as a defect to surface, not round off.
2. **Ending ARR doesn't tie to GL recognized revenue** within whatever tolerance you supply. No
   tolerance means an exact tie is required; a mismatch is reported in dollars, never absorbed.
3. **A customer's `mrr_amount` or `prior_mrr` is missing or unparseable.** That customer is
   excluded and reported, never treated as zero, since zero would misclassify a real customer as
   new or churned.

A customer present in the prior file but **absent entirely** from the current extract is still
picked up and classified Churn — the script indexes off the union of both files, not just the
current one, specifically so a dropped row can't disappear the churn it represents.

## Inputs

1. **Current-period subscription or invoice export** — customer, plan, MRR amount, effective
   date.
2. **Prior-period ending ARR/MRR by customer** — ideally the prior run's own ending state, so the
   waterfall chains period to period without manual reconstruction.
3. **The GL recognized revenue figure** for the period, for the tie-out, and an optional
   tolerance if your revenue recognition timing doesn't move in lockstep with ARR changes.

## Step 1 — Build the worksheet

```bash
python3 scripts/arr_waterfall.py --current current.csv --prior prior.csv \
  --template --out current_worksheet.csv
```

## Step 2 — Run the waterfall

```bash
python3 scripts/arr_waterfall.py --current current_worksheet.csv --prior prior.csv \
  --gl-revenue 512300.00 --tolerance 250.00 \
  --client "Northwind Systems Inc" \
  --out "Northwind - Q4 2025 ARR Waterfall.xlsx"
```

## Deliver

**Workbook tabs:**

1. **Workpaper Summary** — beginning ARR, each movement bucket, ending ARR, the GL tie, and the
   conclusion.
2. **Waterfall** — every customer with prior MRR, current MRR, delta, and movement
   classification, with bucket totals and the full ARR rollforward below the detail.
3. **Revenue Tie** — ending ARR against GL recognized revenue, with the difference.
4. **Exceptions** — every data exception, including customers excluded for unparseable amounts.
5. **Judgment Items** — large contractions (a customer's MRR dropping sharply but not to zero),
   flagged for a human to confirm is a genuine downsell rather than a churn event recorded as a
   reduced renewal.

**Then, in chat:** the headline movement numbers first (new, expansion, contraction, churn,
ending ARR), then the GL tie result, then anything in Judgment Items worth a second look.

## What to escalate

- **A customer with a large contraction** close to full churn — the boundary between "steep
  downsell" and "effectively churned" is a judgment call this skill flags but does not resolve.
- **A paused-but-not-cancelled contract** showing as churn or new depending on how the extract
  represents a $0 MRR row during the pause — confirm which period's data reflects the pause.
- **Any waterfall identity failure**, which indicates a bug in the movement classification or a
  data quality problem in the extract, not a rounding issue to ignore.
- **A persistent, growing gap** between ending ARR and GL recognized revenue, which usually means
  the two are measuring different things (bookings vs. recognized) and the comparison basis needs
  revisiting, not a bigger tolerance.

Report each as a fact with the customer reference. Do not characterize intent.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.
Subscription and customer data never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
