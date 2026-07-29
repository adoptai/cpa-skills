---
name: sales-tax-reconciliation
description: Reconcile filed sales and use tax returns to recorded sales by jurisdiction — proving gross sales agree to the general ledger, that tax collected equals tax remitted through a liability rollforward, and that every jurisdiction with sales activity either has a return filed or a documented nexus conclusion. Use this whenever the user mentions sales tax, use tax, sales tax returns, reconciling sales tax to the GL, sales tax by jurisdiction or state, economic nexus, a sales tax notice or audit, exemption certificates, marketplace facilitator sales, or a sales tax liability account that will not clear. Also trigger on "do our sales tax returns tie to revenue," "which states should we be filing in," "we got a nexus questionnaire," "reconcile sales tax," "check our exempt sales," or when sales tax returns and a sales ledger are provided together. Runs fully local — no client sales data leaves the machine.
---

# Sales and Use Tax Reconciliation

Sales tax is the exposure that grows quietly. It is trust-fund money — collected from customers
and held for a jurisdiction — and in most states an unremitted balance attaches personally to
responsible persons. Unlike income tax, there is often no statute of limitations running in the
taxpayer's favour where no return was ever filed, so an unfiled jurisdiction compounds
indefinitely.

Three failures produce nearly all of the damage:

1. **Collected but not remitted.** The liability account holds a balance nobody clears, and the
   size is only visible in a rollforward.
2. **Sales in a jurisdiction with no return and no documented conclusion.** Not a filing error —
   an unquantified liability with no statute running.
3. **Marketplace sales double-reported or double-taxed.** The facilitator remits, and the seller
   reports the same sales again.

None appear on the face of a return. This skill starts from the ledger and works outward.

## The gate

No clean workpaper unless:

1. **Gross sales per all returns reconcile to GL sales revenue**, with every difference named —
   exempt sales, out-of-state, marketplace-facilitated, intercompany, freight, discounts, returns.
2. **The tax liability rollforward foots.** Beginning liability + tax collected per the GL − tax
   remitted = ending liability, and the ending balance is explainable as amounts not yet due.
3. **Every jurisdiction with sales activity is accounted for** — a return filed, or a written
   nexus conclusion on file. Silence is not an answer.
4. **Effective rates are derived from the returns themselves** and compared across periods.

## Standing rule on rates and thresholds

**Do not state a sales tax rate, an economic nexus threshold, a transaction count, a filing
frequency, or a due date from memory.** Rates change constantly, often mid-year and at the local
level. Economic nexus thresholds vary by state and several have changed since being introduced.

This skill instead:

- **Derives the effective rate** per jurisdiction per period as `tax reported ÷ taxable sales`
  from the filed returns, and compares it across periods. A rate that moves without a stated
  cause is a finding regardless of what the correct rate is — and this catches a rate change the
  client never implemented, which is the common case.
- **Reports sales and transaction counts by jurisdiction** and instructs verification against
  each state's *current* threshold rather than testing against a number. The output is a
  nexus-screening schedule, not a nexus conclusion.
- Flags any jurisdiction where activity exists and no return was filed, for a documented
  conclusion.

A confident wrong threshold here creates real exposure in both directions — a missed filing
obligation, or unnecessary registrations that create permanent filing burden. Screen and route;
do not conclude.

## Inputs

1. **Sales ledger or revenue detail by jurisdiction** — gross sales, taxable sales, exempt sales,
   tax collected, transaction count, per jurisdiction per period. Ship-to based, not bill-to,
   for destination-sourcing states.
2. **Filed returns** — per jurisdiction per period: gross sales, exempt or deducted sales,
   taxable sales, tax reported, tax remitted, and any discount or timely-filing credit taken.
3. **GL detail** — sales revenue, the sales tax liability account with beginning and ending
   balances, and remittances.
4. **Marketplace facilitator reports**, if the client sells through marketplaces. Essential —
   these sales are usually reported and remitted by the facilitator, and including them again on
   the seller's return overstates the liability.
5. **Exemption certificate register**, if available — a count and the customers covered.
6. **Registration list** — where the client is registered and the filing frequency assigned.

## Step 1 — Reconcile and roll forward

```bash
python3 scripts/sales_tax_recon.py \
  --sales sales_by_jurisdiction.csv \
  --returns filed_returns.csv \
  --gl gl.csv \
  --registrations registrations.csv \
  --marketplace marketplace.csv \
  --reconciling-items recon_items.csv \
  --client "Ridgeline Outdoor Co" --period "FY2025" \
  --out "Ridgeline - FY2025 Sales Tax Reconciliation.xlsx"
```

Six tests:

- **Test 1 — Gross sales bridge.** GL sales revenue reconciles to the sum of gross sales per
  returns, through named reconciling items. This is where basis differences surface: a GL kept on
  accrual against returns filed on cash, or freight included in one and not the other.
- **Test 2 — Taxable sales and tax reported.** Per jurisdiction, `taxable sales × derived rate =
  tax reported`, and gross less exempt equals taxable. An arithmetic break here is a return
  preparation error.
- **Test 3 — Tax collected versus tax reported.** Tax collected per the GL agrees to tax reported
  per the returns. A shortfall means tax was collected and not reported; an excess means tax was
  reported that was never collected, which the client funded out of margin.
- **Test 4 — Liability rollforward.** Beginning + collected − remitted = ending, with the ending
  balance compared to the final period's liability. An excess is the collected-not-remitted
  finding, and it is escalated immediately.
- **Test 5 — Jurisdiction completeness.** Every jurisdiction with sales has a return or a
  documented nexus conclusion, and every return filed corresponds to actual activity. Both
  directions: filing where there is no activity and no registration requirement is also worth
  raising, because it creates a permanent obligation.
- **Test 6 — Rate consistency.** Derived effective rates compared across periods per
  jurisdiction. Movement without a stated cause is flagged.

## Step 2 — Read the nexus screening schedule

The workbook produces a screening schedule per jurisdiction: sales, transaction count, whether
registered, whether returns were filed, and marketplace-facilitated sales shown separately.

**Marketplace sales are separated deliberately.** In most states the facilitator's sales count
toward the facilitator's obligation, not the seller's, and in several they are excluded from the
seller's threshold calculation entirely. Aggregating them inflates apparent nexus and produces
registrations the client does not need. The schedule shows direct sales, marketplace sales, and
the total, so the threshold question can be answered on the right base once you have confirmed
which base that state uses.

For each jurisdiction the schedule carries the instruction to confirm the current threshold,
measurement period, and whether it counts gross or taxable sales — all of which vary.

## Step 3 — Exempt sales

Exempt sales are the second most common assessment in a sales tax audit, after unfiled
jurisdictions. The test is not whether the sale was exempt; it is whether a **valid certificate
is on file**. In an audit, an exempt sale without a certificate is a taxable sale, and the tax is
assessed against the seller who can rarely collect it from the customer years later.

The script compares exempt sales value and customer count against the certificate register where
supplied, and reports exempt sales with no matching certificate. Where no register exists, that
is the finding — say so plainly and quantify the exposure as exempt sales at the derived rate.

## Step 4 — Deliver

**Workbook tabs:**

1. **Reconciliation** — the gross sales bridge, the liability rollforward, and the six test
   results. The signable page.
2. **By Jurisdiction** — sales, taxable, exempt, tax collected, tax reported, tax remitted,
   derived rate, and differences, per jurisdiction per period.
3. **Nexus Screening** — direct sales, marketplace sales, transaction counts, registration and
   filing status, and the confirmation instruction per jurisdiction.
4. **Liability Rollforward** — beginning to ending with remittances, and the excess over the
   final period's liability isolated.
5. **Rate Analysis** — derived effective rate by jurisdiction by period, with movement flagged.
6. **Exempt Sales** — exempt sales by jurisdiction and customer, certificate status, and the
   exposure at the derived rate where no certificate exists.
7. **Exceptions & Memo** — unexplained differences, escalations, and the figures requiring
   confirmation against current authority.

**Then, in chat:** whether returns tie to the ledger, the liability rollforward result, any
jurisdiction with activity and no return, and the exempt-sales exposure. Lead with unremitted tax
and unfiled jurisdictions — both grow, and one of them grows without a statute running.

## What to escalate immediately

- **Tax collected and not remitted.** Trust-fund money with personal liability for responsible
  persons in most states. Days matter.
- **Sales in a jurisdiction with no return and no documented nexus conclusion.** Where no return
  was ever filed, the lookback is typically open indefinitely. Voluntary disclosure programs exist
  and generally require the taxpayer to approach the state *before* being contacted — so the
  window closes the moment a notice arrives. That timing is why this is escalated rather than
  scheduled.
- **Exempt sales with no certificate on file**, quantified at the derived rate.
- **Marketplace sales included in the seller's own returns**, which overstates the liability and
  may have been remitted twice.
- **A rate that never changed** in a jurisdiction that changed its rate during the period.
- **Use tax never accrued** on purchases where no sales tax was charged — a self-assessment
  obligation clients routinely miss, and one auditors examine first because it is easy to find.
- **Registration in a jurisdiction with no activity**, creating a filing obligation and penalty
  exposure for non-filing of zero returns.

Report each as a fact with the supporting figures. Whether to file a voluntary disclosure, amend,
or register is a decision for the person signing.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry, no
rate-lookup services. Customer and sales data never leaves the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
