---
name: revenue-trace-to-source
description: Trace a revenue sample from the general ledger through to the invoice, the contract, and the cash receipt — proving each recorded amount agrees across the whole chain, and refusing to record an item as tested unless every required evidence link is present. Use this whenever the user mentions revenue testing, tracing revenue, vouching revenue, revenue substantive procedures, agreeing revenue to invoices or contracts or cash, revenue recognition testing, cutoff testing on revenue, or revenue existence and occurrence. Also trigger on "trace these revenue transactions," "vouch revenue to source," "test revenue recognition," "agree revenue to cash," "did this revenue actually happen," "revenue cutoff testing," or when a revenue population and supporting documents are provided together. Runs fully local — no client revenue or contract data leaves the machine.
---

# Revenue Trace to Source

Before anything else, be clear about which way you are walking, because the two directions test
different assertions and are constantly confused:

- **GL → source** (this skill) tests **occurrence and existence**. Does recorded revenue represent
  a real transaction, at the right amount, in the right period? It finds fictitious and
  overstated revenue.
- **Source → GL** tests **completeness**. Is every real transaction recorded? It finds unrecorded
  revenue and skimming.

**Tracing from the ledger cannot detect unrecorded revenue**, because unrecorded revenue is not in
the population you sampled from. If completeness is the concern, this is the wrong procedure and
the workpaper must say so rather than implying coverage it does not have. The script states the
direction and the assertion on the face of the workpaper for exactly this reason.

## The gate

An item is not "tested" until the evidence exists. The script refuses to produce a clean
workpaper when:

1. **A required evidence link is missing.** Each selection needs the invoice, the contract or
   order, and the cash receipt — or an explicit, documented reason why one does not apply.
   Concluding "agreed" with a blank evidence field is rejected.
2. **Amounts disagree across the chain.** GL amount, invoice amount, contract-derived amount
   (price × quantity), and cash received must agree or carry a named difference.
3. **The population does not tie to GL revenue.** You supply the total; testing a filtered
   extract proves nothing about the account.
4. **A conclusion contradicts its own figures** — marked `agreed` while a difference exists.

## Inputs

1. **The revenue population** and the GL revenue total it should equal. Use the
   `audit-sampling` skill to select from it — that gives you a reproducible seed and a documented
   basis, which this skill assumes rather than reinvents.
2. **The selections to test**, with the GL amount and date for each.
3. **Supporting documents** per selection: invoice, contract or purchase order, shipping or
   delivery evidence, and the cash receipt or bank credit.
4. **The period end date**, for cutoff testing.
5. **Subsequent-period credit memos**, if available. See below — these are the highest-yield
   document in a revenue test and are almost never requested.

## Step 1 — Build the trace worksheet

```bash
python3 scripts/revenue_trace.py --selections selections.csv \
  --population-total 18422900.00 --period-end 2025-12-31 \
  --client "Northwind Systems Inc" \
  --assertion "Occurrence of recorded revenue for FY2025" \
  --template --out trace_worksheet.csv
```

That emits a worksheet with the evidence columns to fill in. Then run the test:

```bash
python3 scripts/revenue_trace.py --selections trace_worksheet.csv \
  --population-total 18422900.00 --period-end 2025-12-31 \
  --credit-memos subsequent_credits.csv \
  --client "Northwind Systems Inc" \
  --assertion "Occurrence of recorded revenue for FY2025" \
  --out "Northwind - FY2025 Revenue Trace.xlsx"
```

## Step 2 — What each link proves, and what it doesn't

Each link in the chain answers a different question. Treating them as interchangeable is how a
revenue test looks complete while proving little.

| Link | What it proves | What it does not prove |
|---|---|---|
| **Invoice** | An amount was billed | That anything was delivered, or that the customer accepted it |
| **Contract or order** | The customer agreed to buy, at a price | That the obligation was satisfied this period |
| **Delivery or shipping** | Something was transferred | That the customer will pay |
| **Cash receipt** | The customer paid | Nothing about period — cash can arrive in any period |

**Cash is the strongest single link**, because a customer paying real money for a real thing is
hard to fabricate. But it is timing-neutral: a year-end receivable subsequently collected is
perfectly normal. What matters is the combination — *no cash, no receivable, and no delivery
evidence* is the signature of fictitious revenue, and no single link would have surfaced it.

The script computes a **completeness-of-evidence score** per item and reports which links are
missing, so a reviewer can see at a glance whether the population was tested on invoices alone.

## Step 3 — Cutoff and recognition

Cutoff is where legitimate businesses make errors and where aggressive ones push. The script
tests:

- **Revenue recorded in the period with delivery after period end** — recognised before the
  performance obligation was satisfied.
- **Revenue recorded in the period with an invoice dated after period end.**
- **Concentration in the final days of the period.** Reports revenue by day for the last stretch
  of the period and flags a spike. A spike is not itself wrong — many businesses have real
  period-end seasonality — but it is where to look.
- **Subsequent-period credit memos reversing tested revenue.** This is the single highest-yield
  test in the set. Revenue recorded at period end and credited shortly afterward was often never
  a valid sale: goods shipped early, a side agreement permitting return, or a bill-and-hold
  arrangement. Supply `--credit-memos` and the script matches them against the selections.

Recognition timing beyond cutoff — multiple performance obligations, variable consideration,
principal versus agent, licences versus services — is judgment. The script flags contracts whose
terms suggest such a question exists; it does not conclude on the treatment. Route those to the
person signing.

## Step 4 — Deliver

**Workbook tabs:**

1. **Workpaper Summary** — the direction and assertion stated explicitly, the population tie,
   coverage, evidence completeness across the sample, exception counts, and the conclusion block.
2. **Trace Detail** — every selection with GL, invoice, contract, delivery, and cash amounts and
   dates, each difference, the evidence score, and the disposition.
3. **Exceptions** — every difference and every missing link, with the dollar effect.
4. **Cutoff Testing** — items failing cutoff, revenue by day near period end, and the
   subsequent-credit-memo matches.
5. **Evidence Coverage** — a matrix of which links were obtained across the sample. If the sample
   was tested on invoices alone, this tab makes it obvious.
6. **Judgment Items** — recognition questions identified but deliberately not concluded, with
   what would need to be determined and by whom.

**Then, in chat:** the direction and assertion first — one sentence, because it bounds everything
else — then the population tie, the exceptions with dollar effect, cutoff findings, and what needs
a judgment call. If completeness is the client's actual concern, say plainly that this procedure
does not address it.

## What to escalate

- **Revenue with no cash, no receivable, and no delivery evidence.** The fictitious-revenue
  pattern. Escalate regardless of amount.
- **Subsequent credit memos reversing period-end revenue**, particularly clustered by customer or
  by salesperson.
- **Contract terms discovered during testing that the accounting does not reflect** — return
  rights, acceptance clauses, contingent fees, bill-and-hold. A side agreement changes the
  accounting and by design does not appear in the accounting records.
- **Invoices with no contract or order** on a population where contracts are the norm.
- **Related-party revenue** within third-party revenue.
- **A customer whose payments consistently come from a different entity** than the one invoiced.
- **Revenue recognised on a contract signed after period end**, which is unambiguous.
- **Round-dollar revenue transactions at period end** to a single customer.

Report each as a fact with the document reference. Do not characterize intent — describe what the
documents show and let the engagement decide.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.
Contract terms and customer data never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
