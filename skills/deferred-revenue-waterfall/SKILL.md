---
name: deferred-revenue-waterfall
description: Builds the deferred revenue waterfall from a contract or invoice population, recomputes straight-line recognition per contract, and ties the ending deferred revenue balance to the GL — refusing to deliver when the ending balance does not tie, a contract is missing its recognition method or service period, or a milestone/usage contract has no stated ending balance to tie against. Use this whenever the user mentions the deferred revenue waterfall, recomputing revenue recognition, tying deferred revenue to the GL, SaaS/subscription recognition schedules, or unearned revenue rollforward. Also trigger on "build the deferred revenue waterfall," "recompute revenue recognition," "tie deferred revenue to the GL," "does our deferred revenue balance make sense," or when a contract/invoice export and a GL deferred revenue balance are supplied together. Pairs with `revenue-trace-to-source`, which proves one sampled transaction is real; this proves the whole contract population's deferred balance. Runs fully local.
---

# Deferred Revenue Waterfall

## The gate

The script refuses to produce a clean waterfall when:

1. **A contract is missing its recognition method, service start, or service end**, or the
   service period is backwards (`service_end` before `service_start`). It is excluded and
   reported as a data exception — never defaulted to `straight_line` or given a guessed period.
2. **A `milestone` or `usage` contract has no `stated_ending_deferred`.** These methods are not
   mechanically recognized by this skill — recognition timing on them is a professional call.
   Without the client's own stated ending balance for that contract, it cannot enter the tie-out.
3. **The computed ending deferred revenue does not tie to the GL ending balance exactly.** The
   difference is reported in dollars, never plugged or absorbed into a residual line.
4. **A contract's recognized revenue for the period exceeds what was available to recognize**
   (its beginning deferred balance plus this period's additions) — a defensive check on the
   recognition math itself.
5. **An invoice number repeats for the same customer**, which usually means a duplicate record
   in the export, not two real invoices.

## Inputs

1. **Contract or invoice export** with `recognition_method` (`straight_line`, `milestone`, or
   `usage`) stated per contract by the client — never inferred from plan name or amount.
2. **Service start/end dates** per contract — the period over which the performance obligation
   is satisfied, not the invoice or payment terms.
3. **The GL deferred revenue beginning and ending balance** for the period being tested.
4. **For `milestone`/`usage` contracts only**, the client's own stated ending deferred balance,
   since this skill does not compute recognition timing for those methods.

## Step 1 — Build the worksheet

```bash
python3 scripts/deferred_revenue.py --contracts contracts.csv \
  --period-start 2025-10-01 --period-end 2025-12-31 \
  --template --out contracts_worksheet.csv
```

## Step 2 — Run the waterfall

```bash
python3 scripts/deferred_revenue.py --contracts contracts_worksheet.csv \
  --period-start 2025-10-01 --period-end 2025-12-31 \
  --gl-beginning 1842500.00 --gl-ending 2015300.00 \
  --client "Northwind Systems Inc" \
  --out "Northwind - Q4 2025 Deferred Revenue Waterfall.xlsx"
```

`straight_line` contracts are recognized day-by-day over their service period, split into
beginning deferred balance, additions (invoiced this period), revenue recognized this period, and
ending deferred balance. `milestone` and `usage` contracts are routed to the Judgment Items tab
with their stated ending balance carried into the tie-out — the skill does not decide when a
milestone is achieved or how usage should be allocated.

## Deliver

**Workbook tabs:**

1. **Workpaper Summary** — the GL beginning balance, the mechanical and judgment-item ending
   balances, the computed ending balance, the GL ending balance, the unreconciled difference, and
   the conclusion.
2. **Waterfall Schedule** — every contract's beginning deferred, additions, recognized, and ending
   deferred balance, with totals.
3. **Recognition JEs** — the proposed Dr Deferred Revenue / Cr Revenue entry per mechanically
   recognized contract with a nonzero recognized amount this period.
4. **GL Tie-Out** — the full rollforward from GL beginning to computed ending, against the GL's
   own ending balance.
5. **Exceptions** — every data exception and gate failure, with the dollar effect where one
   exists.
6. **Judgment Items** — every milestone/usage contract, its stated ending balance, and a note
   that recognition timing was not concluded by this skill.

**Then, in chat:** the tie-out result first — ties or doesn't, and by how much — then the count
and value of judgment items carried at a client-stated balance, then any data exceptions that
excluded a contract from the schedule entirely.

## What to escalate

- **A milestone/usage contract with no stated ending balance**, or one whose stated balance looks
  disconnected from its contract amount and progress to date.
- **A large or growing unreconciled difference** between the computed and GL balances — this is
  the signal that recognition is drifting from what the GL actually shows, not something to
  average away over subsequent periods.
- **A contract invoiced well after its service period began**, which this skill flags as a data
  exception rather than silently recognizing backward.
- **Contract terms found during review that imply a different method than the one stated** —
  usage-based pricing recorded as straight-line, or a straight-line subscription with an embedded
  milestone (e.g., an implementation fee). Report what the contract says and let the engagement
  decide the treatment.

Report each as a fact with the contract reference. Do not characterize intent.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.
Contract and customer data never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
