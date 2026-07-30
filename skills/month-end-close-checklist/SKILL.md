---
name: month-end-close-checklist
description: Build a month-end or period-end close checklist derived from the client's own trial balance rather than a generic template — every account with activity is either reconciled with named support and an owner, waived with a documented reason, or open, and the close is not complete until nothing is left open. Use this whenever the user mentions the month-end close, period-end close, year-end close, close checklist, closing the books, what still needs reconciling, close status, a close that is running late, or which accounts have not been reconciled. Also trigger on "build a close checklist," "where are we on the close," "what's left to close the month," "month end close process," "which accounts still need reconciling," or when a trial balance is provided at period end. Runs fully local — no client data leaves the machine.
---

# Month-End Close Checklist

Every generic close checklist has the same problem: **it is a list of tasks somebody else's business
needed.** It tells you to reconcile the bank and review accruals, which you knew, and it says nothing
about the seven accounts on *your* trial balance that had activity this month and nobody has looked
at.

So this builds the checklist **from the trial balance**. Every account with movement generates a
task with an owner and a status. Accounts you deliberately do not reconcile get waived with a
reason, on the record. And the close is not complete until every account is in one of those two
states — which is a completeness test, not a to-do list.

That distinction matters because the usual close failure is not a forgotten task. It is an account
that nobody ever considered, because it was not on the template.

## The gate

The close is not reportable as complete unless:

1. **Every account with activity is accounted for** — reconciled, waived with a reason, or open.
   Reconciled + waived + open must equal the population of accounts with movement.
2. **Nothing above materiality is waived without a reason.** A waiver is a professional judgement
   and it is recorded as one. A blank reason is not a waiver, it is an omission.
3. **Every open item has an owner and a due date.** An open item with no owner does not close.
4. **Every reconciled account names its support** — which workpaper, statement, or schedule.
   "Reconciled" with no reference is an assertion, not a reconciliation.
5. **The trial balance balances.** If debits do not equal credits, the close has not started. Run
   `trial-balance-integrity` first.

## Inputs

1. **The trial balance** for the period, with prior-period balances so movement can be identified.
   Same schema `trial-balance-integrity` uses, so the two chain directly.
2. **Materiality** for the close — the threshold below which a waiver needs less scrutiny.
3. **The prior close checklist**, if one exists. Carrying forward owners and statuses saves time and
   makes a recurring open item visible.
4. **The team roster** for owner assignment.
5. **The target close date**, so lateness is measurable rather than felt.

## Step 1 — Derive the checklist

```bash
python3 scripts/close_checklist.py \
  --trial-balance tb.csv \
  --materiality 25000 \
  --prior-checklist prior_close.csv \
  --period "2025-11" --target-close-date 2025-12-05 \
  --client "Halstead Regional Foods" \
  --out "Halstead - 2025-11 Close Checklist.xlsx"
```

The script generates a task for every account with movement, and — this is the part a template
cannot do — **assigns the right procedure based on what the account actually is.** A cash account
gets a bank reconciliation. A receivable account gets an aging tie-out. A fixed asset account with
additions gets a depreciation rollforward. Where this repository has a skill for the procedure, the
task names it:

| What the trial balance contains | Task generated | Skill to run |
|---|---|---|
| Cash accounts with movement | Bank reconciliation to the statement | `bank-rec-to-gl` |
| A reconciliation prepared by someone else | Review it rather than re-perform it | `bank-rec-review` |
| Trade receivables | Aging tied to the control account | `ar-aging-tie-out` |
| Trade payables, accrued liabilities | Search for unrecorded liabilities, cutoff | `cutoff-and-unrecorded-liabilities` |
| Fixed assets, accumulated depreciation | Rollforward and depreciation tie-out | `depreciation-tie-out` |
| Wage and payroll tax accounts | Payroll register to filings to GL | `payroll-tax-reconciliation` |
| Sales tax liability | Returns to ledger, liability rollforward | `sales-tax-reconciliation` |
| Revenue accounts with unusual movement | Cutoff and occurrence testing | `revenue-trace-to-source` |
| Expense and card accounts | Policy testing on the population | `expense-policy-testing` |
| Any account, for period-over-period review | Variance with a quantified explanation | `return-yoy-variance` |
| A complete balance sheet and income statement | Statement of cash flows, proved | `cash-flow-tieout` |
| Manual journal entry activity | Full-population anomaly scan | `journal-entry-anomaly-scan` |

So the checklist is not only a status list — it tells you which procedure each account needs and
where to get it.

## Step 2 — Work the statuses

Four statuses, and the distinctions are deliberate:

- **`reconciled`** — support obtained and named. Requires a reference.
- **`waived`** — deliberately not reconciled this period, with a reason. Legitimate for immaterial,
  dormant, or annually-reconciled accounts. Requires a reason.
- **`open`** — still to be done. Requires an owner and a due date.
- **`blocked`** — cannot proceed until something else happens. Requires the blocker to be named,
  because a blocked item with no named blocker is an open item wearing a disguise.

**On waivers.** A waiver is not a shortcut, it is a documented judgement that the account does not
warrant a reconciliation this period. Immaterial and dormant accounts genuinely do not. But a waiver
without a reason is how an account disappears from the close permanently — waived in month one
because it was small, still waived in month eleven when it is not. The script flags any account
waived above materiality, and any account waived in three or more consecutive periods.

## Step 3 — Read the close, not the list

The workbook leads with the things that make a close late or wrong:

- **Accounts with movement and no task history** — new accounts, or accounts that appeared without
  anyone noticing. These are where misposting lands.
- **Recurring open items** — the same account open at the end of three consecutive closes. That is
  not a busy month, it is a process gap.
- **Accounts waived repeatedly**, per above.
- **Owner concentration** — if one person owns most of the open items, the close is
  single-threaded and the date is at risk regardless of effort.
- **Days past the target close date**, with the open items that are driving it.
- **Suspense, clearing, and conversion accounts with a balance.** These are meant to be empty at
  close. A balance means something was parked and the close moved on without it.

## Step 4 — Deliver

**Workbook tabs:**

1. **Close Status** — the gate, counts and dollars by status, days against target, the accounts
   driving lateness, and the escalations. The page a controller reads each morning.
2. **Checklist** — every account with movement: balance, movement, status, procedure required, skill
   to run, support reference, owner, due date, and notes. This is the working document.
3. **Open Items** — sorted by materiality then age, with owners.
4. **Waivers** — every waived account with its reason and how many consecutive periods it has been
   waived.
5. **Coverage Proof** — reconciled + waived + open + blocked equals the population of accounts with
   movement, in count and in dollars.
6. **Recurring Issues** — items open or waived across consecutive periods.
7. **Next Period Carry-Forward** — a pre-populated checklist for the following period, with owners
   retained, so the next close starts from evidence rather than a blank template.

**Then, in chat:** whether the close is complete, the count and value of open items, who owns them,
days against target, and anything waived that should not be. Lead with the open items above
materiality — that is the close.

## What to escalate

- **An account with movement and no owner.** Nobody is looking at it.
- **A suspense, clearing, or conversion account with a balance at close.**
- **The same account open at three consecutive closes.**
- **An account waived above materiality**, and any account waived repeatedly.
- **A reconciliation marked complete with no support reference.**
- **A close completed while items remain open**, which means the financial statements were issued on
  an incomplete basis. Whether that is acceptable is a judgement for the controller, but it should
  be a decision rather than an accident.
- **Journal entries posted after the close was declared complete.** Worth asking for the entry list
  filtered by posting date, and running `journal-entry-anomaly-scan` on it.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
