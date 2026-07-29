---
name: expense-policy-testing
description: Test an expense, travel, or corporate card population against the client's own written policy — missing approvals, amounts over limit, self-approval, split transactions engineered to stay under an approval threshold, duplicate submissions, prohibited categories, missing receipts, and weekend or personal-pattern spend — with the population proven against the general ledger and every transaction given a disposition. Use this whenever the user mentions expense testing, expense report review, T&E audit, travel and entertainment, corporate card or p-card review, expense policy compliance, missing receipts or approvals, duplicate expense claims, or employee reimbursement testing. Also trigger on "test our expense reports," "check the corporate card spend," "who is breaking expense policy," "find duplicate claims," "review T&E," or when an expense population and a policy are provided together. Runs fully local — no employee or expense data leaves the machine.
---

# Expense and T&E Policy Testing

Expense testing is worth doing because the findings are behavioural rather than random. One
person submits the same receipt twice; one manager approves everything without looking; one
department has discovered that two transactions of $2,400 avoid the approval that one of $4,800
would trigger. **So the output is organised by person and by approver before it is organised by
transaction**, because that is the order in which anything gets fixed.

It is also the area where a client's own policy is the standard. There is no external threshold to
apply — if the policy says receipts are required above a stated amount, that is the test. Which
means the first question is whether a written policy exists at all.

## The gate

No clean workpaper unless:

1. **The population ties to the GL.** You supply the expense total per the general ledger for the
   accounts in scope. A filtered card extract proves nothing about the expense line.
2. **Every transaction receives a disposition** — compliant, or a named exception. Matched plus
   exceptions must equal the population in count and in value.
3. **The policy rules are stated, not assumed.** Approval thresholds, receipt thresholds,
   prohibited categories, and per-diem limits come from the client's written policy and are
   recorded on the workpaper.
4. **No transaction is tested twice** under a rule that would double-count its value.

If there is no written policy, that is the headline finding. Run with zero thresholds, say
explicitly that you did, and report the absence of a policy as a control matter — an expense
programme with no written policy has no standard to test against and no basis for discipline.

## Inputs

1. **Expense population** — transaction date, employee, amount, category, merchant, description,
   approver, receipt indicator, report or card reference, cost centre, and payment method.
2. **GL total** for the expense accounts in scope, same period.
3. **The written policy** — approval thresholds by amount and by role, receipt requirements,
   prohibited or restricted categories, per-diem or per-meal caps, advance-approval requirements.
4. **Employee and approver roster**, with reporting lines if available. Reporting lines enable
   the self-approval and circular-approval tests, which are among the most valuable.
5. **Prior-period findings**, if any. A repeat finding by the same person is a different
   conversation from a first occurrence.

## Step 1 — Test

```bash
python3 scripts/expense_test.py \
  --expenses expenses.csv \
  --gl-total 1284550.00 \
  --policy policy.csv \
  --roster roster.csv \
  --period "FY2025" --client "Brightline Media Group" \
  --out "Brightline - FY2025 Expense Policy Testing.xlsx"
```

Tests applied, each producing a named exception type:

**Authorisation**

- *Missing approval* — no approver recorded where the policy requires one
- *Self-approved* — approver equals the claimant. A control finding at any amount
- *Over approval authority* — approved by someone whose stated limit is below the amount
- *Circular approval* — two people approving each other's claims, where reporting lines are supplied

**Threshold behaviour**

- *Over limit* — amount exceeds a policy cap for its category
- *Threshold-adjacent* — amount falls just below an approval threshold. Report the pattern, not
  the single instance; one is a coincidence, a habit is not
- *Split transaction* — several claims by the same person, same day, same category, together
  exceeding a threshold each stays under. **This is the highest-signal test in the set**, because
  it implies knowledge of the control

**Documentation**

- *Missing receipt* where the policy requires one for the amount
- *Missing or generic business purpose* — `client meeting`, `business`, `misc`, `per policy`

**Duplicates**

- Same employee, amount, and date
- Same employee and amount within a short window
- Same merchant, amount, and date across two different employees — one dinner claimed twice

**Category and pattern**

- *Prohibited category* per the policy
- *Weekend and holiday spend* on categories that are not travel
- *Personal-pattern merchants* where the policy prohibits them
- *Round-number claims* at or just under a receipt threshold
- *Claims after termination*, where roster dates are supplied

## Step 2 — Read it by person, not by transaction

The workbook leads with two rollups, because the flat exception list is the least useful view:

- **By employee** — exception count, value, and the dominant type. A single person with fifteen
  missing receipts is a training conversation. A single person with fifteen split transactions is
  not.
- **By approver** — volume approved, exception rate, and self-approvals. **An approver whose
  exception rate is far above peers is approving without reading**, and that is a more valuable
  finding than any individual claim, because it explains the others.

Then look at concentration: if 80% of exception value sits with three people, the programme does
not have a policy problem, it has three problems.

## Step 3 — Deliver

**Workbook tabs:**

1. **Summary** — population proof, policy rules applied, exception counts and values by type,
   concentration, recovery opportunity, and the control observations.
2. **Population Proof** — compliant plus exceptions equals the population, in count and value,
   plus the tie to the GL.
3. **By Employee** and **By Approver** — the rollups above.
4. **Exception Detail** — every exception with the transaction, the rule breached, the policy
   reference, and columns for investigation and disposition.
5. **Splits and Duplicates** — grouped, with the pattern that caught each and the amount at risk.
6. **Compliant Detail** — for re-performance.

**Then, in chat:** whether the population ties, exception count and value by type, the two or
three people or approvers driving it, and the recoverable amount. Lead with the approver pattern
if there is one — it is the cause rather than a symptom.

## What to escalate rather than list

- **Split transactions by the same person**, repeatedly. Implies knowledge of the threshold.
- **An approver with a very high exception rate**, or one who approves their own claims.
- **Claims submitted after a termination date.**
- **The same receipt claimed by two employees**, or the same merchant and amount on the same day
  across two claimants.
- **A reimbursement paid to a bank account matching a vendor** in the AP master, where that data
  is available.
- **Cash advances that were never reconciled** to actual expenses.
- **An expense programme with no written policy**, or one last updated years ago.
- **Personal spend patterns** — regular grocery, fuel near home, or subscription charges on a
  corporate card.

State these as facts with amounts and references. **Do not characterise intent.** An expense
finding lands on an individual rather than on a process, so the language matters more here than
almost anywhere else in the engagement: describe what the record shows, quantify it, and let
management and HR handle it through their own process.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.
Employee-level data is sensitive and personally identifying; it never leaves the machine, and
output uses employee identifiers and names only as supplied.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
