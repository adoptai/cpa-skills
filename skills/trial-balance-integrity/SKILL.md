---
name: trial-balance-integrity
description: Test a trial balance and chart of accounts before anything is built on them — proving debits equal credits, that every account maps to a financial statement line with none left unmapped, and flagging duplicate or near-duplicate accounts, unused accounts, accounts with a balance in the wrong direction, and misclassifications. Use this whenever the user mentions the trial balance, chart of accounts, a TB that does not balance, mapping accounts to financial statements, cleaning up or restructuring a chart of accounts, converting to new accounting software, duplicate or unused accounts, or an account with an unexpected balance. Also trigger on "does our trial balance balance," "review the chart of accounts," "map the TB," "clean up our accounts," "convert our books," "which accounts are unused," or when a trial balance is provided at the start of an engagement. Runs fully local — no client data leaves the machine.
---

# Trial Balance and Chart of Accounts Integrity

This is the first thing to run on a new client, a cleanup engagement, or a software conversion,
because **everything else in this repository assumes the trial balance is sound.** A
reconciliation, a variance analysis, a cash flow statement, and a tax return are all built on it.
If two accounts are duplicates, or an account maps to nothing, the error propagates into every
downstream workpaper and is much harder to find there.

Debits equalling credits is necessary and tells you almost nothing. Every accounting system
enforces it. What it does not enforce:

- **Two accounts for the same thing**, so a balance is split and neither number is right
- **An account mapped to no financial statement line**, which silently drops out of the
  statements while the trial balance still balances
- **An asset with a credit balance**, or a liability with a debit balance — often real, sometimes
  a misposting, always worth a reason
- **Accounts nobody has used in years**, which are where misposting hides because nobody reviews
  them

## The gate

No clean workpaper unless:

1. **Debits equal credits**, to the cent.
2. **Every account maps to a financial statement line.** An unmapped account is reported, never
   defaulted to a caption. A mapping that silently drops an account is the failure this test
   exists to prevent.
3. **The mapped statement subtotals foot back to the trial balance.** Assets, liabilities, equity,
   revenue and expense derived from the mapping must reproduce the trial balance totals.
4. **Every account has a type.** An untyped account cannot be tested for sign or classification.

## Inputs

1. **The trial balance** — account number, account name, debit, credit, and ideally the account
   type. Prior-year comparative balances if available.
2. **The mapping** from account to financial statement line, if one exists. If none exists,
   producing one is the deliverable, and the skill will tell you which accounts you have not yet
   placed.
3. **Activity counts or dates of last use**, if the system provides them. This makes the
   unused-account test meaningful rather than a guess from a zero balance.
4. **The target format**, if you are converting — the account structure the new system expects.

## Step 1 — Test

```bash
python3 scripts/tb_integrity.py \
  --trial-balance tb.csv \
  --mapping mapping.csv \
  --prior-year tb_prior.csv \
  --client "Halstead Regional Foods" --period "12/31/2025" \
  --out "Halstead - 2025 Trial Balance Integrity.xlsx"
```

Eight tests:

- **Test 1 — Debits equal credits.** The necessary condition, checked first and quickly.
- **Test 2 — Every account mapped.** Unmapped accounts listed individually with their balances.
- **Test 3 — Mapping foots to the trial balance.** Statement subtotals derived from the mapping
  reproduce the trial balance. This catches a mapping that double-counts or drops.
- **Test 4 — Accounting equation.** Assets = liabilities + equity + (revenue − expense), from the
  account types. A break here means accounts are typed wrongly even though the TB balances.
- **Test 5 — Balance direction.** Accounts whose balance sits opposite to their type, each
  requiring a reason. Contra accounts are expected to be opposite and are excluded where declared.
- **Test 6 — Duplicate and near-duplicate accounts.** Same name, names differing only by
  punctuation, spacing, case, or a trailing number, and accounts with identical names under
  different numbers.
- **Test 7 — Unused and dormant accounts.** Zero balance with no activity, or no activity since a
  prior period.
- **Test 8 — Structure and numbering.** Accounts outside the ranges implied by their type, gaps in
  the numbering, and inconsistent numbering depth.

## Step 2 — Read the duplicates carefully

Duplicates are the highest-value finding because they corrupt every analysis silently.
`6100 Repairs & Maintenance` and `6105 Repairs and Maintenance` both have balances, both look
reasonable, the trial balance balances, and every variance analysis on either one is wrong. The
script matches on a normalised name — punctuation, case, spacing, and `&`/`and` collapsed — which
is what catches these.

Genuine near-duplicates also exist and are fine: `6100 Repairs - Buildings` and
`6110 Repairs - Vehicles` are legitimately separate. **The script flags candidates and does not
merge anything.** Which duplicates are real is a judgment about how the business wants to see its
results, and merging accounts changes comparatives, so it is a decision, not a cleanup.

## Step 3 — What the sign tests actually mean

An account with a balance opposite to its type is not automatically wrong. Common legitimate cases:

- **Accumulated depreciation and the allowance for doubtful accounts** carry credit balances and
  are assets by classification. Declare them as contra accounts.
- **A bank account in overdraft** legitimately carries a credit balance, and it may need
  reclassifying to a liability for presentation.
- **Accounts payable with a debit balance** usually means a payment recorded against nothing, or a
  credit misapplied. Worth investigating.
- **A revenue account with a debit balance** usually means returns or credits exceeded sales for
  the period, or revenue was reversed into the wrong account.
- **Retained earnings with a debit balance** is an accumulated deficit — normal for many entities.

The script reports the condition and the amount; deciding which are real is judgment, and the
workbook leaves a column for the reason.

## Step 4 — Deliver

**Workbook tabs:**

1. **Integrity Summary** — the eight tests, totals, the accounting equation, and counts by
   finding type. The signable page.
2. **Trial Balance** — every account with debit, credit, net, type, mapped statement line, prior
   year where supplied, movement, and any flag.
3. **Unmapped Accounts** — accounts with no statement line, with balances. Empty is the goal.
4. **Mapping Proof** — statement subtotals derived from the mapping against the trial balance
   totals.
5. **Duplicate Candidates** — grouped by normalised name, with both balances shown, and a column
   for the merge decision.
6. **Sign Exceptions** — accounts with a balance opposite to their type, with a reason column.
7. **Unused and Dormant** — zero-balance and no-activity accounts, with a
   keep/close recommendation column.
8. **Conversion Mapping** — where a target structure is supplied, old account to new account, with
   unmapped accounts blocking the conversion.

**Then, in chat:** whether it balances, how many accounts are unmapped, the duplicate candidates,
and the sign exceptions worth investigating. Lead with unmapped accounts and duplicates — those
are the two that corrupt everything downstream.

## What to escalate

- **An unmapped account with a material balance.** It will disappear from the financial statements
  while the trial balance still balances.
- **Duplicate accounts both carrying balances**, particularly in accounts used for analysis.
- **A suspense, clearing, or "other" account with a material balance** at period end. These are
  meant to be empty; a balance means something was parked and forgotten.
- **Accounts payable or receivable control accounts with a balance opposite to expectation.**
- **A new account created late in the period** with a material balance and few transactions.
- **Accounts with names that describe an adjustment** rather than a category — `Plug`,
  `Difference`, `To balance`, `Conversion`, `Misc`.
- **A conversion balance account still holding a balance** long after the conversion.
- **Retained earnings that do not agree to prior-year retained earnings plus net income less
  distributions.** The script computes this where prior-year figures are supplied, and it is one
  of the fastest ways to detect an entry posted directly to equity.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
