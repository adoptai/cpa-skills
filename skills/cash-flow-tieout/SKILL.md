---
name: cash-flow-tieout
description: Build or verify an indirect-method statement of cash flows and prove it — the net change in cash must equal the movement in the balance sheet cash accounts, each section must tie to the underlying balance sheet changes, and every non-cash and reclassifying adjustment must be named. Use this whenever the user mentions the cash flow statement, statement of cash flows, indirect method, operating investing and financing activities, working capital changes, a cash flow that does not tie or does not foot, free cash flow build-up, or monthly cash flow analysis. Also trigger on "build the cash flow statement," "why doesn't our cash flow tie," "reconcile cash flow to the balance sheet," "cash flow by month," "check the cash flow statement," or when a balance sheet and income statement for two periods are provided together. Runs fully local — no client financial data leaves the machine.
---

# Cash Flow Statement Tie-Out

The statement of cash flows is the only primary statement that must reconcile to another one by
construction. Net income plus adjustments plus working capital movements, across all three
sections, has to equal the change in cash on the balance sheet. There is no judgment in that
requirement — it either foots or it does not.

Which is exactly why it is so often wrong. It is usually prepared last, under time pressure, by
plugging whichever line nobody will question. The plug hides inside "changes in other assets and
liabilities" or "other operating activities," and it survives review because reviewers check that
the bottom line agrees to cash rather than checking that each line agrees to the balance sheet.

**So this skill derives each working capital movement from the balance sheet itself** and compares
it to what the statement claims, rather than accepting the statement's own numbers.

## The gate

No clean statement unless:

1. **Net change in cash equals the balance sheet movement.** Operating + investing + financing
   (+ FX effect) = closing cash − opening cash, per the balance sheet, to the cent.
2. **Every working capital line agrees to the balance sheet movement** for that account, or carries
   a named non-cash adjustment explaining the difference.
3. **The balance sheet balances in both periods.** Assets = liabilities + equity. A cash flow
   built on an unbalanced balance sheet is meaningless, so this is checked first.
4. **There is no unexplained residual.** If a difference exists it appears as an explicit
   "UNEXPLAINED — DO NOT PLUG" line rather than being absorbed into an "other" caption.

## Standing rule on the plug

Every difference must be attributed to a named cause. The legitimate reasons a working capital
movement differs from the raw balance sheet change are finite and knowable:

- **Non-cash additions and disposals** — assets acquired under a lease, an asset received in a
  non-monetary exchange, capitalised interest
- **Acquisitions and disposals of businesses**, which move balance sheet accounts without an
  operating cash flow
- **Foreign currency translation** on balance sheet accounts of foreign operations, which belongs
  in the FX effect line and not in working capital
- **Reclassifications between captions**, including current/non-current reclassification
- **Non-cash write-offs** — bad debt written off against the allowance, inventory written down
- **Accrued but unpaid amounts** moving between accrual accounts
- **Stock compensation, deferred tax, and other non-cash charges** in net income

If a difference cannot be placed in one of those buckets, it is not explained. Report it. The
script will not let it disappear into a caption.

## Inputs

1. **Balance sheet, both periods** — account level or at least caption level, with the cash and
   cash-equivalent accounts identified.
2. **Income statement for the period.**
3. **Known non-cash items** — depreciation, amortisation, stock compensation, deferred tax,
   gains and losses on disposal, impairment, bad debt expense.
4. **Investing and financing detail** — capital expenditure, proceeds from disposals, debt drawn
   and repaid, equity issued, dividends and distributions paid.
5. **The client's own cash flow statement**, if one exists and you are verifying rather than
   building. The comparison against your derived figures is the point of the exercise.

## Step 1 — Prove the balance sheet, then derive

```bash
python3 scripts/cash_flow.py \
  --balance-sheet bs.csv \
  --income-statement is.csv \
  --noncash noncash.csv \
  --investing-financing invfin.csv \
  --adjustments adjustments.csv \
  --client "Latham Industrial Group" --period "FY2025" \
  --out "Latham - FY2025 Cash Flow Tie-Out.xlsx"
```

Six tests:

- **Test 1 — Balance sheet balances in both periods.** Runs before anything else. An unbalanced
  balance sheet makes every downstream figure meaningless.
- **Test 2 — Cash movement identified.** Opening and closing cash per the accounts flagged as
  cash and equivalents, and the resulting change.
- **Test 3 — Working capital derived from the balance sheet.** Each operating asset and liability
  account's movement computed and signed correctly. This is where the derivation replaces trust.
- **Test 4 — The statement foots.** Operating + investing + financing + FX = change in cash.
- **Test 5 — Every adjustment named.** Each difference between a derived movement and the amount
  presented must be attributed to a cause in the list above.
- **Test 6 — Comparison to the client's statement**, where supplied, line by line.

## Step 2 — Get the signs right

Sign errors are the second most common failure after plugging, and they are easy to get wrong
because the logic inverts between assets and liabilities:

| Account moves | Cash effect |
|---|---|
| Operating **asset** increases (AR, inventory, prepaid) | **Outflow** — cash was consumed |
| Operating **asset** decreases | **Inflow** |
| Operating **liability** increases (AP, accruals, deferred revenue) | **Inflow** — payment deferred |
| Operating **liability** decreases | **Outflow** |

The script applies this from each account's declared type, so a misclassified account produces a
sign error you can trace rather than a mystery. **Check the account types before reading the
result** — an account typed as a liability when it is a contra-asset will flip its contribution
and the statement will still foot, because the error is inside the section.

A doubled sign error is the nastiest case: two accounts wrong in opposite directions leave the
statement footing perfectly while both sections are misstated. Test 3's line-by-line derivation is
what surfaces it.

## Step 3 — Read the classification questions

Section classification is judgment in a handful of recurring places, and the script flags them
rather than deciding:

- **Interest paid** — operating under most frameworks, but presentation varies and disclosure may
  be required
- **Dividends received** — operating or investing depending on framework and policy
- **Capitalised interest** — investing, though the expense ran through operating
- **Overdrafts** — financing, or part of cash and equivalents, depending on the arrangement
- **Bank debt with a revolving facility** — gross versus net presentation of draws and repayments
- **Restricted cash** — whether it belongs in cash and equivalents at all
- **Book overdrafts versus bank overdrafts** — these are different and are treated differently

These are flagged as judgment items with the amounts involved, for the person signing.

## Step 4 — Deliver

**Workbook tabs:**

1. **Cash Flow Statement** — the derived statement in proper form, with the foot to the change in
   cash shown, and the unexplained line reading `0.00`. The signable page.
2. **Proof** — the six tests with amounts and differences.
3. **Working Capital Derivation** — every operating account: opening balance, closing balance,
   raw movement, named adjustments, and the cash flow effect with its sign. The tab that catches
   the plug.
4. **Balance Sheet Movement** — every account, both periods, movement, and the section it feeds.
   Nothing is allowed to feed no section.
5. **Comparison** — where the client's statement was supplied, line by line against the derived
   figures, with differences.
6. **Judgment Items** — the classification questions above, with amounts, deliberately not
   concluded.
7. **Monthly Analysis** — where monthly balance sheets are supplied, the cash flow by month with
   unusual movements flagged.

**Then, in chat:** whether it foots, the change in cash, the three sections, and any line where
the derived figure disagrees with what was presented. Lead with the disagreement — that is the
finding, not the total.

## What to escalate

- **A statement that foots only because of an "other" line.** Report the size of that line
  relative to the sections. A material unexplained "other" is a plug regardless of its label.
- **Operating cash flow persistently below net income** while revenue grows — the classic
  earnings-quality signal. Report the relationship across periods; do not characterise it.
- **Capital expenditure with no corresponding fixed asset movement**, or vice versa.
- **Debt movements that do not agree to the debt schedule** or to interest expense at any
  plausible rate.
- **An account that feeds no section.** Every balance sheet movement goes somewhere. An account
  excluded from all three sections is the definition of a hidden plug, and the script refuses to
  let it be silent.
- **Restricted cash included in cash and equivalents** without disclosure.
- **A large FX effect** in an entity with no foreign operations.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
