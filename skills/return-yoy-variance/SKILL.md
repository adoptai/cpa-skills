---
name: return-yoy-variance
description: Compare this year's tax return or financial statements against last year's line by line, quantify every variance, and require a causal explanation for each material movement — including the movements that are suspicious precisely because a number did NOT change. Use this whenever the user mentions comparing years, year-over-year or YoY analysis, prior-year comparison, variance analysis or variance commentary, explaining significant changes, a comparative schedule, or asks why a line moved versus last year. Also trigger on "compare this return to last year," "what changed from PY," "2-year comparison," "why is income down," "prior year vs current year," "flux versus last year," "diff these returns," or when two years of returns, trial balances, or financials are provided together. Runs fully local — no client data leaves the machine.
---

# Year-over-Year Variance Analysis

The purpose is not to list what changed. It is to reach the point where **every material
movement has a named cause, and every material non-movement has been challenged.**

That second half is what separates this from a spreadsheet subtraction. Most reviewers scan
for big changes and stop. The errors that survive review are usually the opposite pattern: a
figure that is identical to last year because nobody updated it — a depreciation amount, an
allocation percentage, a state apportionment factor, an accrual that was rolled forward
because it was easier than recomputing it. A zero variance is a finding until you have
confirmed it should be zero.

## What counts as an explanation

An explanation is causal and quantified. Direction is not an explanation.

- ✗ "Revenue increased." — restates the variance
- ✗ "Higher volume." — no magnitude, no source, not testable
- ✗ "Timing." — the single most common non-explanation in practice; it hides errors
- ✓ "Revenue up $1.24M: $900K from the Riverside contract beginning March 2025, $340K from
  the April price increase, offset by $(60)K from the Delta account lost in Q3." Ties to
  $1.24M and can be tested against the customer detail.

If the components don't sum to the variance, the explanation is incomplete and the remainder
stays visible as unexplained. Do not let the last 15% of a variance disappear into "other."

## Inputs

1. **Current year** — return, trial balance, or financials
2. **Prior year, as filed** — not the software's current state of the prior-year file, which
   may have been amended, adjusted, or corrupted since. If the prior year was amended, use
   the amended figures and say so.
3. **Optional context** — budget, entity events during the year (acquisitions, dispositions,
   new locations, contract wins and losses, headcount changes, method changes), and the
   client's own explanation of the year. Ask for this. Without it you will generate questions
   the client answers in five minutes, which is fine but slower than asking up front.

Confirm both years are on a comparable basis before comparing: same entity, same accounting
method, same period length, same chart of accounts. A short period, a method change, or an
account restructuring makes raw variances meaningless — normalize first and disclose the
normalization, or the entire schedule misleads.

## Step 1 — Map the lines

Align current-year lines to prior-year lines. Three outcomes, and the last two matter most:

- **Matched** — same line both years. Compute the variance.
- **New this year** — a line with a current-year amount and no prior-year line. Every one
  needs a cause. New lines are high-signal: a new revenue stream, a new state, a new
  expense category, a new form. Also the classic tell for a misposting into a freshly
  created account.
- **Disappeared** — a prior-year line with no current-year amount. Higher signal still.
  Either the activity genuinely stopped (name the event) or something was dropped. Dropped
  carryforwards, dropped state filings, and dropped schedules all appear here and nowhere
  else on the return.

Where the chart of accounts was restructured, map old to new explicitly and show the mapping.
Do not silently absorb a restructuring into the variances.

## Step 2 — Run the comparison

```bash
python3 scripts/yoy_compare.py \
  --current cy.csv --prior py.csv \
  --explanations explanations.csv \
  --dollar-threshold 25000 --percent-threshold 10 \
  --absolute-floor 5000 \
  --client "Acme Holdings LLC" --cy-label "TY2025" --py-label "TY2024" \
  --out "Acme Holdings - TY2025 vs TY2024 Variance.xlsx"
```

Input CSVs: `line_id`, `statement`, `caption`, `amount`, and optional `form`, `line_ref`,
`group`, `activity_changed` (`yes`/`no`/blank — see the unchanged test).

**Materiality is dual, and both tests must be satisfied to be immaterial.** A variance is
material if it exceeds the dollar threshold **or** the percent threshold, subject to an
absolute floor below which percent is ignored. This matters: a 400% swing on a $200 account
is noise, while a 3% move on an $8M revenue line is not. A single test alone produces either
an unusable list or a false sense of coverage. Set thresholds by engagement judgment and
record them — they are part of the workpaper.

Regardless of threshold, these are **always** flagged for explanation:

- **Sign flips** — income to loss, asset to liability, debit to credit balance. Direction
  changes are qualitatively material at any magnitude.
- **New and disappeared lines** — per Step 1.
- **Zero-variance lines where activity changed** — the unchanged test. Mark
  `activity_changed = yes` on any line whose underlying facts moved (revenue grew, assets
  were bought or sold, headcount changed, the entity entered a new state). If the amount
  didn't move and the activity did, the script flags it. Depreciation identical to last year
  after a year of capital additions is not a coincidence.
- **Round-number amounts** on lines that should be computed — a depreciation figure of
  exactly $45,000 or an accrual of exactly $100,000 suggests a plug, not a calculation.
- **Lines equal to prior year to the penny** on any account with ongoing activity.

## Step 3 — Relationship tests

Individual line variances miss errors that only show up as broken relationships. Run these
and explain any that break:

- Revenue vs. cost of sales — gross margin percentage by year; a margin that moves several
  points without a stated cause is either a classification error or a real business change
  worth naming
- Revenue vs. payroll, and revenue vs. headcount — revenue up sharply with flat payroll means
  one of them is wrong or something genuinely changed that should be documented
- Revenue vs. receivables, and receivables vs. days sales outstanding
- Purchases vs. payables, and inventory vs. cost of sales
- Fixed-asset additions vs. depreciation expense — the most reliable detector of a stale
  depreciation schedule
- Debt balances vs. interest expense — implied rate by year; a sharp move means missing debt,
  missing interest, or capitalized interest
- Officer compensation vs. distributions (S corps) — a shift from wages to distributions is
  the most examined S corp pattern and should be documented deliberately, not discovered
- Book income vs. taxable income — the total book-to-tax difference by year, with each
  component's movement explained
- Effective tax rate by year, reconciled — the top-down check that catches what the
  bottom-up line review missed. If the ETR moved and no line variance explains it, keep
  looking

## Step 4 — Require explanations

The script writes `unexplained.csv` listing every flagged variance with no explanation. Fill
in `cause`, `components` (which must sum to the variance), `evidence`, and `owner`, then
re-run with `--explanations`. The script reports the **explanation coverage rate** — the
percentage of flagged dollars explained — and refuses to mark the analysis complete below
100% of flagged items. Partial coverage is reported honestly rather than smoothed over.

Where a variance is explained by a **client-provided reason you have not corroborated**, mark
`evidence = client representation`. That is a legitimate but weaker basis and the workpaper
should distinguish it from an explanation tied to a document. Reviewers need to know which
is which.

## Step 5 — Deliver

**Workbook tabs:**

1. **Summary** — thresholds used, lines compared, count and dollar value of material
   variances, explanation coverage rate, top ten variances by absolute dollar, and the
   always-flag exceptions. Read-in-30-seconds page.
2. **Variance Schedule** — every line: caption, form/line reference, prior year, current
   year, dollar variance, percent variance, materiality verdict, flag reason, cause,
   components, evidence, owner. Sorted by absolute variance descending.
3. **Always-Flag Exceptions** — sign flips, new lines, disappeared lines, unchanged-but-
   activity-changed, and round-number suspects, each with its explanation status.
4. **Relationship Tests** — the ratios above, both years, with the movement and its
   explanation.
5. **Unexplained** — what remains. An empty tab is the goal; a missing tab is not acceptable.

**Then, in chat:** the three or four variances that actually matter, in plain language, with
their causes — followed by what remains unexplained and who owns it. Lead with the
unexplained items if there are any. A partner reading this wants to know what to ask the
client, not to re-read the schedule.

## When variances suggest something worse

Some patterns are worth escalating rather than explaining:

- Revenue up with cash flow flat or down
- Margin improving while the business reports no operational change
- Expenses that disappear in a year when the related activity continued
- Round-dollar variances that exactly offset each other across two accounts
- A prior-year figure that does not agree to the prior-year return as filed — this means
  someone changed a closed year, and it needs to be reconciled before any variance is
  meaningful
- Repeated year-over-year "timing" explanations for the same line, which is the signature of
  a systemic error rolling forward rather than a genuine timing difference

Raise these as observations with the facts you have. Do not characterize intent.

## Security posture

Fully local. No client financial or tax data leaves the machine. Inputs read-only.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
