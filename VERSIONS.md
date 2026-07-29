# Versions

Version history for the skills in this repository. Bump the version when a skill's behavior
changes in a way that affects its output, and add a line here.

Repository version: **1.3.0**

| Skill | Version | Last changed | Notes |
|---|---|---|---|
| `ar-aging-tie-out` | 1.0.0 | 2026-07-29 | Added in 1.2.0 |
| `bank-rec-review` | 1.0.0 | 2026-07-29 | Added in 1.3.0 |
| `cash-flow-tieout` | 1.0.0 | 2026-07-29 | Added in 1.3.0 |
| `cutoff-and-unrecorded-liabilities` | 1.0.0 | 2026-07-29 | Added in 1.3.0 |
| `expense-policy-testing` | 1.0.0 | 2026-07-29 | Added in 1.3.0 |
| `trial-balance-integrity` | 1.0.0 | 2026-07-29 | Added in 1.3.0 |
| `audit-sampling` | 1.0.0 | 2026-07-29 | Added in 1.1.0 |
| `bank-statement-to-excel` | 1.0.0 | 2026-07-28 | Initial release |
| `bank-rec-to-gl` | 1.0.0 | 2026-07-28 | Initial release |
| `depreciation-tie-out` | 1.0.0 | 2026-07-29 | Added in 1.1.0 |
| `journal-entry-anomaly-scan` | 1.0.0 | 2026-07-28 | Initial release |
| `k1-extract-summarize` | 1.0.0 | 2026-07-29 | Added in 1.1.0 |
| `payroll-tax-reconciliation` | 1.0.0 | 2026-07-29 | Added in 1.1.0 |
| `return-yoy-variance` | 1.0.0 | 2026-07-28 | Initial release |
| `revenue-trace-to-source` | 1.0.0 | 2026-07-29 | Added in 1.2.0 |
| `sales-tax-reconciliation` | 1.0.0 | 2026-07-29 | Added in 1.2.0 |
| `tax-return-review` | 1.0.0 | 2026-07-28 | Initial release |
| `three-way-match` | 1.0.0 | 2026-07-29 | Added in 1.1.0 |

---

## 1.3.0 — 2026-07-29

Five skills added, bringing the repository to eighteen. No changes to the existing thirteen.

**`cash-flow-tieout`** — Indirect-method statement of cash flows, with working capital movements
**derived from the balance sheet** rather than taken from the statement, because the usual failure
is a plug hidden in an "other" caption that survives review when only the bottom line is checked.
Contra-asset accounts are deliberately excluded from working capital: their movement *is* the
non-cash charge added back separately, so deriving them too would double-count it. Instead each
non-operating movement is cross-checked — contra accounts against the non-cash charge, PP&E against
capital expenditure, debt against draws and repayments, equity against net income and
distributions. In testing, a statement that footed perfectly with a 60,000 plug hidden in "Other
operating activities" was caught twice: once by the PP&E reconciliation and once by the
material-"other" test.

**`expense-policy-testing`** — Tests an expense, T&E, or corporate card population against the
client's own written policy across fourteen exception types, including split transactions
engineered to stay under an approval threshold. Output is organised by employee and by approver
before transaction, because the findings are behavioural: an approver whose exception rate is far
above peers is approving without reading, and that explains the other findings. If no written
policy exists, that is the headline finding.

**`cutoff-and-unrecorded-liabilities`** — Search for unrecorded liabilities and cutoff testing.
This procedure works on what was *not* recorded, so it cannot sample from its own population; the
gates reflect that. Classification depends on the **service date** — when goods transferred or
services were performed — not the invoice or payment date, and a missing service date is treated as
unclassified rather than guessed. Includes a seventeen-item accrual completeness checklist where
"not applicable" is a valid answer with a reason and a blank is not.

**`bank-rec-review`** — Auditing a reconciliation someone else prepared, which is a different
procedure from preparing one. Recomputes from item detail rather than the preparer's subtotals, and
tests **subsequent clearance** — the one procedure that catches a fabricated item, a stale item, and
a plug alike, because a real reconciling item appears in the next period's bank statement and a
fabricated one never does.

**`trial-balance-integrity`** — Run this first on a new client, a cleanup, or a conversion, because
everything else assumes the trial balance is sound. Debits equalling credits is enforced by every
accounting system and proves almost nothing; this tests that every account maps to a statement line,
that the mapping foots, and that duplicates like `Repairs & Maintenance` and `Repairs and
Maintenance` are surfaced. Nothing is merged automatically — merging changes comparatives and is a
decision.

---

## 1.2.0 — 2026-07-29

Three skills added, bringing the repository to thirteen. No changes to the existing ten.

**`sales-tax-reconciliation`** — Reconciles filed returns to the sales ledger and the GL, with a
liability rollforward that isolates collected-but-unremitted tax. Asserts no rate and no nexus
threshold: effective rates are *derived* per jurisdiction per period as tax reported ÷ taxable
sales and compared across periods, and the output is a nexus **screening** schedule with an
instruction to confirm each state's current threshold. Marketplace-facilitated sales are tracked
separately throughout, because in most states they count toward the facilitator's obligation
rather than the seller's, and aggregating them produces registrations the client does not need.
A jurisdiction with activity and no return cannot be explained away with a reconciling item.

**`ar-aging-tie-out`** — Ties the aging to the GL control account, then **recomputes every bucket
from the invoice date** rather than trusting the report. The total can agree while the aging is
wrong, and the aging is what drives the allowance and the credit decisions. Grosses up netted
credit balances, tests cutoff, and where subsequent receipts are supplied distinguishes a
collection-timing issue from a valuation one.

**`revenue-trace-to-source`** — Traces revenue from the GL through invoice, contract, delivery and
cash. States the direction and assertion on the face of the workpaper, because GL→source tests
occurrence and **cannot** detect unrecorded revenue — if completeness is the concern this is the
wrong procedure. Refuses to record an item as tested unless every required evidence link exists,
reports an evidence-coverage matrix so a sample vouched to invoices alone is obvious, and matches
subsequent-period credit memos against tested revenue.

---

## 1.1.0 — 2026-07-29

Five skills added. No changes to the existing five, so no workpaper produced under 1.0.0
disagrees with a re-run.

**`k1-extract-summarize`** — Extracts Schedule K-1 data box by box, preserving the code alongside
every amount, and foots the aggregate K-1s back to the entity's Schedule K. Detects a missing K-1
two ways: ownership percentages that do not total 100.000%, and a recipient list that does not
reconcile in both directions. State schedules are bucketed separately from federal amounts,
because state K-1s reuse the federal box numbers and would otherwise double-count into the
footing test. Full TINs are rejected on input.

**`payroll-tax-reconciliation`** — Four-way tie across the payroll register, the four Forms 941,
W-2/W-3 totals, and the GL, over nine tests including a liability rollforward that isolates
undeposited trust-fund tax. Asserts no wage base or tax rate: the Social Security wage base is
*inferred* from the register as the cap the payroll system actually applied, and effective rates
are computed from the filed figures so rate drift between quarters is detectable without knowing
the correct rate. A `--rounding-tolerance` defaults to zero and reports anything it absorbs.

**`audit-sampling`** — MUS/PPS, stratified, and random attribute selection. The population must
tie to a stated control total, a seed is mandatory so the selection is re-performable, and
negative balances require an explicit treatment decision rather than being silently dropped. The
reliability factor is computed as `-ln(risk)`; expansion factors and attribute sample sizes are
firm methodology and must be supplied. Projection uses tainting and compares to tolerable
misstatement.

**`three-way-match`** — PO to invoice to receiving, with the population gate that matched plus
exceptions must equal the whole invoice population in both count and value. Duplicate detection
runs four patterns, including same-amount-within-a-window and same-PO-line-billed-twice, which
catch the duplicates an exact invoice-number check misses. Aggregates exceptions by vendor first,
because the causes are process defects that cluster. Quantifies GRNI as a period-end accrual.

**`depreciation-tie-out`** — Cost and accumulated depreciation rollforwards, beginning balances
agreed to the prior year *as filed*, and asset-level integrity across seven tests. Recomputes
straight-line exactly from the schedule's own inputs; for accelerated methods it tests internal
consistency only and flags the rate for verification, asserting no recovery period, MACRS
percentage, bonus rate, or Section 179 limit. Catches the four failures that account for most
fixed-asset error: a disposed asset still depreciating, beginning accumulated that does not agree
to the prior year, an asset depreciated past basis, and duplicate assets.

---

## 1.0.0 — 2026-07-28

Initial release. Five skills.

**`bank-statement-to-excel`** — PDF statement to workbook with a mandatory completeness proof:
balance roll-forward, agreement to printed statement control totals, and line-by-line running
balance continuity. Local OCR path for scanned statements. Multi-statement continuity check to
catch a missing month.

**`bank-rec-to-gl`** — Six-pass matching cascade (check number, exact date, date tolerance,
description tokens, many-to-one batched deposits, one-to-many splits), four-column
reconciliation in proper form, proposed journal entries for unrecorded bank items, aging with
stale-check and unclaimed-property flags. Will not plug: an unexplained difference keeps the
reconciliation open.

**`tax-return-review`** — Six review passes beginning with what is absent from the return.
Evidence validation rejects any item concluded without a form/line reference and a source
document. Checklists for 1040, 1120-S, 1065, 1120, depreciation, and amended returns, all
deliberately free of statutory figures.

**`return-yoy-variance`** — Dual materiality (dollar or percent, with an absolute floor).
Always-flag conditions for sign flips, new and disappeared lines, round-number plugs, and
zero-variance lines where the underlying activity changed. Explanation components must sum to
the variance; coverage reported as a percentage of flagged dollars.

**`journal-entry-anomaly-scan`** — Population integrity gate (entry balancing, document number
continuity, field population) before 26 anomaly tests with accumulating risk scores. Includes
approval-threshold circumvention and same-day splitting detection. Reports attributes only,
never intent.

---

## Checking for updates

When using a skill from this repository, check for updates once per session on first use:

1. Fetch `VERSIONS.md` from
   `https://raw.githubusercontent.com/adoptai/cpa-skills/dev/VERSIONS.md`
2. Compare against the local skill versions
3. Only notify if two or more skills have updates, or any skill has a major version bump

Notify non-blockingly at the end of a response:

```
---
Skills update available: X CPA skills have updates.
Say "update skills" to update, or run `git pull` in your cpa-skills folder.
```

If the user asks to update, run `git pull` in the repository and confirm what changed.

**Read the changelog before updating mid-engagement.** These skills produce workpapers. A
version change that alters a materiality default, a test weight, or a scoring threshold will
make a re-run disagree with a workpaper already in the file, and a reviewer will need to know
why. Note the skill version in the workpaper when it matters.
