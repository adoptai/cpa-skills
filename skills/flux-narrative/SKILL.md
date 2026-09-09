---
name: flux-narrative
description: Compute period-over-period movement between two trial balances, apply dual dollar-and-percent materiality, and draft variance explanations for a reviewer to confirm rather than write from scratch. Use this whenever the user mentions flux analysis, flux report, month-over-month or period-over-period variance, explaining why an account moved, a flux narrative or flux commentary, preparing the flux tab for a close package, or asks why a balance changed from last month. Also trigger on "flux this close," "explain the movement on this account," "draft the variance commentary," "two trial balances, what changed," or when two periods of trial balance data are supplied together. Runs fully local — no client or company data leaves the machine.
---

# Flux Narrative Builder

The numbers side of a flux report takes seconds: subtract one trial balance from another. The narrative side is what actually takes the week, because every material movement needs a specific, sourced reason, not a restatement of the number that moved.

This skill computes the movement and drafts the narrative's starting point. It does not invent a cause from nothing. It flags what changed, and where a person already explained the same account last period, it offers that explanation back as a labeled draft for a reviewer to confirm or replace. **A draft never counts as an explanation on its own.** Nothing closes until a reviewer confirms it, re-quantified for this period, with a source attached.

## What counts as an explanation

An explanation is causal and quantified. Direction is not an explanation.

- ✗ "Revenue increased." — restates the variance
- ✗ "Higher volume." — no magnitude, no source, not testable
- ✗ "Timing." — the single most common non-explanation in practice; it hides errors
- ✓ "Revenue up $84K on the January renewal cohort, per the deferred revenue rollforward." Ties to $84K and can be tested against the schedule.

If the components don't sum to the variance, the explanation is incomplete and the remainder stays visible as unexplained. Do not let the last 15% of a movement disappear into "other."

## Inputs

1. **Current period trial balance** — the period just closed
2. **Prior period trial balance, as reported** — not a since-adjusted version of the prior file. If the prior period was restated, use the restated figures and say so.
3. **Prior-period commentary (optional)** — the confirmed cause, evidence, and owner from last period's flux review, keyed by account. This is what makes drafting possible: a recurring driver (a SaaS renewal cadence, a known seasonal pattern, an ongoing ramp) gets recognized as recurring instead of re-investigated cold every month.
4. **This period's explanations (optional, second run)** — the worklist this skill writes, filled in by a reviewer and fed back in to close the loop.

Confirm both periods are on a comparable basis before comparing: same entity, same accounting method, same chart of accounts. A restated prior period or a chart-of-accounts restructuring makes raw variances meaningless — normalize first and disclose the normalization, or the schedule misleads.

## Step 1 — Map the accounts

Align current-period accounts to prior-period accounts by account ID. Three outcomes, and the last two matter most:

- **Matched** — same account both periods. Compute the variance.
- **New this period** — an account with a current balance and no prior one. Every one needs a cause. High-signal: a new vendor, a new revenue stream, a new expense category. Also the classic tell for a misposting into a freshly created account.
- **Disappeared** — an account with a prior balance and nothing this period. Higher signal still. Either the activity genuinely stopped (name the event) or something was dropped.

## Step 2 — Run the comparison

```bash
python3 scripts/flux_compare.py \
  --current cp.csv --prior pp.csv \
  --prior-commentary last_period_explanations.csv \
  --dollar-threshold 10000 --percent-threshold 10 --absolute-floor 2000 \
  --entity "Acme Holdings LLC" --cp-label "Sep 2026" --pp-label "Aug 2026" \
  --out "Acme Holdings - Sep 2026 Flux.xlsx"
```

Input CSVs: `account_id`, `account_name`, `amount`, and optional `account_type` and `activity_changed` (`yes`/`no`/blank, current-period file only — see the unchanged test).

**Materiality is dual, and both tests must be satisfied to be immaterial.** A movement is material if it exceeds the dollar threshold **or** the percent threshold, subject to an absolute floor below which percent is ignored. A 400% swing on a $200 account is noise; a 3% move on an $8M revenue account is not. Set thresholds by engagement judgment and record them.

Regardless of threshold, these are **always** flagged:

- **Sign flips** — a balance that crossed zero. Qualitatively material at any magnitude.
- **New and disappeared accounts** — per Step 1.
- **Zero-variance accounts where activity changed** — mark `activity_changed = yes` on any account whose underlying facts moved even though the reported figure didn't. Depreciation identical to last month after a capital addition is not a coincidence.
- **Round-number amounts** on accounts that should be computed — an accrual of exactly $50,000 suggests a plug, not a calculation.
- **Accounts equal to the prior period to the penny** where there's ongoing activity.

## Step 3 — Review the drafts, not just the blanks

This is where this skill differs from a flat variance calculator. When `--prior-commentary` is supplied and a flagged account was explained last period, the worklist arrives with that explanation already sitting in the `cause` field, prefixed `DRAFT, confirm or replace:`. It carries no dollar components — a component tied to last month's movement would misstate this month's, so those always start blank.

Three things can happen to a draft:

- **It still applies.** Re-quantify the components against this period's actual detail, attach evidence, set `status = confirmed`.
- **It's close but not exact.** Edit the cause, then quantify and confirm.
- **It no longer applies.** Replace it. A stale carried-forward reason marked confirmed anyway is worse than an honest blank, because it looks resolved to the next reviewer.

Everything without a prior-period match starts blank, exactly like a flat variance tool, because there is nothing honest to draft from.

## Step 4 — Require confirmation

The script writes `flux_worklist.csv` listing every open item: unconfirmed drafts and true blanks together. Fill in `cause` (or accept the draft), `components` (which must sum to the variance), `evidence`, `owner`, and **`status = confirmed`**, then re-run with `--explanations`. A row left as `draft` — even with a perfectly good cause sitting in it — does not count. That is the entire point: the skill offers a starting point, a person decides it's actually true this period.

The script reports the **confirmation rate**, both by count and by dollar, and refuses to mark the analysis complete below 100% of flagged items. Where a cause rests on a **client- or business-unit-provided reason you have not corroborated**, mark `evidence = client representation` — a legitimate but weaker basis, and the workbook flags it so a reviewer knows which is which.

## Step 5 — Deliver

**Workbook tabs:**

1. **Summary** — thresholds used, accounts compared, confirmation rate by count and by dollar, top ten movements, and status.
2. **Flux Schedule** — every account: type, prior period, current period, dollar and percent variance, materiality verdict, flag reason, status, cause, components, tie check, evidence, owner. Sorted by absolute variance descending.
3. **Always-Flag Exceptions** — sign flips, new accounts, disappeared accounts, unchanged-but-activity-changed, and round-number suspects.
4. **Draft Queue** — every account carrying an unconfirmed draft, isolated so a reviewer can work through exactly this list first.
5. **Unexplained** — what remains open, drafts and true blanks together. An empty tab is the goal.

**Then, in chat:** the handful of movements that actually matter, in plain language, with their confirmed causes, followed by what's still in the draft queue or genuinely unexplained and who owns it. Lead with anything still open. A controller reading this wants to know what to review, not to re-read the schedule.

## When a movement suggests something worse

Some patterns are worth escalating rather than explaining:

- A confirmed cause that repeats verbatim for the same account for several consecutive periods — the signature of a systemic error rolling forward under a label that says "recurring," not a genuine recurring driver
- A round-dollar movement that exactly offsets an equal and opposite movement in another account
- An account marked "activity changed, amount unchanged" more than once
- A prior-period figure that doesn't agree to the prior period's own reported balance — someone changed a closed period, and it needs to be reconciled before any variance is meaningful

Raise these as observations with the facts you have. Do not characterize intent.

## Security posture

Fully local. No client or company financial data leaves the machine. Inputs read-only.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
