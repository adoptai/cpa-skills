---
name: consolidation-rollup
description: Consolidate multiple entities' trial balances into one consolidated TB — mapping each entity's chart of accounts to a consolidated line, translating through FX with a formula-driven cumulative translation adjustment (CTA), and posting elimination entries at the worksheet level — every consolidated figure traced to its entity source. Use for consolidating entities, a multi-entity or multi-subsidiary close, rolling up trial balances, consolidated financial statements, parent/subsidiary consolidation, FX translation, or CTA. Trigger on "consolidate the entities," "roll up the trial balances," "build the consolidated TB," "translate the foreign sub," "what's our CTA," or when two or more entity TBs are given together for a group close. For proving intercompany AR/AP agree before eliminating, use intercompany-recon first — this skill posts the elimination it's given and checks it balances, it does not derive the amount. For one entity's own TB integrity, use trial-balance-integrity before this. Runs fully local.
---

# Consolidation Rollup

A consolidated trial balance is a parent entity's TB, every subsidiary's TB translated into
one reporting currency, and a layer of elimination entries — summed by consolidated line and
proved to tie. Each of those three pieces has its own, separate failure mode: a chart-of-accounts
mapping gap silently drops an account from the consolidation while every entity's own TB still
balances; a translation that uses one rate for the whole balance sheet manufactures an imbalance
that has nothing to do with CTA; and an elimination entry that does not itself balance corrupts
every consolidated line it touches. This skill tests each piece on its own terms before it lets
the pieces combine.

## The one distinction that matters most

**A cumulative translation adjustment is a formula, not a plug.** Translating assets and
liabilities at the closing rate, revenue and expense at the average rate, and equity at a
historical rate is *required* by the mechanics of foreign currency translation — and it
necessarily leaves each entity's own translated trial balance out of balance by an exact,
derivable amount. That residual **is** the cumulative translation adjustment under ASC 830 / IAS
21. It is computed here strictly from the rates you supplied in `mapping.csv` and `fx_rates.csv`,
reported as its own named equity line, and it is never used to absorb anything else. An unmapped
account and an unbalanced elimination entry are not CTA — they are gate failures, reported as
such, never quietly netted into the one number that is allowed to be a residual.

## The gate

No clean consolidated TB unless:

1. **Every entity's own trial balance balances**, debits to credits, before anything downstream
   touches it. An entity TB that does not balance is not a consolidation problem — route it back
   through `trial-balance-integrity` first.
2. **Every elimination entry balances by its own `entry_id`.** An entry whose debits and credits
   do not agree is excluded from the consolidation entirely — none of its lines post — never
   applied half-complete.
3. **Every entity TB account maps to a consolidated line and a statement type.** An unmapped
   account is reported individually with its own balance, never defaulted into an "other" line,
   and it blocks the clean gate.
4. **Every functional currency other than the reporting currency has both a closing rate and an
   average rate supplied.** This script will not assume a 1:1 rate. A missing historical rate is
   allowed to fall back to the closing rate, but only visibly, flagged on every sheet that uses
   it — never a silent assumption.
5. **The consolidated trial balance ties** — debits equal credits to the cent, with CTA included
   as its own line — and every consolidated figure traces to the entity-level rows that produced
   it in the Entity Detail tab.
6. **Every intercompany-tagged elimination entry is checked against its pair's reconciliation
   status**, when `--ic-status` is supplied. A pair that is not `RECONCILED` keeps that
   elimination provisional, not final.

## Scope: what this skill is not

- It does not split non-controlling interest for a partially-owned entity. An entity with
  `ownership_pct` under 100 is flagged in the Summary and the memo; the NCI entry is a separate
  journal entry you supply through `--elimination-entries`, not something this script computes.
  Deciding how NCI should be presented is a judgment call outside the mechanical layer — see
  `AGENTS.md` rule 4.
- It does not derive an intercompany elimination amount. It posts whatever
  `elimination-entries.csv` gives it and checks that each entry balances. Run `intercompany-recon`
  first to prove the amount, and feed its proven, already-translated position in here.
- It does not perform the full ASC 830 historical rollforward of retained earnings — opening
  retained earnings at its own historical rate, plus current-year net income at the average rate,
  less dividends at the rate when declared. It translates whatever balance sits in the
  equity-mapped GL account at a single historical rate (or the closing rate, flagged, if no
  historical rate was supplied). If GAAP-precision equity translation matters for this engagement,
  split the account before mapping it.
- It does not book journal entries anywhere. Everything this skill produces is presentation of
  what you gave it.
- It does not decide whether an unreconciled intercompany pair's elimination is close enough to
  post anyway. It posts it and calls it provisional; a controller decides whether that is
  acceptable for a flash close.

## Inputs

1. **Entities** — `entity_id, entity_name, functional_currency, ownership_pct`. `ownership_pct`
   is optional and defaults to 100; anything else is flagged as an NCI exception.
2. **Entity trial balances**, all of them in one file — `entity_id, account_number,
   account_name, debit, credit`. Every `entity_id` referenced here must appear in `entities.csv`
   first.
3. **Chart-of-accounts mapping** — `entity_id, account_number, consolidated_code,
   consolidated_line, stmt_type`. `entity_id` may be `*` to apply to any entity not matched more
   specifically by its own ID — useful when entities share a chart of accounts.
   `stmt_type` (`asset | liability | equity | revenue | expense`) is what selects the translation
   rate, so get it right: it is a mechanical lookup, not a label.
4. **FX rates**, required for every functional currency other than the reporting currency —
   `currency, rate_type, rate_to_reporting` where `rate_type` is `closing`, `average`, or
   `historical`. Closing and average are mandatory; historical falls back to closing, flagged,
   if omitted.
5. **Elimination entries**, optional — `entry_id, consolidated_line, stmt_type, debit, credit,
   description, entity_a, entity_b, source`. These post directly against a consolidated line in
   reporting currency — the standard worksheet treatment for investment elimination, intercompany
   AR/AP elimination, intercompany profit elimination, and any other consolidating entry.
   `entity_a`/`entity_b` are optional; populate both on an intercompany elimination so it can be
   checked against `--ic-status`.
6. **IC pair status**, optional — `entity_a, entity_b, status`, e.g. `intercompany-recon`'s own
   Pair Summary tab exported to CSV. Confirm before running: is this the *current* period's
   recon, not a prior period's?

## Step 1 — Run

```bash
python3 scripts/consolidation_rollup.py \
  --entities entities.csv --trial-balances entity_tb.csv --mapping mapping.csv \
  --fx-rates fx_rates.csv --reporting-currency USD \
  --elimination-entries elimination_entries.csv --ic-status ic_pair_status.csv \
  --client "Acme Group" --period "2026-09" \
  --out "Acme Group - 2026-09 Consolidated TB.xlsx"
```

The script runs the six gate tests in order, translates every mapped account using the rate its
`stmt_type` implies, computes each currency-translated entity's CTA as the exact balancing
residual, applies the elimination entries that balance, and refuses a clean result if anything
above failed. The filename gets an `[INCOMPLETE]` suffix whenever it does.

## Step 2 — Read the CTA line, don't just accept it

The FX & CTA tab shows, per entity, the translated debit and credit totals that produced its CTA
figure. Two things are worth checking before you move on:

- **An entity on the reporting currency should show a zero CTA.** If it does not, something is
  wrong with the mapping or the data, not with translation — there is no real translation effect
  to explain it.
- **An entity with unmapped accounts gets a note calling its CTA provisional.** The CTA for that
  entity was computed off an incomplete subset of its accounts; treat the number as directional
  until the mapping gap in Test 3 is closed and you re-run.

## Step 3 — Close the unmapped gap

`consolidation_unmapped.csv` lists every TB account the mapping did not recognize, by entity,
with its raw functional-currency balance. Add the missing `(entity_id, account_number)` rows to
`mapping.csv` — or a `*` row if the account exists identically across entities — and re-run.
Nothing is translated or consolidated for an unmapped account; it simply does not appear in the
Consolidated Trial Balance until it is placed correctly. A wrong placement is worse than a missing
one, so the script never guesses a `stmt_type` or a line for you.

## Step 4 — Deliver

**Workbook tabs:**

1. **Consolidation Summary** — the six gate tests with pass/fail and detail, consolidated
   debit/credit totals, and a by-entity table (currency, ownership %, TB balanced, unmapped count,
   CTA amount and side). The signable page.
2. **Consolidated Trial Balance** — every consolidated line with debit, credit, and net, grouped
   and subtotaled by statement type, with the consolidated total at the foot.
3. **Entity Detail** — every mapped account, by entity: functional-currency balance, the rate
   type and rate applied, whether that rate was an approximation, and the translated
   reporting-currency balance. This is what "every figure traced" means in practice — any
   consolidated number can be traced back here to the entity rows that built it.
4. **Unmapped Accounts** — accounts with no consolidated line, by entity, in functional currency.
   Empty is the goal.
5. **Elimination Entries** — every entry supplied, whether it posted, and the IC reconciliation
   status for any entry tagged with both `entity_a` and `entity_b`.
6. **FX & CTA** — the rate table (closing/average/historical per currency, with approximations
   flagged) and the per-entity CTA computation.
7. **Exceptions & Memo** — failing gate tests, NCI entities out of scope, provisional
   eliminations, and the CTA-is-not-a-plug reminder.

**Then, in chat:** whether the gate passed, the consolidated total, the CTA by entity, and any
unmapped accounts or provisional eliminations — lead with whichever of those is blocking a clean
result.

## What to escalate

- **A material CTA swing from the prior period** without a corresponding FX rate movement of
  similar scale — that points at a rate or mapping error, not a currency event.
- **An unmapped account with a material balance.** It is excluded from the consolidation
  entirely until mapped, which understates the group's position in exactly the direction nobody
  notices until an auditor does.
- **A `stmt_type` that disagrees across entities for what should be the same consolidated line**
  (flagged as a conflict on the Consolidated Trial Balance tab) — this means two entities'
  mappings disagree about whether an account is, say, a liability or equity, which the script
  cannot resolve on its own.
- **A partially-owned entity with no NCI entry ever supplied**, period after period. The
  consolidated equity this skill produces is not final for that entity until NCI is booked.
- **An elimination entry tagged to an intercompany pair that stays unreconciled two periods
  running.** Same pattern `intercompany-recon` flags, one layer up.
- **A historical-rate approximation on an equity account with a large, old balance** — the older
  the balance, the more a closing-rate approximation can distort CTA relative to the true
  historical-rate translation.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry. Entity
names, balances, and ownership structure never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
