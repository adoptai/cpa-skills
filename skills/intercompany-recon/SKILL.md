---
name: intercompany-recon
description: Reconcile intercompany AR/AP balances across two or more entities — match each entity's leg of every IC transaction against its counterparty's mirror leg (tolerant of FX noise), flag mismatches with a likely cause (timing, FX difference, missing entry, data error), and propose elimination entries only for the portion that ties. Use whenever the user mentions intercompany reconciliation, IC recon, due to/due from accounts, intercompany AR or AP, reconciling entities before consolidation, intercompany eliminations or mismatches, "the entities don't agree," shared-services or management-fee charges between subsidiaries, or multi-entity/multi-currency close. Trigger on "reconcile the intercompanies," "tie out IC balances," "why don't parent and sub agree," or trial balances/IC subledgers from two or more entities given together. For one entity's cash-to-GL rec use bank-rec-to-gl; for one entity's full balance sheet use balance-sheet-recon-pack. Runs fully local — no client data leaves the machine.
---

# Intercompany Reconciliation

Two entities' books independently record the same economic event, usually in two different
currencies, and they are expected to disagree slightly for reasons you can name — then agree
exactly once those reasons are accounted for. That is the same discipline as `bank-rec-to-gl`,
generalized across entities instead of across a bank statement and a GL. The elimination entry a
controller needs for consolidation is worthless until the mismatch behind it is explained; this
skill produces the explanation first and the entry second.

## The one rule that matters most

**Never plug, and never let FX tolerance quietly substitute for proof.** Matching across
currencies has to be forgiving — two entities translating the same invoice through different
rate dates will never agree to the penny — but "close enough to match" is not the same as
"explained." Every amount left over after matching, and every matched pair that paired only
within tolerance rather than exactly, is carried forward as a visible discrepancy that needs a
confirmed classification before it is excluded from the unexplained total. A pair that happens to
net to zero with unclassified items underneath it is coincidence, not proof, and this skill will
not call it reconciled.

## The gate

A pair is not `RECONCILED`, and the pack is not `COMPLETE`, unless:

1. **Every discrepancy has a confirmed classification** — an unmatched leg, or a matched pair
   that tied only within FX tolerance (not exactly), both need one. `timing_in_transit`,
   `fx_difference`, `missing_entry`, or `data_error`. Nothing is excluded from the unexplained
   total just because a human hasn't looked at it yet.
2. **`unidentified` never clears the gate.** It is the one classification that is itself the
   finding, not a resolution — same posture as `bank-rec-to-gl`'s unidentified withdrawal.
3. **Pure rounding (one cent or less) is the only thing auto-excused.** Anything larger that
   survived tolerant matching is disclosed, not absorbed.
4. **Elimination entries are based on the adjusted — not raw — position**, and only for the
   portion of a pair that is fully explained. A residual that hasn't been classified yet is shown
   separately and is never folded into the proposed entry.
5. **If `--expected-pairs` is supplied, every expected pair must appear** — an entity pair with no
   activity on either side is `NOT RECEIVED`, not silently absent.

## Scope: what this skill is not

- It does not decide whether an unmatched item really is timing versus a missing entry versus a
  genuine error. The script proposes a hypothesis (age-based, or "paired by doc ID but amounts
  differ") in the Mismatch Queue's suggested-cause column; the preparer confirms the real
  classification. That boundary is deliberate — see `AGENTS.md` rule 4.
- It does not decide whether an FX rate is correct, current, or the right rate type to use. You
  supply the rate; the skill applies it and surfaces, as `fx_difference`, exactly the gap that
  using a single period rate instead of each entity's own transaction-date rate will create.
- It does not book journal entries. Missing-entry and data-error items get a suggested `je_entity`
  and `je_required` flag in the Mismatch Queue; posting them is the preparer's or the close
  process's job.
- It does not replace `balance-sheet-recon-pack` for proving one entity's own balance sheet ties
  to its subledgers. This skill proves two or more entities agree with *each other*; the
  elimination entries it proposes are an input to consolidation, not the consolidation itself.

## Inputs

You need at least one of transactions or balances, plus the entity list:

1. **Entities** — `entity_id, entity_name, functional_currency`. Every entity referenced
   anywhere else must appear here first.
2. **FX rates**, required whenever more than one functional currency is in play —
   `currency, rate_to_reporting`. One rate per currency for the period, applied uniformly. Never
   let the script assume 1:1 for a currency it has no rate for; it refuses to run instead. This
   is a translation convenience for comparison, not a transaction-date rate, and the gap between
   the two is exactly what `fx_difference` exists to catch — do not treat it as a bug when a
   genuinely matched pair doesn't tie to the penny.
3. **IC subledger / transactions**, the detailed leg-by-leg data (preferred — enables real
   matching) — `entity, counterparty, doc_id, date, description, amount`. Amount is signed from
   the recording entity's own perspective: positive means the counterparty owes this entity,
   negative means this entity owes the counterparty.
4. **IC balances**, the stated trial-balance-level position per entity per counterparty — same
   sign convention, `entity, counterparty, amount, as_of`. Works alone (coarse, pair-level tie-out
   with no transaction detail) or alongside transactions, in which case the subledger must sum to
   the stated balance — a gap there means the subledger is incomplete or the balance is wrong,
   and it is flagged before you rely on anything downstream.
5. **Classifications**, fed back after reviewing the Mismatch Queue — see Step 2.
6. **Expected pairs**, optional — `entity_a, entity_b` — for coverage, so a relationship with zero
   activity submitted by either side is caught rather than silently skipped.

Confirm before running: every entity that appears in the data is in `entities.csv`, and the
reporting currency is the one consolidation actually uses.

## Step 1 — Match

```bash
python3 scripts/ic_recon.py \
  --entities entities.csv --fx-rates fx_rates.csv --reporting-currency USD \
  --transactions transactions.csv --balances balances.csv \
  --as-of 2026-01-31 --period "2026-01" --client "Acme Group" \
  --out "Acme Group - 2026-01 Intercompany Recon.xlsx"
```

For every pair of entities with any recorded activity between them, the script mirrors one
side's legs (negates them) so a genuinely matching transaction lines up as equal amounts, then
runs the same cascade `bank-rec-to-gl` uses — strongest evidence first, and a leg leaves the pool
once matched so a weaker rule can never steal it from a stronger one:

| Pass | Rule | Notes |
|---|---|---|
| 1 | Doc ID match | Matches regardless of amount — a doc-ID match with a differing amount is itself a finding, carried forward as a variance, not silently accepted or rejected. |
| 2 | Amount (within FX tolerance) + exact date | |
| 3 | Amount within tolerance + date within tolerance | Default ±5 days. |
| 4 | Amount within tolerance + description token overlap | |
| 5 / 6 | Many-to-one / one-to-many | Batched intercompany settlements, common with shared-services recharges. |

**The script does not decide what an unmatched item or a matched-but-inexact pair means.**
Classification is judgment and it is where the value of a controller's review actually sits.

## Step 2 — Classify every discrepancy

Run the script once, open `ic_unmatched.csv` (same rows as the workbook's Mismatch Queue tab),
and classify each row:

- **`timing_in_transit`** — recorded by one entity, not yet by the other, and expected to true up
  next period on its own. Record the expected resolution date so a reviewer can confirm it
  actually cleared.
- **`fx_difference`** — both entities recorded it; the translated amounts differ because each
  side's books (or this reconciliation's single period rate) used a different rate. Legitimate,
  but it needs a destination — usually the correcting entity's own books or cumulative
  translation adjustment, not this pack's elimination entry.
- **`missing_entry`** — the counterparty never recorded it at all. Needs a JE at the deficient
  entity; set `je_entity` to whichever entity should book it (defaults to the counterparty of the
  leg that does exist, if left blank).
- **`data_error`** — wrong amount, wrong period, wrong account, duplicate. Needs a correcting JE,
  not a reconciling item that rolls forward untouched.
- **`unidentified`** — the one you must never bury. Escalate regardless of amount; the absence of
  an explanation is the test, not materiality, same as `bank-rec-to-gl`.

Feed the completed file back in:

```bash
python3 scripts/ic_recon.py ... --classifications ic_unmatched.csv --out "..."
```

## Step 3 — Read the Pair Summary literally

Each pair shows both entities' positions already mirrored into the same frame — "per A: B owes
A" and "per B: B owes A" — so a tie means the two numbers are equal, not opposite. Status moves
through `NOT RECEIVED` (no activity from either side, only meaningful with `--expected-pairs`),
`UNCLASSIFIED ITEMS`, `UNIDENTIFIED`, `EXCEPTION`, up to `RECONCILED` only once every discrepancy
for that pair has a confirmed, non-`unidentified` classification. Do not read a zero "unexplained
difference" as reconciled if the pair still shows unclassified items underneath it — that
combination means the zero is coincidental, and the workbook will say so.

## Step 4 — Proposed eliminations

For each pair, once reconciled (or for whatever portion of an unreconciled pair is already
proven), the script proposes a consolidation elimination at the **adjusted** position — the
balance each entity's books will show once every confirmed `timing_in_transit` / `missing_entry`
/ `data_error` item posts, not the raw, as-recorded balance. An `fx_difference` classified on a
matched pair is deliberately excluded from the entry itself and called out separately, since it
represents translation noise on an already-agreed transaction, not an amount to eliminate or
correct through this pack. Any amount that isn't yet explained stays out of the proposed entry
entirely and is shown as an unresolved residual — this skill will eliminate the proven amount and
tell you what's left to resolve, but it will not eliminate a guess.

## Step 5 — Deliver

**Workbook tabs:**

1. **Pair Summary** — every entity pair, both sides' positions in the same frame, raw and
   unexplained difference, unidentified total, discrepancy count, and status. The signable page.
2. **Mismatch Queue** — every unmatched leg and every matched-but-inexact pair, with a suggested
   cause, the confirmed classification, which entity should book a JE, and expected resolution.
   This is the working document and the one that answers "what's different and why."
3. **Matched Detail** — every matched pair, the rule that matched it, and the variance even when
   it matched (so a reviewer can see exactly how much FX noise a "clean" match absorbed).
4. **Proposed Eliminations** — the Dr/Cr entities and amount for each pair's proven position,
   with any unresolved residual or excluded FX variance called out by name.
5. **Exceptions & Memo** — completeness warnings (subledger-to-balance ties, missing FX rates,
   not-received pairs), escalations, and the never-plug reminder.

**Then write a short memo in chat**: how many pairs, how many reconciled vs. not, the total
unidentified amount (if any — lead with this), the proposed eliminations and their total, and any
pair with no activity submitted when one was expected.

## What to escalate

- **Any `unidentified` item**, regardless of size.
- **A pair that keeps showing the same unresolved item two periods running.** An intercompany
  timing difference that never clears is rarely timing.
- **A missing-entry pattern concentrated at one entity.** One subsidiary consistently failing to
  record intercompany charges is a control gap worth naming, not just a recurring reconciling
  item.
- **An `fx_difference` large enough to suggest the entities aren't actually using a consistent
  rate policy** — that's a process finding for the close calendar, not just this period's noise.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry. FX
rates, entity names, and balances never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
