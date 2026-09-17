---
name: balance-sheet-recon-pack
description: Build the balance sheet reconciliation pack across the full chart, not just cash — for every asset, liability, and equity account, prove GL balance equals supporting schedule plus itemized items, with an unexplained difference that must be zero and a sign-off requiring a named preparer and reviewer before an account counts as reconciled. Use this whenever the user mentions a balance sheet reconciliation pack, the rec pack, the close binder or tie-out binder, reconciling every balance sheet account, which accounts are reconciled this month, or asks which accounts still need a schedule or a sign-off. Also trigger on "build the rec pack," "reconcile the whole balance sheet," "tie every account to its schedule," or when a trial balance and account-level schedules are provided together. For deep transaction matching on one cash account, use bank-rec-to-gl instead — this pack assembles that result, and every other account's, into one full-chart proof. Runs fully local — no client data leaves the machine.
---

# Balance Sheet Reconciliation Pack

`bank-rec-to-gl` proves one account: cash. This proves the rest of the balance sheet the
same way, account by account, and it is **the artifact auditors ask for first** — a rec
pack tells them in one workbook which balances are proven, which are open, and which
nobody has looked at yet.

The discipline is identical to a bank reconciliation, generalized. Every account's GL
balance must equal its supporting schedule plus whatever dated, itemized reconciling
items explain the gap. What is left over is the unexplained difference, and it must read
zero. Not close to zero. Zero, or the account is not reconciled.

## The one rule that matters most

**Never plug, and never accept a tie with no signature as reconciled.** Two different
failures corrupt a rec pack and this skill refuses both:

- A "reconciling item" invented or rounded to make the math work. If it cannot be
  itemized with a date and a dollar amount, it does not belong in the reconciling items
  column — it belongs in the unexplained difference, visible, where a reviewer will see
  it.
- An account marked "reconciled" because the preparer believes it, with no name and date
  actually on file. A tie with no preparer and reviewer of record is progress. It is not
  a reconciled account, and this skill will not call it one.

## The gate

No account counts as reconciled, and the pack is not reportable as complete, unless:

1. **Every balance sheet account in the trial balance appears in the pack** — reconciled,
   exception, tied-but-unsigned, or not received. An account absent entirely is the
   failure mode this exists to prevent; it is never silently dropped.
2. **Unexplained = GL balance − (schedule balance + reconciling items).** Any account
   where that is not exactly zero is an exception and stays visible — never plugged,
   never absorbed into "other," never rounded away.
3. **RECONCILED requires both a preparer and a reviewer, each with a name and a date, on
   record.** Without both, the best status an account can reach is `TIED - NOT SIGNED`
   or `PREPARED - NOT REVIEWED`.
4. **The pack is COMPLETE only when every account is RECONCILED.** The script still
   produces the workbook when it is not — a controller needs it mid-close — but the file
   is named `[INCOMPLETE]` and the run exits non-zero.

## Scope: what this skill is not

- It does not decide whether an account should be reconciled *at all* this period.
  Waiving an immaterial or dormant account is a close-checklist judgment — see
  `month-end-close-checklist`. `NOT RECEIVED` here means exactly that: no schedule was
  submitted, not that someone decided it didn't need one.
- It does not perform the account-specific procedure. For cash, run `bank-rec-to-gl`
  (six-pass transaction matching) or `bank-rec-review` if someone else prepared it; for
  AR, `ar-aging-tie-out`; for fixed assets, `depreciation-tie-out`; for payables and
  accruals, `cutoff-and-unrecorded-liabilities`. This pack takes each of those
  procedures' *output* — a schedule balance and a list of reconciling items — and proves
  the chart is covered. The script names a suggested skill per account so you know which
  one to run first.
- It does not classify what a reconciling item *is*. Same boundary as `bank-rec-to-gl`:
  matching and footing are mechanical, classification is judgment.

## Inputs

1. **Trial balance** — account, name, debit, credit, type, and contra flag. Same schema
   `trial-balance-integrity` and `month-end-close-checklist` use, so all three chain
   directly. Only asset, contra-asset, liability, contra-liability, and equity rows are
   in scope; revenue and expense rows are skipped. If the trial balance itself does not
   tie or has untyped accounts, run `trial-balance-integrity` first — this pack assumes a
   sound trial balance and will error out on an unrecognized account type rather than
   guess it.
2. **Supporting schedules**, one row per account that has one: schedule balance, source,
   and preparer/reviewer name and date. This is usually the output of running the
   account-specific skill above, or a subledger export, or a manually prepared schedule.
   An account with no row here is `NOT RECEIVED`.
3. **Reconciling items**, zero or more per account: date, description, signed amount.
   Every item needs a date — an undated item cannot be aged, and aging is what turns a
   reconciling item from a label into evidence.
4. **Prior period's unresolved items**, if a prior pack exists — used only to detect an
   item that keeps reappearing unresolved, which is rarely timing.
5. **Materiality**, if you want the `ESCALATE` flag to mean something. It does not affect
   the gate — every exception and every not-received account is shown regardless of
   size — it only marks which ones need attention first.

## Step 1 — Build the pack

```bash
python3 scripts/recon_pack.py \
  --trial-balance tb.csv \
  --schedules schedules.csv \
  --items items.csv \
  --prior-pack prior_items.csv \
  --materiality 5000 \
  --as-of 2025-12-31 \
  --client "Halstead Regional Foods" --period "2025-12" \
  --out "Halstead - 2025-12 Balance Sheet Recon Pack.xlsx"
```

For each balance sheet account the script computes:

- **Raw difference** — GL balance minus the schedule balance.
- **Unexplained** — raw difference minus the total of that account's reconciling items.
  If this is not zero, the account is an `EXCEPTION`, full stop.
- **Status**, in order of least to most complete: `NOT RECEIVED` (no schedule at all),
  `EXCEPTION` (schedule received, math doesn't tie), `TIED - NOT SIGNED` (ties, no
  preparer on file), `PREPARED - NOT REVIEWED` (preparer on file, no reviewer),
  `RECONCILED` (ties, preparer and reviewer both on file).

An account number in `schedules.csv` or `items.csv` that does not appear on the trial
balance is reported as a data warning rather than silently ignored — it is almost always
a typo, and a typo here means a real account's schedule never gets matched to it.

## Step 2 — Read the sign-off column literally

The distinction between `TIED - NOT SIGNED`, `PREPARED - NOT REVIEWED`, and `RECONCILED`
is the entire point of tracking sign-off separately from the math. A workbook full of
accounts where the difference is zero looks finished. It is not finished until a named
person has prepared it and a different named person has reviewed it — that is what makes
the pack an audit artifact rather than a spreadsheet someone ran once. Do not backfill a
preparer or reviewer name to close out a status; if the sign-off didn't happen, the
account stays open.

## Step 3 — Read the aging and recurrence, not just the total

- **Items over 90 days** on the Aging tab are flagged. An open item that old on a
  balance-sheet-wide basis (not just cash) usually means either a stale reconciling item
  that never should have carried forward, or a real problem nobody escalated.
- **The same item — same account, same description, same amount — unresolved for three
  or more consecutive periods** is called out by name on both the Reconciling Items tab
  and in the memo. This is the balance-sheet-wide version of the recurring-item pattern
  `bank-rec-to-gl` watches for, and it is rarely timing.

## Step 4 — Deliver

**Workbook tabs:**

1. **Pack Summary** — the gate, the coverage proof (every status count against the total
   population of balance sheet accounts), and completeness. The signable page.
2. **Chart Status** — every balance sheet account: GL balance, schedule balance,
   reconciling items total, unexplained, status, escalate flag, preparer/reviewer names
   and dates, source, and the suggested skill to run if the schedule is still missing.
   Sorted worst-status-first so the accounts needing attention sort to the top. This is
   the working document.
3. **Reconciling Items** — every itemized reconciling item across every account, with age
   and recurrence flagged.
4. **Aging** — reconciling items bucketed 0–30 / 31–60 / 61–90 / 91–180 / 180+ by
   account, with stale and recurring flags.
5. **Exceptions & Memo** — exceptions, not-received accounts, unsigned ties, recurring
   items, and data warnings, in one readable memo.

**Then write a short memo in chat**: whether the pack is complete, the count of
exceptions and their dollar total, the count of not-received accounts and which skills
would supply their schedules, and anything recurring or escalated. Lead with exceptions
and not-received accounts above materiality — those are the pack.

## What to escalate

- **Any exception above materiality.** An unexplained difference on a balance sheet
  account is the same finding as an unexplained bank reconciliation difference, just on
  an account nobody watches as closely as cash.
- **A material account that is `NOT RECEIVED`.** Nobody has proven that balance this
  period, regardless of what the trial balance says it is.
- **A suspense, clearing, or conversion account carrying any balance at all**, reconciled
  or not — these are meant to be empty. A tied reconciliation on a suspense account still
  means something was parked there.
- **The same reconciling item recurring three or more consecutive periods.** Treat it as
  an open finding, not an accepted balance, the same way you would on a bank rec.
- **An account reconciled last period that is `NOT RECEIVED` or `EXCEPTION` this
  period**, with no obvious cause — a control that existed and stopped operating is worth
  more attention than one that never existed.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no
telemetry. Client balances and preparer/reviewer names never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
