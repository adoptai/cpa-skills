---
name: prepaid-amortization
description: Build and verify a prepaid expense amortization schedule, recompute straight-line amortization from each item's own start date, end date, cost, and convention, tie the ending balance and current-period amortization to the GL, and propose catch-up or reversing journal entries for any gap - never posting, only proposing. Use whenever the user mentions prepaid expenses, the prepaid schedule, prepaid amortization, prepaid insurance, prepaid rent, a 12-month rule item, unamortized prepaid balances, or checking whether the prepaid schedule ties to the GL. Also trigger on "does the prepaid schedule tie," "roll forward prepaids," "build the prepaid schedule," "what's the catch-up on this prepaid," "amortize this invoice," or when a prepaid schedule or invoice list is supplied with GL detail for the prepaid accounts. Same tie-out pattern as depreciation-tie-out, applied to the other recurring schedule every close carries. Runs fully local - no client data leaves the machine.
---

# Prepaid Amortization Schedule

A prepaid schedule rots the same way a fixed asset register does: it is rolled forward every
close, rarely rebuilt, and whatever was wrong last month is still wrong this month plus
whatever broke since. Three failures account for most of what this tie-out finds:

1. **An item whose service period already ended, still carrying a balance.** The vendor was
   paid, the coverage or service expired, and nobody ran the last entry. The balance sheet
   still shows an asset that no longer represents future economic benefit.
2. **A beginning balance that doesn't agree to what was actually closed last period.** Usually
   a spreadsheet overwritten instead of rolled, or a formula that didn't drag down for a new
   row. It misstates every period after it until someone notices.
3. **An item amortizing on the wrong cadence** — too fast, too slow, or not started at all —
   so this period's entry doesn't match what straight-line amortization of the item's own
   cost, start date, and end date actually requires.

None of these show up by reading the amortization expense line. They show up in a rollforward
and a recomputation, which is why this skill runs both before proposing anything.

## The gate

No clean workpaper unless all of these hold:

1. **The schedule rollforward foots.** Beginning balance + additions − current-period
   amortization = ending balance, in total and by GL account.
2. **The beginning balance agrees to the prior period as closed** — supplied from the prior
   period's schedule or GL, not from whatever the current file happens to show, which is what
   introduced the error in the first place.
3. **The schedule agrees to the GL** for the ending balance and the current-period
   amortization expense.
4. **No item has cumulative amortization exceeding its total cost.**
5. **Recomputed straight-line amortization agrees to what the schedule states** for every item
   with enough data to recompute.
6. **Every item is free of integrity findings** — missing dates or cost, a period that ended
   with a balance still outstanding, amortization running before the service period begins, or
   a duplicate.

## Standing rule on capitalization

**This skill does not decide whether a payment should be capitalized as a prepaid asset at
all.** That question turns on the 12-month rule and the entity's own capitalization policy
threshold, both of which change and neither of which belongs in a schedule-arithmetic tool from
memory. Where it matters, the workpaper says *"verify whether this item qualifies for current
deduction against current instructions for Reg. 1.263(a)-4(f) and the entity's capitalization
policy."* It assumes every item already on the schedule was properly capitalized, and checks
only whether the arithmetic of amortizing it is right.

What the script does instead:

- **Recomputes straight-line amortization exactly**, from the item's own cost, start date, end
  date, and convention — whole-month (even by calendar month, the common prepaid convention)
  or daily (prorated by exact days). This is arithmetic, not judgment, and it is checked
  precisely.
- **Computes the expected cumulative amortization through period end** and nets out what the
  schedule already shows as amortized through the prior period. The difference is this period's
  proposed entry — the ordinary monthly amount for a clean item, or the full catch-up for an
  item that fell behind (or a reversal for one that ran ahead) in a period before this one.
- **Never guesses a start date, end date, or cost.** An item missing any of these becomes a
  visible exception, excluded from the recomputation, not silently zeroed or carried forward.

## Inputs

1. **Prepaid schedule or invoice list**, per item: ID, vendor, description, the prepaid (GL)
   account, the expense account it amortizes into, payment date, service start and end dates,
   total cost, convention, cumulative amortization through the prior period, and — only if the
   item is already on an existing schedule — this period's amortization and ending balance as
   currently shown. Leave the last two blank for a new item and the script builds them.
2. **Prior-period ending balances**, by GL account, **as actually closed** — from the prior
   period's schedule or the GL, not the current file's own idea of where it started.
3. **GL or trial balance figures** to tie to: the prepaid account ending balance and the
   current-period amortization expense.
4. **Book and tax schedules separately**, if they differ — for example because an item
   qualifies for the 12-month rule for tax but is amortized over its service period for book.
   Do not reconcile one to the other and call the difference an error; it's a deferred item.

## Step 1 — Build or verify, and tie out

```bash
python3 scripts/prepaid_amortization.py \
  --schedule prepaid_schedule.csv \
  --prior-period prior_balances.csv \
  --tb-balance 184500.00 --tb-amortization 42300.00 \
  --period-end 2026-09-30 \
  --client "Kestrel Fabrication LLC" \
  --out "Kestrel - September 2026 Prepaid Amortization Tie-Out.xlsx"
```

`--period-start` defaults to the first of the `--period-end` month; pass it explicitly for a
quarter-end or other non-monthly close. Every argument except `--schedule`, `--period-end`, and
`--out` is optional — supply what you have. Fewer inputs means fewer tests run (each says so
explicitly) and never a guessed number filling the gap.

Six tests run:

- **Test 1** — Schedule rollforward foots, in total and by GL account
- **Test 2** — Beginning balance agrees to the prior period as closed
- **Test 3** — Schedule agrees to the GL
- **Test 4** — No item amortized beyond its total cost
- **Test 5** — Recomputed amortization agrees to the schedule (see below)
- **Test 6** — Item-level integrity (see below)

## Step 2 — Item-level review

Test 6 checks each item for the conditions that indicate a stale or broken schedule:

- **Service period ended, balance remains** — the most common real finding. An unrecorded
  expense sitting on the balance sheet as if it still had future value.
- **Amortizing before the service period begins** — a cutoff problem, not an amortization
  problem.
- **Over-amortized** — cumulative amortization exceeding cost, arithmetically impossible.
- **Missing start date, end date, or cost** — the item cannot be recomputed or built, which is
  itself a finding, not a reason to assume zero.
- **Missing GL or expense account** — nothing can be tied out or posted for this item.
- **Duplicate items** — same vendor, description, cost, and start date under different IDs. A
  double-recorded prepaid doubles amortization and is invisible in the totals.

## Step 3 — Work the proposed entries, not apply them

Every item that disagrees with its recomputation, and every item the schedule left blank,
produces one line on the **Proposed JEs** tab: a first-time build, a catch-up for a prior-period
shortfall, or a reversal for a prior-period excess. Each line states its reason. **Nothing is
posted.** This is a batch for a Controller or preparer to approve or reject line by line, the
same discipline `accrual-proposals` applies to accrual entries.

## Step 4 — Deliver

**Workbook tabs:**

1. **Tie-Out Summary** — the six tests with amounts and differences, the verdict, and the
   capitalization-rule note. The signable page.
2. **Rollforward** — beginning through ending balance by GL account, with the prior-period
   agreement shown.
3. **Schedule Detail** — every item with its recomputed figure, whether it was built or
   verified, the difference, and flags.
4. **Exceptions** — every item-level finding with the item and the condition.
5. **Proposed JEs** — the debit, credit, amount, and reason for every entry needed to bring the
   schedule's current-period figure to what the recomputation says it should be.

**Then, in chat:** whether the rollforward foots, whether the beginning balance agrees to the
prior period, the current-period amortization figure and whether it ties, and the proposed
entries worth acting on. Lead with any item whose service period already ended but still
carries a balance — it is unrecorded expense sitting on the balance sheet right now.

## What to escalate

- **Beginning balance not agreeing to the prior period as closed.** Determine whether the prior
  period, the current period, or both are wrong before proceeding — this misstates every period
  going forward and does not resolve itself.
- **A large or repeated catch-up on the same item.** One missed month is a timing slip; several
  months in a row on the same vendor is a process gap worth fixing at the source, not just
  correcting in the schedule.
- **An item whose capitalization is questionable** — paid in a lump sum, short service period,
  or small dollar amount relative to the entity's policy threshold. Route the "should this even
  be a prepaid" question to the signer; this skill only amortizes what's already on the
  schedule.
- **Book and tax schedules diverging** for the same item. That's a deferred tax item to track,
  not an error to force into agreement.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry. The
schedule is read-only. Output filenames carry client name and period only.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
