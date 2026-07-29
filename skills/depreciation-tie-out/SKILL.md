---
name: depreciation-tie-out
description: Tie the fixed asset register to the depreciation schedule and to the tax return or financial statements — proving the cost and accumulated depreciation rollforwards foot, that beginning balances equal prior-year ending, that disposals were removed with gain or loss computed, and that no asset is depreciated beyond its basis. Use this whenever the user mentions fixed assets, the fixed asset register, depreciation schedule, Form 4562, depreciation tie-out, accumulated depreciation, asset additions or disposals, a fixed asset rollforward, book versus tax depreciation, or checking depreciation against the return. Also trigger on "does our depreciation schedule tie," "reconcile fixed assets," "check the 4562," "did we remove the assets we sold," "depreciation looks wrong," "roll forward the FA register," or when a fixed asset register and a return or trial balance are provided together. Runs fully local — no client asset data leaves the machine.
---

# Depreciation and Fixed Asset Tie-Out

Fixed assets accumulate errors quietly because the schedule is rarely rebuilt — it is rolled
forward, year after year, often through a change of software or preparer. Each roll carries
forward whatever was already wrong.

Four failures account for most of what this tie-out finds:

1. **A disposed asset still depreciating.** The asset was sold, the gain was recorded, and the
   register was never updated. Depreciation continues on something the client no longer owns.
2. **Beginning accumulated depreciation that does not agree to the prior-year return.** Almost
   always introduced by a software conversion, and it silently misstates every subsequent year.
3. **An asset depreciated past its cost.** Arithmetically impossible and surprisingly common,
   usually from a life change applied retroactively.
4. **Additions never added.** Capital expenditure expensed or sitting in construction in
   progress, so depreciation is understated and the balance sheet is wrong.

None of these show up by reading the depreciation expense figure. They show up in a rollforward,
which is why this skill starts there.

## The gate

No clean workpaper unless all of these hold:

1. **The cost rollforward foots.** Beginning cost + additions − disposals = ending cost, in total
   and by asset class.
2. **The accumulated depreciation rollforward foots.** Beginning accumulated + current-year
   expense − accumulated on disposals = ending accumulated.
3. **Beginning balances equal prior-year ending balances.** Supplied from the prior-year return
   or financial statements — not from the software's current state of the file, which is what
   introduced the error in the first place.
4. **Register totals agree to the return or trial balance** for cost, accumulated depreciation,
   and current-year depreciation expense.
5. **No asset has accumulated depreciation exceeding its depreciable basis.**

## Standing rule on tax figures

**Do not state a recovery period, MACRS percentage, bonus depreciation rate, Section 179 limit,
threshold, or phase-out from memory.** These change, and several have changed more than once
recently.

What the script does instead:

- **Recomputes straight-line depreciation exactly** where the schedule states a straight-line
  method, using the cost, salvage, life, and convention *on the schedule*. This is arithmetic and
  is checked precisely.
- **For MACRS and other accelerated methods, tests internal consistency only** — that
  accumulated depreciation does not exceed basis, that prior accumulated plus current equals
  ending, that the asset is in service, that the life and class are populated and consistent
  across years — and flags each asset for verification of the *rate* against current
  instructions. It does not reproduce percentage tables.
- **Flags every asset whose life or method changed from the prior year**, because that is either
  an error or a method change requiring proper procedure rather than a schedule edit.

Where a figure must be confirmed, the workpaper says *"verify recovery period and convention for
[asset class] placed in service [date] against current instructions for Form 4562."* That is more
useful than a number recalled from an earlier year, and it cannot go stale.

## Inputs

1. **Fixed asset register**, current year, per asset: description, class, acquisition date, in-service
   date, cost, salvage, method, life, convention, prior accumulated depreciation, current-year
   depreciation, ending accumulated, disposal date and proceeds where applicable.
2. **Prior-year ending balances** — total cost and total accumulated depreciation, by class, per
   the prior-year return or financial statements **as filed**.
3. **The return or trial balance figures** to tie to: cost, accumulated depreciation, current-year
   depreciation expense. Form 4562 totals if available.
4. **Book and tax registers separately** if the client maintains both. Do not reconcile a book
   register to a tax return — the difference is a deferred tax item, not an error, and mixing
   them produces a meaningless variance.

## Step 1 — Roll forward and tie

```bash
python3 scripts/depr_tieout.py \
  --register fa_register.csv \
  --prior-year prior_balances.csv \
  --tb-cost 4820115.00 --tb-accum 2140880.00 --tb-depreciation 412655.00 \
  --basis tax --year 2025 \
  --client "Kestrel Fabrication LLC" \
  --out "Kestrel - 2025 Fixed Asset Tie-Out.xlsx"
```

Seven tests run:

- **Test 1** — Cost rollforward foots, in total and by class
- **Test 2** — Accumulated depreciation rollforward foots
- **Test 3** — Beginning balances agree to prior-year ending
- **Test 4** — Register totals agree to the trial balance or return
- **Test 5** — No asset over-depreciated; accumulated ≤ depreciable basis for every asset
- **Test 6** — Straight-line assets recomputed and agreed
- **Test 7** — Asset-level integrity (see below)

## Step 2 — Asset-level review

Test 7 checks each asset for the conditions that indicate a stale or broken register:

- **Disposed but still depreciating** — a disposal date with current-year depreciation running
  past it. The most common real finding.
- **Disposed but not removed** — a disposal date with cost still in ending balances.
- **Gain or loss not computed** — a disposal with proceeds but no gain or loss, or a gain
  computed on the wrong accumulated depreciation figure. The script recomputes proceeds less net
  book value.
- **In service before acquisition**, or an in-service date after year end.
- **Fully depreciated but still held** — legitimate, and it should remain on the register.
  Depreciation still running on it is not legitimate.
- **Zero or negative cost**, or depreciation with no cost.
- **Missing method, life, or convention** — the schedule cannot be recomputed or reviewed, which
  is itself a finding.
- **Life or method changed from the prior year** — flagged for every affected asset.
- **An addition dated in a prior year appearing for the first time** — either a late-recorded
  addition needing prior-year consideration, or a duplicate.
- **Duplicate assets** — same description, cost, and in-service date under different asset IDs.
  A double-recorded asset doubles depreciation and is invisible in totals.

## Step 3 — Deliver

**Workbook tabs:**

1. **Tie-Out Summary** — the seven tests with amounts and differences, the verdict, and the
   figures requiring confirmation against current authority. The signable page.
2. **Rollforward** — cost and accumulated depreciation, beginning through ending, in total and by
   asset class, with the prior-year agreement shown.
3. **Asset Detail** — every asset with recomputed depreciation, difference, net book value, and
   flags.
4. **Exceptions** — every asset-level finding with the asset, the condition, and the dollar
   effect where quantifiable.
5. **Additions** — current-year additions with in-service dates, method, life, and the note to
   verify recovery period and any bonus or Section 179 election against current instructions.
6. **Disposals** — each disposal with proceeds, cost, accumulated depreciation removed, net book
   value, and recomputed gain or loss, plus whether depreciation correctly stopped.
7. **Book vs Tax** — where both registers are supplied, the difference by class, which is the
   deferred tax input rather than an error.

**Then, in chat:** whether the rollforwards foot, whether beginning balances agree to the prior
year, the current-year depreciation figure and whether it ties, and the asset-level exceptions
worth acting on. Lead with any beginning-balance disagreement — it affects every year going
forward and is the hardest to unwind later.

## What to escalate

- **Beginning accumulated depreciation not agreeing to the prior-year return.** Determine whether
  the prior year, the current year, or both are wrong before proceeding. This can require an
  amended return or a method-change filing, and it does not resolve itself.
- **A retroactive life or method change made by editing the schedule.** A change in depreciation
  method or life generally requires a prescribed procedure, not a spreadsheet edit. Report what
  changed and route the treatment question to the signer.
- **Assets on the register that no longer physically exist.** If no physical inventory has been
  performed in years, say so — the register is unverified as to existence, which is a scope point
  as much as a control one.
- **Construction in progress carried for multiple years** without being placed in service.
  Depreciation may be understated and the classification may be wrong.
- **Repairs and maintenance containing capital items**, or additions that look like repairs. Worth
  scanning the expense account; this is a frequent examination adjustment in both directions.
- **Bonus or Section 179 elections that appear inconsistent** across assets placed in service in
  the same year, or state treatment applied without regard to state decoupling — many states do
  not conform, and each state has to be tested separately.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry. The
register is read-only. Output filenames carry client name and year only.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
