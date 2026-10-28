---
name: lease-schedule-tie-out
description: Recompute ASC 842 lease amortization schedules from the lease register, tie the recomputed ROU asset and lease liability balances as of a specific period end to the GL, and flag any lease with a modification for professional remeasurement rather than guessing at it. Use this whenever the user mentions lease accounting, ASC 842, the lease register, right-of-use asset, ROU asset, lease liability, operating lease versus finance lease, lease amortization schedule, or tying the lease schedule to the GL. Also trigger on "does our lease schedule tie," "recompute the lease amortization," "check the ROU asset balance," "is this lease a finance or operating lease," "was this lease remeasured correctly," or when a lease register and GL lease account detail are provided together. Pairs with depreciation-tie-out, which shares the same recompute-and-tie pattern for fixed assets; use that skill instead when the asset in question is owned, not leased. Runs fully local — no client lease data leaves the machine.
---

# Lease Schedule Tie-Out (ASC 842)

A lease schedule is rolled forward period after period, usually in a spreadsheet nobody
rebuilds from the contract terms once it's set up. Errors from a mis-keyed rate, a missed
modification, or a copy-paste that didn't move with the period persist silently for years.
This skill recomputes the schedule from the lease's own terms — the ones in the agreement, not
a remembered default — and proves whether the recomputed balance, as of a specific date,
actually equals the GL.

## The gate

The script refuses to produce a clean tie-out when:

1. **Any required term is missing or unparseable.** `incremental_borrowing_rate`, `term_months`,
   `payment_amount`, `classification`, and `commencement_date` must all be populated per lease.
   None of these is ever assumed or defaulted — a blank rate does not become a round number, it
   becomes a visible exception naming the lease and the missing field.
2. **A lease has a modification.** Any lease with a `modification_date` is excluded from the
   mechanical recompute entirely and routed to the Modification Flags tab. Remeasuring a
   modified lease requires a professional determination of the revised discount rate, term, and
   payments — this script will not guess at any of those.
3. **The recomputed balance does not tie to the GL.** For every non-modified lease, the
   liability and ROU asset recomputed as of the stated period end must equal the GL balance
   supplied for that lease, or the script reports a named, dollar-quantified difference. Nothing
   is plugged to force a match.

## Inputs

1. **Lease register** — one row per lease, with the full term, payment, frequency, the
   incremental borrowing rate taken from the lease agreement or the client's own present-value
   workpaper (never assumed by this skill), the classification, and a modification date if one
   occurred.
2. **GL detail** for the ROU asset and lease liability accounts, by lease, as of the period end
   being tied out.
3. **The period end date.** The schedule is recomputed through lease end to get the
   amortization pattern right, but only the balance as of this date is tied to the GL — tying to
   the lease-end balance for a lease still in progress would prove nothing.

## Step 1 — Build the lease register

```bash
python3 scripts/lease_schedule.py --template --out lease_register.csv
```

Fill in one row per lease. `incremental_borrowing_rate` is an annual rate as a decimal (`0.06`
for 6%) — pull it from the lease agreement or the amortization table the client already
maintains. Leave `modification_date` blank unless the lease was actually modified.

## Step 2 — Run the tie-out

```bash
python3 scripts/lease_schedule.py --leases lease_register.csv --gl gl_lease_balances.csv \
  --period-end 2025-12-31 --client "Northwind Systems Inc" \
  --out "Northwind - FY2025 Lease Tie-Out.xlsx"
```

The script recomputes each non-modified lease using the effective-interest method: interest
accrues on the beginning liability at the periodic rate derived from the annual rate and payment
frequency, principal reduces the liability, and the ROU asset amortizes on a schedule
appropriate to the classification — straight-line total lease cost for operating leases (so
interest plus amortization is level each period, the single-lease-cost pattern ASC 842 requires
for operating leases), straight-line over the term for finance leases. It then determines how
many periods have elapsed between `commencement_date` and `--period-end` and ties the balance at
that point — not the lease-end balance — to the GL.

## Step 3 — Deliver

**Workbook tabs:**

1. **Workpaper Summary** — the period end, result counts, exceptions, and a conclusion block.
2. **Schedule** — every elapsed period for every lease: beginning and ending liability,
   interest, payment, principal, beginning and ending ROU, and amortization.
3. **GL Tie-Out** — each lease's recomputed balance against the GL, with the difference called
   out and highlighted where it does not tie.
4. **Modification Flags** — every lease excluded from the mechanical recompute, with its terms
   shown, for the preparer to remeasure.
5. **Exceptions** — missing terms, unparseable fields, and ties that failed.

**Then, in chat:** the result counts first, then every exception with its dollar effect, then
the modifications routed for remeasurement.

## What to escalate

- **A modification discovered during testing that the client's schedule never remeasured** — the
  recorded balances are stale by construction and the gap compounds every period until fixed.
- **A lease classified as operating that reads like a finance lease on its terms** (a bargain
  purchase option, a term covering substantially all of the asset's useful life, a transfer of
  ownership at the end) — the classification call itself is judgment and belongs to the
  preparer, not this script.
- **A large, unexplained tie-out difference on a lease with no modification** — this points at
  either a bad rate, a missing payment, or an amendment that was never flagged as one.
- **Embedded leases inside a services or supply contract** that the register does not capture at
  all — this skill only tests what's in the register it's given.

Report each as a fact with the lease ID and the dollar amount. Do not characterize intent.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.
Lease terms and GL balances never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
