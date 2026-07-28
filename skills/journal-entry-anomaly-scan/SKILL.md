---
name: journal-entry-anomaly-scan
description: Scan a full-population general ledger or journal entry extract for duplicate, round-dollar, weekend and holiday, after-hours, manually posted, self-approved, and threshold-adjacent entries — then produce a risk-ranked exception listing and an audit workpaper documenting the procedure, population, criteria, exceptions, and conclusion. Use this whenever the user mentions journal entry testing, JE testing, GL scan, searching for unusual or manual journal entries, duplicate entries, round-dollar entries, weekend postings, topside adjustments, segregation of duties in posting, fraud risk procedures on entries, or asks to find anomalies in the general ledger. Also trigger on "test journal entries," "scan the GL," "find suspicious entries," "AU-C 240 journal entry procedures," "who posted this," "look for plugs," or when a GL detail or JE export is provided for examination. Runs fully local — no client ledger data leaves the machine.
---

# Journal Entry Anomaly Scan

Journal entry testing is a required procedure, not a discretionary one, and it is the
procedure most often performed as theater — a sample of twenty entries pulled from a
population of four hundred thousand, with no stated basis for selection. Full-population
scanning is available and cheap, which changes what a reasonable procedure looks like.

Two things to hold onto throughout.

**The output is an attribute listing, never a conclusion about intent.** An entry posted at
2:14 a.m. on a Sunday for exactly $250,000 by the person who also approved it is a
*characteristic*, and characteristics require investigation, not accusation. Write "posted
outside business hours by the approving user" and let the follow-up determine what it means.
Every exception gets a disposition after inquiry — most turn out to be a batch process, a
time-zone artifact, or an overseas shared-service team. Document that too; a cleared
exception is evidence.

**A single flag is weak; converging flags are strong.** The value is not in the round-dollar
list or the weekend list. It is in the entry that appears on four lists at once. Rank by
accumulated risk, and read the top of that ranking rather than working through each test
sequentially.

## Step 0 — Prove the population before testing it

Testing an incomplete extract produces false comfort, and this is where most JE testing
quietly fails. Before any anomaly test runs:

1. **Every entry balances.** Debits equal credits for each `entry_id`. An unbalanced entry
   means truncated extraction, not a ledger error — no ERP posts an unbalanced entry.
2. **The population ties to the trial balance.** Net activity by account, plus opening
   balances, must equal closing balances per the TB. If it doesn't, the extract is missing
   entries and no anomaly test is meaningful.
3. **Entry number continuity.** Gaps in a sequential document numbering scheme are either an
   incomplete extract or deleted entries. Both matter; they need different follow-ups.
4. **Period boundaries.** Confirm the extract covers the full period with nothing on either
   side, and that entries dated in-period but *posted* after period end are visible — that
   population is tested separately below.
5. **Fields actually populated.** Missing `posted_by`, `posted_timestamp`, or `source` fields
   silently disable entire tests. If a system doesn't capture posting user or timestamp, that
   is itself a control finding worth reporting, and the scope limitation must be stated.

The script performs 1, 3, and 5 and reports 2's inputs so you can tie it out.

```bash
python3 scripts/je_scan.py --entries gl.csv --validate-only
```

## Step 1 — Run the scan

```bash
python3 scripts/je_scan.py \
  --entries gl.csv \
  --period-start 2025-01-01 --period-end 2025-12-31 \
  --materiality 75000 \
  --approval-thresholds "10000,50000,250000" \
  --business-hours 7-19 \
  --holidays holidays.csv \
  --authorized-users authorized.csv \
  --client "Acme Holdings LLC" \
  --out "Acme - FY2025 JE Scan Workpaper.xlsx"
```

Input CSV — one row per **line**, grouped by `entry_id`:

| Field | Required | Notes |
|---|---|---|
| `entry_id` | yes | Journal entry number/document number |
| `line_no` | | Line within the entry |
| `effective_date` | yes | The accounting date |
| `posted_timestamp` | | When it entered the system. Drives weekend/after-hours/post-period tests |
| `account` / `account_name` | yes | |
| `debit` / `credit` | yes | One or the other, never both |
| `description` | | Line or header memo |
| `posted_by` | | User ID |
| `approved_by` | | User ID |
| `source` | | `MANUAL`, `AP`, `AR`, `PAYROLL`, `INTERFACE`, etc. |
| `reversal_of` | | Entry ID this reverses, if the system tracks it |

## Step 2 — The tests

Each carries a weight; scores accumulate per entry. Thresholds are engagement judgment —
record what you used, because the criteria are part of the workpaper.

**Manual and source-based**

- *Manual entries* — posted by a person rather than a subsystem. This is the base population
  for most of what follows. In a well-controlled environment manual entries are a small
  fraction of the total; if they are a large fraction, that is the finding.
- *Topside and consolidation entries* — posted above the subledgers, especially near period
  end.
- *Entries to accounts that are rarely used* — an account with two entries all year, one of
  which is material, is worth a look regardless of any other attribute.

**Timing**

- *Weekend and holiday postings* — by `posted_timestamp`, not effective date.
- *After-hours postings* — outside configured business hours. Expect false positives from
  batch jobs and offshore teams; confirm the client's actual processing pattern before
  treating these as exceptions, and note the time zone the timestamps are recorded in.
- *Period-end concentration* — entries in the last few days of the period, and in the first
  days of the next period dated back into the closed one. Sort by effective-date-to-posting-
  date lag; a long lag on a material manual entry is one of the highest-yield attributes in
  the entire scan.
- *Entries posted after the books were closed* but dated within the period.

**Amount patterns**

- *Round dollar* — scaled to magnitude, so $1,000,000.00 flags but $47,000 does not merely
  for ending in three zeros. Round amounts on computed accounts (depreciation, accruals,
  reserves, allowances) score higher, because those should be the output of a calculation.
- *Threshold-adjacent* — entries falling just below an approval limit. Supply the client's
  actual authorization thresholds. **An entry of $9,987 under a $10,000 approval limit is
  among the strongest single indicators available**, because it implies knowledge of the
  control and intent to stay under it. Splitting is the related pattern: several entries just
  below a limit, same day, same account, same user.
- *Benford's law on leading digits* — a population-level test, not an entry-level one. A
  deviation tells you where to look, not what is wrong, and it is unreliable on small
  populations or on amounts with natural bounds (anything priced, capped, or contractual).
  Report the distribution; do not draw a conclusion from it alone.

**Authorization and segregation**

- *Preparer equals approver* — self-approved entries. A control finding regardless of the
  amount.
- *Unapproved entries* — blank `approved_by` where the control requires approval.
- *Users outside the authorized list* — supply the list. Terminated employees still posting
  is the finding this test exists for, and it is worth checking `posted_by` against HR
  termination dates.
- *Users posting outside their normal accounts or normal volume* — a per-user profile that
  shifts mid-year.

**Duplicates**

- *Exact duplicates* — same account, amount, and effective date under different entry IDs.
- *Near duplicates* — same account and amount within a short date window. Usually a genuine
  double-posting, which is a real misstatement, not a fraud indicator.

**Description quality**

- *Blank descriptions* on manual entries.
- *Generic or self-incriminating descriptions* — `adjustment`, `reclass`, `plug`, `to
  balance`, `per client`, `per [name]`, `correction`, `true-up`, `misc`, `see attached`, `as
  discussed`, `do not reverse`. These are worth reading in full rather than pattern-matching:
  a description is the entry's own explanation of itself, and a manual material entry that
  cannot explain itself in words is the thing you are looking for.

**Account relationships**

- *Direct revenue or reserve manipulation patterns* — a manual entry crediting revenue with
  an unusual counter-account, or debiting a reserve/allowance with a counter-account outside
  the normal cycle.
- *Entries touching cash and revenue directly* without an AR leg.
- *Suspense, clearing, and "other" accounts* used as the counter-account on material manual
  entries.
- *Entries that cross unrelated cycles* in a single document.

**Reversals**

- *Reversing entries that never reversed* — an accrual reversed in the following period is
  normal; one that was never reversed inflates or deflates the period and stays that way.
- *Reversals of a different amount* than the original.
- *Same-day reverse-and-repost*, which can be a legitimate correction or an attempt to
  reshape an entry after review.

## Step 3 — Rank, select, and investigate

The script produces a risk score per entry and a ranked exception listing. Then:

1. **Read the top of the ranking first**, not the individual test tabs.
2. **Select for testing** on a stated basis: all entries above a score threshold, all
   material manual entries, all self-approved entries above a floor, plus a random sample
   from the remainder so the unflagged population is not untested. Say what the basis was —
   an unstated selection basis is the defect that makes JE testing indefensible.
3. **Obtain support** for each selection: the entry, its supporting documentation, and an
   explanation from the preparer.
4. **Disposition each one**: `supported` · `supported with control observation` ·
   `misstatement identified` · `unresolved`. Unresolved items escalate. So does any pattern of
   the same user, account, or description recurring across dispositions — the pattern is
   often more significant than any single entry.

## Step 4 — Deliver the workpaper

**Workbook tabs:**

1. **Workpaper Summary** — objective, population and how it was obtained, completeness tests
   and their results, criteria and thresholds used, tests performed, exceptions by test,
   selections and their basis, conclusion, preparer/reviewer signature block. This is the tab
   that has to stand on its own if someone re-reads the file in three years.
2. **Population Integrity** — balancing test by entry, entry-number continuity, field
   population rates, activity by account for tie-out to the TB, and any scope limitation.
3. **Ranked Exceptions** — every flagged entry with its accumulated score, every flag it hit,
   amount, date, posting lag, user, approver, and columns for selection, support obtained, and
   disposition.
4. **One tab per test** — so a reviewer can re-perform any single test.
5. **User Activity Profile** — entries and dollars by user, by month, with manual percentage
   and self-approval count. Reveals the mid-year behavior change no single-entry test catches.
6. **Benford** — leading-digit distribution against expected, clearly labeled as directional
   only.
7. **Not Selected** — the population that was flagged but not selected, with the reason. The
   absence of this tab is what makes a selection look arbitrary.

**Then, in chat:** population size and dollars, completeness result, the five or six entries
that actually warrant attention and why, control observations, and what you need from the
client. Lead with anything unresolved.

## Reporting language

Write attributes and facts. "Entry JE-40118 for $250,000.00 was posted on Sunday 2025-12-28
at 02:14 by user rkm, who also appears as the approver, with the description 'reclass per
discussion'." That sentence is defensible and actionable. Anything that characterizes intent
is neither. Where a control weakness is evident, report it as a control observation to
management; where a misstatement is identified, quantify it and route it through the normal
misstatement process. If a matter suggests possible fraud, follow the engagement's
communication protocol — that is a firm decision, not a scan output.

## Security posture

Fully local. General ledger data never leaves the machine — no network calls, no uploads, no
telemetry. Inputs read-only.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
