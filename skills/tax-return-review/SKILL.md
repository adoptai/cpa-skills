---
name: tax-return-review
description: Review a prepared tax return against a structured review checklist, producing evidence for every item tested — form and line reference, the supporting document, and the page it came from — plus a preparer-review summary of open items, risks, and recommended follow-ups. Use this whenever the user mentions reviewing a return, second-review or detail review, signing off on a return, a review checklist, review notes, missing schedules, an inconsistent or unsupported return, tying a return to source documents, or asks "is this return ready to file." Also trigger on "review this 1040/1120/1120-S/1065/990," "check this return," "run my review checklist," "quality review," "what's missing from this return," "prep this for signature," or when a completed return PDF is provided with source documents. Runs fully local — no client tax data leaves the machine.
---

# Tax Return Review

You are the second set of eyes. The preparer's job was to produce a return; your job is to
determine whether it is **supported, internally consistent, and complete** — and to leave
behind a record that says exactly what you tested and what you found.

Two standards govern everything below.

**First: an item is not "tested" without a citation.** Every conclusion carries a form and
line reference on one side and a source document with page on the other. `Verified` with no
citation is not a review note; it is a claim. A reviewer who signs off on unsupported
testing has transferred the preparer's risk onto themselves without reducing it.

**Second: the errors that matter most are invisible on the face of the return.** A wrong
number is on the page and can be caught by reading. A *missing* form, a dropped
carryforward, an unfiled information return — none of those appear anywhere on the return
you are holding. That is why the passes below start with what is absent, not with what is
present. Most reviewers do this backwards.

## Standing rule on tax figures

**Do not state a current-year dollar threshold, phase-out range, rate, standard deduction,
contribution limit, mileage rate, or filing deadline from memory.** These change annually
and some change mid-year. When a review step depends on one, do this instead:

- Recompute using the figure **as shown on the return or in the software's own worksheet**,
  and test internal consistency and arithmetic — which is where errors actually live.
- Where the figure itself must be confirmed, write the review note as: *"Verify [limit] for
  TY[year] against current IRS instructions for Form X line Y"* and flag it for the
  preparer. Cite the form and line, not a number you recall.

Being confidently wrong about a threshold is worse than flagging it, because it will be
relied on. Say what needs checking and let the current authority answer.

## Inputs

Ask for whatever is missing before starting:

1. **The return as prepared** — full PDF including all forms, schedules, statements,
   elections, and the software's diagnostics/warnings list. The diagnostics list is
   frequently the highest-yield document in the file and is frequently not provided.
2. **The source documents** — W-2s, 1099s (all types), K-1s, brokerage statements and 1099-B
   detail, closing statements, depreciation schedules, trial balance or financials,
   organizer, and payment records for estimates and extensions.
3. **The prior-year return as filed** — and the prior-year carryforward schedules.
4. **The engagement's review checklist**, if the firm has one. If not, use
   `reference/checklists.md` and say which you used.
5. **Entity and year** — return type, tax year, states, and whether this is an original or
   amended return.

Note in the memo anything you did not receive. A review performed without the diagnostics
list or without prior-year carryforwards is a limited-scope review, and it should say so.

## Pass 1 — What isn't here (do this first)

Read the source documents and the prior-year return and ask what the current return *should*
contain. Build the expected inventory, then compare to what exists.

- **Every source document maps to a return line.** Take each W-2, 1099, and K-1 and locate
  where it lands. A 1099 in the file with no home on the return is unreported income —
  matched against IRS records automatically, and the most common source of a notice.
- **Prior-year forms that vanished.** Any form present last year and absent this year needs a
  reason: business closed, property sold, election terminated, account closed. "The software
  didn't carry it" is not a reason.
- **Carryforwards.** NOL, capital loss, charitable contribution, credit carryforwards
  (general business, foreign tax, AMT), suspended passive losses, at-risk carryovers,
  Section 179 disallowed amounts, basis (stock, debt, partnership, IRA). **This is where the
  most money quietly disappears**, especially in the year a client changes preparers or the
  firm changes software. Trace each carryforward from the prior-year return to the current
  return and cite both. Do not accept the software's number as the source — the software is
  what dropped it.
- **Trigger-based forms.** Read the source documents for facts that create a filing
  requirement the preparer may not have been told about: foreign accounts or assets, foreign
  entities or gifts, cryptocurrency activity, new state nexus, a new entity or subsidiary,
  employee vs. contractor questions, retirement plan filings, information returns the client
  owed. Foreign-reporting penalties in particular are assessed per form per year and are
  severe, so an unanswered question here is escalated even when the balance due is small.
- **Elections and statements.** Elections made in prior years that must be disclosed or
  continued; new elections the transactions require; required disclosure statements.
- **Preparer due-diligence forms** where the return claims credits requiring them.

Anything absent becomes an open item with an owner. This pass typically produces the most
valuable review notes in the entire engagement.

## Pass 2 — Identity and consistency across forms

Mechanical, fast, and it catches rejections before the e-file does. Check that these agree
everywhere they appear — federal forms, every state, and every schedule:

Legal names and spelling · TINs (SSN/EIN/ITIN) · addresses · filing status · entity type ·
tax year begin and end dates · dependent names, TINs, dates of birth, and relationship codes
· state residency and part-year dates · business codes · the same entity's EIN across the
return and its K-1s.

A single transposed TIN digit causes an e-file rejection at best and a misapplied payment at
worst. Compare against the source document, not against another page of the return — an
error copied consistently is still an error, and consistency-only checking is exactly how it
survives review.

## Pass 3 — Tie material numbers to source

For every material line, cite both sides. Materiality is set by engagement judgment, but
these are tested regardless of amount because they are high-risk or high-frequency:

- Wages, withholding, and Social Security/Medicare wages to each W-2
- Interest, dividends, and qualified dividends to each 1099-INT/DIV, including
  nominee/adjustment items
- Capital transactions: proceeds *and* basis to 1099-B, and separately confirm the
  short/long-term split and any noncovered-security basis the client had to supply
- Retirement distributions to 1099-R, including the distribution code, taxable amount, and
  whether a rollover was reported as such
- K-1 items to the K-1 **as issued** — box by box, including separately stated items,
  self-employment amounts, foreign items, and the state schedules. Where you also have the
  entity return, confirm the K-1 in the individual return matches the K-1 the entity
  actually filed
- Business income and expenses to the trial balance, and reconcile book-to-tax differences
  line by line
- Depreciation to the fixed-asset detail: current-year additions and dispositions, method,
  life, convention, bonus/Section 179 treatment, and prior accumulated depreciation. Confirm
  the schedule's totals agree to the return and that disposed assets were actually removed
- **Estimated payments and extension payments to the payment records** — bank debits, EFTPS
  or state confirmations. Never to the software's expectation of what was paid. Wrong
  estimates are among the most common causes of a post-filing notice, and they are trivially
  preventable at review
- Prior-year overpayment applied forward, to the prior-year return as filed
- State income allocation and apportionment to the underlying schedules

## Pass 4 — Recompute, don't re-read

Independently recompute the return's spine. Reading a number confirms it was typed;
recomputing it confirms it is right.

Gross income → adjustments → AGI → deductions → taxable income → tax → credits → other taxes
→ payments → balance due or refund. Then the subtotals that carry limitations: itemized
deduction limits, charitable limits and any carryover created, investment interest, passive
activity and at-risk limits, basis limitations for flow-throughs, SE tax and its deduction,
QBI including the components that feed it, credit ordering and limitation, and AMT/NIIT/
additional Medicare where applicable.

Watch specifically for: subtotals that don't foot to their components, a limitation applied
to the wrong base, a carryover *created* this year that was never scheduled forward, signs
reversed on a loss, and amounts that appear on the return but on no supporting schedule.

Where a computation depends on a statutory figure, apply the standing rule above.

## Pass 5 — Year-over-year reasonableness

Compare to the prior year and require an explanation for every significant movement — a
number that moved without a reason is either an error or an unrecorded event. (The
`return-yoy-variance` skill does this systematically and produces the variance schedule.)
Also flag the opposite pattern: a number that is *identical* to last year when the
underlying activity changed. That usually means a rolled-forward figure nobody updated —
depreciation, an allocation percentage, a state apportionment factor.

## Pass 6 — Presentation and filing mechanics

Signatures and dates, preparer identification and PTIN, e-file authorizations, bank
information for direct deposit or debit (digit by digit — a wrong routing number is a
months-long problem), payment vouchers and amounts, next-year estimates and their basis,
required disclosures and statements attached, state-specific forms and copies of the federal
return where the state requires them, and the software diagnostics list cleared with each
override explained. An unexplained override is a review note, always.

## Output

### 1. Review notes — one row per item tested

Every row carries evidence. Write to a CSV, then validate and build the workbook:

```bash
python3 scripts/build_review.py \
  --notes review_notes.csv \
  --client "Acme Holdings LLC" --return-type "1065" --tax-year 2025 \
  --checklist "Firm partnership checklist v3" \
  --out "Acme Holdings - 2025 1065 - Review Notes.xlsx"
```

Columns: `ref`, `pass`, `area`, `form`, `line`, `item_tested`, `expected`, `per_return`,
`difference`, `source_document`, `source_page`, `conclusion`, `severity`, `disposition`,
`owner`, `due_date`, `notes`.

`conclusion` is one of `agreed`, `exception`, `open_item`, `n/a_explained`. `severity` is
`must_fix_before_filing`, `should_fix`, `advisory`, or `informational`.

**The script rejects the file** if any row concludes `agreed` or `exception` without both a
`source_document` and a `form`/`line` reference, if any `must_fix_before_filing` item has no
owner, or if a stated `difference` doesn't equal `per_return − expected`. This is deliberate:
the discipline is the deliverable.

### 2. Preparer-review summary

Front page of the workbook and a short version in chat:

- Scope: what was reviewed, checklist used, documents received, and **what was not received**
- Verdict: `ready to file` · `ready subject to listed items` · `not ready`
- Must-fix items before filing, each with form/line, dollar effect where quantifiable, and owner
- Open items requiring client information, with what to request (drafting the client request
  is a natural next step — say so)
- Risk observations: positions taken, disclosure considerations, penalty exposure areas,
  anything that would matter if the return were examined
- Recommended follow-ups for next year: elections to consider, records to start keeping,
  estimated-payment adjustments, entity or method questions
- Carryforward schedule confirmed and forwarded — list them with amounts. This is the single
  most useful artifact for next year's preparer

Lead with the verdict and the must-fix count. Three sentences, then detail.

## What not to do

- Do not fix the return. You are reviewing. Write the note, quantify the effect, assign it.
  A reviewer who silently corrects removes the preparer's feedback loop and the audit trail.
- Do not clear a diagnostic because it looks familiar.
- Do not accept the software's carryforward as evidence of the carryforward.
- Do not conclude on a position's technical merits from memory. Identify the issue, state
  what authority would need to be confirmed, and route it to the person who will sign.
- Do not let "immaterial" cover an item you did not test. Untested and immaterial are
  different conclusions and the workpaper should say which.

## Security posture

Fully local. No return data, TIN, or client document is transmitted anywhere. Source PDFs are
opened read-only. Output filenames use client name and year only — never a full TIN.

## Dependencies

```bash
pip install openpyxl
# pdfplumber if you are extracting the return PDF: pip install pdfplumber
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
