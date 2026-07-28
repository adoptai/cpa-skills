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

---

# Appendix — bundled files

If you installed the `.skill` package these files are already in place and you can ignore this appendix. If you copied the skill as text, create the files below at the paths shown, alongside your `SKILL.md`. The skill will not run without them.

## `reference/checklists.md`

```markdown
# Review Checklists

Use these when the firm has no checklist of its own. Say in the memo which checklist was
used. Each item below is a *test to perform*, and each must produce a review-note row with
evidence on both sides.

**These checklists deliberately contain no dollar thresholds, rates, or limits.** Those
change annually. Where an item depends on one, the test is written as "confirm against
current-year instructions for Form X" — cite the form and line, never a remembered figure.

---

## A. Universal — every return type, every year

**Identity and mechanics**
- Legal name spelling matches the TIN record, on every federal and state form
- TIN correct and consistent everywhere it appears (compare to the source document, not to another page)
- Address current; note if it differs from prior year (affects state filings and notices)
- Tax year begin/end dates correct on every form; short period handled if applicable
- Entity type and business code consistent across federal and all states
- Original vs. amended correctly designated; if amended, the "as originally filed" column ties to the return as actually filed
- Signatures, dates, preparer identification and PTIN present
- E-file authorization signed and dated; date not before the return was complete
- Direct deposit/debit routing and account numbers verified digit by digit
- Payment vouchers present with correct amounts and payee
- Software diagnostics list obtained, every diagnostic cleared, every override explained in writing

**Completeness**
- Each source document in the file traced to a return line (a document with no home = unreported item)
- Every form present in the prior year either present now or its absence explained
- All required statements, elections, and disclosures attached
- All state returns required by residency, nexus, and source income identified and filed
- Extension filed timely if applicable; extension payment agreed to the payment record

**Carryforwards — trace prior-year return → current-year return, cite both**
- Net operating loss (federal and each state separately; state amounts often differ)
- Capital loss carryover, split short-term and long-term
- Charitable contribution carryover by category and year of origin
- Credit carryforwards: general business, foreign tax, AMT, and any state credits
- Suspended passive activity losses, by activity
- At-risk carryovers, by activity
- Section 179 amounts disallowed and carried forward
- Basis: stock and debt basis (S corp), partnership outside basis, IRA basis
- Prior-year overpayment applied forward, agreed to the prior-year return as filed
- Any carryover *created this year* is scheduled forward and documented for next year

**Payments**
- Every estimated payment agreed to a bank debit, EFTPS confirmation, or state confirmation — not to the software's expectation
- Withholding agreed to source documents in total
- Payments applied to the correct year and the correct entity

**Trigger questions — read the source documents, not just the organizer**
- Foreign bank/financial accounts, foreign assets, foreign entities, foreign gifts or trusts (penalties here are per form per year and severe — escalate any unanswered question)
- Digital asset activity, including transfers that are not sales
- New state activity, remote employees, or economic nexus
- New entity, subsidiary, or entity classification change
- Worker classification (employee vs. contractor)
- Retirement plan filings the client may owe
- Information returns the client was required to issue
- Related-party transactions requiring disclosure

**Year-over-year reasonableness**
- Every significant movement explained
- Every line *identical* to prior year confirmed as still correct (rolled-forward figures nobody updated: depreciation, allocation percentages, apportionment factors)

---

## B. Form 1040 — individual

- Filing status supported by facts; if head of household, the qualifying-person and cost tests documented
- Dependents: names, TINs, dates of birth, relationship codes; confirm no dependent is claimed on another return you also prepare
- W-2s: wages, all withholding boxes, Social Security and Medicare wages; box 12 and 14 items considered for their return effects; multiple W-2s aggregated correctly and excess Social Security withholding considered
- Interest and dividends agreed to each 1099; nominee and accrued-interest adjustments shown; tax-exempt interest reported and its state treatment considered; private activity portion identified
- Capital transactions: proceeds and basis to 1099-B; short/long split; noncovered securities where the client supplied basis (document the support); wash sales; worthless securities; any carryover generated
- Schedule C: income agreed to 1099-NEC/1099-K plus other receipts; expenses reviewed for personal elements; home office; vehicle with mileage support; SE tax and its deduction recomputed
- Schedule E rentals: per-property income and expenses; days of personal use; depreciation; passive loss limitation and any suspended amount scheduled forward
- Schedule E flow-throughs: each K-1 box by box, including separately stated items, SE amounts, foreign items, and state schedules; basis and at-risk limitations applied; where the entity return is also available, the K-1 matches what the entity filed
- Retirement: each 1099-R distribution code and taxable amount; rollovers reported as rollovers, not as income; any early-distribution exception documented; required distributions considered
- Social Security taxable portion recomputed
- Adjustments to income: each supported; retirement and HSA contributions confirmed eligible and within limits (confirm current-year limits against instructions)
- Itemized vs. standard: the comparison shown, and the state consequence of the federal choice considered
- Itemized detail: medical, taxes with the state-and-local limitation, mortgage interest agreed to 1098 with any acquisition-debt limitation applied, charitable with substantiation and any carryover, casualty where applicable
- QBI: each activity's status; the components feeding the computation; aggregation elections; recompute
- Credits: eligibility documented; ordering and limitations recomputed; refundable vs. nonrefundable correct; due-diligence form present where the claimed credit requires one
- Other taxes: SE tax, NIIT, additional Medicare, early-distribution penalty, household employment
- Estimated tax penalty: confirm the exception applied is the right one and supported
- Next-year estimates computed on a stated basis, and that basis documented for the client

---

## C. Form 1120-S — S corporation

- Valid S election in effect; no terminating event during the year (ineligible shareholder, second class of stock, excess passive income with prior C corp E&P)
- Shareholder list, ownership percentages, and days held; allocations recomputed on a per-share per-day basis; any mid-year transfer handled correctly
- Book-to-tax reconciliation complete, each difference identified
- Reasonable compensation to shareholder-employees addressed; distributions vs. wages considered and the position documented (this is the most frequently examined S corp issue)
- Distributions tested against stock basis and AAA; any distribution in excess of basis reported correctly
- Loans to and from shareholders documented; debt basis supported; any repayment effect on basis
- Separately stated items correctly separated, not buried in ordinary income
- Schedule L, M-1, M-2 internally consistent and agreeing to the trial balance; AAA rollforward foots
- Each K-1 foots to the return totals; K-1s agree in aggregate to 100% of each item
- Shareholder basis schedules maintained and provided to shareholders
- Built-in gains, passive income, and accumulated E&P considered if the entity was ever a C corp
- Fringe benefits to >2% shareholders treated correctly
- State composite and withholding filings for nonresident shareholders

---

## D. Form 1065 — partnership

- Partner list, ownership and capital percentages, and any change during the year; allocations recomputed
- Allocation method consistent with the partnership agreement; special allocations traced to the agreement and tested for substantial economic effect
- Capital accounts maintained on the required basis; beginning capital agrees to prior-year ending; the rollforward foots
- Book-to-tax reconciliation complete; Schedules L, M-1, M-2 consistent and agreeing to the trial balance
- Guaranteed payments identified and treated correctly, not commingled with distributive share
- Liabilities allocated between recourse and nonrecourse; allocation supported; effect on basis and at-risk
- Contributions and distributions of property reviewed for gain recognition and disguised-sale exposure
- Section 704(c) items tracked where property was contributed with built-in gain or loss
- Section 754 election status and any basis adjustments; transfers of interests during the year identified
- Separately stated items correctly separated; foreign items and state schedules complete
- Each K-1 foots to return totals; aggregate K-1s equal 100% of each item; K-1 capital accounts agree to the partnership's records
- Partner-level items the partnership must report supplied (self-employment amounts, credits, state detail)
- State composite and nonresident withholding filings

---

## E. Form 1120 — C corporation

- Book-to-tax reconciliation complete; Schedules L, M-1/M-3 consistent and agreeing to the trial balance
- Tax provision, current and deferred, agreed to the financial statements; deferred balances rolled forward
- NOL carryforward traced and any current-year limitation applied
- Charitable contribution limitation and carryover
- Officer compensation reported; related-party transactions disclosed
- Accrued expenses to related parties tested for deductibility timing
- Meals, entertainment, transportation fringe, and other statutory limitations applied
- Interest expense limitation considered
- Accumulated E&P tracked; dividends and distributions treated correctly
- Consolidated or controlled-group considerations; separate-company detail supporting the consolidation
- State apportionment factors agreed to underlying schedules and consistent across states; changes from prior year explained
- Estimated payments and the safe-harbor basis used

---

## F. Depreciation and fixed assets — all entity types

- Fixed-asset schedule totals agree to the return and to the balance sheet
- Current-year additions: in service during the year, correct class life, method, and convention
- Bonus depreciation and Section 179: eligibility, elections made or out, and any state decoupling (many states do not conform — test each state separately)
- Dispositions: assets removed from the schedule, gain/loss computed with correct accumulated depreciation, recapture character correct, installment treatment if applicable
- Listed property with substantiation; personal-use percentage
- Prior accumulated depreciation agrees to the prior-year return
- Assets fully depreciated but still held remain on the schedule
- Any method change made through the proper procedure, not by simply changing the schedule
- Repairs vs. capitalization decisions consistent with the client's policy and prior years

---

## G. Amended returns

- The "as originally filed" column agrees to the return as actually filed, not to the software's current state of the file
- Explanation of changes specific and complete
- All affected years and all affected states amended, not just the federal year with the largest change
- Carryforward effects on subsequent years traced and those years amended if already filed
- Statute of limitations for refund confirmed as still open before work proceeds
- Interest and penalty effects estimated for the client
- Original return's payment history reflected correctly

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
```

## `scripts/build_review.py`

```python
#!/usr/bin/env python3
"""
Validate tax return review notes and build the review workbook.

The validation is the point. This script REFUSES to build a workbook when the
review notes fail the evidence discipline:

  - a row concluding 'agreed' or 'exception' with no source_document
  - a row concluding 'agreed' or 'exception' with no form/line reference
  - a 'must_fix_before_filing' item with no owner
  - a stated difference that does not equal per_return - expected
  - an unrecognised conclusion or severity value

Usage:
    python3 build_review.py --notes review_notes.csv \
        --client "Acme Holdings LLC" --return-type 1065 --tax-year 2025 \
        --checklist "Firm partnership checklist v3" \
        --out "Acme Holdings - 2025 1065 - Review Notes.xlsx"

Optional:
    --scope-received  "Return PDF, diagnostics, TB, FA schedule, PY return"
    --scope-missing   "Payment confirmations for Q3 estimate; K-1 from Meridian LP"
    --preparer "J. Alvarez"  --reviewer "S. Rao"
    --verdict "ready_subject_to_items"    ready | ready_subject_to_items | not_ready
    --force   build anyway, marked FAILED VALIDATION (for debugging only)

Notes CSV columns:
    ref, pass, area, form, line, item_tested, expected, per_return, difference,
    source_document, source_page, conclusion, severity, disposition, owner,
    due_date, notes

  conclusion : agreed | exception | open_item | n/a_explained
  severity   : must_fix_before_filing | should_fix | advisory | informational
"""

from __future__ import annotations

import argparse
import csv
import sys
from collections import Counter, defaultdict
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is required.  pip install openpyxl")

MONEY = '#,##0.00;[Red](#,##0.00)'

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
ADV_FILL = PatternFill("solid", fgColor="DDEBF7")

CONCLUSIONS = {"agreed", "exception", "open_item", "n/a_explained"}
SEVERITIES = ["must_fix_before_filing", "should_fix", "advisory", "informational"]
SEV_FILL = {
    "must_fix_before_filing": BAD_FILL,
    "should_fix": WARN_FILL,
    "advisory": ADV_FILL,
    "informational": None,
}
VERDICTS = {
    "ready": "READY TO FILE",
    "ready_subject_to_items": "READY SUBJECT TO THE LISTED ITEMS",
    "not_ready": "NOT READY TO FILE",
}

COLS = [
    ("ref", "Ref", 8), ("pass", "Pass", 8), ("area", "Area", 24),
    ("form", "Form", 12), ("line", "Line", 10),
    ("item_tested", "Item tested", 44),
    ("expected", "Expected / per source", 18),
    ("per_return", "Per return", 16),
    ("difference", "Difference", 14),
    ("source_document", "Source document (evidence)", 34),
    ("source_page", "Pg", 6),
    ("conclusion", "Conclusion", 16),
    ("severity", "Severity", 22),
    ("disposition", "Disposition / what to do", 40),
    ("owner", "Owner", 14), ("due_date", "Due", 12),
    ("notes", "Notes", 46),
]


def dec(raw):
    s = str(raw or "").strip()
    if s in ("", "-", "--", "n/a", "N/A", "None"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace(",", "").replace("$", "").strip()
    if s.endswith("-"):
        neg, s = True, s[:-1]
    try:
        v = Decimal(s)
    except InvalidOperation:
        return None          # non-numeric "expected" is legitimate (e.g. "Yes")
    return -v if neg else v


def load(path: Path) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No review notes in {path}")
    out = []
    for i, r in enumerate(rows, start=1):
        rec = {k: " ".join(str(r.get(k, "") or "").split()) for k, _, _ in COLS}
        rec["_row"] = i
        rec["ref"] = rec["ref"] or f"R{i:03d}"
        rec["conclusion"] = rec["conclusion"].lower()
        rec["severity"] = rec["severity"].lower()
        out.append(rec)
    return out


def validate(rows: list[dict]) -> list[str]:
    errs: list[str] = []
    seen: dict[str, int] = {}

    for r in rows:
        n, ref = r["_row"], r["ref"]

        if ref in seen:
            errs.append(f"Row {n}: duplicate ref '{ref}' (also row {seen[ref]}).")
        seen[ref] = n

        if r["conclusion"] not in CONCLUSIONS:
            errs.append(
                f"Row {n} ({ref}): conclusion '{r['conclusion']}' is not one of "
                f"{sorted(CONCLUSIONS)}."
            )
        if r["severity"] and r["severity"] not in SEVERITIES:
            errs.append(
                f"Row {n} ({ref}): severity '{r['severity']}' is not one of "
                f"{SEVERITIES}."
            )
        if not r["item_tested"]:
            errs.append(f"Row {n} ({ref}): item_tested is empty. Say what was tested.")

        # --- the evidence discipline
        if r["conclusion"] in ("agreed", "exception"):
            if not r["source_document"]:
                errs.append(
                    f"Row {n} ({ref}): concluded '{r['conclusion']}' with no "
                    f"source_document. An item is not tested without evidence - cite "
                    f"the document, or change the conclusion to 'open_item'."
                )
            if not r["form"] and not r["line"]:
                errs.append(
                    f"Row {n} ({ref}): concluded '{r['conclusion']}' with no form or "
                    f"line reference. A reviewer must be able to find it on the return."
                )

        if r["severity"] == "must_fix_before_filing" and not r["owner"]:
            errs.append(
                f"Row {n} ({ref}): must_fix_before_filing with no owner. Every "
                f"blocking item needs a named person."
            )
        if r["conclusion"] == "exception" and not r["disposition"]:
            errs.append(
                f"Row {n} ({ref}): exception with no disposition. State what should "
                f"be done about it."
            )
        if r["conclusion"] == "n/a_explained" and not r["notes"]:
            errs.append(
                f"Row {n} ({ref}): concluded 'n/a_explained' but gave no explanation "
                f"in notes. 'Not applicable' is a conclusion and needs a reason."
            )

        # --- arithmetic
        exp, per, diff = dec(r["expected"]), dec(r["per_return"]), dec(r["difference"])
        if exp is not None and per is not None:
            computed = per - exp
            if diff is None:
                r["difference"] = f"{computed:.2f}"
            elif diff != computed:
                errs.append(
                    f"Row {n} ({ref}): stated difference {diff} does not equal "
                    f"per_return - expected ({per} - {exp} = {computed})."
                )
            if r["conclusion"] == "agreed" and computed != Decimal("0"):
                errs.append(
                    f"Row {n} ({ref}): concluded 'agreed' but per_return differs from "
                    f"expected by {computed}. That is an exception, not an agreement."
                )
    return errs


def hdr(ws, n, row=1):
    for c in range(1, n + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def sheet_summary(wb, rows, meta, verdict_key, scope_recv, scope_miss, valid):
    ws = wb.active
    ws.title = "Review Summary"
    for c, w in {1: 4, 2: 46, 3: 60, 4: 16, 5: 14, 6: 12}.items():
        ws.column_dimensions[get_column_letter(c)].width = w
    r = 1

    def line(txt, *, bold=False, size=11, col=2, fill=None):
        nonlocal r
        c = ws.cell(row=r, column=col, value=txt)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c.fill = fill
        r += 1

    line("PREPARER REVIEW SUMMARY", bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    if not valid:
        line("VALIDATION FAILED - this workbook does not meet the evidence standard. "
             "See the Validation tab.", bold=True, fill=BAD_FILL)
        r += 1

    sev = Counter(x["severity"] for x in rows)
    con = Counter(x["conclusion"] for x in rows)
    mustfix = [x for x in rows if x["severity"] == "must_fix_before_filing"]

    v = ws.cell(row=r, column=2, value="VERDICT: " + VERDICTS.get(verdict_key, "(not stated)"))
    v.font = OK_FONT if verdict_key == "ready" else BAD_FONT
    if verdict_key != "ready":
        v.fill = WARN_FILL if verdict_key == "ready_subject_to_items" else BAD_FILL
    r += 1
    line(f"{len(mustfix)} item(s) must be resolved before filing.", bold=True)
    r += 1

    line("SCOPE", bold=True, size=12)
    line(f"Items tested: {len(rows)}")
    line("Documents received: " + (scope_recv or "(not stated)"))
    mc = ws.cell(row=r, column=2,
                 value="Documents NOT received: " + (scope_miss or "(none noted)"))
    mc.font = Font(bold=bool(scope_miss))
    if scope_miss:
        mc.fill = WARN_FILL
        ws.cell(row=r, column=3,
                value="A review performed without the diagnostics list or prior-year "
                      "carryforwards is a limited-scope review and should say so."
                ).font = Font(italic=True, size=9)
    r += 2

    line("RESULTS", bold=True, size=12)
    for key in ("agreed", "exception", "open_item", "n/a_explained"):
        ws.cell(row=r, column=2, value=f"  {key}")
        ws.cell(row=r, column=4, value=con.get(key, 0))
        r += 1
    r += 1
    for s in SEVERITIES:
        ws.cell(row=r, column=2, value=f"  {s}")
        c = ws.cell(row=r, column=4, value=sev.get(s, 0))
        if s == "must_fix_before_filing" and sev.get(s):
            c.font = BAD_FONT
        r += 1
    r += 1

    if mustfix:
        line("MUST FIX BEFORE FILING", bold=True, size=12)
        for h, col in zip(["Ref", "Form / line", "Item", "Effect", "Owner", "Due"],
                          range(2, 8)):
            hc = ws.cell(row=r, column=col, value=h)
            hc.fill, hc.font = HDR_FILL, HDR_FONT
        r += 1
        for x in mustfix:
            ws.cell(row=r, column=2, value=x["ref"])
            ws.cell(row=r, column=3,
                    value=" ".join(p for p in (x["form"], x["line"]) if p))
            ws.cell(row=r, column=4, value=x["item_tested"])
            d = dec(x["difference"])
            dc = ws.cell(row=r, column=5, value=float(d) if d is not None else "")
            dc.number_format = MONEY
            ws.cell(row=r, column=6, value=x["owner"])
            ws.cell(row=r, column=7, value=x["due_date"])
            r += 1
        r += 1

    openitems = [x for x in rows if x["conclusion"] == "open_item"]
    if openitems:
        line("OPEN ITEMS - CLIENT INFORMATION NEEDED", bold=True, size=12)
        for x in openitems:
            line(f"  [{x['ref']}] {x['item_tested']} — {x['disposition'] or 'request from client'}")
        r += 1
        line("Next step: turn these into a plain-language client request.", bold=True)
        r += 1

    r += 1
    line("Prepared by / date:  __________________  ____________", bold=True)
    line("Reviewed by / date:  __________________  ____________", bold=True)


def sheet_notes(wb, rows):
    ws = wb.create_sheet("Review Notes")
    ws.append([label for _, label, _ in COLS])
    hdr(ws, len(COLS))
    for x in rows:
        out = []
        for key, _, _ in COLS:
            if key in ("expected", "per_return", "difference"):
                d = dec(x[key])
                out.append(float(d) if d is not None else (x[key] or None))
            else:
                out.append(x[key] or None)
        ws.append(out)

    idx = {k: i for i, (k, _, _) in enumerate(COLS)}
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for k in ("expected", "per_return", "difference"):
            cell = row[idx[k]]
            if isinstance(cell.value, (int, float)):
                cell.number_format = MONEY
        sv = row[idx["severity"]].value
        fill = SEV_FILL.get(sv) if sv else None
        if fill:
            row[idx["severity"]].fill = fill
        if row[idx["conclusion"]].value == "exception":
            row[idx["conclusion"]].font = BAD_FONT
        if not row[idx["source_document"]].value:
            row[idx["source_document"]].fill = WARN_FILL

    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{ws.max_row}"
    for i, (_, _, w) in enumerate(COLS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def sheet_by_area(wb, rows):
    ws = wb.create_sheet("Coverage by Area")
    heads = ["Area", "Tested", "Agreed", "Exceptions", "Open items",
             "N/A explained", "Must fix"]
    ws.append(heads)
    hdr(ws, len(heads))
    agg = defaultdict(Counter)
    for x in rows:
        a = agg[x["area"] or "(unstated)"]
        a["n"] += 1
        a[x["conclusion"]] += 1
        if x["severity"] == "must_fix_before_filing":
            a["mustfix"] += 1
    for area, a in sorted(agg.items()):
        ws.append([area, a["n"], a["agreed"], a["exception"], a["open_item"],
                   a["n/a_explained"], a["mustfix"]])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        if row[3].value:
            row[3].font = BAD_FONT
        if row[6].value:
            row[6].fill = BAD_FILL
    ws.freeze_panes = "A2"
    for c, w in {1: 34, 2: 10, 3: 10, 4: 12, 5: 12, 6: 15, 7: 11}.items():
        ws.column_dimensions[get_column_letter(c)].width = w


def sheet_validation(wb, errs):
    ws = wb.create_sheet("Validation")
    for c, w in {1: 4, 2: 118}.items():
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.cell(row=1, column=2, value="EVIDENCE VALIDATION").font = Font(bold=True, size=14)
    if not errs:
        c = ws.cell(row=3, column=2,
                    value="PASSED - every tested item carries a form/line reference and "
                          "a source document, every blocking item has an owner, and all "
                          "stated differences foot.")
        c.font = OK_FONT
        return
    c = ws.cell(row=3, column=2, value=f"FAILED - {len(errs)} problem(s).")
    c.font, c.fill = BAD_FONT, BAD_FILL
    r = 5
    for e in errs:
        cell = ws.cell(row=r, column=2, value=e)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--notes", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--client", default="")
    ap.add_argument("--return-type", default="")
    ap.add_argument("--tax-year", default="")
    ap.add_argument("--checklist", default="")
    ap.add_argument("--scope-received", default="")
    ap.add_argument("--scope-missing", default="")
    ap.add_argument("--preparer", default="")
    ap.add_argument("--reviewer", default="")
    ap.add_argument("--verdict", default="", choices=[""] + list(VERDICTS))
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    rows = load(Path(args.notes))
    errs = validate(rows)
    valid = not errs

    mustfix = sum(1 for x in rows if x["severity"] == "must_fix_before_filing")
    openn = sum(1 for x in rows if x["conclusion"] == "open_item")
    verdict = args.verdict or (
        "not_ready" if mustfix else
        "ready_subject_to_items" if openn else "ready"
    )

    print("=" * 70)
    print("TAX RETURN REVIEW")
    print("=" * 70)
    print(f"Client      : {args.client or '(not stated)'}")
    print(f"Return      : {args.return_type or '(not stated)'}  TY {args.tax_year or '?'}")
    print(f"Checklist   : {args.checklist or '(not stated)'}")
    print(f"Items tested: {len(rows)}")
    con = Counter(x["conclusion"] for x in rows)
    print(f"  agreed {con['agreed']}   exception {con['exception']}   "
          f"open_item {con['open_item']}   n/a_explained {con['n/a_explained']}")
    print(f"Must fix before filing: {mustfix}")
    print(f"Verdict     : {VERDICTS[verdict]}")

    if errs:
        print(f"\nEVIDENCE VALIDATION FAILED - {len(errs)} problem(s):")
        for e in errs[:40]:
            print(f"  ! {e}")
        if len(errs) > 40:
            print(f"  ... and {len(errs) - 40} more")
        if not args.force:
            print("\n" + "=" * 70)
            print("WORKBOOK NOT WRITTEN.")
            print("An item concluded without a citation is not a tested item. Fix the")
            print("notes and re-run. Use --force only to inspect a marked-FAILED file.")
            print("=" * 70)
            return 1
    else:
        print("\nEvidence validation: PASSED")

    meta = {
        "Client": args.client or "(not stated)",
        "Return type": args.return_type or "(not stated)",
        "Tax year": args.tax_year or "(not stated)",
        "Checklist used": args.checklist or "(none stated - say which was used)",
        "Preparer": args.preparer or "(not stated)",
        "Reviewer": args.reviewer or "(not stated)",
        "Review date": datetime.now().strftime("%Y-%m-%d"),
    }

    wb = Workbook()
    sheet_summary(wb, rows, meta, verdict, args.scope_received, args.scope_missing, valid)
    sheet_notes(wb, rows)
    sheet_by_area(wb, rows)
    sheet_validation(wb, errs)
    wb.active = 0

    out = Path(args.out)
    if not valid:
        out = out.with_name(out.stem + " [FAILED VALIDATION]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\nWorkbook: {out}")
    return 0 if valid else 1


if __name__ == "__main__":
    sys.exit(main())
```

