---
name: k1-extract-summarize
description: Extract Schedule K-1 data from partnership, S corporation, and trust tax packages into a standardized box-by-box summary — including separately stated items, self-employment amounts, foreign items, and state schedules — then prove the extraction by footing the aggregate K-1s back to the entity return and confirming ownership percentages total 100%. Use this whenever the user mentions K-1s, Schedule K-1, K-1 packages, extracting or summarizing K-1 data, a K-1 tracker or matrix, entering K-1s into a return, tying K-1s to the entity return, or has a stack of K-1 PDFs to process. Also trigger on "summarize these K-1s," "build a K-1 schedule," "do the K-1s foot," "which K-1s are we still missing," "K-1 to 1040 input," or when several K-1 PDFs are provided together. Runs fully local — no K-1 or TIN data leaves the machine.
---

# Schedule K-1 → Standardized Summary

K-1 season fails in two directions, and neither is a data-entry problem.

The first is **transcription into the wrong box.** A K-1 has dozens of boxes and codes, many of
which look interchangeable and are not. Box 13 code W and box 13 code A go to entirely
different places on a 1040. The number gets typed correctly and lands in the wrong return line,
which is invisible on review because the figure agrees to the K-1.

The second is **a K-1 that never arrives.** Nobody notices an absent K-1 by reading the K-1s
they have. You find it by comparing what you received to what you expected, which means
building the expectation first.

So this skill does two things a spreadsheet does not: it preserves the **code** alongside every
amount, and it reconciles the population you received against the population you should have.

## The gate

This skill will not produce a clean summary unless:

1. **Aggregate K-1s foot to the entity return.** Where the entity return (or its Schedule K) is
   available, the sum of each box across all K-1s must equal the entity's total for that line.
   This is the single strongest test available and most preparers never run it.
2. **Ownership percentages total 100%.** Profit, loss, and capital percentages each sum to
   100.000% across the full set. A set that sums to 94% is missing a K-1.
3. **Every amount carries its box and code**, plus the source file and page.
4. **No amount is inferred.** An illegible or ambiguous figure is an exception, never a
   back-solved plug.

If the entity return is not available, say so — the summary is then unfooted and its scope is
limited. Do not present an unfooted summary as though it were reconciled.

## Inputs

1. **The K-1 PDFs.** Ideally the full set for an entity. A partial set can be summarized but
   cannot be footed.
2. **The entity return or its Schedule K** — enables the footing test. Ask for it.
3. **The expected recipient list** — partner, shareholder, or beneficiary schedule. Without it
   you cannot detect a missing K-1 except through the percentage test.
4. **Prior-year K-1s**, if available. Year-over-year comparison at the recipient level catches
   a partner who dropped out, an allocation that changed without an agreement amendment, and a
   capital account that does not roll forward.

## Step 1 — Extract

```bash
python3 scripts/extract_text.py --input k1_packages/ --out work/
# add --ocr for scanned K-1s (requires local ocrmypdf; nothing is uploaded)
```

This writes, per PDF, a page-delimited text layer, every word with its coordinates, and a
metadata file flagging pages with no usable text layer. Scanned K-1s must go through the local
OCR path before any figure is read from them — reading amounts off a scanned page without OCR
produces silent blanks, not errors.

K-1 layouts vary by software (Lacerte, UltraTax, CCH, Drake, ProSystem) but the box numbering
is fixed by the form, so **anchor on box numbers, not on position.** Software-specific
supplemental statements are where the detail lives and are the most commonly skipped pages.

Capture, per K-1:

| Field | Notes |
|---|---|
| `source_file`, `source_page` | Required on every row |
| `entity_name`, `entity_ein` | The issuing entity |
| `entity_type` | `1065`, `1120-S`, `1041` |
| `tax_year`, `period_begin`, `period_end` | Short periods matter |
| `final_k1`, `amended_k1` | Both change downstream treatment |
| `recipient_name`, `recipient_tin_last4` | **Last four only.** Never write a full TIN to an output file. |
| `recipient_type` | Individual, partnership, corporation, trust, IRA, disregarded, exempt |
| `pct_profit`, `pct_loss`, `pct_capital` | Beginning and ending, where shown |
| `box`, `code`, `amount`, `description` | One row per box/code combination |
| `statement_ref` | Which supplemental statement the amount came from |

**Rules that prevent the common errors:**

- **Never merge two codes into one line.** Box 13 with codes A and W is two rows. Collapsing
  them destroys the only information that determines where each lands.
- **A box with a code and no amount** points to a supplemental statement. Go find it. This is
  the most frequent cause of an understated return.
- **Capture footnotes and supplemental statements as their own rows** with
  `statement_ref` populated. Section 199A information, foreign detail, and at-risk data
  almost always live there rather than on the face of the K-1.
- **Negative amounts and parentheses** — preserve the sign exactly as presented. A loss entered
  as income is a two-sided error.
- **Do not normalize codes across entity types.** A 1065 K-1 box 13 and an 1120-S K-1 box 12
  are different forms. Keep `entity_type` on every row and never map codes between them.

Write to `k1_lines.csv`.

## Step 2 — Foot and validate

```bash
python3 scripts/build_k1_summary.py \
  --lines k1_lines.csv \
  --entity-totals entity_schedule_k.csv \
  --expected-recipients partners.csv \
  --client "Meridian Holdings LP" --tax-year 2025 \
  --out "Meridian Holdings LP - 2025 K-1 Summary.xlsx"
```

The script runs five tests and refuses to write a clean workbook if any fail:

- **Test A — Aggregate footing.** For every box/code, the sum across all K-1s equals the
  entity's Schedule K total. Reported per box, to the cent.
- **Test B — Ownership percentages.** Profit, loss, and capital each total 100.000%.
- **Test C — Recipient completeness.** Every expected recipient has a K-1; every K-1 maps to an
  expected recipient. Both directions — an unexpected K-1 matters as much as a missing one.
- **Test D — Internal consistency.** Duplicate box/code rows for one recipient, amounts with no
  box, boxes with a code and no amount and no statement reference, and any full TIN appearing
  in the data.
- **Test E — Entity consistency.** One EIN, one tax year, one entity type across the set.
  Mixed values mean K-1s from different entities or years were commingled, which invalidates
  the footing test.

## Step 3 — Read the output like a reviewer

The workbook has seven tabs:

1. **Summary** — entity, year, K-1 count, footing result, percentage totals, missing
   recipients, exception count. The page a reviewer reads first.
2. **Footing** — box by box: sum of K-1s, entity Schedule K total, difference. Differences read
   `0.00` or the workbook is marked failed.
3. **K-1 Matrix** — recipients down, box/code across. The working view: one glance shows which
   recipient is missing an item everyone else received, which is how a skipped supplemental
   statement surfaces.
4. **Detail** — every extracted row with source file and page, for re-performance.
5. **State Schedules** — state-by-state apportioned amounts and withholding by recipient.
   Nonresident withholding and composite-return eligibility get decided here, and this is the
   tab most often missing from a manual summary.
6. **Recipient Reconciliation** — expected versus received, with status and, where prior-year
   K-1s were supplied, the change in each percentage and in capital.
7. **Exceptions** — illegible figures, unresolved codes, ambiguities, and judgment calls with
   page references.

## What to escalate rather than summarize

These surface during extraction and matter more than the summary itself:

- **A K-1 marked final for a recipient who still holds an interest**, or not marked final for
  one who exited. Changes basis and gain recognition.
- **An amended K-1** superseding one already entered on a filed return — that return needs
  amending, and the client may not know.
- **Percentages that changed mid-year** with no corresponding agreement amendment or transfer
  document.
- **A capital account that does not roll forward** from prior-year ending to current-year
  beginning.
- **A negative ending capital account** — implies distributions or losses in excess of basis and
  a deficit restoration question.
- **Box 20 code Z / Section 199A information that is absent or incomplete.** Without it the QBI
  deduction cannot be computed correctly, and the entity has to be asked. Late in the season
  this is the item that holds returns hostage, so raise it the moment you see it.
- **Foreign items present** — may trigger recipient-level reporting the recipient has not been
  told about, with per-form penalties.
- **A recipient type that changes the treatment**: an IRA receiving UBTI, an exempt
  organization, a disregarded entity whose owner is the real taxpayer.

Report these as facts with the K-1 and page. Do not conclude on treatment — route it to the
person signing.

## Standing rule on tax figures

Do not state a threshold, rate, phase-out, or limit from memory. Where a code's treatment
depends on one, write the note as *"verify treatment of box X code Y for TY[year] against
current instructions for Schedule K-1 (Form [entity type])"* and cite the box and code. Codes
and their meanings are revised between years; a code letter recalled from an earlier year is a
silent misposting.

## Security posture

Fully local. **Full TINs are never written to output** — recipient identification uses name and
last four digits only, and the validator fails if a full TIN pattern appears in the data.
Output filenames carry entity name and year only. Source PDFs open read-only. No network calls.

## Dependencies

```bash
pip install openpyxl pdfplumber
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
