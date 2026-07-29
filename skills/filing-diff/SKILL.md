---
name: filing-diff
description: Compare two versions of a tax return or filing field by field and show every change — proving that every field in both versions is accounted for as changed, unchanged, added or removed, that the net numeric effect reconciles to the movement in the bottom line, and that carryforwards agree to the prior year as filed. Use this whenever the user mentions comparing two versions of a return, an amended return against the original, a superseding filing, what changed between drafts, changes made after review sign-off, validating a software conversion, or confirming that carryforwards were brought forward correctly. Also trigger on "diff these returns," "what changed from the original," "compare as-filed to amended," "did anything change after I reviewed it," "check the carryforwards came across," or when two versions of the same filing are provided together. Runs fully local — no return data leaves the machine.
---

# Filing Diff

Four situations need this, and all four are places where a change slips through unnoticed:

1. **An amended return.** The "as originally filed" column must agree to what was *actually* filed
   — not to the software's current state of the file, which may have been adjusted since. This is
   the single most common defect in amended-return work.
2. **Changes made after review sign-off.** A reviewer approves a return, someone opens it again,
   and something moves. Firms rarely have any way to detect this, and it is a real quality-control
   failure.
3. **Carryforwards.** Dropped carryforwards are the highest-frequency finding in return review, and
   the check is a comparison: prior-year-as-filed against what the current year carried forward.
   Doing it by eye is why it gets missed.
4. **A software conversion.** The same return in two systems should produce the same numbers.

The mistake people make is diffing only the numbers they thought to look at. **A diff is only
useful if it accounts for every field**, including the ones that appeared, vanished, or stayed put
when they should have moved.

## The gate

No clean diff unless:

1. **Every field is accounted for.** Matched + added + removed must equal the union of fields
   across both versions. A field present in one version and silently absent from the comparison is
   the failure this test exists to prevent.
2. **Both versions are identified** — which is the earlier and which the later, with dates or labels.
   A diff with no direction cannot be read.
3. **The net numeric effect reconciles.** Where a total's components are identified, the sum of
   changes to those components must equal the change in the total. A change that does not propagate
   is either an error or an override.
4. **No change is reported without both values.** A field showing a new value and no old value is
   an addition, not a change, and the distinction matters.

## Inputs

1. **Two field-level extracts of the same filing** — form, line, description, value. From the
   return PDFs (extract with `pdfplumber`) or, better, from the software's own export, which gives
   you the field identifiers rather than the printed labels.
2. **Version labels and dates** — `as filed 2026-03-14` and `amended 2026-08-02`, or
   `reviewed 2026-03-10` and `filed 2026-03-14`.
3. **Optional: a totals map** — which fields are the components of which total. This enables the
   propagation test and is worth building once per return type.
4. **Optional: a review timestamp**, to isolate changes made after sign-off.
5. **Optional: the prior-year carryforward schedule as filed**, for the carryforward test.

Match on **field identifier where available, falling back to form plus line.** Do not match on the
printed description — descriptions change between software versions and years, which produces
phantom additions and removals.

## Step 1 — Diff

```bash
python3 scripts/filing_diff.py \
  --version-a as_filed.csv --label-a "As filed 2026-03-14" \
  --version-b amended.csv --label-b "Amended 2026-08-02" \
  --totals-map totals.csv \
  --prior-year-carryforwards py_carryforwards.csv \
  --materiality 1000 \
  --client "Brannon Holdings LLC" --return-type 1065 --tax-year 2025 \
  --out "Brannon - 2025 1065 Amended vs Original.xlsx"
```

Seven tests:

- **Test 1 — Field accounting.** Matched + added + removed = the union of fields in both versions.
- **Test 2 — Versions identified and directional.** Labels present, and which is later established.
- **Test 3 — Changes quantified.** Every numeric change carries an old value, a new value, and a
  delta. Non-numeric changes (names, addresses, elections, checkboxes) are reported separately
  because they cannot be netted and are frequently the more consequential change.
- **Test 4 — Propagation.** Where the totals map identifies components, the sum of component
  changes equals the change in the total. A total that did not move while its components did is
  flagged, and so is the reverse.
- **Test 5 — Carryforward continuity.** Each prior-year carryforward as filed against the value
  carried into the current year.
- **Test 6 — Post-review changes.** Where a review timestamp is supplied, every change made after
  it, listed individually regardless of amount.
- **Test 7 — Suspicious stability.** Fields that did *not* change where a dependent field did. A
  taxable income figure that is identical while AGI moved is either an override or an error, and no
  ordinary diff surfaces it.

## Step 2 — Read the non-numeric changes first

The instinct is to read the largest dollar change first. Resist it. **A changed identification
number, filing status, entity type, address, election, or checkbox often matters more than a
five-figure numeric change**, because those either invalidate the filing, change the entity's tax
character, or misdirect correspondence and payments. They also cannot be netted, so they will not
appear anywhere in the dollar summary.

The workbook puts them on their own tab, ahead of the numeric detail, for that reason.

## Step 3 — On amended returns specifically

If you are comparing an original to an amended return, two things are worth stating plainly on the
workpaper:

- The "as originally filed" figures must come from **the return as actually filed** — a copy of the
  filed return or the acknowledgement, not the software file. If the software file was adjusted
  after filing, the amended return's comparative column will be wrong and the diff is what reveals
  it.
- **Every affected year and every affected state** needs the same treatment. A change that alters a
  carryforward changes subsequent years too, and those may already be filed. The script reports
  which changed fields feed a carryforward so the downstream years can be identified.

Whether to amend, and which years, is a decision for the person signing. This produces the
evidence.

## Step 4 — Deliver

**Workbook tabs:**

1. **Diff Summary** — the seven tests, counts of changed, added, removed and unchanged, the net
   numeric effect, and the changes above materiality. The signable page.
2. **Non-Numeric Changes** — identification, status, elections, checkboxes, names, addresses.
   Deliberately ahead of the numbers.
3. **Numeric Changes** — every changed field with form, line, old value, new value, delta, and
   whether it exceeds materiality. Sorted by absolute delta.
4. **Added and Removed** — fields present in only one version. A removed field with a value is
   usually the more serious of the two.
5. **Propagation** — totals against the sum of their components, with any break.
6. **Carryforward Continuity** — prior year as filed against current year carried forward.
7. **Post-Review Changes** — where a review timestamp was supplied.
8. **Unchanged Detail** — for completeness, so the field accounting can be re-performed.

**Then, in chat:** the direction of the comparison, count of changes by kind, the non-numeric
changes, the net numeric effect, and anything that failed to propagate. Lead with non-numeric
changes and propagation breaks — those are the findings; the largest dollar movement is usually
the one everybody already knows about.

## What to escalate

- **A removed field that carried a value.** Something stopped being reported. That is different
  from a value changing and is easier to miss.
- **A total that did not move while its components did**, or moved by a different amount. An
  override, an error, or a rounding convention nobody documented.
- **Any change to an identification number, filing status, or entity type.**
- **A carryforward that does not agree to the prior year as filed.**
- **Any change made after review sign-off**, regardless of amount. The amount is not the point;
  the fact that the control was bypassed is.
- **A change to a field that feeds a carryforward**, because subsequent years are affected and may
  already be filed.
- **An amended return whose "as originally filed" column does not agree** to the return as actually
  filed.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry. Where
identification numbers appear as changed fields, the workbook masks all but the last four digits.

## Dependencies

```bash
pip install openpyxl
# pdfplumber if extracting field values from return PDFs
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
