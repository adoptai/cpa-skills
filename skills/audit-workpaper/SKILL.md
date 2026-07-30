---
name: audit-workpaper
description: Build an audit workpaper that documents the objective, the population and how it was obtained, the procedures performed, the evidence examined, the exceptions found, and a conclusion — and validate that the conclusion is actually supported by what was documented, refusing to accept "no exceptions noted" when exceptions exist or a conclusion reached with no evidence recorded. Use this whenever the user mentions audit workpapers, working papers, audit documentation, a workpaper template, documenting a procedure, evidence and conclusions, tickmarks, workpaper review notes, sign-off, or preparing documentation for inspection or peer review. Also trigger on "write up this workpaper," "document this procedure," "build a workpaper," "does this workpaper support the conclusion," "prepare this for review," or when procedure results need to be turned into documentation. Runs fully local — no client data leaves the machine.
---

# Audit Workpaper

A workpaper has one job: **demonstrate that the work was done, and that the conclusion follows from
it.** An experienced auditor with no prior involvement should be able to read it and understand what
was tested, on what population, against what evidence, what was found, and why the conclusion is
what it is — without asking anyone.

The most frequently cited documentation deficiency in inspection is not a missing signature or a
formatting problem. It is **a conclusion that the documentation does not support** — most often
"no exceptions noted" on a workpaper that never records what was examined, or a conclusion reached
while exceptions sit unaddressed elsewhere in the file.

So the validation here is not a completeness checklist. It is a **consistency check between the
conclusion and everything above it.**

## The gate

The workpaper is not accepted unless:

1. **The objective, population, procedure, evidence, and conclusion are all present.** A workpaper
   missing any one of them is incomplete by definition.
2. **The population is stated with its source and its total**, and how it was obtained. A procedure
   performed on an unstated population supports nothing.
3. **Every procedure step records the evidence examined.** A step with a result and no evidence is
   an assertion.
4. **The conclusion is consistent with the exceptions.** This is the test that matters:
   - It may not state or imply "no exceptions" when exceptions are recorded
   - It may not assert agreement where a recorded difference is non-zero
   - It may not be reached at all where a step has no evidence
   - An exception with no disposition blocks the conclusion
5. **Preparer and reviewer are named and different people.** A workpaper reviewed by its preparer
   has not been reviewed.
6. **Every tickmark used is defined**, and every tickmark defined is used.

## Inputs

1. **The objective** — what this workpaper is meant to establish, tied to an assertion or a control.
2. **The population** — description, source, how obtained, total, and item count. Where it should
   agree to a general ledger or trial balance figure, that figure.
3. **The procedure steps** actually performed — not the audit programme's wording, what was done.
4. **The evidence examined** per step, with document references.
5. **The exceptions found**, each with an amount where quantifiable and a disposition.
6. **The conclusion**, in your own words.
7. **Preparer, reviewer, dates**, and the workpaper reference.

## Step 1 — Build and validate

```bash
python3 scripts/build_workpaper.py \
  --header header.csv \
  --procedures procedures.csv \
  --evidence evidence.csv \
  --exceptions exceptions.csv \
  --tickmarks tickmarks.csv \
  --conclusion conclusion.txt \
  --out "C-210 Trade Receivables Existence.xlsx"
```

Seven tests:

- **Test 1 — Required elements present.** Objective, population, procedure, evidence, conclusion,
  preparer, reviewer.
- **Test 2 — Population stated and tied.** Source, how obtained, total, count; and where a control
  figure is supplied, agreement to it.
- **Test 3 — Every step has evidence.** No step concluded without a document reference.
- **Test 4 — Conclusion consistent with exceptions.** The core test. See below.
- **Test 5 — Every exception dispositioned.** An exception left open blocks the conclusion.
- **Test 6 — Preparer and reviewer are different, and both dated.**
- **Test 7 — Tickmarks defined and used.** Both directions.

## Step 2 — How the conclusion consistency test works

The script reads the conclusion text and compares its claims against the recorded facts. It looks
for language asserting an absence of problems — `no exceptions`, `no exceptions noted`, `without
exception`, `agreed`, `no differences`, `properly stated`, `fairly stated`, `no misstatements`,
`satisfactory`, `effective` — and fails the workpaper if any of it is present while:

- exceptions are recorded, or
- a step reports a non-zero difference, or
- an exception has no disposition

It also fails a conclusion that draws a positive assurance while any step has no evidence recorded,
because a conclusion cannot rest on work that was not documented.

**This is deliberately hard to satisfy by rewording.** If exceptions exist, the conclusion must
acknowledge them and explain why they do not change the outcome — which is what a defensible
conclusion looks like anyway. "Three exceptions totalling $4,120 were identified, all relating to
timing and cleared in the subsequent period; the balance is supported" passes. "No exceptions noted"
on the same facts does not.

## Step 3 — On writing the conclusion

A conclusion is a logical statement connecting the evidence to the objective. Four failure modes,
all common:

- **Restating the procedure.** "We tested 40 invoices" is not a conclusion; it is a procedure.
- **Concluding on the wrong thing.** A sample selected proportional to recorded value supports
  occurrence, not completeness. If the objective was completeness, the conclusion cannot claim it.
- **Concluding beyond the evidence.** Testing existence does not support valuation.
- **Burying an exception.** An exception mentioned nowhere in the conclusion is an exception the
  reviewer has to find themselves.

Write the preliminary conclusion immediately after performing the procedure rather than at
assembly. Capturing it while the work is fresh preserves the judgement as it actually was, and if it
needs revising later, revise it — the revision is itself part of the record.

## Step 4 — Deliver

**Workbook tabs:**

1. **Workpaper** — the document itself, in proper form: header block, objective, population and
   source, procedures performed, evidence examined, exceptions, conclusion, and the sign-off block.
   This is the thing that goes in the file.
2. **Validation** — the seven tests with what failed and why. Not part of the workpaper; part of
   preparing it.
3. **Procedures and Evidence** — step by step, with the evidence reference and result against each.
4. **Exceptions** — each with amount, disposition, and whether it affects the conclusion.
5. **Tickmark Legend** — every mark defined, with usage count.
6. **Cross-References** — other workpapers referenced, so the file hangs together.

**Then, in chat:** whether the workpaper is accepted, and if not, exactly which element is missing or
which claim the documentation does not support. Be specific — "the conclusion states no exceptions
while three are recorded" is actionable; "documentation deficiency" is not.

## What this deliberately does not do

- **It does not write your conclusion.** It validates that the conclusion you wrote is supported.
  Reaching the conclusion is the professional judgement and it is the one thing that cannot be
  delegated.
- **It does not decide whether an exception is material** or isolated. It requires that you state a
  disposition.
- **It does not assess whether the procedure was appropriate** for the assertion. A well-documented
  wrong procedure is still a wrong procedure, and that is a planning and review question.

## What to escalate

- **A conclusion the documentation does not support.** The most-cited deficiency; fix it before the
  file is assembled rather than explaining it afterwards.
- **An exception with no disposition** at the point the workpaper is signed.
- **A workpaper reviewed by its preparer.**
- **Evidence described generically** — "per client", "per discussion", "reviewed" — without a
  document reference. Inquiry alone is rarely sufficient evidence, and where it is the sole basis the
  workpaper should say so explicitly rather than implying more.
- **A population that does not agree to the general ledger**, which means the procedure covered
  something other than the account.
- **A significant matter where the evidence contradicts the conclusion** and the contradiction is not
  addressed. Documenting how a contradiction was resolved is required, and omitting it is worse than
  the contradiction itself.
- **Documentation added after the assembly deadline** without being marked as such.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
