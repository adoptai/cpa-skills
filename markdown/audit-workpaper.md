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

---

# Appendix — bundled files

If you installed the `.skill` package these files are already in place and you can ignore this appendix. If you copied the skill as text, create the files below at the paths shown, alongside your `SKILL.md`. The skill will not run without them.

## `scripts/build_workpaper.py`

```python
#!/usr/bin/env python3
"""
Build an audit workpaper and validate that its conclusion is actually supported.

The most frequently cited documentation deficiency is not a missing signature. It is
a conclusion the documentation does not support - most often "no exceptions noted" on
a workpaper that never records what was examined, or a conclusion reached while
exceptions sit unaddressed.

So the validation here is a CONSISTENCY CHECK between the conclusion and everything
above it, not a completeness checklist.

Gates:
  * objective, population, procedure, evidence, conclusion, preparer, reviewer present
  * the population is stated with its source and total, and agrees to a control figure
  * every procedure step records the evidence examined
  * the conclusion may not assert an absence of problems while exceptions exist, a
    difference is non-zero, or a step has no evidence
  * every exception has a disposition
  * preparer and reviewer are different people
  * every tickmark is both defined and used

It does NOT write the conclusion. Reaching it is the professional judgement, and that
is the one thing that cannot be delegated.

Usage:
    python3 build_workpaper.py --header header.csv --procedures procedures.csv \
        --evidence evidence.csv --exceptions exceptions.csv \
        --tickmarks tickmarks.csv --conclusion conclusion.txt \
        --out "C-210 Trade Receivables Existence.xlsx"

--header CSV (key,value rows):
    client, period_end, workpaper_ref, title, objective, assertion,
    population_description, population_source, population_how_obtained,
    population_total, population_count, control_figure, control_figure_source,
    preparer, preparer_date, reviewer, reviewer_date, cross_references

--procedures CSV:
    step, procedure_performed, result, difference, evidence_ref, tickmark
--evidence CSV:
    evidence_ref, description, document_type, obtained_from, date_obtained
--exceptions CSV:
    exception_id, step, description, amount, disposition, affects_conclusion
--tickmarks CSV:
    mark, definition
--conclusion: plain text file
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is required.  pip install openpyxl")

ZERO = Decimal("0.00")
MONEY = '#,##0.00;[Red](#,##0.00)'

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
TOP = Border(top=Side(style="thin"))

REQUIRED_HEADER = ["client", "period_end", "workpaper_ref", "title", "objective",
                   "population_description", "population_source",
                   "population_how_obtained", "population_total",
                   "preparer", "reviewer"]

# Language asserting an absence of problems. Present in the conclusion while
# exceptions exist = the most-cited documentation deficiency.
NO_PROBLEM_PHRASES = [
    "no exception", "without exception", "no exceptions noted", "none noted",
    "no differences", "no difference", "no misstatement", "no errors", "no issues",
    "properly stated", "fairly stated", "accurately stated", "correctly stated",
    "agreed without", "all agreed", "fully agreed", "no variances",
    "satisfactory", "operating effectively", "operated effectively", "no deviations",
    "appears reasonable with no", "nothing came to our attention",
]
# Positive-assurance language, which cannot rest on undocumented work. Matched as
# word stems rather than exact phrases: an earlier version required "supports the" and
# rejected a well-written conclusion saying "to support the existence of...", which is
# the kind of false positive that gets a tool abandoned.
ASSURANCE_STEMS = [
    "support", "sufficient", "appropriate", "conclude", "conclusion", "fairly stated",
    "properly stated", "accurately stated", "reasonable", "objective was met",
    "objective has been met", "adequate", "satisfied", "effective", "no further work",
    "evidence obtained", "in our opinion", "we are satisfied",
]
GENERIC_EVIDENCE = ("per client", "per discussion", "per management", "reviewed",
                    "inquiry", "as discussed", "verbal", "per conversation",
                    "noted", "n/a", "none")


def dec(raw):
    if raw is None:
        return None
    s = str(raw).strip()
    if s in ("", "-", "--", "n/a", "N/A", "None"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace(",", "").replace("$", "").strip()
    if s.endswith("-"):
        neg, s = True, s[:-1]
    try:
        v = Decimal(s)
    except InvalidOperation:
        return None
    return -v if neg else v


def clean(s) -> str:
    return " ".join(str(s or "").split())


def rows_of(path: Path, label: str):
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No rows in {label} ({path})")
    return rows


def load_header(path: Path) -> dict:
    out = {}
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rdr = csv.reader(fh)
        for row in rdr:
            if len(row) >= 2 and clean(row[0]):
                k = clean(row[0]).lower().replace(" ", "_")
                if k in ("key", "field"):
                    continue
                out[k] = clean(row[1])
    return out


def load_optional(path: Path | None, label: str) -> list[dict]:
    if not path or not path.exists():
        return []
    return [{k: clean(v) for k, v in r.items()} for r in rows_of(path, label)]


# ------------------------------------------------------------------ validation

def validate(header, procs, evidence, exceptions, tickmarks, conclusion) -> list[dict]:
    tests = []
    ev_refs = {e.get("evidence_ref", "").lower() for e in evidence
               if e.get("evidence_ref")}

    # ---- Test 1: required elements
    missing = [f for f in REQUIRED_HEADER if not header.get(f)]
    if not procs:
        missing.append("procedures (none supplied)")
    if not evidence:
        missing.append("evidence (none supplied)")
    if not conclusion.strip():
        missing.append("conclusion")
    tests.append({
        "num": 1, "name": "Required elements present", "passed": not missing,
        "detail": "" if not missing else
                  "missing: " + ", ".join(missing) +
                  ". A workpaper missing any of these is incomplete by definition.",
    })

    # ---- Test 2: population stated and tied
    pop_total = dec(header.get("population_total"))
    control = dec(header.get("control_figure"))
    pop_problems = []
    if pop_total is None:
        pop_problems.append("population total not stated - a procedure performed on an "
                            "unstated population supports nothing")
    if not header.get("population_how_obtained"):
        pop_problems.append("how the population was obtained is not stated")
    if control is not None and pop_total is not None and control != pop_total:
        pop_problems.append(f"population total {pop_total:,.2f} does not agree to the "
                            f"control figure {control:,.2f} (difference "
                            f"{pop_total - control:,.2f}) - the procedure covered "
                            f"something other than the account")
    if control is None:
        pop_problems.append("no control figure supplied, so the population is untied "
                            "(state one, or say explicitly why none applies)")
    tests.append({
        "num": 2, "name": "Population stated, sourced, and tied",
        "passed": not pop_problems,
        "detail": "; ".join(pop_problems),
    })

    # ---- Test 3: every step has evidence
    step_problems = []
    for p in procs:
        step = p.get("step") or "(unnumbered)"
        ref = p.get("evidence_ref", "")
        if not ref:
            step_problems.append(f"step {step}: no evidence reference - a step with a "
                                 f"result and no evidence is an assertion")
        elif ref.lower() not in ev_refs:
            step_problems.append(f"step {step}: evidence reference '{ref}' is not in the "
                                 f"evidence schedule")
        if not p.get("procedure_performed"):
            step_problems.append(f"step {step}: no description of what was performed")
        if not p.get("result"):
            step_problems.append(f"step {step}: no result recorded")
    for e in evidence:
        d = (e.get("description") or "").lower()
        if any(g == d or d.startswith(g) for g in GENERIC_EVIDENCE):
            step_problems.append(
                f"evidence {e.get('evidence_ref')}: described generically as "
                f"\"{e.get('description')}\". Inquiry alone is rarely sufficient - name "
                f"the document, or state explicitly that inquiry was the sole basis")
    tests.append({
        "num": 3, "name": "Every procedure step records its evidence",
        "passed": not step_problems, "detail": "; ".join(step_problems[:8]),
    })

    # ---- Test 4: conclusion consistency  (the core test)
    concl = conclusion.lower()
    exc_count = len(exceptions)
    nonzero_diffs = [p for p in procs
                     if dec(p.get("difference")) not in (None, ZERO)]
    undispositioned = [e for e in exceptions if not e.get("disposition")]
    no_evidence_steps = [p for p in procs if not p.get("evidence_ref")]

    found_no_problem = [ph for ph in NO_PROBLEM_PHRASES if ph in concl]
    found_assurance = [st for st in ASSURANCE_STEMS if st in concl]
    word_count = len(concl.split())

    concl_problems = []
    if found_no_problem and exc_count:
        concl_problems.append(
            f"the conclusion asserts an absence of problems (\"{found_no_problem[0]}\") "
            f"while {exc_count} exception(s) are recorded. This is the most frequently "
            f"cited documentation deficiency. The conclusion must acknowledge the "
            f"exceptions and explain why they do not change the outcome")
    if found_no_problem and nonzero_diffs:
        concl_problems.append(
            f"the conclusion asserts an absence of differences (\"{found_no_problem[0]}\") "
            f"while step(s) "
            f"{', '.join(str(p.get('step')) for p in nonzero_diffs[:4])} report a non-zero "
            f"difference")
    if undispositioned:
        concl_problems.append(
            f"{len(undispositioned)} exception(s) have no disposition. An exception left "
            f"open blocks the conclusion")
    if found_assurance and no_evidence_steps:
        concl_problems.append(
            f"the conclusion draws positive assurance (\"{found_assurance[0]}\") while "
            f"{len(no_evidence_steps)} step(s) record no evidence. A conclusion cannot "
            f"rest on work that was not documented")
    # Only fault a conclusion for stating no outcome when it is genuinely thin. A long
    # conclusion that discusses exceptions and materiality is a real conclusion even if
    # it avoids this tool's vocabulary; a fifteen-word one that restates the procedure
    # is not.
    if conclusion.strip() and not found_assurance and not found_no_problem \
            and word_count < 25:
        concl_problems.append(
            f"the conclusion is {word_count} words and states no outcome. A conclusion is "
            f"a logical statement connecting the evidence to the objective - not a "
            f"restatement of the procedure")
    if exc_count and not any(w in concl for w in
                             ("exception", "difference", "misstatement", "error",
                              "deviation", "finding", "adjust")):
        concl_problems.append(
            f"{exc_count} exception(s) are recorded but the conclusion does not mention "
            f"exceptions at all. An exception the conclusion ignores is one the reviewer "
            f"has to find themselves")
    tests.append({
        "num": 4, "name": "Conclusion is consistent with the exceptions and the evidence",
        "passed": not concl_problems, "detail": "; ".join(concl_problems),
    })

    # ---- Test 5: exceptions dispositioned
    exc_problems = []
    for e in exceptions:
        eid = e.get("exception_id") or "(unnumbered)"
        if not e.get("disposition"):
            exc_problems.append(f"{eid}: no disposition")
        if not e.get("description"):
            exc_problems.append(f"{eid}: no description")
        if not e.get("affects_conclusion"):
            exc_problems.append(f"{eid}: does not state whether it affects the conclusion")
    tests.append({
        "num": 5, "name": "Every exception has a disposition",
        "passed": not exc_problems, "detail": "; ".join(exc_problems[:8]),
    })

    # ---- Test 6: preparer and reviewer
    prep, rev = header.get("preparer", ""), header.get("reviewer", "")
    signoff = []
    if prep and rev and prep.strip().lower() == rev.strip().lower():
        signoff.append(f"preparer and reviewer are the same person ({prep}) - a workpaper "
                       f"reviewed by its preparer has not been reviewed")
    if prep and not header.get("preparer_date"):
        signoff.append("preparer date missing")
    if rev and not header.get("reviewer_date"):
        signoff.append("reviewer date missing")
    tests.append({
        "num": 6, "name": "Preparer and reviewer named, different, and dated",
        "passed": not signoff, "detail": "; ".join(signoff),
    })

    # ---- Test 7: tickmarks
    defined = {t.get("mark", "").strip() for t in tickmarks if t.get("mark")}
    used = {p.get("tickmark", "").strip() for p in procs if p.get("tickmark")}
    tick = []
    for u in sorted(used - defined):
        tick.append(f"tickmark '{u}' is used but not defined")
    for d in sorted(defined - used):
        tick.append(f"tickmark '{d}' is defined but never used")
    for t in tickmarks:
        if t.get("mark") and not t.get("definition"):
            tick.append(f"tickmark '{t['mark']}' has no definition")
    tests.append({
        "num": 7, "name": "Every tickmark defined and used",
        "passed": not tick, "detail": "; ".join(tick[:8]),
    })

    return tests


# -------------------------------------------------------------------- workbook

def hdr(ws, n, row=1):
    for c in range(1, n + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 28


def widths(ws, w):
    for c, v in w.items():
        ws.column_dimensions[get_column_letter(c)].width = v


def sheet_workpaper(wb, header, procs, evidence, exceptions, conclusion, accepted):
    ws = wb.active
    ws.title = "Workpaper"
    widths(ws, {1: 4, 2: 26, 3: 96})
    r = 1

    def block(label, value="", *, bold=False, size=11, fill=None, wrap=True):
        nonlocal r
        lc = ws.cell(row=r, column=2, value=label)
        lc.font = Font(bold=True, size=size)
        lc.alignment = Alignment(vertical="top")
        if fill:
            lc.fill = fill
        vc = ws.cell(row=r, column=3, value=value)
        vc.font = Font(bold=bold, size=size)
        vc.alignment = Alignment(wrap_text=wrap, vertical="top")
        r += 1

    def rule(title):
        nonlocal r
        c = ws.cell(row=r, column=2, value=title)
        c.font, c.fill = Font(bold=True, size=12), SUB_FILL
        ws.cell(row=r, column=3).fill = SUB_FILL
        r += 1

    ws.cell(row=r, column=2, value=header.get("title") or "AUDIT WORKPAPER").font = \
        Font(bold=True, size=14)
    r += 1
    if not accepted:
        c = ws.cell(row=r, column=2,
                    value="DRAFT - VALIDATION FAILED. See the Validation tab. Do not file "
                          "until the conclusion is supported by the documentation.")
        c.font, c.fill = BAD_FONT, BAD_FILL
        c.alignment = Alignment(wrap_text=True)
        r += 1
    r += 1

    rule("IDENTIFICATION")
    for k, label in [("client", "Client"), ("period_end", "Period end"),
                     ("workpaper_ref", "Workpaper reference"),
                     ("assertion", "Assertion / control tested")]:
        block(label, header.get(k, "") or "(not stated)")
    r += 1

    rule("OBJECTIVE")
    block("Objective", header.get("objective", "") or "(NOT STATED)")
    r += 1

    rule("POPULATION")
    for k, label in [("population_description", "Description"),
                     ("population_source", "Source"),
                     ("population_how_obtained", "How obtained"),
                     ("population_count", "Item count"),
                     ("population_total", "Total"),
                     ("control_figure", "Control figure"),
                     ("control_figure_source", "Control figure source")]:
        block(label, header.get(k, "") or "(not stated)")
    pt, cf = dec(header.get("population_total")), dec(header.get("control_figure"))
    if pt is not None and cf is not None:
        diff = pt - cf
        block("Difference to control", f"{diff:,.2f}"
              + ("  — AGREED" if diff == ZERO else "  — DOES NOT AGREE"),
              bold=True, fill=None if diff == ZERO else BAD_FILL)
    r += 1

    rule("PROCEDURES PERFORMED")
    for h, col in zip(["Step", "Procedure performed", "Result", "Difference",
                       "Evidence", "TM"], (2, 3, 4, 5, 6, 7)):
        c = ws.cell(row=r, column=col, value=h)
        c.fill, c.font = HDR_FILL, HDR_FONT
    r += 1
    for p in procs:
        ws.cell(row=r, column=2, value=p.get("step"))
        c = ws.cell(row=r, column=3, value=p.get("procedure_performed"))
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c2 = ws.cell(row=r, column=4, value=p.get("result"))
        c2.alignment = Alignment(wrap_text=True, vertical="top")
        d = dec(p.get("difference"))
        if d is not None:
            dc = ws.cell(row=r, column=5, value=float(d))
            dc.number_format = MONEY
            if d != ZERO:
                dc.font, dc.fill = BAD_FONT, BAD_FILL
        ec = ws.cell(row=r, column=6, value=p.get("evidence_ref") or None)
        if not p.get("evidence_ref"):
            ec.fill = BAD_FILL
        ws.cell(row=r, column=7, value=p.get("tickmark") or None)
        r += 1
    r += 1

    rule("EVIDENCE EXAMINED")
    for h, col in zip(["Ref", "Description", "Type", "Obtained from", "Date"],
                      (2, 3, 4, 5, 6)):
        c = ws.cell(row=r, column=col, value=h)
        c.fill, c.font = HDR_FILL, HDR_FONT
    r += 1
    for e in evidence:
        ws.cell(row=r, column=2, value=e.get("evidence_ref"))
        c = ws.cell(row=r, column=3, value=e.get("description"))
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=r, column=4, value=e.get("document_type"))
        ws.cell(row=r, column=5, value=e.get("obtained_from"))
        ws.cell(row=r, column=6, value=e.get("date_obtained"))
        r += 1
    r += 1

    rule("EXCEPTIONS")
    if not exceptions:
        block("", "None identified.")
    else:
        for h, col in zip(["ID", "Description", "Amount", "Disposition",
                           "Affects conclusion"], (2, 3, 4, 5, 6)):
            c = ws.cell(row=r, column=col, value=h)
            c.fill, c.font = HDR_FILL, HDR_FONT
        r += 1
        for e in exceptions:
            ws.cell(row=r, column=2, value=e.get("exception_id"))
            c = ws.cell(row=r, column=3, value=e.get("description"))
            c.alignment = Alignment(wrap_text=True, vertical="top")
            a = dec(e.get("amount"))
            if a is not None:
                ac = ws.cell(row=r, column=4, value=float(a))
                ac.number_format = MONEY
            dc = ws.cell(row=r, column=5, value=e.get("disposition") or None)
            if not e.get("disposition"):
                dc.fill = BAD_FILL
            ws.cell(row=r, column=6, value=e.get("affects_conclusion") or None)
            r += 1
    r += 1

    rule("CONCLUSION")
    c = ws.cell(row=r, column=3, value=conclusion.strip() or "(NOT STATED)")
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = max(30, min(160, 16 * (len(conclusion) // 90 + 2)))
    r += 2

    rule("SIGN-OFF")
    block("Prepared by", f"{header.get('preparer', '(not stated)')}    "
                         f"{header.get('preparer_date', '')}")
    block("Reviewed by", f"{header.get('reviewer', '(not stated)')}    "
                         f"{header.get('reviewer_date', '')}")
    if header.get("cross_references"):
        r += 1
        block("Cross-references", header["cross_references"])


def sheet_validation(wb, tests, accepted):
    ws = wb.create_sheet("Validation")
    widths(ws, {1: 4, 2: 8, 3: 52, 4: 12, 5: 96})
    r = 1
    ws.cell(row=r, column=2, value="VALIDATION").font = Font(bold=True, size=14)
    r += 1
    c = ws.cell(row=r, column=2,
                value="ACCEPTED - the conclusion is supported by the documentation"
                if accepted else
                "NOT ACCEPTED - the documentation does not support the conclusion")
    c.font = OK_FONT if accepted else BAD_FONT
    if not accepted:
        c.fill = BAD_FILL
    r += 2
    for h, col in zip(["Test", "Name", "Result", "What failed and why"], (2, 3, 4, 5)):
        cc = ws.cell(row=r, column=col, value=h)
        cc.fill, cc.font = HDR_FILL, HDR_FONT
    r += 1
    for t in tests:
        ws.cell(row=r, column=2, value=t["num"])
        ws.cell(row=r, column=3, value=t["name"]).alignment = \
            Alignment(wrap_text=True, vertical="top")
        v = ws.cell(row=r, column=4, value="PASS" if t["passed"] else "FAIL")
        v.font = OK_FONT if t["passed"] else BAD_FONT
        if not t["passed"]:
            v.fill = BAD_FILL
        d = ws.cell(row=r, column=5, value=t["detail"] or None)
        d.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    r += 2
    for note in [
        "This tab is part of preparing the workpaper, not part of the workpaper itself.",
        "",
        "Test 4 is the one that matters. The most frequently cited documentation",
        "deficiency is a conclusion the documentation does not support - typically",
        "\"no exceptions noted\" on a workpaper that never records what was examined.",
        "",
        "It is deliberately hard to satisfy by rewording. If exceptions exist, the",
        "conclusion must acknowledge them and explain why they do not change the outcome,",
        "which is what a defensible conclusion looks like anyway.",
    ]:
        ws.cell(row=r, column=2, value=note).font = Font(italic=True)
        r += 1


def sheet_list(wb, title, header, rows, fmt, empty, footer=""):
    ws = wb.create_sheet(title)
    ws.append(header)
    hdr(ws, len(header))
    for row in rows:
        ws.append(row)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i, f in fmt.items():
            if i < len(row):
                row[i].number_format = f
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value=empty)
    elif footer:
        rr = ws.max_row + 2
        c = ws.cell(row=rr, column=1, value=footer)
        c.font = Font(italic=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    widths(ws, {i: 22 for i in range(1, len(header) + 1)} | {2: 48,
                                                            len(header): 40})


# ------------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--header", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--procedures")
    ap.add_argument("--evidence")
    ap.add_argument("--exceptions")
    ap.add_argument("--tickmarks")
    ap.add_argument("--conclusion", help="plain text file containing the conclusion")
    ap.add_argument("--force", action="store_true",
                    help="write a workbook marked DRAFT even when validation fails")
    args = ap.parse_args()

    header = load_header(Path(args.header))
    procs = load_optional(Path(args.procedures) if args.procedures else None,
                          "procedures")
    evidence = load_optional(Path(args.evidence) if args.evidence else None, "evidence")
    exceptions = load_optional(Path(args.exceptions) if args.exceptions else None,
                               "exceptions")
    tickmarks = load_optional(Path(args.tickmarks) if args.tickmarks else None,
                              "tickmarks")
    conclusion = ""
    if args.conclusion and Path(args.conclusion).exists():
        conclusion = Path(args.conclusion).read_text(encoding="utf-8")

    tests = validate(header, procs, evidence, exceptions, tickmarks, conclusion)
    accepted = all(t["passed"] for t in tests)

    # ---- console
    print("=" * 78)
    print("AUDIT WORKPAPER")
    print("=" * 78)
    print(f"Reference : {header.get('workpaper_ref', '(not stated)')}")
    print(f"Title     : {header.get('title', '(not stated)')}")
    print(f"Client    : {header.get('client', '(not stated)')}   "
          f"Period end: {header.get('period_end', '(not stated)')}")
    print(f"Steps {len(procs)}   evidence items {len(evidence)}   "
          f"exceptions {len(exceptions)}   tickmarks {len(tickmarks)}")
    print()
    for t in tests:
        print(f"Test {t['num']}: {'PASS' if t['passed'] else '*** FAIL ***':<14} "
              f"{t['name']}")
        if t["detail"]:
            for chunk in t["detail"].split("; "):
                print(f"         - {chunk}")
    print()
    if accepted:
        print("ACCEPTED - the conclusion is supported by the documentation.")
    else:
        print("NOT ACCEPTED - the documentation does not support the conclusion.")

    if not accepted and not args.force:
        print("\n" + "=" * 78)
        print("WORKBOOK NOT WRITTEN.")
        print("A workpaper has one job: demonstrate the work was done and that the")
        print("conclusion follows from it. Fix the specific failure above - this tool does")
        print("not write your conclusion, it validates that the one you wrote is")
        print("supported. Use --force to produce a DRAFT-marked file for editing.")
        print("=" * 78)
        return 1

    wb = Workbook()
    sheet_workpaper(wb, header, procs, evidence, exceptions, conclusion, accepted)
    sheet_validation(wb, tests, accepted)
    sheet_list(wb, "Procedures and Evidence",
               ["Step", "Procedure performed", "Result", "Difference", "Evidence ref",
                "Evidence description", "Tickmark"],
               [[p.get("step"), p.get("procedure_performed"), p.get("result"),
                 float(dec(p.get("difference"))) if dec(p.get("difference")) is not None
                 else None,
                 p.get("evidence_ref"),
                 next((e.get("description") for e in evidence
                       if e.get("evidence_ref", "").lower()
                       == p.get("evidence_ref", "").lower()), "(NOT IN SCHEDULE)"),
                 p.get("tickmark")] for p in procs],
               {3: MONEY}, "No procedures recorded.",
               footer="A step with a result and no evidence is an assertion, not a "
                      "procedure.")
    sheet_list(wb, "Exceptions",
               ["ID", "Step", "Description", "Amount", "Disposition",
                "Affects conclusion"],
               [[e.get("exception_id"), e.get("step"), e.get("description"),
                 float(dec(e.get("amount"))) if dec(e.get("amount")) is not None else None,
                 e.get("disposition") or "(NONE - BLOCKS THE CONCLUSION)",
                 e.get("affects_conclusion")] for e in exceptions],
               {3: MONEY}, "None identified.",
               footer="An exception with no disposition blocks the conclusion. An exception "
                      "the conclusion does not mention is one the reviewer has to find.")
    sheet_list(wb, "Tickmark Legend", ["Mark", "Definition", "Times used"],
               [[t.get("mark"), t.get("definition"),
                 len([p for p in procs
                      if p.get("tickmark", "").strip() == t.get("mark", "").strip()])]
                for t in tickmarks],
               {}, "No tickmarks used.",
               footer="Every mark used must be defined, and every mark defined must be "
                      "used.")
    xrefs = [x.strip() for x in (header.get("cross_references") or "").split(",")
             if x.strip()]
    sheet_list(wb, "Cross-References", ["Workpaper referenced", "Purpose"],
               [[x, None] for x in xrefs], {},
               "None recorded. If this workpaper depends on another, reference it so the "
               "file hangs together.")
    wb.active = 0
    out = Path(args.out)
    if not accepted:
        out = out.with_name(out.stem + " [DRAFT - NOT ACCEPTED]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\nWorkbook: {out}")
    return 0 if accepted else 1


if __name__ == "__main__":
    sys.exit(main())
```

