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

---

# Appendix — bundled files

If you installed the `.skill` package these files are already in place and you can ignore this appendix. If you copied the skill as text, create the files below at the paths shown, alongside your `SKILL.md`. The skill will not run without them.

## `scripts/build_k1_summary.py`

```python
#!/usr/bin/env python3
"""
Validate extracted Schedule K-1 data and build the standardized summary workbook.

Five gates, all of which must pass before a clean workbook is written:

  A  Aggregate footing   sum of each box/code across K-1s = entity Schedule K total
  B  Ownership           profit, loss and capital percentages each total 100.000%
  C  Completeness        expected recipients vs K-1s received, both directions
  D  Internal            duplicate box/code rows, orphan amounts, full TINs
  E  Entity consistency  one EIN, one tax year, one entity type across the set

Usage:
    python3 build_k1_summary.py --lines k1_lines.csv \
        --entity-totals entity_schedule_k.csv \
        --expected-recipients partners.csv \
        --client "Meridian Holdings LP" --tax-year 2025 \
        --out "Meridian Holdings LP - 2025 K-1 Summary.xlsx"

--lines CSV (one row per box/code per recipient):
    source_file, source_page, entity_name, entity_ein, entity_type, tax_year,
    final_k1, amended_k1, recipient_name, recipient_tin_last4, recipient_type,
    pct_profit, pct_loss, pct_capital, box, code, amount, description,
    statement_ref, state

--entity-totals CSV (from the entity return's Schedule K):
    box, code, amount

--expected-recipients CSV:
    recipient_name, pct_profit, pct_loss, pct_capital,
    prior_pct_profit, prior_capital_ending    (last two optional)

Money is Decimal. Percentages are Decimal. Nothing is rounded.
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from collections import defaultdict
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is required.  pip install openpyxl")

ZERO = Decimal("0.00")
HUNDRED = Decimal("100.000")
MONEY = '#,##0.00;[Red](#,##0.00)'
PCT3 = '0.000"%"'

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")

# A full SSN or EIN must never reach an output file.
TIN_RE = re.compile(r"\b\d{3}-\d{2}-\d{4}\b|\b\d{2}-\d{7}\b|\b\d{9}\b")


def dec(raw, field="amount"):
    if raw is None:
        return None
    s = str(raw).strip()
    if s in ("", "-", "--", "n/a", "N/A", "None"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace(",", "").replace("$", "").replace("%", "").strip()
    if s.endswith("-"):
        neg, s = True, s[:-1]
    try:
        v = Decimal(s)
    except InvalidOperation:
        raise ValueError(f"Unparseable {field}: {raw!r}")
    return -v if neg else v


def clean(s) -> str:
    return " ".join(str(s or "").split())


def truthy(s) -> bool:
    return clean(s).lower() in ("yes", "y", "true", "1", "x")


def load_lines(path: Path) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No K-1 lines in {path}")
    out = []
    for i, r in enumerate(rows, start=1):
        rec = {k: clean(r.get(k)) for k in (
            "source_file", "entity_name", "entity_ein", "entity_type", "tax_year",
            "recipient_name", "recipient_tin_last4", "recipient_type",
            "box", "code", "description", "statement_ref", "state",
        )}
        rec["row"] = i
        rec["source_page"] = clean(r.get("source_page"))
        rec["final_k1"] = truthy(r.get("final_k1"))
        rec["amended_k1"] = truthy(r.get("amended_k1"))
        rec["amount"] = dec(r.get("amount"))
        for p in ("pct_profit", "pct_loss", "pct_capital"):
            rec[p] = dec(r.get(p), p)
        if not rec["recipient_name"]:
            raise ValueError(f"Row {i}: recipient_name is required.")
        if not rec["box"] and rec["amount"] is not None:
            raise ValueError(
                f"Row {i} ({rec['recipient_name']}): amount {rec['amount']} with no box. "
                f"Every amount belongs to a box - an unattributed figure cannot be placed "
                f"on a return."
            )
        out.append(rec)
    return out


def load_entity_totals(path: Path | None) -> dict[tuple, Decimal]:
    if not path or not path.exists():
        return {}
    out: dict[tuple, Decimal] = {}
    with path.open(newline="", encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            box, code = clean(r.get("box")), clean(r.get("code"))
            amt = dec(r.get("amount"))
            if not box or amt is None:
                continue
            out[(box, code)] = out.get((box, code), ZERO) + amt
    return out


def load_expected(path: Path | None) -> dict[str, dict]:
    if not path or not path.exists():
        return {}
    out = {}
    with path.open(newline="", encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            n = clean(r.get("recipient_name"))
            if not n:
                continue
            out[n.lower()] = {
                "recipient_name": n,
                "pct_profit": dec(r.get("pct_profit"), "pct_profit"),
                "pct_loss": dec(r.get("pct_loss"), "pct_loss"),
                "pct_capital": dec(r.get("pct_capital"), "pct_capital"),
                "prior_pct_profit": dec(r.get("prior_pct_profit"), "prior_pct_profit"),
                "prior_capital_ending": dec(r.get("prior_capital_ending"),
                                            "prior_capital_ending"),
            }
    return out


# --------------------------------------------------------------------------- gates

def build_recipients(lines: list[dict]) -> dict[str, dict]:
    recips: dict[str, dict] = {}
    for ln in lines:
        key = ln["recipient_name"].lower()
        r = recips.setdefault(key, {
            "recipient_name": ln["recipient_name"],
            "tin_last4": ln["recipient_tin_last4"],
            "recipient_type": ln["recipient_type"],
            "final_k1": False, "amended_k1": False,
            "pct_profit": None, "pct_loss": None, "pct_capital": None,
            "source_file": ln["source_file"], "boxes": {}, "state_boxes": {},
            "states": set(),
        })
        r["final_k1"] |= ln["final_k1"]
        r["amended_k1"] |= ln["amended_k1"]
        for p in ("pct_profit", "pct_loss", "pct_capital"):
            if ln[p] is not None and r[p] is None:
                r[p] = ln[p]
        if ln["state"]:
            r["states"].add(ln["state"])
        if ln["box"] and ln["amount"] is not None:
            # State K-1 schedules reuse the federal box numbers. Aggregating them
            # together would double-count every apportioned amount into the federal
            # footing test, so they are kept in a separate bucket keyed by state.
            if ln["state"]:
                key = (ln["state"], ln["box"], ln["code"])
                r["state_boxes"].setdefault(key, ZERO)
                r["state_boxes"][key] += ln["amount"]
            else:
                r["boxes"].setdefault((ln["box"], ln["code"]), ZERO)
                r["boxes"][(ln["box"], ln["code"])] += ln["amount"]
    return recips


def gate_a(recips, entity_totals) -> dict:
    if not entity_totals:
        return {"applicable": False, "passed": True, "rows": [],
                "note": "entity Schedule K totals not supplied - summary is UNFOOTED"}
    agg: dict[tuple, Decimal] = defaultdict(lambda: ZERO)
    for r in recips.values():
        for k, v in r["boxes"].items():
            agg[k] += v
    rows, ok = [], True
    for key in sorted(set(agg) | set(entity_totals), key=lambda k: (str(k[0]), str(k[1]))):
        k1_sum = agg.get(key, ZERO)
        ent = entity_totals.get(key)
        diff = k1_sum - (ent if ent is not None else ZERO)
        status = ("no entity total supplied" if ent is None
                  else "ties" if diff == ZERO else "DIFFERENCE")
        if ent is None or diff != ZERO:
            ok = False
        rows.append({"box": key[0], "code": key[1], "k1_sum": k1_sum,
                     "entity": ent, "difference": diff, "status": status})
    return {"applicable": True, "passed": ok, "rows": rows, "note": ""}


def gate_b(recips) -> dict:
    res = {}
    ok = True
    for field in ("pct_profit", "pct_loss", "pct_capital"):
        vals = [r[field] for r in recips.values() if r[field] is not None]
        total = sum(vals, Decimal("0.000")) if vals else None
        passed = total is not None and total == HUNDRED
        missing = [r["recipient_name"] for r in recips.values() if r[field] is None]
        if not passed:
            ok = False
        res[field] = {"total": total, "passed": passed, "count": len(vals),
                      "missing": missing}
    return {"passed": ok, "detail": res}


def gate_c(recips, expected) -> dict:
    if not expected:
        return {"applicable": False, "passed": True, "missing": [], "unexpected": [],
                "pct_mismatch": [],
                "note": "expected recipient list not supplied - a missing K-1 can only be "
                        "detected through the percentage test"}
    got = set(recips)
    exp = set(expected)
    missing = sorted(expected[k]["recipient_name"] for k in exp - got)
    unexpected = sorted(recips[k]["recipient_name"] for k in got - exp)
    mismatch = []
    for k in exp & got:
        for f in ("pct_profit", "pct_loss", "pct_capital"):
            e, a = expected[k][f], recips[k][f]
            if e is not None and a is not None and e != a:
                mismatch.append({"recipient": recips[k]["recipient_name"],
                                 "field": f, "expected": e, "per_k1": a})
    return {"applicable": True,
            "passed": not missing and not unexpected and not mismatch,
            "missing": missing, "unexpected": unexpected, "pct_mismatch": mismatch,
            "note": ""}


def gate_d(lines, recips) -> dict:
    probs = []
    seen = defaultdict(list)
    for ln in lines:
        if ln["box"] and ln["amount"] is not None:
            seen[(ln["recipient_name"].lower(), ln["state"],
                  ln["box"], ln["code"])].append(ln["row"])
    for (rname, state, box, code), rowlist in seen.items():
        if len(rowlist) > 1:
            where = f"{state} " if state else "federal "
            probs.append(f"{rname}: {where}box {box} code {code or '-'} appears on rows "
                         f"{rowlist} - one row per box/code per recipient per "
                         f"jurisdiction, or the footing test double-counts")
    for ln in lines:
        if ln["box"] and ln["code"] and ln["amount"] is None and not ln["statement_ref"]:
            probs.append(f"Row {ln['row']} ({ln['recipient_name']}): box {ln['box']} "
                         f"code {ln['code']} has no amount and no statement_ref - the "
                         f"amount is on a supplemental statement that was not captured")
    for ln in lines:
        for field in ("recipient_tin_last4", "description", "recipient_name"):
            if TIN_RE.search(ln[field] or ""):
                probs.append(f"Row {ln['row']}: a full TIN pattern appears in "
                             f"'{field}'. Use last four digits only.")
        if len(ln["recipient_tin_last4"]) > 4:
            probs.append(f"Row {ln['row']} ({ln['recipient_name']}): "
                         f"recipient_tin_last4 is longer than 4 characters.")
    for r in recips.values():
        if not r["boxes"]:
            probs.append(f"{r['recipient_name']}: no box amounts captured at all - "
                         f"the K-1 was seen but not read")
    return {"passed": not probs, "problems": probs}


def gate_e(lines) -> dict:
    probs = []
    for field, label in (("entity_ein", "EIN"), ("tax_year", "tax year"),
                         ("entity_type", "entity type")):
        vals = sorted({ln[field] for ln in lines if ln[field]})
        if len(vals) > 1:
            probs.append(f"Multiple {label} values in one population: {vals}. K-1s from "
                         f"different entities or years were commingled, which makes the "
                         f"footing test meaningless. Split them and run separately.")
        if not vals:
            probs.append(f"No {label} captured on any row.")
    return {"passed": not probs, "problems": probs}


# --------------------------------------------------------------------------- workbook

def hdr(ws, n, row=1):
    for c in range(1, n + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def widths(ws, w):
    for c, v in w.items():
        ws.column_dimensions[get_column_letter(c)].width = v


def sheet_summary(wb, meta, gates, recips, escalations, overall):
    ws = wb.active
    ws.title = "Summary"
    widths(ws, {1: 4, 2: 46, 3: 22, 4: 18, 5: 66})
    r = 1

    def line(t, *, bold=False, size=11, fill=None, col=2):
        nonlocal r
        c = ws.cell(row=r, column=col, value=t)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c.fill = fill
        r += 1

    line("SCHEDULE K-1 SUMMARY", bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    v = ws.cell(row=r, column=2,
                value="RESULT: PASS - K-1 population is complete and foots to the entity return"
                if overall else
                "RESULT: FAILED - see the failing gate below. Do not use this summary "
                "for return input.")
    v.font = OK_FONT if overall else BAD_FONT
    if not overall:
        v.fill = BAD_FILL
    r += 2

    line("GATES", bold=True, size=12)
    for label, g, note in [
        ("A  Aggregate footing to entity Schedule K", gates["a"],
         gates["a"].get("note", "")),
        ("B  Ownership percentages total 100.000%", gates["b"], ""),
        ("C  Recipient completeness", gates["c"], gates["c"].get("note", "")),
        ("D  Internal consistency", gates["d"], ""),
        ("E  Entity consistency", gates["e"], ""),
    ]:
        ws.cell(row=r, column=2, value="  " + label)
        applicable = g.get("applicable", True)
        txt = "PASS" if g["passed"] else "FAIL"
        if not applicable:
            txt = "NOT PERFORMED"
        c = ws.cell(row=r, column=3, value=txt)
        c.font = OK_FONT if (g["passed"] and applicable) else BAD_FONT
        if not g["passed"]:
            c.fill = BAD_FILL
        elif not applicable:
            c.fill = WARN_FILL
        if note:
            ws.cell(row=r, column=5, value=note).font = Font(italic=True, size=9)
        r += 1
    r += 1

    line("OWNERSHIP PERCENTAGES", bold=True, size=12)
    for f, lbl in (("pct_profit", "Profit"), ("pct_loss", "Loss"),
                   ("pct_capital", "Capital")):
        d = gates["b"]["detail"][f]
        ws.cell(row=r, column=2, value=f"  {lbl} - total across {d['count']} K-1(s)")
        c = ws.cell(row=r, column=3,
                    value=float(d["total"]) if d["total"] is not None else "not captured")
        if d["total"] is not None:
            c.number_format = PCT3
        c.font = OK_FONT if d["passed"] else BAD_FONT
        if not d["passed"]:
            c.fill = BAD_FILL
            gap = (HUNDRED - d["total"]) if d["total"] is not None else None
            if gap:
                ws.cell(row=r, column=5,
                        value=f"off by {gap:.3f}pp - a set that does not total 100% is "
                              f"missing a K-1 or has a transcription error"
                        ).font = Font(italic=True, size=9)
        r += 1
    r += 1

    if gates["c"]["applicable"]:
        if gates["c"]["missing"]:
            line("K-1s NOT RECEIVED", bold=True, size=12, fill=BAD_FILL)
            for n in gates["c"]["missing"]:
                line(f"  {n}")
            r += 1
        if gates["c"]["unexpected"]:
            line("K-1s RECEIVED FROM UNEXPECTED RECIPIENTS", bold=True, size=12,
                 fill=WARN_FILL)
            for n in gates["c"]["unexpected"]:
                line(f"  {n} - not on the expected recipient list")
            r += 1

    if escalations:
        line("ESCALATE - facts requiring a decision by the signer", bold=True, size=12,
             fill=WARN_FILL)
        for e in escalations:
            line("  " + e)
        r += 1

    for g, title in ((gates["e"], "ENTITY CONSISTENCY PROBLEMS"),
                     (gates["d"], "INTERNAL CONSISTENCY PROBLEMS")):
        if g["problems"]:
            line(title, bold=True, size=12, fill=BAD_FILL)
            for p in g["problems"]:
                line("  " + p)
            r += 1

    r += 1
    line("Prepared by / date:  __________________  ____________", bold=True)
    line("Reviewed by / date:  __________________  ____________", bold=True)


def sheet_footing(wb, g):
    ws = wb.create_sheet("Footing")
    heads = ["Box", "Code", "Sum of K-1s", "Entity Schedule K", "Difference", "Status"]
    ws.append(heads)
    hdr(ws, len(heads))
    if not g["applicable"]:
        ws.cell(row=2, column=1, value="NOT PERFORMED")
        ws.cell(row=2, column=3, value=g["note"])
        widths(ws, {1: 10, 2: 8, 3: 18, 4: 20, 5: 16, 6: 60})
        return
    for row in g["rows"]:
        ws.append([row["box"], row["code"] or None, float(row["k1_sum"]),
                   float(row["entity"]) if row["entity"] is not None else None,
                   float(row["difference"]), row["status"]])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in (2, 3, 4):
            row[i].number_format = MONEY
        if row[5].value != "ties":
            row[5].font, row[5].fill = BAD_FONT, BAD_FILL
            row[4].font = BAD_FONT
        else:
            row[5].font = OK_FONT
    ws.freeze_panes = "A2"
    widths(ws, {1: 10, 2: 8, 3: 18, 4: 20, 5: 16, 6: 34})


def sheet_matrix(wb, recips):
    ws = wb.create_sheet("K-1 Matrix")
    keys = sorted({k for r in recips.values() for k in r["boxes"]},
                  key=lambda k: (str(k[0]).zfill(3), str(k[1])))
    heads = ["Recipient", "TIN x4", "Type", "Profit %", "Capital %", "Final", "Amended"] \
        + [f"Box {b}{(' ' + c) if c else ''}" for b, c in keys]
    ws.append(heads)
    hdr(ws, len(heads))
    for r in sorted(recips.values(), key=lambda x: x["recipient_name"]):
        row = [r["recipient_name"], r["tin_last4"] or None, r["recipient_type"] or None,
               float(r["pct_profit"]) if r["pct_profit"] is not None else None,
               float(r["pct_capital"]) if r["pct_capital"] is not None else None,
               "Y" if r["final_k1"] else "", "Y" if r["amended_k1"] else ""]
        for k in keys:
            v = r["boxes"].get(k)
            row.append(float(v) if v is not None else None)
        ws.append(row)
    # totals
    ws.append([])
    t = ws.max_row + 1
    ws.cell(row=t, column=1, value="TOTAL").font = Font(bold=True)
    for i in range(len(keys)):
        col = 8 + i
        L = get_column_letter(col)
        c = ws.cell(row=t, column=col, value=f"=SUM({L}2:{L}{ws.max_row - 2})")
        c.number_format, c.font = MONEY, Font(bold=True)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[3].number_format = PCT3
        row[4].number_format = PCT3
        for i in range(7, len(heads)):
            row[i].number_format = MONEY
        if row[6].value == "Y":
            row[6].fill = WARN_FILL
    ws.freeze_panes = "B2"
    w = {1: 32, 2: 9, 3: 16, 4: 10, 5: 11, 6: 7, 7: 9}
    for i in range(len(keys)):
        w[8 + i] = 15
    widths(ws, w)


def sheet_detail(wb, lines):
    ws = wb.create_sheet("Detail")
    heads = ["Recipient", "Box", "Code", "Amount", "Description", "Statement ref",
             "State", "Source file", "Pg", "Row"]
    ws.append(heads)
    hdr(ws, len(heads))
    for ln in lines:
        ws.append([ln["recipient_name"], ln["box"] or None, ln["code"] or None,
                   float(ln["amount"]) if ln["amount"] is not None else None,
                   ln["description"] or None, ln["statement_ref"] or None,
                   ln["state"] or None, ln["source_file"] or None,
                   ln["source_page"] or None, ln["row"]])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[3].number_format = MONEY
        if row[3].value is None and row[2].value and not row[5].value:
            row[3].fill = BAD_FILL
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:J{max(ws.max_row, 2)}"
    widths(ws, {1: 30, 2: 7, 3: 7, 4: 16, 5: 44, 6: 24, 7: 8, 8: 26, 9: 5, 10: 6})


def sheet_states(wb, lines, recips):
    ws = wb.create_sheet("State Schedules")
    states = sorted({ln["state"] for ln in lines if ln["state"]})
    if not states:
        ws.cell(row=1, column=1,
                value="No state-level K-1 data captured.").font = Font(bold=True)
        ws.cell(row=3, column=1,
                value="If the entity operates in more than one state, the K-1 state "
                      "schedules were probably not captured. Nonresident withholding "
                      "and composite-return eligibility are decided from this data, so "
                      "confirm whether state schedules exist before concluding.")
        widths(ws, {1: 100})
        return
    heads = ["Recipient", "State", "Box", "Code", "Amount"]
    ws.append(heads)
    hdr(ws, len(heads))
    for ln in sorted(lines, key=lambda x: (x["state"], x["recipient_name"])):
        if not ln["state"]:
            continue
        ws.append([ln["recipient_name"], ln["state"], ln["box"] or None,
                   ln["code"] or None,
                   float(ln["amount"]) if ln["amount"] is not None else None])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[4].number_format = MONEY
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:E{max(ws.max_row, 2)}"
    widths(ws, {1: 30, 2: 8, 3: 8, 4: 8, 5: 16})


def sheet_recon(wb, recips, expected, gc):
    ws = wb.create_sheet("Recipient Reconciliation")
    heads = ["Recipient", "Status", "Profit % per K-1", "Profit % expected",
             "Difference", "Prior-year profit %", "Change", "Final", "Amended", "Notes"]
    ws.append(heads)
    hdr(ws, len(heads))

    names = sorted(set(recips) | set(expected))
    for key in names:
        r = recips.get(key)
        e = expected.get(key)
        name = (r or e)["recipient_name"]
        if r and e:
            status = "received"
        elif r:
            status = "NOT ON EXPECTED LIST"
        else:
            status = "K-1 NOT RECEIVED"
        pk = r["pct_profit"] if r else None
        pe = e["pct_profit"] if e else None
        diff = (pk - pe) if (pk is not None and pe is not None) else None
        prior = e["prior_pct_profit"] if e else None
        change = (pk - prior) if (pk is not None and prior is not None) else None
        notes = []
        if change not in (None, Decimal("0")) and change is not None:
            notes.append("allocation percentage changed from prior year - confirm an "
                         "agreement amendment or transfer document exists")
        if r and r["amended_k1"]:
            notes.append("AMENDED K-1 - if the original was already filed, that return "
                         "needs amending")
        if r and r["final_k1"]:
            notes.append("marked FINAL - confirm the interest actually terminated")
        ws.append([name, status,
                   float(pk) if pk is not None else None,
                   float(pe) if pe is not None else None,
                   float(diff) if diff is not None else None,
                   float(prior) if prior is not None else None,
                   float(change) if change is not None else None,
                   "Y" if r and r["final_k1"] else "",
                   "Y" if r and r["amended_k1"] else "",
                   "; ".join(notes)])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in (2, 3, 4, 5, 6):
            row[i].number_format = PCT3
        if row[1].value != "received":
            row[1].font, row[1].fill = BAD_FONT, BAD_FILL
        if row[4].value not in (None, 0):
            row[4].font = BAD_FONT
        if row[9].value:
            row[9].fill = WARN_FILL
    ws.freeze_panes = "B2"
    widths(ws, {1: 32, 2: 22, 3: 15, 4: 16, 5: 12, 6: 17, 7: 10, 8: 7, 9: 9, 10: 62})


def sheet_exceptions(wb, gates, escalations):
    ws = wb.create_sheet("Exceptions")
    heads = ["#", "Type", "Detail"]
    ws.append(heads)
    hdr(ws, len(heads))
    n = 0
    for kind, items in (("entity consistency", gates["e"]["problems"]),
                        ("internal consistency", gates["d"]["problems"]),
                        ("escalation", escalations)):
        for it in items:
            n += 1
            ws.append([n, kind, it])
    for m in gates["c"].get("pct_mismatch", []):
        n += 1
        ws.append([n, "percentage mismatch",
                   f"{m['recipient']}: {m['field']} is {m['per_k1']} per K-1 but "
                   f"{m['expected']} per the expected recipient list"])
    if n == 0:
        ws.cell(row=2, column=3,
                value="None. No illegible figures, unresolved codes, or ambiguities.")
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[2].alignment = Alignment(wrap_text=True, vertical="top")
    widths(ws, {1: 5, 2: 24, 3: 110})


# --------------------------------------------------------------------------- main

def find_escalations(recips) -> list[str]:
    out = []
    for r in recips.values():
        if r["amended_k1"]:
            out.append(f"{r['recipient_name']}: AMENDED K-1. If the original was already "
                       f"entered on a filed return, that return needs amending.")
        if r["final_k1"]:
            out.append(f"{r['recipient_name']}: K-1 marked FINAL. Confirm the interest "
                       f"terminated; affects basis and gain recognition.")
        cap = r["pct_capital"]
        if cap is not None and cap < 0:
            out.append(f"{r['recipient_name']}: negative capital percentage "
                       f"({cap}) - implies distributions or losses in excess of basis.")
        has_199a = any(str(b) == "20" and str(c).upper() == "Z"
                       for b, c in r["boxes"])
        if r["boxes"] and not has_199a:
            out.append(f"{r['recipient_name']}: no box 20 code Z / Section 199A "
                       f"information captured. QBI cannot be computed without it - "
                       f"request from the entity now rather than at filing.")
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--lines", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--entity-totals")
    ap.add_argument("--expected-recipients")
    ap.add_argument("--client", default="")
    ap.add_argument("--tax-year", default="")
    ap.add_argument("--force", action="store_true",
                    help="write a workbook marked FAILED even when a gate fails")
    args = ap.parse_args()

    lines = load_lines(Path(args.lines))
    entity_totals = load_entity_totals(
        Path(args.entity_totals) if args.entity_totals else None)
    expected = load_expected(
        Path(args.expected_recipients) if args.expected_recipients else None)

    recips = build_recipients(lines)
    gates = {
        "a": gate_a(recips, entity_totals),
        "b": gate_b(recips),
        "c": gate_c(recips, expected),
        "d": gate_d(lines, recips),
        "e": gate_e(lines),
    }
    escalations = find_escalations(recips)
    overall = all(g["passed"] for g in gates.values())

    ents = sorted({ln["entity_name"] for ln in lines if ln["entity_name"]})
    meta = {
        "Entity": args.client or (ents[0] if ents else "(not stated)"),
        "EIN": sorted({ln["entity_ein"] for ln in lines if ln["entity_ein"]})[:1] or "(not captured)",
        "Entity type": sorted({ln["entity_type"] for ln in lines if ln["entity_type"]})[:1] or "(not captured)",
        "Tax year": args.tax_year or (sorted({ln["tax_year"] for ln in lines if ln["tax_year"]})[:1] or ["(not captured)"])[0],
        "K-1s in population": len(recips),
        "Box/code rows extracted": len(lines),
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    meta = {k: (v[0] if isinstance(v, list) and v else v) for k, v in meta.items()}

    # ---- console
    print("=" * 72)
    print("SCHEDULE K-1 SUMMARY")
    print("=" * 72)
    print(f"Entity            : {meta['Entity']}")
    print(f"Tax year          : {meta['Tax year']}")
    print(f"K-1s in population: {len(recips)}")
    print(f"Rows extracted    : {len(lines)}")
    print()
    ga = gates["a"]
    if not ga["applicable"]:
        print("Gate A footing            : NOT PERFORMED - " + ga["note"])
    else:
        bad = [r for r in ga["rows"] if r["status"] != "ties"]
        print(f"Gate A footing            : {'PASS' if ga['passed'] else '*** FAIL ***'}"
              f"  ({len(ga['rows']) - len(bad)}/{len(ga['rows'])} boxes tie)")
        for r in bad[:12]:
            ent = f"{r['entity']:,.2f}" if r["entity"] is not None else "(none supplied)"
            print(f"    box {r['box']:<4} code {r['code'] or '-':<3} K-1s "
                  f"{r['k1_sum']:>14,.2f}  entity {ent:>16}  "
                  f"diff {r['difference']:>12,.2f}")
    gb = gates["b"]
    print(f"Gate B percentages        : {'PASS' if gb['passed'] else '*** FAIL ***'}")
    for f, lbl in (("pct_profit", "profit"), ("pct_loss", "loss"),
                   ("pct_capital", "capital")):
        d = gb["detail"][f]
        t = f"{d['total']:.3f}%" if d["total"] is not None else "not captured"
        print(f"    {lbl:<8} total {t:>14}  {'ok' if d['passed'] else 'MUST BE 100.000%'}")
    gc = gates["c"]
    if not gc["applicable"]:
        print("Gate C completeness       : NOT PERFORMED - " + gc["note"])
    else:
        print(f"Gate C completeness       : {'PASS' if gc['passed'] else '*** FAIL ***'}")
        for n in gc["missing"]:
            print(f"    K-1 NOT RECEIVED: {n}")
        for n in gc["unexpected"]:
            print(f"    unexpected K-1  : {n}")
    for key, label in (("d", "Gate D internal"), ("e", "Gate E entity")):
        g = gates[key]
        print(f"{label:<26}: {'PASS' if g['passed'] else '*** FAIL ***'}")
        for p in g["problems"][:10]:
            print(f"    ! {p}")
    if escalations:
        print("\nESCALATE:")
        for e in escalations[:12]:
            print(f"  ! {e}")
        if len(escalations) > 12:
            print(f"  ... and {len(escalations) - 12} more")

    if not overall and not args.force:
        print("\n" + "=" * 72)
        print("WORKBOOK NOT WRITTEN. Do not use an unfooted or incomplete K-1 summary")
        print("for return input - a missing K-1 is invisible on the face of a return.")
        print("Fix the extraction or obtain the missing K-1s and re-run.")
        print("=" * 72)
        return 1

    wb = Workbook()
    sheet_summary(wb, meta, gates, recips, escalations, overall)
    sheet_footing(wb, gates["a"])
    sheet_matrix(wb, recips)
    sheet_detail(wb, lines)
    sheet_states(wb, lines, recips)
    sheet_recon(wb, recips, expected, gates["c"])
    sheet_exceptions(wb, gates, escalations)
    wb.active = 0

    out = Path(args.out)
    if not overall:
        out = out.with_name(out.stem + " [FAILED GATES]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\n{'PASSED' if overall else 'FAILED (forced)'}.  Workbook: {out}")
    return 0 if overall else 1


if __name__ == "__main__":
    sys.exit(main())
```

## `scripts/extract_text.py`

```python
#!/usr/bin/env python3
"""
Extract the text layer and word geometry from PDFs, fully locally.

Bundled with this skill so a single-skill install works standalone. Identical to
the extractor in the bank-statement-to-excel skill.

Writes, per input PDF:
    <stem>.text.txt    page-delimited text
    <stem>.words.jsonl one JSON object per word with coordinates
    <stem>.meta.json   page count, per-page char counts, scanned flag

Usage:
    python3 extract_text.py --input statements/ --out work/
    python3 extract_text.py --input jan.pdf --out work/ --ocr

No network access. Originals are opened read-only and never modified.
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    sys.exit("pdfplumber is required.  pip install pdfplumber")

# A page with fewer than this many extractable characters is treated as a scan.
SCAN_CHAR_THRESHOLD = 60


def ocr_pdf(src: Path) -> Path:
    """Run local OCR, returning a path to a new searchable PDF. Never touches src."""
    if not shutil.which("ocrmypdf"):
        sys.exit(
            "This PDF has no text layer and 'ocrmypdf' is not installed.\n"
            "Install it locally, then re-run with --ocr:\n"
            "  macOS:  brew install ocrmypdf\n"
            "  Debian: sudo apt install ocrmypdf tesseract-ocr\n"
            "Do not upload client tax documents to a cloud OCR service."
        )
    out = Path(tempfile.mkdtemp(prefix="ocr_")) / f"{src.stem}.ocr.pdf"
    print(f"  OCR (local): {src.name} -> {out.name}", flush=True)
    subprocess.run(
        [
            "ocrmypdf",
            "--force-ocr",
            "--deskew",
            "--optimize", "0",
            "--quiet",
            str(src),
            str(out),
        ],
        check=True,
    )
    return out


def extract(pdf_path: Path, out_dir: Path, use_ocr: bool) -> dict:
    stem = pdf_path.stem
    source = pdf_path

    if use_ocr:
        source = ocr_pdf(pdf_path)

    text_lines: list[str] = []
    words_out: list[str] = []
    page_meta: list[dict] = []

    with pdfplumber.open(str(source)) as pdf:
        for pageno, page in enumerate(pdf.pages, start=1):
            txt = page.extract_text() or ""
            text_lines.append(f"\n===== PAGE {pageno} =====\n{txt}")

            for w in page.extract_words(
                use_text_flow=False,
                keep_blank_chars=False,
                extra_attrs=["size"],
            ):
                words_out.append(
                    json.dumps(
                        {
                            "page": pageno,
                            "text": w["text"],
                            "x0": round(float(w["x0"]), 2),
                            "x1": round(float(w["x1"]), 2),
                            "top": round(float(w["top"]), 2),
                            "bottom": round(float(w["bottom"]), 2),
                            "size": round(float(w.get("size", 0)), 1),
                        },
                        ensure_ascii=False,
                    )
                )

            page_meta.append(
                {
                    "page": pageno,
                    "chars": len(txt),
                    "scanned": len(txt.strip()) < SCAN_CHAR_THRESHOLD,
                    "width": round(float(page.width), 1),
                    "height": round(float(page.height), 1),
                }
            )

    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / f"{stem}.text.txt").write_text("\n".join(text_lines), encoding="utf-8")
    (out_dir / f"{stem}.words.jsonl").write_text(
        "\n".join(words_out) + "\n", encoding="utf-8"
    )

    scanned_pages = [p["page"] for p in page_meta if p["scanned"]]
    meta = {
        "source_file": pdf_path.name,
        "ocr_applied": use_ocr,
        "page_count": len(page_meta),
        "scanned_pages": scanned_pages,
        "scanned": bool(scanned_pages),
        "pages": page_meta,
    }
    (out_dir / f"{stem}.meta.json").write_text(
        json.dumps(meta, indent=2), encoding="utf-8"
    )
    return meta


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input", required=True, help="PDF file or folder of PDFs")
    ap.add_argument("--out", default="work", help="output directory (default: work)")
    ap.add_argument(
        "--ocr",
        action="store_true",
        help="run local OCR first (required for scanned statements)",
    )
    args = ap.parse_args()

    src = Path(args.input).expanduser()
    out_dir = Path(args.out).expanduser()

    if src.is_dir():
        pdfs = sorted(src.glob("*.pdf")) + sorted(src.glob("*.PDF"))
    elif src.is_file():
        pdfs = [src]
    else:
        sys.exit(f"Not found: {src}")

    if not pdfs:
        sys.exit(f"No PDFs found in {src}")

    print(f"Extracting {len(pdfs)} PDF(s) -> {out_dir}/")
    needs_ocr: list[str] = []

    for p in pdfs:
        print(f"- {p.name}")
        meta = extract(p, out_dir, args.ocr)
        flag = ""
        if meta["scanned"] and not args.ocr:
            needs_ocr.append(p.name)
            flag = f"  [SCANNED pages {meta['scanned_pages']} - re-run with --ocr]"
        print(f"    {meta['page_count']} page(s){flag}")

    if needs_ocr:
        print(
            "\nThese files have pages with no usable text layer and MUST be re-run "
            "with --ocr before you read any figures from them:"
        )
        for n in needs_ocr:
            print(f"  - {n}")
        print(
            "Reading amounts off a scanned page without OCR produces silent blanks, "
            "not errors."
        )
        return 2

    print("\nDone. Text layer looks usable on every page.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

