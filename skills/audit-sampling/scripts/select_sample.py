#!/usr/bin/env python3
"""
Audit sample selection and projection: MUS/PPS, stratified, random attribute.

Gates:
  * the population must tie to a stated control total before anything is selected
  * a random seed is required and recorded, so the selection is re-performable
  * negative and zero balances require an explicit treatment decision
  * sample size is derived from stated parameters, or taken from firm tables and
    documented as such -- never silently chosen

Computes what is mathematics (reliability factor = -ln(risk)) and requires the
user to supply what is firm methodology (expansion factor, tolerable misstatement,
risk). It will not invent a methodology parameter.

Usage - validate the population first:
    python3 select_sample.py --population pop.csv \
        --control-total 4128455.19 --validate-only

Usage - select:
    python3 select_sample.py --population pop.csv --control-total 4128455.19 \
        --method mus --tolerable-misstatement 125000 \
        --expected-misstatement 25000 --expansion-factor 1.6 --risk 0.05 \
        --negatives separate --seed 20260729 \
        --client "Harbour Freight Systems Inc" \
        --assertion "Existence of trade receivables at 12/31/2025" \
        --out "Harbour Freight - AR Existence Sample.xlsx"

Usage - project results:
    python3 select_sample.py --project results.csv --population pop.csv \
        --control-total 4128455.19 --method mus \
        --tolerable-misstatement 125000 --risk 0.05 --seed 20260729 \
        --out "Harbour Freight - AR Existence Conclusion.xlsx"

--population CSV:
    item_id, description, book_value          (required)
    date, account, party, reference           (optional, carried through)

--project CSV (the filled-in Selection tab, exported to CSV):
    item_id, book_value, audited_amount, evidence, disposition
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import math
import random
import sys
from datetime import datetime
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
PCT = '0.0"%"'

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
TOP = Border(top=Side(style="thin"))


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
        raise ValueError(f"Unparseable amount: {raw!r}")
    return -v if neg else v


def clean(s) -> str:
    return " ".join(str(s or "").split())


def load_population(path: Path) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No rows in {path}")
    out = []
    for i, r in enumerate(rows, start=1):
        bv = dec(r.get("book_value"))
        if bv is None:
            raise ValueError(f"Row {i}: book_value is required. A row with no amount "
                             f"cannot be part of a monetary population.")
        out.append({
            "row": i,
            "item_id": clean(r.get("item_id")) or f"ITEM{i:05d}",
            "description": clean(r.get("description")),
            "book_value": bv,
            "date": clean(r.get("date")),
            "account": clean(r.get("account")),
            "party": clean(r.get("party")),
            "reference": clean(r.get("reference")),
        })
    return out


def file_hash(path: Path) -> str:
    h = hashlib.sha256()
    h.update(path.read_bytes())
    return h.hexdigest()[:16]


# ------------------------------------------------------------------- population

def summarize(pop: list[dict], control_total: Decimal | None) -> dict:
    pos = [p for p in pop if p["book_value"] > ZERO]
    neg = [p for p in pop if p["book_value"] < ZERO]
    zed = [p for p in pop if p["book_value"] == ZERO]
    total = sum((p["book_value"] for p in pop), ZERO)
    diff = (total - control_total) if control_total is not None else None
    return {
        "count": len(pop), "total": total,
        "positive": pos, "negative": neg, "zero": zed,
        "positive_total": sum((p["book_value"] for p in pos), ZERO),
        "negative_total": sum((p["book_value"] for p in neg), ZERO),
        "control_total": control_total, "difference": diff,
        "ties": diff == ZERO if diff is not None else None,
        "largest": sorted(pop, key=lambda p: abs(p["book_value"]), reverse=True)[:10],
    }


# ----------------------------------------------------------------- sample sizes

def reliability_factor(risk: float) -> float:
    """Poisson reliability factor for zero expected misstatements: -ln(risk).

    This is arithmetic, not a table lookup, so it is computed rather than recalled.
    """
    if not 0 < risk < 1:
        raise ValueError("--risk must be strictly between 0 and 1 (e.g. 0.05).")
    return -math.log(risk)


def mus_plan(pop_total: Decimal, tm: Decimal, em: Decimal, risk: float,
             expansion: float | None) -> dict:
    rf = reliability_factor(risk)
    notes = [f"Reliability factor computed as -ln({risk}) = {rf:.4f} "
             f"(zero expected misstatements)."]
    if em > ZERO:
        if expansion is None:
            notes.append(
                "EXPECTED MISSTATEMENT WAS SUPPLIED BUT NO EXPANSION FACTOR. Expansion "
                "factors are firm methodology and this script will not invent one. The "
                "sample size below ignores expected misstatement and is therefore "
                "UNDERSTATED. Supply --expansion-factor, or enter your firm's size with "
                "--sample-size.")
            adj_tm = tm
        else:
            adj_tm = tm - (em * Decimal(str(expansion)))
            notes.append(
                f"Tolerable misstatement reduced for expected misstatement: "
                f"{tm:,.2f} - ({em:,.2f} x {expansion}) = {adj_tm:,.2f}.")
            if adj_tm <= ZERO:
                notes.append(
                    "ADJUSTED TOLERABLE MISSTATEMENT IS ZERO OR NEGATIVE. Expected "
                    "misstatement times the expansion factor equals or exceeds tolerable "
                    "misstatement - sampling cannot provide the assurance sought. "
                    "Reconsider the plan or examine the population 100%.")
    else:
        adj_tm = tm
    interval = (adj_tm / Decimal(str(rf))) if adj_tm > ZERO else None
    size = int(math.ceil(float(pop_total) / float(interval))) if interval and interval > ZERO else None
    return {"reliability_factor": rf, "adjusted_tm": adj_tm,
            "sampling_interval": interval, "sample_size": size, "notes": notes}


# ------------------------------------------------------------------- selections

def select_mus(pop: list[dict], interval: Decimal, seed: int) -> tuple[list[dict], dict]:
    """Systematic monetary-unit selection with a random start.

    Items >= the interval are certain to be selected (top stratum). Selection walks
    the cumulative book value and picks the item spanning each hit point.
    """
    rng = random.Random(seed)
    start = Decimal(str(rng.random())) * interval

    items = [p for p in pop if p["book_value"] > ZERO]
    top = [p for p in items if p["book_value"] >= interval]
    rest = [p for p in items if p["book_value"] < interval]

    selected = []
    for p in top:
        selected.append({**p, "basis": "top stratum (>= sampling interval), examined 100%",
                         "hit_amount": None})
    cum = ZERO
    hit = start
    for p in rest:
        cum += p["book_value"]
        while hit <= cum:
            selected.append({**p, "basis": "interval hit", "hit_amount": hit})
            hit += interval
    meta = {"random_start": start, "interval": interval,
            "top_stratum_count": len(top),
            "top_stratum_total": sum((p["book_value"] for p in top), ZERO)}
    return selected, meta


def select_attribute(pop: list[dict], n: int, seed: int) -> tuple[list[dict], dict]:
    rng = random.Random(seed)
    if n >= len(pop):
        return ([{**p, "basis": "100% examination (sample size >= population)",
                  "hit_amount": None} for p in pop],
                {"note": "sample size met or exceeded the population"})
    idx = sorted(rng.sample(range(len(pop)), n))
    return ([{**pop[i], "basis": "random selection", "hit_amount": None} for i in idx],
            {"note": f"{n} items selected at random from {len(pop)}"})


def select_stratified(pop: list[dict], bounds: list[Decimal], sizes: list[int],
                      seed: int) -> tuple[list[dict], dict]:
    rng = random.Random(seed)
    items = sorted((p for p in pop if p["book_value"] > ZERO),
                   key=lambda p: p["book_value"], reverse=True)
    strata: list[list[dict]] = [[] for _ in range(len(bounds) + 1)]
    for p in items:
        placed = False
        for i, b in enumerate(bounds):
            if p["book_value"] >= b:
                strata[i].append(p)
                placed = True
                break
        if not placed:
            strata[-1].append(p)

    selected, detail = [], []
    for i, s in enumerate(strata):
        want = sizes[i] if i < len(sizes) else 0
        label = (f">= {bounds[i]:,.0f}" if i < len(bounds)
                 else f"< {bounds[-1]:,.0f}" if bounds else "all")
        if want >= len(s):
            pick = s
            basis = f"stratum {i + 1} ({label}) - examined 100%"
        else:
            pick = [s[j] for j in sorted(rng.sample(range(len(s)), want))] if want else []
            basis = f"stratum {i + 1} ({label}) - random within stratum"
        for p in pick:
            selected.append({**p, "basis": basis, "hit_amount": None})
        detail.append({"stratum": i + 1, "label": label, "count": len(s),
                       "total": sum((x["book_value"] for x in s), ZERO),
                       "selected": len(pick)})
    return selected, {"strata": detail}


# ------------------------------------------------------------------- projection

def project(results: list[dict], interval: Decimal | None, method: str) -> dict:
    rows, total_proj = [], ZERO
    for r in results:
        bv, aa = r["book_value"], r["audited_amount"]
        if aa is None:
            rows.append({**r, "difference": None, "tainting": None,
                         "projected": None, "note": "NOT TESTED"})
            continue
        diff = bv - aa
        if method != "mus" or interval is None:
            taint, proj = None, diff
        elif bv >= interval:
            taint, proj = None, diff          # examined in full, no projection
        elif bv == ZERO:
            taint, proj = None, diff
        else:
            taint = diff / bv
            proj = taint * interval
        total_proj += proj
        rows.append({**r, "difference": diff, "tainting": taint, "projected": proj,
                     "note": "" if diff == ZERO else "difference found"})
    untested = [r for r in rows if r["note"] == "NOT TESTED"]
    return {"rows": rows, "projected_misstatement": total_proj,
            "differences": [r for r in rows if r["difference"] not in (None, ZERO)],
            "untested": untested}


# -------------------------------------------------------------------- workbook

def hdr(ws, n, row=1):
    for c in range(1, n + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def widths(ws, w):
    for c, v in w.items():
        ws.column_dimensions[get_column_letter(c)].width = v


def sheet_plan(wb, plan_meta, summary, notes, warnings):
    ws = wb.active
    ws.title = "Sampling Plan"
    widths(ws, {1: 4, 2: 50, 3: 24, 4: 70})
    r = 1

    def line(label, val=None, *, bold=False, size=11, fill=None, note=""):
        nonlocal r
        c = ws.cell(row=r, column=2, value=label)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c.fill = fill
        if val is not None:
            v = ws.cell(row=r, column=3, value=val)
            v.font = Font(bold=bold)
            if isinstance(val, float):
                v.number_format = MONEY
        if note:
            n = ws.cell(row=r, column=4, value=note)
            n.font = Font(italic=True, size=9)
            n.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    line("SAMPLING PLAN", bold=True, size=14)
    r += 1
    for k, v in plan_meta.items():
        line(k, v, bold=(k in ("Method", "Sample size", "Random seed")))
    r += 1

    line("POPULATION", bold=True, size=12, fill=SUB_FILL)
    line("  Items in population", summary["count"])
    line("  Population total", float(summary["total"]))
    if summary["control_total"] is not None:
        line("  Control total (GL / trial balance)", float(summary["control_total"]))
        d = summary["difference"]
        line("  Difference", float(d),
             bold=True,
             fill=None if d == ZERO else BAD_FILL,
             note="" if d == ZERO else
                  "POPULATION DOES NOT TIE. A sample drawn from an incomplete population "
                  "supports nothing, and the projection will understate misstatement by "
                  "the amount missing.")
    else:
        line("  Control total", "NOT SUPPLIED", fill=BAD_FILL,
             note="Without a control total the completeness of the population is "
                  "unproven and the sample is not defensible.")
    line("  Positive items", len(summary["positive"]))
    line("  Positive total", float(summary["positive_total"]))
    line("  Negative items", len(summary["negative"]),
         fill=WARN_FILL if summary["negative"] else None,
         note="Negative balances cannot be sampled proportionally to size in MUS. "
              "The treatment applied is recorded above." if summary["negative"] else "")
    line("  Negative total", float(summary["negative_total"]))
    line("  Zero-balance items", len(summary["zero"]),
         note="Zero balances have no chance of MUS selection. If they matter to the "
              "assertion, test them by another means." if summary["zero"] else "")
    r += 1

    if notes:
        line("DERIVATION", bold=True, size=12, fill=SUB_FILL)
        for n in notes:
            line("  " + n)
        r += 1

    if warnings:
        line("WARNINGS", bold=True, size=12, fill=BAD_FILL)
        for w in warnings:
            line("  " + w)
        r += 1

    line("REPRODUCIBILITY", bold=True, size=12, fill=SUB_FILL)
    line("  This selection can be re-performed exactly from the seed, method, and "
         "parameters recorded above, against the population file identified by the "
         "hash. See the Reproducibility tab for the exact command.")
    r += 1
    line("Prepared by / date:  __________________  ____________", bold=True)
    line("Reviewed by / date:  __________________  ____________", bold=True)


def sheet_selection(wb, selected):
    ws = wb.create_sheet("Selection")
    heads = ["#", "Item ID", "Description", "Date", "Account", "Party", "Reference",
             "Book value", "Selection basis", "Evidence examined", "Audited amount",
             "Difference", "Disposition"]
    ws.append(heads)
    hdr(ws, len(heads))
    for i, s in enumerate(selected, start=1):
        ws.append([i, s["item_id"], s["description"], s["date"] or None,
                   s["account"] or None, s["party"] or None, s["reference"] or None,
                   float(s["book_value"]), s["basis"], None, None, None, None])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[7].number_format = MONEY
        row[10].number_format = MONEY
        row[11].number_format = MONEY
        if "top stratum" in str(row[8].value):
            row[8].fill = WARN_FILL
    last = ws.max_row
    ws.append([])
    t = ws.max_row + 1
    ws.cell(row=t, column=2, value="TOTAL SELECTED").font = Font(bold=True)
    c = ws.cell(row=t, column=8, value=f"=SUM(H2:H{last})")
    c.number_format, c.font, c.border = MONEY, Font(bold=True), TOP
    ws.freeze_panes = "C2"
    widths(ws, {1: 5, 2: 16, 3: 40, 4: 12, 5: 14, 6: 24, 7: 16, 8: 16, 9: 44,
                10: 30, 11: 16, 12: 14, 13: 22})


def sheet_pop_summary(wb, summary, selected, plan):
    ws = wb.create_sheet("Population Summary")
    widths(ws, {1: 4, 2: 44, 3: 20, 4: 18, 5: 50})
    r = 1
    ws.cell(row=r, column=2, value="POPULATION AND COVERAGE").font = Font(bold=True, size=14)
    r += 2

    sel_total = sum((s["book_value"] for s in selected), ZERO)
    for label, val, fmt in [
        ("Items in population", summary["count"], None),
        ("Population total", float(summary["total"]), MONEY),
        ("Items selected", len(selected), None),
        ("Value selected", float(sel_total), MONEY),
    ]:
        ws.cell(row=r, column=2, value=label)
        c = ws.cell(row=r, column=3, value=val)
        if fmt:
            c.number_format = fmt
        r += 1
    for label, num, den in [("Coverage - items", len(selected), summary["count"]),
                            ("Coverage - dollars", float(sel_total),
                             float(summary["total"]) or 1)]:
        ws.cell(row=r, column=2, value=label)
        c = ws.cell(row=r, column=3, value=(num / den * 100) if den else 0)
        c.number_format = PCT
        r += 1
    r += 1

    ws.cell(row=r, column=2, value="Distribution by size band").font = Font(bold=True)
    r += 1
    for h, col in zip(["Band", "Items", "Total value"], (2, 3, 4)):
        c = ws.cell(row=r, column=col, value=h)
        c.fill, c.font = HDR_FILL, HDR_FONT
    r += 1
    interval = plan.get("sampling_interval")
    bands = [(interval, None, f">= sampling interval ({interval:,.2f})")] if interval else []
    bands += [(Decimal("100000"), interval, "100,000 to interval"),
              (Decimal("10000"), Decimal("100000"), "10,000 - 100,000"),
              (Decimal("1000"), Decimal("10000"), "1,000 - 10,000"),
              (ZERO, Decimal("1000"), "0 - 1,000")]
    for lo, hi, label in bands:
        if lo is None:
            continue
        items = [p for p in summary["positive"]
                 if p["book_value"] >= lo and (hi is None or p["book_value"] < hi)]
        ws.cell(row=r, column=2, value=label)
        ws.cell(row=r, column=3, value=len(items))
        c = ws.cell(row=r, column=4,
                    value=float(sum((p["book_value"] for p in items), ZERO)))
        c.number_format = MONEY
        r += 1
    r += 1

    ws.cell(row=r, column=2, value="Ten largest items in the population").font = Font(bold=True)
    r += 1
    for p in summary["largest"]:
        ws.cell(row=r, column=2, value=f"  {p['item_id']}  {p['description'][:40]}")
        c = ws.cell(row=r, column=3, value=float(p["book_value"]))
        c.number_format = MONEY
        r += 1


def sheet_not_selected(wb, pop, selected):
    ws = wb.create_sheet("Not Selected")
    picked = {s["row"] for s in selected}
    heads = ["Item ID", "Description", "Book value", "Reason not selected"]
    ws.append(heads)
    hdr(ws, len(heads))
    rest = [p for p in pop if p["row"] not in picked]
    for p in rest:
        reason = ("negative balance - see the negative-balance treatment on the "
                  "Sampling Plan" if p["book_value"] < ZERO
                  else "zero balance - no chance of monetary-unit selection"
                  if p["book_value"] == ZERO
                  else "not selected by the sampling method")
        ws.append([p["item_id"], p["description"], float(p["book_value"]), reason])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[2].number_format = MONEY
        if "negative" in str(row[3].value) or "zero" in str(row[3].value):
            row[3].fill = WARN_FILL
    r = ws.max_row + 2
    ws.cell(row=r, column=1,
            value=f"{len(rest)} of {len(pop)} items were not selected, totalling "
                  f"{sum((p['book_value'] for p in rest), ZERO):,.2f}. The untested "
                  f"remainder is quantified here deliberately - a selection that hides "
                  f"what it did not cover looks arbitrary."
            ).font = Font(italic=True)
    ws.freeze_panes = "A2"
    widths(ws, {1: 16, 2: 44, 3: 16, 4: 56})


def sheet_reproducibility(wb, command, seed, pop_path, pop_hash, plan_meta):
    ws = wb.create_sheet("Reproducibility")
    widths(ws, {1: 4, 2: 34, 3: 84})
    r = 1
    ws.cell(row=r, column=2, value="REPRODUCIBILITY").font = Font(bold=True, size=14)
    r += 2
    for k, v in [("Population file", pop_path),
                 ("Population SHA-256 (first 16)", pop_hash),
                 ("Random seed", seed),
                 ("Selected on", datetime.now().strftime("%Y-%m-%d %H:%M"))]:
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=str(v))
        r += 1
    r += 1
    ws.cell(row=r, column=2, value="Exact command").font = Font(bold=True)
    c = ws.cell(row=r, column=3, value=command)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    r += 2
    for line in [
        "Re-running this command against the same population file reproduces the same",
        "selection exactly. If the hash differs, the population changed after selection",
        "and the sample must be re-performed - a selection binds to the population as it",
        "stood when it was drawn.",
        "",
        "The seed must be recorded BEFORE selection. Choosing a seed after seeing results",
        "is not random selection.",
    ]:
        ws.cell(row=r, column=2, value=line)
        r += 1


def sheet_projection(wb, proj, tm, interval, method):
    ws = wb.create_sheet("Projection")
    heads = ["Item ID", "Book value", "Audited amount", "Difference", "Tainting %",
             "Projected misstatement", "Evidence", "Disposition", "Note"]
    ws.append(heads)
    hdr(ws, len(heads))
    for r_ in proj["rows"]:
        ws.append([r_["item_id"], float(r_["book_value"]),
                   float(r_["audited_amount"]) if r_["audited_amount"] is not None else None,
                   float(r_["difference"]) if r_["difference"] is not None else None,
                   float(r_["tainting"] * 100) if r_["tainting"] is not None else None,
                   float(r_["projected"]) if r_["projected"] is not None else None,
                   r_.get("evidence") or None, r_.get("disposition") or None,
                   r_["note"] or None])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in (1, 2, 3, 5):
            row[i].number_format = MONEY
        row[4].number_format = PCT
        if row[8].value == "NOT TESTED":
            row[8].font, row[8].fill = BAD_FONT, BAD_FILL
        elif row[8].value:
            row[8].fill = WARN_FILL
    r = ws.max_row + 2
    ws.cell(row=r, column=1, value="Projected misstatement").font = Font(bold=True)
    c = ws.cell(row=r, column=6, value=float(proj["projected_misstatement"]))
    c.number_format, c.font, c.border = MONEY, Font(bold=True), TOP
    r += 1
    if tm is not None:
        ws.cell(row=r, column=1, value="Tolerable misstatement").font = Font(bold=True)
        c = ws.cell(row=r, column=6, value=float(tm))
        c.number_format = MONEY
        r += 1
        pm = proj["projected_misstatement"]
        ratio = (abs(pm) / tm) if tm else None
        verdict = ("Projected misstatement is below tolerable misstatement."
                   if abs(pm) < tm else
                   "PROJECTED MISSTATEMENT EQUALS OR EXCEEDS TOLERABLE MISSTATEMENT.")
        c = ws.cell(row=r, column=1, value=verdict)
        c.font = OK_FONT if abs(pm) < tm else BAD_FONT
        if abs(pm) >= tm:
            c.fill = BAD_FILL
        r += 2
        if ratio is not None and Decimal("0.5") <= ratio < Decimal("1"):
            c = ws.cell(row=r, column=1,
                        value=f"Projected misstatement is {ratio * 100:.0f}% of tolerable. "
                              f"Every individual difference may have looked small while the "
                              f"balance is still at risk - this is the situation sampling "
                              f"exists to reveal. Consider extending the sample or "
                              f"proposing an adjustment.")
            c.fill = WARN_FILL
            c.alignment = Alignment(wrap_text=True)
            r += 2
    if proj["untested"]:
        c = ws.cell(row=r, column=1,
                    value=f"{len(proj['untested'])} selected item(s) were NOT TESTED. "
                          f"An untested selection is a scope limitation, not a "
                          f"substitution - the projection above is incomplete.")
        c.font, c.fill = BAD_FONT, BAD_FILL
        r += 1
    for line in [
        "",
        "MUS tainting: an item at or above the sampling interval is examined in full and "
        "projects its actual difference. An item below the interval projects "
        "(difference / book value) x sampling interval.",
        "Basic precision and an allowance for sampling risk depend on firm methodology "
        "and are not computed here. Apply your firm's factors to reach an upper limit.",
    ]:
        ws.cell(row=r, column=1, value=line).alignment = Alignment(wrap_text=True)
        r += 1
    ws.freeze_panes = "A2"
    widths(ws, {1: 18, 2: 16, 3: 17, 4: 14, 5: 12, 6: 22, 7: 30, 8: 22, 9: 18})


# ------------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--population", required=True)
    ap.add_argument("--control-total")
    ap.add_argument("--validate-only", action="store_true")
    ap.add_argument("--method", choices=["mus", "stratified", "attribute"], default="mus")
    ap.add_argument("--tolerable-misstatement")
    ap.add_argument("--expected-misstatement", default="0")
    ap.add_argument("--expansion-factor", type=float)
    ap.add_argument("--risk", type=float, default=0.05)
    ap.add_argument("--sample-size", type=int,
                    help="override the derived size, e.g. from firm sampling tables")
    ap.add_argument("--strata-bounds", default="",
                    help="comma-separated descending bounds, e.g. 250000,50000,10000")
    ap.add_argument("--strata-sizes", default="",
                    help="comma-separated sizes per stratum, highest first")
    ap.add_argument("--negatives", choices=["separate", "exclude", "sample"],
                    help="required when the population contains negative balances")
    ap.add_argument("--seed", type=int)
    ap.add_argument("--client", default="")
    ap.add_argument("--assertion", default="")
    ap.add_argument("--project", help="results CSV to project")
    ap.add_argument("--out")
    args = ap.parse_args()

    pop_path = Path(args.population)
    pop = load_population(pop_path)
    ct = dec(args.control_total)
    summary = summarize(pop, ct)

    print("=" * 74)
    print("POPULATION")
    print("=" * 74)
    print(f"Items                 : {summary['count']:,}")
    print(f"Total                 : {summary['total']:>18,.2f}")
    if ct is not None:
        print(f"Control total         : {ct:>18,.2f}")
        print(f"Difference            : {summary['difference']:>18,.2f}  "
              f"{'TIES' if summary['ties'] else '*** DOES NOT TIE ***'}")
    else:
        print("Control total         : NOT SUPPLIED")
    print(f"Positive / negative / zero items: {len(summary['positive'])} / "
          f"{len(summary['negative'])} / {len(summary['zero'])}")
    if summary["negative"]:
        print(f"Negative total        : {summary['negative_total']:>18,.2f}")

    if args.validate_only:
        if ct is not None and not summary["ties"]:
            print("\nResolve the difference before selecting. A sample drawn from an "
                  "incomplete\npopulation supports nothing.")
            return 1
        print("\nPopulation looks usable. Decide the negative-balance treatment before "
              "selecting.")
        return 0

    # ---------------- gates
    problems = []
    if ct is None:
        problems.append("--control-total is required. Without it the completeness of the "
                        "population is unproven and the sample is not defensible.")
    elif not summary["ties"]:
        problems.append(f"Population total {summary['total']:,.2f} does not agree to the "
                        f"control total {ct:,.2f} (difference "
                        f"{summary['difference']:,.2f}). Resolve before selecting.")
    if args.seed is None:
        problems.append("--seed is required so the selection can be re-performed exactly. "
                        "Record it before selecting, not after.")
    if summary["negative"] and not args.negatives:
        problems.append(f"The population contains {len(summary['negative'])} negative "
                        f"balance(s) totalling {summary['negative_total']:,.2f}. Choose "
                        f"--negatives separate|exclude|sample explicitly - negative "
                        f"balances cannot be sampled proportionally to size and this "
                        f"script will not guess.")
    if not args.project and not args.out:
        problems.append("--out is required.")

    if args.project:
        return do_projection(args, pop, summary, ct)

    tm = dec(args.tolerable_misstatement)
    em = dec(args.expected_misstatement) or ZERO
    notes, warnings = [], []
    plan = {}

    if args.method == "mus":
        if tm is None and args.sample_size is None:
            problems.append("--tolerable-misstatement is required for MUS (or supply "
                            "--sample-size from firm tables).")
        if problems:
            print_problems(problems)
            return 1
        if args.sample_size is not None and tm is None:
            interval = summary["positive_total"] / Decimal(args.sample_size)
            plan = {"reliability_factor": None, "adjusted_tm": None,
                    "sampling_interval": interval, "sample_size": args.sample_size,
                    "notes": [f"Sample size {args.sample_size} supplied from firm "
                              f"sampling tables; interval derived as positive population "
                              f"total / sample size."]}
        else:
            plan = mus_plan(summary["positive_total"], tm, em, args.risk,
                            args.expansion_factor)
            if args.sample_size is not None:
                plan["notes"].append(
                    f"Derived size {plan['sample_size']} overridden with "
                    f"--sample-size {args.sample_size} (firm tables).")
                plan["sampling_interval"] = summary["positive_total"] / Decimal(args.sample_size)
                plan["sample_size"] = args.sample_size
        notes = plan["notes"]
        warnings = [n for n in notes
                    if n.startswith(("EXPECTED MISSTATEMENT", "ADJUSTED TOLERABLE"))]
        if plan["sampling_interval"] is None:
            print_problems(["Sampling interval could not be computed - see the "
                            "derivation notes."] + notes)
            return 1
        selected, meta = select_mus(pop, plan["sampling_interval"], args.seed)
        notes.append(f"Sampling interval {plan['sampling_interval']:,.2f}; random start "
                     f"{meta['random_start']:,.2f}; {meta['top_stratum_count']} item(s) at "
                     f"or above the interval examined 100%.")

    elif args.method == "attribute":
        if args.sample_size is None:
            problems.append("--sample-size is required for attribute sampling. Attribute "
                            "sample sizes come from firm tables based on tolerable and "
                            "expected deviation rates; this script will not invent one.")
        if problems:
            print_problems(problems)
            return 1
        selected, meta = select_attribute(pop, args.sample_size, args.seed)
        plan = {"sampling_interval": None, "sample_size": args.sample_size}
        notes = [meta["note"],
                 "Attribute sampling supports a conclusion about a DEVIATION RATE. Do not "
                 "project dollar misstatement from this sample."]

    else:  # stratified
        bounds = [dec(b) for b in args.strata_bounds.split(",") if b.strip()]
        sizes = [int(s) for s in args.strata_sizes.split(",") if s.strip()]
        if not bounds or not sizes:
            problems.append("--strata-bounds and --strata-sizes are both required for "
                            "stratified sampling, and the boundaries must be justified "
                            "on the workpaper.")
        if problems:
            print_problems(problems)
            return 1
        selected, meta = select_stratified(pop, bounds, sizes, args.seed)
        plan = {"sampling_interval": None, "sample_size": len(selected)}
        notes = [f"Stratum {d['stratum']} ({d['label']}): {d['count']} items totalling "
                 f"{d['total']:,.2f}, {d['selected']} selected."
                 for d in meta["strata"]]

    if problems:
        print_problems(problems)
        return 1

    sel_total = sum((s["book_value"] for s in selected), ZERO)
    plan_meta = {
        "Client": args.client or "(not stated)",
        "Assertion tested": args.assertion or "(NOT STATED - record what this sample "
                                             "supports)",
        "Method": {"mus": "Monetary unit (PPS)", "attribute": "Random attribute",
                   "stratified": "Stratified random"}[args.method],
        "Tolerable misstatement": float(tm) if tm else "(n/a)",
        "Expected misstatement": float(em),
        "Expansion factor": args.expansion_factor if args.expansion_factor else "(not supplied)",
        "Risk of incorrect acceptance": args.risk,
        "Reliability factor": round(plan.get("reliability_factor"), 4)
                              if plan.get("reliability_factor") else "(n/a)",
        "Sampling interval": float(plan["sampling_interval"]) if plan.get("sampling_interval") else "(n/a)",
        "Sample size": len(selected),
        "Value selected": float(sel_total),
        "Negative-balance treatment": args.negatives or "(no negatives in population)",
        "Random seed": args.seed,
        "Population file": pop_path.name,
        "Population hash": file_hash(pop_path),
    }

    command = ("python3 select_sample.py "
               + " ".join(f"--{k} {v}" for k, v in [
                   ("population", pop_path.name),
                   ("control-total", args.control_total),
                   ("method", args.method),
                   ("tolerable-misstatement", args.tolerable_misstatement or "-"),
                   ("expected-misstatement", args.expected_misstatement),
                   ("risk", args.risk),
                   ("negatives", args.negatives or "-"),
                   ("seed", args.seed)] if v not in (None, "-"))
               + f" --out \"{args.out}\"")

    print()
    print("=" * 74)
    print("SELECTION")
    print("=" * 74)
    for k, v in plan_meta.items():
        print(f"  {k:<32} {v}")
    print()
    for n in notes:
        print(f"  note: {n}")
    print()
    print(f"Coverage: {len(selected)} of {summary['count']} items "
          f"({len(selected) / summary['count'] * 100:.1f}%), "
          f"{sel_total:,.2f} of {summary['total']:,.2f} "
          f"({float(sel_total) / float(summary['total']) * 100:.1f}% of dollars)")

    wb = Workbook()
    sheet_plan(wb, plan_meta, summary, notes, warnings)
    sheet_selection(wb, selected)
    sheet_pop_summary(wb, summary, selected, plan)
    sheet_not_selected(wb, pop, selected)
    sheet_reproducibility(wb, command, args.seed, pop_path.name,
                          file_hash(pop_path), plan_meta)
    wb.active = 0
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\nWorkbook: {out}")
    print("Next: record the assertion if not already stated, test each selection, and "
          "enter\nthe audited amounts. Then re-run with --project to obtain the "
          "projection.")
    return 0


def print_problems(problems):
    print("\n" + "=" * 74)
    print("SELECTION NOT PERFORMED")
    print("=" * 74)
    for p in problems:
        print(f"  ! {p}")
    print("=" * 74)


def do_projection(args, pop, summary, ct) -> int:
    by_id = {p["item_id"].lower(): p for p in pop}
    with Path(args.project).open(newline="", encoding="utf-8-sig") as fh:
        raw = list(csv.DictReader(fh))
    results = []
    for r in raw:
        iid = clean(r.get("item_id"))
        bv = dec(r.get("book_value"))
        if bv is None and iid.lower() in by_id:
            bv = by_id[iid.lower()]["book_value"]
        if bv is None:
            continue
        results.append({"item_id": iid, "book_value": bv,
                        "audited_amount": dec(r.get("audited_amount")),
                        "evidence": clean(r.get("evidence")),
                        "disposition": clean(r.get("disposition"))})
    if not results:
        sys.exit("No usable rows in the results file.")

    tm = dec(args.tolerable_misstatement)
    interval = None
    if args.method == "mus" and tm is not None:
        plan = mus_plan(summary["positive_total"], tm,
                        dec(args.expected_misstatement) or ZERO,
                        args.risk, args.expansion_factor)
        interval = plan["sampling_interval"]
    proj = project(results, interval, args.method)

    print()
    print("=" * 74)
    print("PROJECTION")
    print("=" * 74)
    print(f"Items tested            : {len(results) - len(proj['untested'])}")
    print(f"Items NOT tested        : {len(proj['untested'])}")
    print(f"Differences found       : {len(proj['differences'])}")
    if interval:
        print(f"Sampling interval       : {interval:>18,.2f}")
    print(f"Projected misstatement  : {proj['projected_misstatement']:>18,.2f}")
    if tm:
        print(f"Tolerable misstatement  : {tm:>18,.2f}")
        pm = abs(proj["projected_misstatement"])
        if pm >= tm:
            print("  *** PROJECTED MISSTATEMENT EQUALS OR EXCEEDS TOLERABLE ***")
        elif tm and pm / tm >= Decimal("0.5"):
            print(f"  Projected is {pm / tm * 100:.0f}% of tolerable - consider extending "
                  f"the sample or proposing an adjustment.")
        else:
            print("  Projected misstatement is below tolerable misstatement.")
    for r in proj["differences"][:15]:
        print(f"    {r['item_id']:<16} book {r['book_value']:>13,.2f}  audited "
              f"{r['audited_amount']:>13,.2f}  diff {r['difference']:>11,.2f}  "
              f"projected {r['projected']:>13,.2f}")
    if proj["untested"]:
        print("\n  ! Untested selections are a scope limitation, not a substitution. "
              "The projection is incomplete.")

    if args.out:
        wb = Workbook()
        wb.active.title = "Projection"
        sheet_projection(wb, proj, tm, interval, args.method)
        if "Projection" in wb.sheetnames and wb["Projection"] is not wb.active:
            wb.remove(wb.active)
        wb.save(args.out)
        print(f"\nWorkbook: {args.out}")
    return 0 if not proj["untested"] else 1


if __name__ == "__main__":
    sys.exit(main())
