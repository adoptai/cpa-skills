---
name: depreciation-tie-out
description: Tie the fixed asset register to the depreciation schedule and to the tax return or financial statements — proving the cost and accumulated depreciation rollforwards foot, that beginning balances equal prior-year ending, that disposals were removed with gain or loss computed, and that no asset is depreciated beyond its basis. Use this whenever the user mentions fixed assets, the fixed asset register, depreciation schedule, Form 4562, depreciation tie-out, accumulated depreciation, asset additions or disposals, a fixed asset rollforward, book versus tax depreciation, or checking depreciation against the return. Also trigger on "does our depreciation schedule tie," "reconcile fixed assets," "check the 4562," "did we remove the assets we sold," "depreciation looks wrong," "roll forward the FA register," or when a fixed asset register and a return or trial balance are provided together. Runs fully local — no client asset data leaves the machine.
---

# Depreciation and Fixed Asset Tie-Out

Fixed assets accumulate errors quietly because the schedule is rarely rebuilt — it is rolled
forward, year after year, often through a change of software or preparer. Each roll carries
forward whatever was already wrong.

Four failures account for most of what this tie-out finds:

1. **A disposed asset still depreciating.** The asset was sold, the gain was recorded, and the
   register was never updated. Depreciation continues on something the client no longer owns.
2. **Beginning accumulated depreciation that does not agree to the prior-year return.** Almost
   always introduced by a software conversion, and it silently misstates every subsequent year.
3. **An asset depreciated past its cost.** Arithmetically impossible and surprisingly common,
   usually from a life change applied retroactively.
4. **Additions never added.** Capital expenditure expensed or sitting in construction in
   progress, so depreciation is understated and the balance sheet is wrong.

None of these show up by reading the depreciation expense figure. They show up in a rollforward,
which is why this skill starts there.

## The gate

No clean workpaper unless all of these hold:

1. **The cost rollforward foots.** Beginning cost + additions − disposals = ending cost, in total
   and by asset class.
2. **The accumulated depreciation rollforward foots.** Beginning accumulated + current-year
   expense − accumulated on disposals = ending accumulated.
3. **Beginning balances equal prior-year ending balances.** Supplied from the prior-year return
   or financial statements — not from the software's current state of the file, which is what
   introduced the error in the first place.
4. **Register totals agree to the return or trial balance** for cost, accumulated depreciation,
   and current-year depreciation expense.
5. **No asset has accumulated depreciation exceeding its depreciable basis.**

## Standing rule on tax figures

**Do not state a recovery period, MACRS percentage, bonus depreciation rate, Section 179 limit,
threshold, or phase-out from memory.** These change, and several have changed more than once
recently.

What the script does instead:

- **Recomputes straight-line depreciation exactly** where the schedule states a straight-line
  method, using the cost, salvage, life, and convention *on the schedule*. This is arithmetic and
  is checked precisely.
- **For MACRS and other accelerated methods, tests internal consistency only** — that
  accumulated depreciation does not exceed basis, that prior accumulated plus current equals
  ending, that the asset is in service, that the life and class are populated and consistent
  across years — and flags each asset for verification of the *rate* against current
  instructions. It does not reproduce percentage tables.
- **Flags every asset whose life or method changed from the prior year**, because that is either
  an error or a method change requiring proper procedure rather than a schedule edit.

Where a figure must be confirmed, the workpaper says *"verify recovery period and convention for
[asset class] placed in service [date] against current instructions for Form 4562."* That is more
useful than a number recalled from an earlier year, and it cannot go stale.

## Inputs

1. **Fixed asset register**, current year, per asset: description, class, acquisition date, in-service
   date, cost, salvage, method, life, convention, prior accumulated depreciation, current-year
   depreciation, ending accumulated, disposal date and proceeds where applicable.
2. **Prior-year ending balances** — total cost and total accumulated depreciation, by class, per
   the prior-year return or financial statements **as filed**.
3. **The return or trial balance figures** to tie to: cost, accumulated depreciation, current-year
   depreciation expense. Form 4562 totals if available.
4. **Book and tax registers separately** if the client maintains both. Do not reconcile a book
   register to a tax return — the difference is a deferred tax item, not an error, and mixing
   them produces a meaningless variance.

## Step 1 — Roll forward and tie

```bash
python3 scripts/depr_tieout.py \
  --register fa_register.csv \
  --prior-year prior_balances.csv \
  --tb-cost 4820115.00 --tb-accum 2140880.00 --tb-depreciation 412655.00 \
  --basis tax --year 2025 \
  --client "Kestrel Fabrication LLC" \
  --out "Kestrel - 2025 Fixed Asset Tie-Out.xlsx"
```

Seven tests run:

- **Test 1** — Cost rollforward foots, in total and by class
- **Test 2** — Accumulated depreciation rollforward foots
- **Test 3** — Beginning balances agree to prior-year ending
- **Test 4** — Register totals agree to the trial balance or return
- **Test 5** — No asset over-depreciated; accumulated ≤ depreciable basis for every asset
- **Test 6** — Straight-line assets recomputed and agreed
- **Test 7** — Asset-level integrity (see below)

## Step 2 — Asset-level review

Test 7 checks each asset for the conditions that indicate a stale or broken register:

- **Disposed but still depreciating** — a disposal date with current-year depreciation running
  past it. The most common real finding.
- **Disposed but not removed** — a disposal date with cost still in ending balances.
- **Gain or loss not computed** — a disposal with proceeds but no gain or loss, or a gain
  computed on the wrong accumulated depreciation figure. The script recomputes proceeds less net
  book value.
- **In service before acquisition**, or an in-service date after year end.
- **Fully depreciated but still held** — legitimate, and it should remain on the register.
  Depreciation still running on it is not legitimate.
- **Zero or negative cost**, or depreciation with no cost.
- **Missing method, life, or convention** — the schedule cannot be recomputed or reviewed, which
  is itself a finding.
- **Life or method changed from the prior year** — flagged for every affected asset.
- **An addition dated in a prior year appearing for the first time** — either a late-recorded
  addition needing prior-year consideration, or a duplicate.
- **Duplicate assets** — same description, cost, and in-service date under different asset IDs.
  A double-recorded asset doubles depreciation and is invisible in totals.

## Step 3 — Deliver

**Workbook tabs:**

1. **Tie-Out Summary** — the seven tests with amounts and differences, the verdict, and the
   figures requiring confirmation against current authority. The signable page.
2. **Rollforward** — cost and accumulated depreciation, beginning through ending, in total and by
   asset class, with the prior-year agreement shown.
3. **Asset Detail** — every asset with recomputed depreciation, difference, net book value, and
   flags.
4. **Exceptions** — every asset-level finding with the asset, the condition, and the dollar
   effect where quantifiable.
5. **Additions** — current-year additions with in-service dates, method, life, and the note to
   verify recovery period and any bonus or Section 179 election against current instructions.
6. **Disposals** — each disposal with proceeds, cost, accumulated depreciation removed, net book
   value, and recomputed gain or loss, plus whether depreciation correctly stopped.
7. **Book vs Tax** — where both registers are supplied, the difference by class, which is the
   deferred tax input rather than an error.

**Then, in chat:** whether the rollforwards foot, whether beginning balances agree to the prior
year, the current-year depreciation figure and whether it ties, and the asset-level exceptions
worth acting on. Lead with any beginning-balance disagreement — it affects every year going
forward and is the hardest to unwind later.

## What to escalate

- **Beginning accumulated depreciation not agreeing to the prior-year return.** Determine whether
  the prior year, the current year, or both are wrong before proceeding. This can require an
  amended return or a method-change filing, and it does not resolve itself.
- **A retroactive life or method change made by editing the schedule.** A change in depreciation
  method or life generally requires a prescribed procedure, not a spreadsheet edit. Report what
  changed and route the treatment question to the signer.
- **Assets on the register that no longer physically exist.** If no physical inventory has been
  performed in years, say so — the register is unverified as to existence, which is a scope point
  as much as a control one.
- **Construction in progress carried for multiple years** without being placed in service.
  Depreciation may be understated and the classification may be wrong.
- **Repairs and maintenance containing capital items**, or additions that look like repairs. Worth
  scanning the expense account; this is a frequent examination adjustment in both directions.
- **Bonus or Section 179 elections that appear inconsistent** across assets placed in service in
  the same year, or state treatment applied without regard to state decoupling — many states do
  not conform, and each state has to be tested separately.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry. The
register is read-only. Output filenames carry client name and year only.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*

---

# Appendix — bundled files

If you installed the `.skill` package these files are already in place and you can ignore this appendix. If you copied the skill as text, create the files below at the paths shown, alongside your `SKILL.md`. The skill will not run without them.

## `scripts/depr_tieout.py`

```python
#!/usr/bin/env python3
"""
Fixed asset register tie-out: rollforwards, prior-year agreement, and asset-level
integrity.

Seven tests. No clean workpaper unless the rollforwards foot, beginning balances
agree to the prior year as filed, the register agrees to the trial balance or
return, and no asset is depreciated beyond its basis.

Asserts NO recovery periods, MACRS percentages, bonus rates, or Section 179
limits. Straight-line assets are recomputed exactly from the cost, salvage, life
and convention ON THE SCHEDULE. Accelerated methods are tested for internal
consistency only and flagged for rate verification against current instructions.

Usage:
    python3 depr_tieout.py --register fa_register.csv \
        --prior-year prior_balances.csv \
        --tb-cost 4820115.00 --tb-accum 2140880.00 --tb-depreciation 412655.00 \
        --basis tax --year 2025 --client "Kestrel Fabrication LLC" \
        --out "Kestrel - 2025 Fixed Asset Tie-Out.xlsx"

--register CSV:
    asset_id, description, asset_class, acquisition_date, in_service_date,
    cost, salvage, method, life_years, convention, prior_accum, current_depr,
    ending_accum, disposal_date, proceeds, gain_loss,
    prior_method (optional), prior_life_years (optional)

  method: SL / straight-line recomputed exactly. MACRS, DB, DDB, SYD and others
  tested for consistency only.

--prior-year CSV (from the prior-year return or financial statements AS FILED):
    asset_class, ending_cost, ending_accum
    (a row with asset_class "TOTAL" is accepted for totals only)
"""

from __future__ import annotations

import argparse
import csv
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
DATEF = "yyyy-mm-dd"

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
TOP = Border(top=Side(style="thin"))
DBL = Border(top=Side(style="thin"), bottom=Side(style="double"))

SL_METHODS = {"sl", "s/l", "straight-line", "straight line", "straightline"}
HALF_YEAR = {"hy", "half-year", "half year", "mq", "mid-quarter", "mid quarter",
             "mm", "mid-month", "mid month"}


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


def d0(raw):
    v = dec(raw)
    return v if v is not None else ZERO


def clean(s) -> str:
    return " ".join(str(s or "").split())


def pdate(raw):
    s = clean(raw)
    if not s:
        return None
    for f in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d-%b-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s.split(" ")[0], f).date()
        except ValueError:
            continue
    raise ValueError(f"Unrecognized date {raw!r}. Prefer ISO YYYY-MM-DD.")


def load_register(path: Path) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No assets in {path}")
    out = []
    for i, r in enumerate(rows, start=1):
        a = {
            "row": i,
            "asset_id": clean(r.get("asset_id")) or f"FA{i:05d}",
            "description": clean(r.get("description")),
            "asset_class": clean(r.get("asset_class")) or "(unclassified)",
            "acquisition_date": pdate(r.get("acquisition_date")),
            "in_service_date": pdate(r.get("in_service_date")),
            "cost": d0(r.get("cost")),
            "salvage": d0(r.get("salvage")),
            "method": clean(r.get("method")),
            "life_years": dec(r.get("life_years")),
            "convention": clean(r.get("convention")),
            "prior_accum": d0(r.get("prior_accum")),
            "current_depr": d0(r.get("current_depr")),
            "ending_accum": dec(r.get("ending_accum")),
            "disposal_date": pdate(r.get("disposal_date")),
            "proceeds": dec(r.get("proceeds")),
            "gain_loss": dec(r.get("gain_loss")),
            "prior_method": clean(r.get("prior_method")),
            "prior_life_years": dec(r.get("prior_life_years")),
            "flags": [], "recomputed": None, "recomp_diff": None,
        }
        if a["ending_accum"] is None:
            a["ending_accum"] = a["prior_accum"] + a["current_depr"]
        out.append(a)
    return out


def load_prior(path: Path | None) -> dict:
    out = {"by_class": {}, "total_cost": None, "total_accum": None}
    if not path or not path.exists():
        return out
    with path.open(newline="", encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            cls = clean(r.get("asset_class"))
            c, ac = dec(r.get("ending_cost")), dec(r.get("ending_accum"))
            if cls.upper() == "TOTAL":
                out["total_cost"], out["total_accum"] = c, ac
            elif cls:
                out["by_class"][cls] = {"cost": c or ZERO, "accum": ac or ZERO}
    if out["total_cost"] is None and out["by_class"]:
        out["total_cost"] = sum((v["cost"] for v in out["by_class"].values()), ZERO)
        out["total_accum"] = sum((v["accum"] for v in out["by_class"].values()), ZERO)
    return out


# ------------------------------------------------------------------ recompute

def recompute_sl(a: dict, year_end: date) -> Decimal | None:
    """Exact straight-line recomputation from the schedule's own inputs."""
    if a["method"].lower() not in SL_METHODS:
        return None
    if not a["life_years"] or a["life_years"] <= ZERO:
        return None
    basis = a["cost"] - a["salvage"]
    if basis <= ZERO:
        return ZERO
    annual = basis / a["life_years"]
    isd = a["in_service_date"]
    first_year = bool(isd and isd.year == year_end.year)
    factor = Decimal(1)
    if first_year:
        conv = a["convention"].lower()
        if conv in HALF_YEAR or not conv:
            factor = Decimal("0.5")
        else:
            months = 12 - isd.month + 1
            factor = Decimal(months) / Decimal(12)
    remaining = basis - a["prior_accum"]
    if remaining <= ZERO:
        return ZERO
    expected = annual * factor
    if expected > remaining:
        expected = remaining
    if a["disposal_date"] and a["disposal_date"].year == year_end.year:
        # disposal-year convention varies; report rather than assert
        return None
    return expected.quantize(Decimal("0.01"))


def asset_flags(assets: list[dict], year: int, year_end: date) -> list[str]:
    problems = []
    dupes = defaultdict(list)

    for a in assets:
        f = a["flags"]
        basis = a["cost"] - a["salvage"]

        if a["ending_accum"] > basis and basis > ZERO:
            f.append(f"OVER-DEPRECIATED: accumulated {a['ending_accum']:,.2f} exceeds "
                     f"depreciable basis {basis:,.2f} by "
                     f"{a['ending_accum'] - basis:,.2f} - arithmetically impossible, "
                     f"usually a life change applied retroactively")
        expected_end = a["prior_accum"] + a["current_depr"]
        if a["disposal_date"] is None and a["ending_accum"] != expected_end:
            f.append(f"accumulated does not roll: prior {a['prior_accum']:,.2f} + current "
                     f"{a['current_depr']:,.2f} = {expected_end:,.2f}, but ending accum "
                     f"is {a['ending_accum']:,.2f}")
        if a["disposal_date"]:
            if a["disposal_date"].year < year and a["current_depr"] != ZERO:
                f.append(f"DISPOSED {a['disposal_date']} in a prior year but still "
                         f"depreciating {a['current_depr']:,.2f} - depreciation is "
                         f"running on an asset the client no longer owns")
            if a["disposal_date"].year < year and a["cost"] != ZERO:
                f.append(f"disposed {a['disposal_date']} in a prior year but cost "
                         f"{a['cost']:,.2f} remains in the register")
            nbv = a["cost"] - a["ending_accum"]
            if a["proceeds"] is not None:
                expected_gl = a["proceeds"] - nbv
                if a["gain_loss"] is None:
                    f.append(f"disposal with proceeds {a['proceeds']:,.2f} but no gain or "
                             f"loss recorded - recomputed {expected_gl:,.2f}")
                elif a["gain_loss"] != expected_gl:
                    f.append(f"gain/loss {a['gain_loss']:,.2f} does not agree to proceeds "
                             f"less net book value ({a['proceeds']:,.2f} - {nbv:,.2f} = "
                             f"{expected_gl:,.2f})")
        if a["acquisition_date"] and a["in_service_date"] and \
                a["in_service_date"] < a["acquisition_date"]:
            f.append(f"in service {a['in_service_date']} before acquisition "
                     f"{a['acquisition_date']}")
        if a["in_service_date"] and a["in_service_date"] > year_end:
            f.append(f"in-service date {a['in_service_date']} is after year end "
                     f"{year_end} - not yet depreciable")
        if a["cost"] <= ZERO and a["current_depr"] != ZERO:
            f.append(f"depreciation {a['current_depr']:,.2f} with cost {a['cost']:,.2f}")
        if a["cost"] < ZERO:
            f.append(f"negative cost {a['cost']:,.2f}")
        if basis > ZERO and a["prior_accum"] >= basis and a["current_depr"] > ZERO \
                and not a["disposal_date"]:
            f.append(f"fully depreciated at the start of the year but still taking "
                     f"{a['current_depr']:,.2f} of depreciation")
        missing = [k for k in ("method", "convention") if not a[k]]
        if not a["life_years"]:
            missing.append("life_years")
        if missing and a["cost"] > ZERO:
            f.append(f"missing {', '.join(missing)} - the schedule cannot be recomputed "
                     f"or reviewed, which is itself a finding")
        if a["prior_method"] and a["method"] and \
                a["prior_method"].lower() != a["method"].lower():
            f.append(f"METHOD CHANGED from '{a['prior_method']}' to '{a['method']}' - a "
                     f"method change generally requires a prescribed procedure, not a "
                     f"schedule edit. Route the treatment question to the signer.")
        if a["prior_life_years"] and a["life_years"] and \
                a["prior_life_years"] != a["life_years"]:
            f.append(f"LIFE CHANGED from {a['prior_life_years']} to {a['life_years']} "
                     f"years - same consideration as a method change")
        if a["in_service_date"] and a["in_service_date"].year < year and \
                a["prior_accum"] == ZERO and a["cost"] > ZERO:
            f.append(f"in service {a['in_service_date']} (a prior year) but zero prior "
                     f"accumulated depreciation - a late-recorded addition needing "
                     f"prior-year consideration, or a duplicate")

        if a["recomputed"] is not None and a["recomp_diff"] not in (None, ZERO):
            f.append(f"straight-line recomputation differs: schedule "
                     f"{a['current_depr']:,.2f} vs recomputed {a['recomputed']:,.2f} "
                     f"(difference {a['recomp_diff']:,.2f})")

        dupes[(clean(a["description"]).lower(), a["cost"], a["in_service_date"])].append(a)

        for msg in f:
            problems.append(f"{a['asset_id']} {a['description'][:36]}: {msg}")

    for key, group in dupes.items():
        if len(group) > 1 and key[1] > ZERO and key[0]:
            ids = ", ".join(x["asset_id"] for x in group)
            for x in group:
                x["flags"].append(f"possible DUPLICATE asset: same description, cost and "
                                  f"in-service date as {ids} - a double-recorded asset "
                                  f"doubles depreciation and is invisible in totals")
            problems.append(f"Possible duplicate assets ({ids}): {group[0]['description']} "
                            f"at {key[1]:,.2f}")
    return problems


# ---------------------------------------------------------------------- tests

def rollforward(assets: list[dict], prior: dict, year: int):
    by_class = defaultdict(lambda: {
        "begin_cost": ZERO, "additions": ZERO, "disposals": ZERO, "end_cost": ZERO,
        "begin_accum": ZERO, "current": ZERO, "disp_accum": ZERO, "end_accum": ZERO})
    for a in assets:
        c = by_class[a["asset_class"]]
        is_addition = a["in_service_date"] and a["in_service_date"].year == year
        is_disposal = a["disposal_date"] and a["disposal_date"].year == year

        if is_addition:
            c["additions"] += a["cost"]
        else:
            c["begin_cost"] += a["cost"]
        if is_disposal:
            c["disposals"] += a["cost"]
            c["disp_accum"] += a["ending_accum"]
        else:
            c["end_cost"] += a["cost"]

        if not is_addition:
            c["begin_accum"] += a["prior_accum"]
        c["current"] += a["current_depr"]
        if not is_disposal:
            c["end_accum"] += a["ending_accum"]

    rows = []
    for cls, c in sorted(by_class.items()):
        cost_computed = c["begin_cost"] + c["additions"] - c["disposals"]
        accum_computed = c["begin_accum"] + c["current"] - c["disp_accum"]
        rows.append({
            "asset_class": cls, **c,
            "cost_computed": cost_computed,
            "cost_diff": cost_computed - c["end_cost"],
            "accum_computed": accum_computed,
            "accum_diff": accum_computed - c["end_accum"],
            "prior_cost": (prior["by_class"].get(cls) or {}).get("cost"),
            "prior_accum_bal": (prior["by_class"].get(cls) or {}).get("accum"),
        })
    return rows


def totals_of(rows, key):
    return sum((r[key] for r in rows), ZERO)


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


def sheet_summary(wb, meta, tests, overall, escalations, verify_notes):
    ws = wb.active
    ws.title = "Tie-Out Summary"
    widths(ws, {1: 4, 2: 52, 3: 18, 4: 18, 5: 62})
    r = 1

    def line(label, a=None, b=None, *, bold=False, size=11, fill=None, note="",
             border=None):
        nonlocal r
        c = ws.cell(row=r, column=2, value=label)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c.fill = fill
        for col, val in ((3, a), (4, b)):
            if val is None:
                continue
            cc = ws.cell(row=r, column=col,
                         value=float(val) if isinstance(val, Decimal) else val)
            if isinstance(val, Decimal):
                cc.number_format = MONEY
            cc.font = Font(bold=bold)
            if border:
                cc.border = border
        if note:
            n = ws.cell(row=r, column=5, value=note)
            n.font = Font(italic=True, size=9)
            n.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    line("FIXED ASSET AND DEPRECIATION TIE-OUT", bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    v = ws.cell(row=r, column=2,
                value="TIED OUT - rollforwards foot, beginning balances agree, register "
                      "agrees to the return"
                if overall else
                "NOT TIED OUT - see the failing test(s) below")
    v.font = OK_FONT if overall else BAD_FONT
    if not overall:
        v.fill = BAD_FILL
    r += 2

    for t in tests:
        line(f"TEST {t['num']} - {t['name']}", bold=True, size=12, fill=SUB_FILL)
        if t.get("skipped"):
            line("  " + t["note"], fill=WARN_FILL)
            r += 1
            continue
        counts = set(t.get("counts", []))
        for label, val in t["lines"]:
            line("  " + label, int(val) if label in counts else val)
        line("  Assets flagged" if t.get("is_count") else "  Difference",
             int(t["difference"]) if t.get("is_count") else t["difference"],
             bold=True, border=TOP, note=t.get("note", ""))
        c = ws.cell(row=r - 1, column=3)
        c.font = OK_FONT if t["passed"] else BAD_FONT
        if not t["passed"]:
            c.fill = BAD_FILL
        r += 1

    if escalations:
        line("ESCALATE", bold=True, size=12, fill=BAD_FILL)
        for e in escalations:
            line("  " + e)
        r += 1

    line("FIGURES REQUIRING CONFIRMATION AGAINST CURRENT AUTHORITY", bold=True, size=12,
         fill=WARN_FILL)
    for n in verify_notes:
        line("  " + n)
    r += 1
    line("This workpaper asserts no recovery period, MACRS percentage, bonus rate, or "
         "Section 179 limit. Straight-line assets are recomputed exactly from the "
         "schedule's own inputs; accelerated methods are tested for internal "
         "consistency only.", bold=True)
    r += 2
    line("Prepared by / date:  __________________  ____________", bold=True)
    line("Reviewed by / date:  __________________  ____________", bold=True)


def sheet_rollforward(wb, rows, prior):
    ws = wb.create_sheet("Rollforward")
    heads = ["Asset class", "Begin cost", "Additions", "Disposals", "Computed end cost",
             "End cost per register", "Cost diff", "PY end cost", "PY cost diff",
             "Begin accum", "Current depr", "Accum on disposals", "Computed end accum",
             "End accum per register", "Accum diff", "PY end accum", "PY accum diff"]
    ws.append(heads)
    hdr(ws, len(heads))
    for x in rows:
        pcd = (x["begin_cost"] - x["prior_cost"]) if x["prior_cost"] is not None else None
        pad = (x["begin_accum"] - x["prior_accum_bal"]) if x["prior_accum_bal"] is not None else None
        ws.append([x["asset_class"], float(x["begin_cost"]), float(x["additions"]),
                   float(x["disposals"]), float(x["cost_computed"]),
                   float(x["end_cost"]), float(x["cost_diff"]),
                   float(x["prior_cost"]) if x["prior_cost"] is not None else None,
                   float(pcd) if pcd is not None else None,
                   float(x["begin_accum"]), float(x["current"]),
                   float(x["disp_accum"]), float(x["accum_computed"]),
                   float(x["end_accum"]), float(x["accum_diff"]),
                   float(x["prior_accum_bal"]) if x["prior_accum_bal"] is not None else None,
                   float(pad) if pad is not None else None])
    last = ws.max_row
    ws.append(["TOTAL"] + [f"=SUM({get_column_letter(c)}2:{get_column_letter(c)}{last})"
                           for c in range(2, len(heads) + 1)])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in range(1, len(heads)):
            row[i].number_format = MONEY
        for i in (6, 8, 14, 16):
            if isinstance(row[i].value, (int, float)) and row[i].value:
                row[i].font, row[i].fill = BAD_FONT, BAD_FILL
    for c in range(1, len(heads) + 1):
        cell = ws.cell(row=ws.max_row, column=c)
        cell.font, cell.border = Font(bold=True), TOP
    ws.freeze_panes = "B2"
    widths(ws, {i: 16 for i in range(1, len(heads) + 1)} | {1: 26})


def sheet_assets(wb, assets):
    ws = wb.create_sheet("Asset Detail")
    heads = ["Asset ID", "Description", "Class", "Acquired", "In service", "Cost",
             "Salvage", "Method", "Life", "Conv", "Prior accum", "Current depr",
             "Recomputed (SL only)", "Diff", "End accum", "NBV", "Disposed",
             "Proceeds", "Gain/loss", "Flags"]
    ws.append(heads)
    hdr(ws, len(heads))
    for a in assets:
        ws.append([a["asset_id"], a["description"], a["asset_class"],
                   a["acquisition_date"], a["in_service_date"], float(a["cost"]),
                   float(a["salvage"]), a["method"] or None,
                   float(a["life_years"]) if a["life_years"] else None,
                   a["convention"] or None, float(a["prior_accum"]),
                   float(a["current_depr"]),
                   float(a["recomputed"]) if a["recomputed"] is not None else None,
                   float(a["recomp_diff"]) if a["recomp_diff"] is not None else None,
                   float(a["ending_accum"]),
                   float(a["cost"] - a["ending_accum"]),
                   a["disposal_date"],
                   float(a["proceeds"]) if a["proceeds"] is not None else None,
                   float(a["gain_loss"]) if a["gain_loss"] is not None else None,
                   "; ".join(a["flags"])])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in (3, 4, 16):
            row[i].number_format = DATEF
        for i in (5, 6, 10, 11, 12, 13, 14, 15, 17, 18):
            row[i].number_format = MONEY
        if row[13].value not in (None, 0):
            row[13].font, row[13].fill = BAD_FONT, BAD_FILL
        if row[19].value:
            row[19].fill = BAD_FILL
        if isinstance(row[15].value, (int, float)) and row[15].value < 0:
            row[15].font, row[15].fill = BAD_FONT, BAD_FILL
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:T{max(ws.max_row, 2)}"
    widths(ws, {1: 12, 2: 34, 3: 20, 4: 12, 5: 12, 6: 14, 7: 11, 8: 10, 9: 7,
                10: 7, 11: 14, 12: 14, 13: 18, 14: 12, 15: 14, 16: 14, 17: 12,
                18: 13, 19: 13, 20: 80})


def sheet_list(wb, title, rows, heads, fmt_map, empty_msg, footer=""):
    ws = wb.create_sheet(title)
    ws.append(heads)
    hdr(ws, len(heads))
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i, f in fmt_map.items():
            if i < len(row):
                row[i].number_format = f
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value=empty_msg)
    elif footer:
        rr = ws.max_row + 2
        c = ws.cell(row=rr, column=1, value=footer)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.font = Font(italic=True)
    ws.freeze_panes = "A2"
    widths(ws, {i: 18 for i in range(1, len(heads) + 1)} | {2: 34, len(heads): 70})


# ------------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--register", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--prior-year")
    ap.add_argument("--tb-cost")
    ap.add_argument("--tb-accum")
    ap.add_argument("--tb-depreciation")
    ap.add_argument("--basis", choices=["tax", "book"], default="tax")
    ap.add_argument("--year", type=int, required=True)
    ap.add_argument("--year-end", help="YYYY-MM-DD, defaults to Dec 31 of --year")
    ap.add_argument("--client", default="")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    year_end = pdate(args.year_end) if args.year_end else date(args.year, 12, 31)
    assets = load_register(Path(args.register))
    prior = load_prior(Path(args.prior_year) if args.prior_year else None)

    for a in assets:
        rc = recompute_sl(a, year_end)
        a["recomputed"] = rc
        a["recomp_diff"] = (a["current_depr"] - rc) if rc is not None else None

    problems = asset_flags(assets, args.year, year_end)
    rows = rollforward(assets, prior, args.year)

    tot = {k: totals_of(rows, k) for k in
           ("begin_cost", "additions", "disposals", "end_cost", "cost_computed",
            "begin_accum", "current", "disp_accum", "end_accum", "accum_computed")}
    tb_cost, tb_accum, tb_depr = dec(args.tb_cost), dec(args.tb_accum), dec(args.tb_depreciation)

    tests = []

    tests.append({"num": 1, "name": "Cost rollforward foots",
                  "lines": [("Beginning cost", tot["begin_cost"]),
                            ("Add: additions", tot["additions"]),
                            ("Less: disposals", -tot["disposals"]),
                            ("= Computed ending cost", tot["cost_computed"]),
                            ("Ending cost per register", tot["end_cost"])],
                  "difference": tot["cost_computed"] - tot["end_cost"],
                  "passed": tot["cost_computed"] == tot["end_cost"]})

    tests.append({"num": 2, "name": "Accumulated depreciation rollforward foots",
                  "lines": [("Beginning accumulated", tot["begin_accum"]),
                            ("Add: current-year depreciation", tot["current"]),
                            ("Less: accumulated on disposals", -tot["disp_accum"]),
                            ("= Computed ending accumulated", tot["accum_computed"]),
                            ("Ending accumulated per register", tot["end_accum"])],
                  "difference": tot["accum_computed"] - tot["end_accum"],
                  "passed": tot["accum_computed"] == tot["end_accum"]})

    if prior["total_cost"] is not None:
        dc = tot["begin_cost"] - prior["total_cost"]
        da = tot["begin_accum"] - (prior["total_accum"] or ZERO)
        tests.append({"num": 3,
                      "name": "Beginning balances agree to prior year as filed",
                      "lines": [("Beginning cost per register", tot["begin_cost"]),
                                ("Prior-year ending cost as filed", prior["total_cost"]),
                                ("Beginning accumulated per register", tot["begin_accum"]),
                                ("Prior-year ending accumulated as filed",
                                 prior["total_accum"] or ZERO)],
                      "difference": dc + da,
                      "passed": dc == ZERO and da == ZERO,
                      "note": "" if (dc == ZERO and da == ZERO) else
                              f"cost off by {dc:,.2f}, accumulated off by {da:,.2f}. "
                              f"Almost always a software conversion. This silently "
                              f"misstates every subsequent year and does not resolve "
                              f"itself."})
    else:
        tests.append({"num": 3, "name": "Beginning balances agree to prior year",
                      "skipped": True, "passed": True,
                      "note": "NOT PERFORMED - prior-year balances not supplied. This is "
                              "the test that catches a software conversion error, and it "
                              "is the hardest error to unwind later. Obtain the "
                              "prior-year return as filed.",
                      "lines": [], "difference": ZERO})

    if tb_cost is not None or tb_accum is not None or tb_depr is not None:
        lines, diff = [], ZERO
        for label, reg, tb in (("Cost", tot["end_cost"], tb_cost),
                               ("Accumulated depreciation", tot["end_accum"], tb_accum),
                               ("Current-year depreciation", tot["current"], tb_depr)):
            if tb is None:
                continue
            lines += [(f"{label} per register", reg), (f"{label} per return / TB", tb)]
            diff += reg - tb
        tests.append({"num": 4, "name": "Register agrees to the return or trial balance",
                      "lines": lines, "difference": diff, "passed": diff == ZERO})
    else:
        tests.append({"num": 4, "name": "Register agrees to the return or trial balance",
                      "skipped": True, "passed": True,
                      "note": "NOT PERFORMED - no return or trial balance figures supplied",
                      "lines": [], "difference": ZERO})

    over = [a for a in assets
            if (a["cost"] - a["salvage"]) > ZERO
            and a["ending_accum"] > (a["cost"] - a["salvage"])]
    tests.append({"num": 5, "name": "No asset depreciated beyond its basis",
                  "counts": ["Assets over-depreciated"],
                  "lines": [("Assets over-depreciated", Decimal(len(over))),
                            ("Excess depreciation",
                             sum((a["ending_accum"] - (a["cost"] - a["salvage"])
                                  for a in over), ZERO))],
                  "difference": sum((a["ending_accum"] - (a["cost"] - a["salvage"])
                                     for a in over), ZERO),
                  "passed": not over})

    sl = [a for a in assets if a["recomputed"] is not None]
    sl_bad = [a for a in sl if a["recomp_diff"] != ZERO]
    tests.append({"num": 6, "name": "Straight-line assets recomputed and agreed",
                  "counts": ["Straight-line assets recomputed", "Assets disagreeing"],
                  "lines": [("Straight-line assets recomputed", Decimal(len(sl))),
                            ("Assets disagreeing", Decimal(len(sl_bad))),
                            ("Net difference",
                             sum((a["recomp_diff"] for a in sl_bad), ZERO))],
                  "difference": sum((a["recomp_diff"] for a in sl_bad), ZERO),
                  "passed": not sl_bad,
                  "note": "Accelerated methods are not recomputed - their rates must be "
                          "verified against current instructions."})

    n_flagged = len([a for a in assets if a["flags"]])
    tests.append({"num": 7, "name": "Asset-level integrity",
                  "counts": ["Assets in register", "Assets with findings"],
                  "is_count": True,
                  "lines": [("Assets in register", Decimal(len(assets))),
                            ("Assets with findings", Decimal(n_flagged))],
                  "difference": Decimal(n_flagged),
                  "passed": n_flagged == 0,
                  "note": "" if not n_flagged else
                          f"{n_flagged} asset(s) carry findings - see Exceptions and "
                          f"Asset Detail. The count is items, not dollars."})

    overall = all(t["passed"] for t in tests)

    escalations = []
    t3 = tests[2]
    if not t3.get("skipped") and not t3["passed"]:
        escalations.append(
            "Beginning balances do not agree to the prior year as filed. Determine "
            "whether the prior year, the current year, or both are wrong before "
            "proceeding - this may require an amended return or a method-change filing.")
    still = [a for a in assets if a["disposal_date"]
             and a["disposal_date"].year < args.year and a["current_depr"] != ZERO]
    if still:
        escalations.append(
            f"{len(still)} asset(s) disposed in a prior year are still depreciating, "
            f"totalling {sum((a['current_depr'] for a in still), ZERO):,.2f} of "
            f"current-year depreciation on assets the client no longer owns.")
    if over:
        escalations.append(
            f"{len(over)} asset(s) depreciated beyond basis by "
            f"{sum((a['ending_accum'] - (a['cost'] - a['salvage']) for a in over), ZERO):,.2f}.")
    changed = [a for a in assets
               if any(f.startswith(("METHOD CHANGED", "LIFE CHANGED")) for f in a["flags"])]
    if changed:
        escalations.append(
            f"{len(changed)} asset(s) had a method or life change. A change in method or "
            f"recovery period generally requires a prescribed procedure, not a schedule "
            f"edit. Route the treatment question to the signer.")

    additions = [a for a in assets
                 if a["in_service_date"] and a["in_service_date"].year == args.year]
    verify = [
        "Recovery period and convention for each current-year addition, against current "
        "instructions for Form 4562",
        "Bonus depreciation rate and eligibility for the year the asset was placed in "
        "service",
        "Section 179 limit, phase-out threshold, and any business-income limitation",
        "State conformity - many states decouple from federal bonus and Section 179, so "
        "each state must be tested separately",
        "MACRS percentages for any accelerated asset - this workpaper tests consistency "
        "only, not the rate",
    ]

    # ---- console
    print("=" * 74)
    print(f"FIXED ASSET AND DEPRECIATION TIE-OUT ({args.basis} basis)")
    print("=" * 74)
    print(f"Client : {args.client or '(not stated)'}    Year: {args.year}  "
          f"(year end {year_end})")
    print(f"Assets : {len(assets)}    Additions: {len(additions)}    "
          f"Disposals: {len([a for a in assets if a['disposal_date'] and a['disposal_date'].year == args.year])}")
    print()
    for t in tests:
        if t.get("skipped"):
            print(f"Test {t['num']}: NOT PERFORMED - {t['note'][:80]}")
            continue
        if t.get("is_count"):
            print(f"Test {t['num']}: {'PASS' if t['passed'] else '*** FAIL ***':<14} "
                  f"{int(t['difference']):>10} asset(s) flagged   {t['name']}")
        else:
            print(f"Test {t['num']}: {'PASS' if t['passed'] else '*** FAIL ***':<14} "
                  f"diff {t['difference']:>16,.2f}   {t['name']}")
    print()
    print(f"Cost      : begin {tot['begin_cost']:>14,.2f} + add {tot['additions']:>13,.2f} "
          f"- disp {tot['disposals']:>13,.2f} = {tot['cost_computed']:>14,.2f}")
    print(f"Accum     : begin {tot['begin_accum']:>14,.2f} + curr {tot['current']:>13,.2f} "
          f"- disp {tot['disp_accum']:>13,.2f} = {tot['accum_computed']:>14,.2f}")
    print(f"Net book value: {tot['end_cost'] - tot['end_accum']:>16,.2f}")

    if problems:
        print(f"\nASSET-LEVEL FINDINGS ({len(problems)}):")
        for p in problems[:20]:
            print(f"  ! {p}")
        if len(problems) > 20:
            print(f"  ... and {len(problems) - 20} more - see Asset Detail")
    if escalations:
        print("\nESCALATE:")
        for e in escalations:
            print(f"  ! {e}")

    if not overall and not args.force:
        print("\n" + "=" * 74)
        print("WORKBOOK NOT WRITTEN. The register does not tie out.")
        print("Fixed asset errors compound - a wrong beginning balance misstates every")
        print("subsequent year. Resolve before relying on the depreciation figure.")
        print("=" * 74)
        return 1

    meta = {
        "Client": args.client or "(not stated)",
        "Year": args.year,
        "Year end": str(year_end),
        "Basis": args.basis,
        "Assets in register": len(assets),
        "Current-year additions": len(additions),
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    wb = Workbook()
    sheet_summary(wb, meta, tests, overall, escalations, verify)
    sheet_rollforward(wb, rows, prior)
    sheet_assets(wb, assets)
    sheet_list(wb, "Exceptions",
               [[i + 1, p] for i, p in enumerate(problems)],
               ["#", "Finding"], {}, "None - no asset-level findings.")
    sheet_list(wb, "Additions",
               [[a["asset_id"], a["description"], a["asset_class"], a["in_service_date"],
                 float(a["cost"]), a["method"], float(a["life_years"]) if a["life_years"] else None,
                 a["convention"], float(a["current_depr"]),
                 "verify recovery period, convention, and any bonus or Section 179 "
                 "election against current instructions for Form 4562"]
                for a in additions],
               ["Asset ID", "Description", "Class", "In service", "Cost", "Method",
                "Life", "Conv", "Current depr", "To verify"],
               {3: DATEF, 4: MONEY, 8: MONEY},
               "No current-year additions.")
    disposals = [a for a in assets if a["disposal_date"]]
    sheet_list(wb, "Disposals",
               [[a["asset_id"], a["description"], a["disposal_date"], float(a["cost"]),
                 float(a["ending_accum"]), float(a["cost"] - a["ending_accum"]),
                 float(a["proceeds"]) if a["proceeds"] is not None else None,
                 float(a["gain_loss"]) if a["gain_loss"] is not None else None,
                 float((a["proceeds"] or ZERO) - (a["cost"] - a["ending_accum"])),
                 "depreciation stopped" if a["current_depr"] == ZERO
                 else f"STILL DEPRECIATING {a['current_depr']:,.2f}"]
                for a in disposals],
               ["Asset ID", "Description", "Disposed", "Cost", "Accum removed", "NBV",
                "Proceeds", "Gain/loss per register", "Recomputed gain/loss", "Status"],
               {2: DATEF, 3: MONEY, 4: MONEY, 5: MONEY, 6: MONEY, 7: MONEY, 8: MONEY},
               "No disposals.",
               footer="Recomputed gain or loss is proceeds less net book value. A "
                      "difference means the gain was computed on the wrong accumulated "
                      "depreciation figure, or the asset was not fully removed.")
    wb.active = 0
    out = Path(args.out)
    if not overall:
        out = out.with_name(out.stem + " [DOES NOT TIE]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\n{'TIED OUT' if overall else 'DOES NOT TIE (forced)'}.  Workbook: {out}")
    return 0 if overall else 1


if __name__ == "__main__":
    sys.exit(main())
```

