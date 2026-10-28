---
name: budget-vs-actual
description: Compares actuals to budget by department and account, applies dollar and percentage thresholds, and refuses to call the report complete until every variance above its applicable threshold has an explanation whose components sum to the variance exactly. Use this whenever the user mentions budget vs actual, BvA report, variance analysis against budget, flux against budget, department spend vs budget, or "why did we go over budget." Trigger on "build the BvA," "run budget vs actual for this quarter," "which departments are over budget," "explain this variance," or when actuals, a budget export, and threshold config are supplied together. Reuses the enforced-explanation gate pattern from return-yoy-variance, applied to budget rather than prior-year comparisons — use return-yoy-variance for year-over-year variance, this one for actual-vs-budget variance.
---

# Budget vs Actual Report

A BvA report that lists variances without forcing an explanation for the ones that matter is a
list, not a report. This skill applies the same enforced-explanation discipline as
`return-yoy-variance`, scoped to actual-vs-budget instead of year-over-year.

## The gate

The script refuses to call the report complete when:

1. **A variance above its threshold has no explanation.** It lands on the Unexplained Variance
   tab, and the workbook summary states plainly the report is not complete while that tab is
   non-empty.
2. **An explanation's components don't sum to the variance.** If component rows are supplied, they
   must add up exactly to the variance they claim to explain — a partial explanation is an
   exception, not a pass.
3. **Aggregation drops a row.** Total actual and total budget on the report must equal the sums of
   the source files exactly.

## Inputs

1. **Actuals** — `department, account, actual_amount`.
2. **Budget export** — `department, account, budget_amount`.
3. **Threshold config** — `department, account, dollar_threshold, pct_threshold`. Either field may
   be blank to disable that test, and `department`/`account` may be blank to scope a threshold more
   broadly — blank-both is the global default. The most specific threshold that matches wins:
   account+department beats account-only beats department-only beats global. **No threshold is
   ever hardcoded** — a department/account with no matching threshold row at all is flagged for
   review rather than silently passed.
4. **Explanations** — `department, account, reason_category, narrative, component_description,
   component_amount`. One row per component; an account can have several rows explaining pieces
   of one variance (a travel overage split into conference airfare and hotel, say).

## Step 1 — Get a blank explanations worksheet for what currently exceeds threshold

```bash
python3 scripts/budget_vs_actual.py --actuals actuals.csv --budget budget.csv \
  --thresholds thresholds.csv --template --out explanations_template.csv
```

This evaluates thresholds first and emits one row per department/account pair found in the
actuals or budget files, ready for explanation narratives to be filled in.

## Step 2 — Build the report

```bash
python3 scripts/budget_vs_actual.py --actuals actuals.csv --budget budget.csv \
  --thresholds thresholds.csv --explanations explanations.csv \
  --period "FY2026 Q1" --client "Northwind Systems Inc" \
  --out "Northwind - Q1 BvA.xlsx"
```

## Deliver

**Workbook tabs:**

1. **BvA Report** — department, account, budget, actual, variance $, variance %, whether it
   exceeds its threshold, whether it's explained, and the explanation narrative.
2. **Unexplained Variance** — the gate. Empty means the report is complete; non-empty means it
   isn't, and the workbook summary says so.
3. **Threshold Config** — which threshold basis (account+department / account / department /
   global) actually applied to each line, so a reviewer can audit the thresholds themselves.

**Then, in chat:** total budget and total actual first, then the count exceeding threshold and the
count still unexplained, then the two or three largest unexplained variances by dollar amount.

## What to escalate

- **A department/account pair with no matching threshold row at all.** The script flags it rather
  than inventing a default — the threshold config is incomplete and needs a decision, not a guess.
- **A variance explained the same way every period** (e.g., "timing" every month for a year) — a
  pattern worth a conversation about whether the budget itself needs revisiting.
- **A component total that keeps falling short of the variance** across periods for the same
  account — suggests an unidentified driver, not a reporting error.
- **Large unexplained variances concentrated in one department.** Report the fact; don't
  characterize whether it reflects a budgeting miss or an operational one.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.
Department spend and budget data never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*

---

# Appendix — bundled files

If you installed the `.skill` package these files are already in place and you can ignore this appendix. If you copied the skill as text, create the files below at the paths shown, alongside your `SKILL.md`. The skill will not run without them.

## `scripts/budget_vs_actual.py`

```python
#!/usr/bin/env python3
"""
Compare actuals to budget by department and account, apply thresholds, and
enforce an explanation for every variance above them.

Gates:
  * every variance that exceeds its applicable threshold must have at least
    one explanation row, or the report states plainly it is NOT complete.
  * where component rows are supplied for a variance, the components must sum
    to the variance exactly - a partial explanation is an exception, never a
    silent pass.
  * total actual and total budget on the report must equal the sums of the
    source files exactly - no row may be dropped in aggregation.

No dollar or percentage threshold is hardcoded. Every threshold is supplied in
--thresholds, scoped as specifically as the user wants (by account, by
department, or a blank-both global default row).

Usage - emit a blank explanations worksheet for whatever already exceeds threshold:
    python3 budget_vs_actual.py --actuals actuals.csv --budget budget.csv \
        --thresholds thresholds.csv --template --out explanations_template.csv

Usage - build the report:
    python3 budget_vs_actual.py --actuals actuals.csv --budget budget.csv \
        --thresholds thresholds.csv --explanations explanations.csv \
        --period "FY2026 Q1" --client "Northwind Systems Inc" \
        --out "Northwind - Q1 BvA.xlsx"

--actuals CSV:       department, account, actual_amount
--budget CSV:        department, account, budget_amount
--thresholds CSV:    department, account, dollar_threshold, pct_threshold
                     department and/or account may be blank - blank-both is
                     the global default row. Most specific match wins:
                     account+department > account-only > department-only >
                     global. A variance is flagged if it exceeds EITHER
                     supplied threshold (either may be blank to disable it).
--explanations CSV:  department, account, reason_category, narrative,
                     component_description, component_amount
                     one row per component; multiple rows may explain one
                     department/account pair.
"""

from __future__ import annotations

import argparse
import csv
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
MONEY = '#,##0.00;[Red](#,##0.00)'
PCT = '0.0"%"'

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")

EXPL_COLS = ["department", "account", "reason_category", "narrative",
             "component_description", "component_amount"]


def dec(raw):
    if raw is None:
        return None
    s = str(raw).strip()
    if s in ("", "-", "--", "n/a", "N/A", "None"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace(",", "").replace("$", "").replace("%", "").strip()
    try:
        v = Decimal(s)
    except InvalidOperation:
        raise ValueError(f"Unparseable amount: {raw!r}")
    return -v if neg else v


def clean(s) -> str:
    return " ".join(str(s or "").split())


def load(path: Path, cols: list[str]) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    out = []
    for i, r in enumerate(rows, start=1):
        x = {k: clean(r.get(k)) for k in cols}
        x["row"] = i
        out.append(x)
    return out


def threshold_for(dept: str, acct: str, thresholds: list[dict]) -> tuple[dict | None, str]:
    """Most specific match wins: account+dept > account-only > dept-only > global."""
    candidates = {
        4: lambda t: t["department"].lower() == dept.lower() and t["account"].lower() == acct.lower() and t["department"] and t["account"],
        3: lambda t: not t["department"] and t["account"].lower() == acct.lower() and t["account"],
        2: lambda t: t["department"].lower() == dept.lower() and not t["account"] and t["department"],
        1: lambda t: not t["department"] and not t["account"],
    }
    for rank in (4, 3, 2, 1):
        for t in thresholds:
            if candidates[rank](t):
                label = {4: "account+department", 3: "account", 2: "department",
                         1: "global default"}[rank]
                return t, label
    return None, "none"


def build(actuals, budget, thresholds, explanations):
    errors = []
    act_idx = defaultdict(Decimal)
    bud_idx = defaultdict(Decimal)
    keys = []
    for a in actuals:
        amt = dec(a["actual_amount"])
        if amt is None:
            errors.append(f"Actuals row {a['row']} ({a['department']}/{a['account']}): "
                          f"actual_amount is required.")
            continue
        k = (a["department"], a["account"])
        if k not in act_idx:
            keys.append(k)
        act_idx[k] += amt
    for b in budget:
        amt = dec(b["budget_amount"])
        if amt is None:
            errors.append(f"Budget row {b['row']} ({b['department']}/{b['account']}): "
                          f"budget_amount is required.")
            continue
        k = (b["department"], b["account"])
        if k not in act_idx and k not in bud_idx:
            keys.append(k)
        bud_idx[k] += amt

    for t in thresholds:
        t["department"] = t["department"]
        t["account"] = t["account"]
        t["dollar_threshold"] = dec(t["dollar_threshold"]) if t["dollar_threshold"] else None
        t["pct_threshold"] = dec(t["pct_threshold"]) if t["pct_threshold"] else None

    expl_idx = defaultdict(list)
    for e in explanations:
        k = (e["department"], e["account"])
        expl_idx[k].append(e)

    report, unexplained, threshold_log = [], [], []
    for k in sorted(set(keys)):
        dept, acct = k
        a = act_idx.get(k, ZERO)
        b = bud_idx.get(k, ZERO)
        variance = a - b
        pct = (variance / b * 100) if b != ZERO else None

        t, basis = threshold_for(dept, acct, thresholds)
        exceeds = False
        if t:
            if t["dollar_threshold"] is not None and abs(variance) > t["dollar_threshold"]:
                exceeds = True
            if t["pct_threshold"] is not None and pct is not None and abs(pct) > t["pct_threshold"]:
                exceeds = True
        else:
            # no threshold at all matches -> cannot evaluate, treated as flagged
            # so it surfaces rather than silently passing unreviewed
            exceeds = variance != ZERO

        threshold_log.append({"department": dept, "account": acct, "basis": basis,
                              "dollar_threshold": t["dollar_threshold"] if t else None,
                              "pct_threshold": t["pct_threshold"] if t else None})

        expl_rows = expl_idx.get(k, [])
        comp_total = sum((dec(e["component_amount"]) or ZERO) for e in expl_rows
                         if e["component_amount"])
        has_narrative = any(e["narrative"] or e["reason_category"] for e in expl_rows)
        explained = bool(expl_rows) and has_narrative

        if exceeds and not explained:
            unexplained.append({"department": dept, "account": acct, "variance": variance,
                                "pct": pct, "basis": basis})
            errors.append(f"{dept}/{acct}: variance {variance:,.2f} exceeds its "
                          f"{basis} threshold with no explanation on file.")
        elif exceeds and expl_rows and any(e["component_amount"] for e in expl_rows):
            if comp_total != variance:
                errors.append(f"{dept}/{acct}: explanation components sum to "
                              f"{comp_total:,.2f} but the variance is {variance:,.2f}. "
                              f"Components don't explain the whole variance.")

        report.append({"department": dept, "account": acct, "budget": b, "actual": a,
                       "variance": variance, "pct": pct, "exceeds": exceeds,
                       "explained": explained, "basis": basis,
                       "narrative": "; ".join(e["narrative"] for e in expl_rows if e["narrative"])})

    total_actual_src = sum((dec(a["actual_amount"]) or ZERO) for a in actuals)
    total_budget_src = sum((dec(b["budget_amount"]) or ZERO) for b in budget)
    total_actual_rep = sum(r["actual"] for r in report)
    total_budget_rep = sum(r["budget"] for r in report)
    if total_actual_src != total_actual_rep:
        errors.append(f"Total actual on report ({total_actual_rep:,.2f}) does not equal "
                      f"the sum of the actuals file ({total_actual_src:,.2f}).")
    if total_budget_src != total_budget_rep:
        errors.append(f"Total budget on report ({total_budget_rep:,.2f}) does not equal "
                      f"the sum of the budget file ({total_budget_src:,.2f}).")

    return {"report": report, "unexplained": unexplained, "threshold_log": threshold_log,
            "errors": errors, "total_actual": total_actual_rep, "total_budget": total_budget_rep}


# -------------------------------------------------------------------- workbook

def hdr(ws, n, row=1):
    for c in range(1, n + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 26


def widths(ws, w):
    for c, v in w.items():
        ws.column_dimensions[get_column_letter(c)].width = v


def sheet_report(wb, meta, result, valid):
    ws = wb.active
    ws.title = "BvA Report"
    widths(ws, {1: 20, 2: 20, 3: 16, 4: 16, 5: 16, 6: 10, 7: 14, 8: 10, 9: 40})
    cols = ["Department", "Account", "Budget", "Actual", "Variance $", "Variance %",
            "Exceeds Threshold", "Explained", "Explanation"]
    for c, v in enumerate(cols, start=1):
        ws.cell(row=1, column=c, value=v)
    hdr(ws, len(cols))
    r = 2
    for row in result["report"]:
        vals = [row["department"], row["account"], float(row["budget"]), float(row["actual"]),
                float(row["variance"]), float(row["pct"]) if row["pct"] is not None else None,
                "YES" if row["exceeds"] else "no", "yes" if row["explained"] else "NO",
                row["narrative"]]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            if c in (3, 4, 5) and v is not None:
                cell.number_format = MONEY
            if c == 6 and v is not None:
                cell.number_format = PCT
            if c == 7 and v == "YES":
                cell.font = BAD_FONT
            if c == 8 and v == "NO":
                cell.font = BAD_FONT
                cell.fill = BAD_FILL
        r += 1
    r += 1
    ws.cell(row=r, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=r, column=3, value=float(result["total_budget"])).number_format = MONEY
    ws.cell(row=r, column=4, value=float(result["total_actual"])).number_format = MONEY
    r += 2
    v = ws.cell(row=r, column=1,
                value="REPORT COMPLETE - every above-threshold variance is explained" if valid
                else "REPORT NOT COMPLETE - see Unexplained Variance tab")
    v.font = OK_FONT if valid else BAD_FONT
    if not valid:
        v.fill = BAD_FILL


def sheet_unexplained(wb, result):
    ws = wb.create_sheet("Unexplained Variance")
    widths(ws, {1: 20, 2: 20, 3: 16, 4: 10, 5: 18})
    cols = ["Department", "Account", "Variance $", "Variance %", "Threshold Basis"]
    for c, v in enumerate(cols, start=1):
        ws.cell(row=1, column=c, value=v)
    hdr(ws, len(cols))
    r = 2
    for u in result["unexplained"]:
        ws.cell(row=r, column=1, value=u["department"])
        ws.cell(row=r, column=2, value=u["account"])
        c3 = ws.cell(row=r, column=3, value=float(u["variance"])); c3.number_format = MONEY
        if u["pct"] is not None:
            c4 = ws.cell(row=r, column=4, value=float(u["pct"])); c4.number_format = PCT
        ws.cell(row=r, column=5, value=u["basis"])
        for c in range(1, 6):
            ws.cell(row=r, column=c).fill = BAD_FILL
        r += 1
    if not result["unexplained"]:
        ws.cell(row=2, column=1, value="None - every above-threshold variance is explained.").font = OK_FONT


def sheet_thresholds(wb, result):
    ws = wb.create_sheet("Threshold Config")
    widths(ws, {1: 20, 2: 20, 3: 18, 4: 16, 5: 14})
    cols = ["Department", "Account", "Threshold Basis", "Dollar Threshold", "Pct Threshold"]
    for c, v in enumerate(cols, start=1):
        ws.cell(row=1, column=c, value=v)
    hdr(ws, len(cols))
    r = 2
    for t in result["threshold_log"]:
        ws.cell(row=r, column=1, value=t["department"])
        ws.cell(row=r, column=2, value=t["account"])
        ws.cell(row=r, column=3, value=t["basis"])
        if t["dollar_threshold"] is not None:
            ws.cell(row=r, column=4, value=float(t["dollar_threshold"])).number_format = MONEY
        if t["pct_threshold"] is not None:
            ws.cell(row=r, column=5, value=float(t["pct_threshold"])).number_format = PCT
        r += 1


def write_template(keys, out: Path):
    with out.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(EXPL_COLS)
        for dept, acct in keys:
            w.writerow([dept, acct, "", "", "", ""])
    print(f"Blank explanations template written to {out} ({len(keys)} rows)")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                  formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--actuals", required=True)
    ap.add_argument("--budget", required=True)
    ap.add_argument("--thresholds", required=True)
    ap.add_argument("--explanations")
    ap.add_argument("--period", default="")
    ap.add_argument("--client", default="")
    ap.add_argument("--out", required=True)
    ap.add_argument("--template", action="store_true",
                    help="emit a blank explanations worksheet instead of the report")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    actuals = load(Path(args.actuals), ["department", "account", "actual_amount"])
    budget = load(Path(args.budget), ["department", "account", "budget_amount"])
    thresholds = load(Path(args.thresholds),
                      ["department", "account", "dollar_threshold", "pct_threshold"])

    if args.template:
        keys = sorted(set((a["department"], a["account"]) for a in actuals) |
                      set((b["department"], b["account"]) for b in budget))
        write_template(keys, Path(args.out))
        return 0

    explanations = load(Path(args.explanations), EXPL_COLS) if args.explanations else []
    result = build(actuals, budget, thresholds, explanations)
    errors = result["errors"]
    valid = not errors

    print("=" * 76)
    print("BUDGET VS ACTUAL")
    print("=" * 76)
    print(f"Client: {args.client or '(not stated)'}    Period: {args.period or '(not stated)'}")
    print(f"Total budget: {result['total_budget']:,.2f}   Total actual: {result['total_actual']:,.2f}")
    print(f"Lines: {len(result['report'])}   Exceeding threshold: "
          f"{sum(1 for r in result['report'] if r['exceeds'])}   "
          f"Unexplained: {len(result['unexplained'])}")
    if errors:
        print(f"\nGATE FAILURES ({len(errors)}):")
        for e in errors[:25]:
            print(f"  ! {e}")

    if not valid and not args.force:
        print("\n" + "=" * 76)
        print("WORKBOOK NOT WRITTEN.")
        print("Every variance above its threshold needs an explanation, and component")
        print("amounts must sum to the variance they claim to explain. Fix the")
        print("Unexplained Variance / component issues above and re-run.")
        print("=" * 76)
        return 1

    wb = Workbook()
    sheet_report(wb, {"client": args.client, "period": args.period}, result, valid)
    sheet_unexplained(wb, result)
    sheet_thresholds(wb, result)
    wb.active = 0

    out = Path(args.out)
    if not valid:
        out = out.with_name(out.stem + " [GATE NOT MET]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\n{'GATE MET' if valid else 'NOT MET (forced)'}.  Workbook: {out}")
    return 0 if valid else 1


if __name__ == "__main__":
    sys.exit(main())
```

