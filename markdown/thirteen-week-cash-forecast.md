---
name: thirteen-week-cash-forecast
description: Builds a 13-week direct cash forecast from AR aging, AP aging, the payroll calendar, and recurring items, with a weekly actual-vs-forecast reforecast mode. Every input row lands in a specific week or an explicit Beyond Horizon bucket, and beginning balance plus net flows must tie exactly to the week-13 ending balance before the workbook is produced. Use this whenever the user mentions a 13-week cash forecast, direct cash forecast, weekly cash forecast, cash runway, cash flow projection, liquidity forecast, or reforecasting cash against actuals. Trigger on "build a 13-week forecast," "how much cash will we have in 8 weeks," "forecast our cash position," "reforecast this week's cash," "what's our cash runway," or when AR aging, AP aging, a payroll calendar, and recurring items are supplied together. The canonical CFO forecasting artifact, recurring weekly, not a one-shot. Pairs with `cash-flow-tieout`, which reconciles a historical statement rather than projecting forward.
---

# 13-Week Cash Forecast

A direct cash forecast is only as credible as its placement logic. Every dollar in AR, AP,
payroll, and recurring items has to land in a specific week — or in an explicit catch-all for
anything beyond the 13-week window — and the script proves that placement rather than asserting
it.

## The gate

The script refuses to produce a clean workbook when:

1. **A row cannot be placed.** Every AR, AP, payroll, and recurring-generated occurrence must land
   in one of the 13 weeks or the Beyond Horizon bucket. A receivable whose aging bucket has no
   matching row in the collection-timing assumptions is an exception, not a guess.
2. **The arithmetic identity breaks.** Beginning balance plus the sum of all 13 weeks' net flows
   must equal the week-13 ending balance exactly. If it doesn't, the week-by-week rollforward has
   a bug and nobody should be looking at the output yet.
3. **A row is missing a required field.** Amount, date, or direction fields that can't be parsed
   are reported by row number, never defaulted.
4. **An actual is compared that was never supplied.** In reforecast mode, only weeks with a
   genuine actual in `--actuals` are compared — a missing week is left out of the Reforecast
   Variance tab, never shown as a zero variance.

## Inputs

1. **AR aging** — `customer, amount, due_date`.
2. **AP aging** — `vendor, amount, due_date`.
3. **Payroll calendar** — `pay_date, amount`.
4. **Recurring items** — `description, amount, frequency, next_date, direction`. `frequency` is
   `weekly`, `biweekly`, or `monthly`; `direction` is `inflow` or `outflow`.
5. **Collection-timing assumptions** — `aging_bucket, week_offset, pct_collected_in_week`. This is
   the client's own judgment about when each aging bucket actually converts to cash, not something
   the script invents. Every bucket that appears in the AR file needs a row here.
6. **`--beginning-balance`** and **`--forecast-start`** — required.
7. **`--buckets`** (optional, default `30,60,90`) — the aging bucket boundaries in days past due,
   same convention as `ar-aging-tie-out`. Change it if the client ages receivables differently.
8. **`--actuals`** (optional) — `week_ending, actual_net_flow`, for the weekly reforecast cadence
   once a week has closed.

## Step 1 — Get blank templates for anything you don't have yet

```bash
python3 scripts/cash_forecast_13wk.py --template ar --out ar_template.csv
python3 scripts/cash_forecast_13wk.py --template assumptions --out assumptions_template.csv
```

`--template` accepts `ar`, `ap`, `payroll`, `recurring`, `assumptions`, or `actuals`.

## Step 2 — Build the forecast

```bash
python3 scripts/cash_forecast_13wk.py \
  --ar ar_aging.csv --ap ap_aging.csv --payroll payroll_calendar.csv \
  --recurring recurring_items.csv --assumptions collection_timing.csv \
  --beginning-balance 850000.00 --forecast-start 2026-01-05 \
  --client "Northwind Systems Inc" \
  --out "Northwind - 13-Week Cash Forecast.xlsx"
```

AR is spread across weeks per its aging bucket's assumption rows (a bucket can split across
multiple weeks — the percentages don't have to sum to 100, and whatever's left over at week 13 is
swept into Beyond Horizon with the shortfall shown, never silently absorbed). AP, payroll, and
recurring items land in the week of their due date, pay date, or projected occurrence.

## Step 3 — The weekly reforecast

Once a week closes, run again with `--actuals` to see where the forecast missed:

```bash
python3 scripts/cash_forecast_13wk.py \
  --ar ar_aging.csv --ap ap_aging.csv --payroll payroll_calendar.csv \
  --recurring recurring_items.csv --assumptions collection_timing.csv \
  --actuals week1_actuals.csv \
  --beginning-balance 850000.00 --forecast-start 2026-01-05 \
  --client "Northwind Systems Inc" \
  --out "Northwind - 13-Week Reforecast W1.xlsx"
```

This is a recurring-use skill by design — re-run it every week against the same forecast horizon,
and the Reforecast Variance tab shows how assumption drift is accumulating, which is the signal to
revisit the collection-timing table itself rather than patch the forecast by hand.

## Deliver

**Workbook tabs:**

1. **13-Week Forecast** — weekly beginning balance, AR collections, AP disbursements, payroll,
   net recurring, and ending balance, plus the gate status.
2. **Assumptions** — every collection-timing row actually used, so a reviewer can see the basis
   for the AR placement without opening the script.
3. **Beyond Horizon** — AR, AP, payroll, and recurring amounts that fall outside the 13-week
   window, by category.
4. **Source Detail** — every input row, its amount and date, and exactly which week (or Beyond
   Horizon) it was placed in — the row-count reconciliation that proves nothing was dropped.
5. **Reforecast Variance** (only when `--actuals` supplied) — forecast vs. actual net flow per
   week with an actual on record.

**Then, in chat:** the beginning and week-13 ending balance first, then any gate failures, then
the Beyond Horizon totals so the reader knows what the 13-week window doesn't cover.

## What to escalate

- **A collection-timing assumption that doesn't match recent actuals** once the reforecast mode
  has a few weeks of history — that's a conversation about the assumption, not a script fix.
- **A week where AP or payroll alone exceeds the beginning balance** — a liquidity event, not a
  forecasting detail.
- **Recurring items whose `next_date` is stale** (far in the past) — likely means the item already
  occurred and the source file wasn't updated.
- **Large, unscheduled AR remainders landing in Beyond Horizon** — usually means the
  collection-timing table doesn't cover that bucket's full percentage and needs a review.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry. AR,
AP, payroll, and vendor data never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*

---

# Appendix — bundled files

If you installed the `.skill` package these files are already in place and you can ignore this appendix. If you copied the skill as text, create the files below at the paths shown, alongside your `SKILL.md`. The skill will not run without them.

## `scripts/cash_forecast_13wk.py`

```python
#!/usr/bin/env python3
"""
Build a 13-week direct cash forecast from AR aging, AP aging, the payroll
calendar, and recurring items - and prove every input row landed somewhere.

Gates:
  * every AR, AP, payroll, and recurring-generated row must land in exactly one
    of the 13 weeks or the Beyond Horizon bucket - proven by a row-count
    reconciliation on the Source Detail tab. A row that cannot be dated is an
    exception, never silently dropped.
  * --beginning-balance + the sum of all 13 weeks' net flows must equal the
    week-13 ending balance exactly (closed arithmetic identity).
  * every collection-timing assumption actually used must be reproduced on the
    Assumptions tab - never buried in code.
  * in --actuals reforecast mode, only weeks with a supplied actual are
    compared; a week with no actual is never fabricated as zero variance.

No aging-bucket boundary, collection percentage, or threshold is hardcoded.
All of it is supplied by the user, sourced from the client's own assumptions.

Usage - emit a blank input template:
    python3 cash_forecast_13wk.py --template ar --out ar_template.csv
    (choices for --template: ar, ap, payroll, recurring, assumptions, actuals)

Usage - build the forecast:
    python3 cash_forecast_13wk.py \
        --ar ar_aging.csv --ap ap_aging.csv --payroll payroll_calendar.csv \
        --recurring recurring_items.csv --assumptions collection_timing.csv \
        --beginning-balance 850000.00 --forecast-start 2026-01-05 \
        --client "Northwind Systems Inc" \
        --out "Northwind - 13-Week Cash Forecast.xlsx"

Usage - weekly reforecast against actuals:
    python3 cash_forecast_13wk.py ... --actuals week1_actuals.csv \
        --out "Northwind - 13-Week Reforecast W1.xlsx"

--ar CSV:            customer, amount, due_date
--ap CSV:             vendor, amount, due_date
--payroll CSV:        pay_date, amount
--recurring CSV:      description, amount, frequency, next_date, direction
                      frequency: weekly | biweekly | monthly
                      direction: inflow | outflow
--assumptions CSV:    aging_bucket, week_offset, pct_collected_in_week
                      aging_bucket must match a bucket produced by --buckets
                      (default 30,60,90 days past due, same convention as
                      ar-aging-tie-out). week_offset is weeks from
                      --forecast-start (0 = week 1). A bucket's percentages
                      need not sum to 100 - whatever is left over at week 13
                      is swept into Beyond Horizon with the shortfall noted.
--actuals CSV:        week_ending, actual_net_flow
"""

from __future__ import annotations

import argparse
import csv
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta
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
DATEF = "yyyy-mm-dd"

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")

TEMPLATES = {
    "ar": ["customer", "amount", "due_date"],
    "ap": ["vendor", "amount", "due_date"],
    "payroll": ["pay_date", "amount"],
    "recurring": ["description", "amount", "frequency", "next_date", "direction"],
    "assumptions": ["aging_bucket", "week_offset", "pct_collected_in_week"],
    "actuals": ["week_ending", "actual_net_flow"],
}


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


def load_rows(path: Path | None, cols: list[str]) -> list[dict]:
    if not path:
        return []
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    out = []
    for i, r in enumerate(rows, start=1):
        x = {k: clean(r.get(k)) for k in cols}
        x["row"] = i
        out.append(x)
    return out


def bucket_bounds_labels(bounds: list[int]):
    labels = []
    prev = 0
    for b in bounds:
        labels.append(f"{prev + 1}-{b}" if prev else f"1-{b}")
        prev = b
    labels = ["current"] + labels + [f"{bounds[-1]}+"]
    return labels


def bucket_of(days_past_due: int, bounds: list[int], labels: list[str]) -> str:
    if days_past_due <= 0:
        return labels[0]
    for i, b in enumerate(bounds):
        if days_past_due <= b:
            return labels[i + 1]
    return labels[-1]


def week_index(d: date, start: date) -> int | None:
    """Return 0-based week index (0..12) or None if before start / beyond 13 weeks."""
    if d < start:
        return 0  # already due/overdue at forecast start lands in week 1
    delta_days = (d - start).days
    idx = delta_days // 7
    return idx if idx < 13 else None


def project_recurring(rec: dict, start: date, weeks_end: date) -> list[date]:
    freq = rec["frequency"].lower()
    nd = pdate(rec["next_date"])
    if nd is None:
        return []
    step = {"weekly": timedelta(days=7), "biweekly": timedelta(days=14)}.get(freq)
    occurrences = []
    if freq == "monthly":
        cur = nd
        guard = 0
        while cur <= weeks_end and guard < 36:
            occurrences.append(cur)
            # advance one calendar month, clamped to month length
            month = cur.month + 1
            year = cur.year + (month - 1) // 12
            month = (month - 1) % 12 + 1
            day = min(cur.day, 28)
            cur = date(year, month, day)
            guard += 1
    elif step:
        cur = nd
        guard = 0
        while cur <= weeks_end and guard < 60:
            occurrences.append(cur)
            cur = cur + step
            guard += 1
    else:
        raise ValueError(f"Unrecognized frequency {rec['frequency']!r} "
                          f"(row {rec['row']}): must be weekly, biweekly, or monthly.")
    return occurrences


def build(ar, ap, payroll, recurring, assumptions, actuals, beginning_balance,
          start: date, bounds: list[int]):
    errors = []
    labels = bucket_bounds_labels(bounds)
    weeks_end = start + timedelta(weeks=13) - timedelta(days=1)

    assump_idx: dict[str, list[tuple[int, Decimal]]] = defaultdict(list)
    for a in assumptions:
        wk = int(a["week_offset"]) if a["week_offset"] else None
        pct = dec(a["pct_collected_in_week"])
        if wk is None or pct is None:
            errors.append(f"Assumptions row {a['row']}: week_offset and "
                          f"pct_collected_in_week are both required.")
            continue
        assump_idx[a["aging_bucket"]].append((wk, pct))

    weeks = [{"beginning": None, "ar": ZERO, "ap": ZERO, "payroll": ZERO,
             "recurring": ZERO, "ending": None} for _ in range(13)]
    beyond = {"ar": ZERO, "ap": ZERO, "payroll": ZERO, "recurring": ZERO}
    source_detail = []

    # AR: spread each receivable per its aging bucket's assumption rows
    for x in ar:
        amt = dec(x["amount"])
        due = pdate(x["due_date"])
        if amt is None or due is None:
            errors.append(f"AR row {x['row']} ({x['customer']}): amount and "
                          f"due_date are both required.")
            continue
        days_past_due = (start - due).days
        b = bucket_of(days_past_due, bounds, labels)
        sched = assump_idx.get(b, [])
        if not sched:
            errors.append(f"AR row {x['row']} ({x['customer']}): aging bucket "
                          f"'{b}' has no collection-timing assumption supplied. "
                          f"Every bucket present in the AR file must have a row "
                          f"in --assumptions, or this receivable cannot be placed.")
            continue
        placed = ZERO
        for wk, pct in sched:
            share = (amt * pct / 100).quantize(Decimal("0.01"))
            placed += share
            if wk < 13:
                weeks[wk]["ar"] += share
                source_detail.append(("AR", x["row"], x["customer"], amt, due,
                                       f"week {wk + 1}", share))
            else:
                beyond["ar"] += share
                source_detail.append(("AR", x["row"], x["customer"], amt, due,
                                       "beyond horizon", share))
        leftover = amt - placed
        if leftover != ZERO:
            beyond["ar"] += leftover
            source_detail.append(("AR", x["row"], x["customer"], amt, due,
                                   "beyond horizon (unscheduled remainder)", leftover))

    # AP: lands in the week of its due_date
    for x in ap:
        amt = dec(x["amount"])
        due = pdate(x["due_date"])
        if amt is None or due is None:
            errors.append(f"AP row {x['row']} ({x['vendor']}): amount and "
                          f"due_date are both required.")
            continue
        wk = week_index(due, start)
        if wk is None:
            beyond["ap"] += amt
            source_detail.append(("AP", x["row"], x["vendor"], amt, due,
                                   "beyond horizon", amt))
        else:
            weeks[wk]["ap"] += amt
            source_detail.append(("AP", x["row"], x["vendor"], amt, due,
                                   f"week {wk + 1}", amt))

    # Payroll: lands in the week of pay_date
    for x in payroll:
        amt = dec(x["amount"])
        pd_ = pdate(x["pay_date"])
        if amt is None or pd_ is None:
            errors.append(f"Payroll row {x['row']}: amount and pay_date are "
                          f"both required.")
            continue
        wk = week_index(pd_, start)
        if wk is None:
            beyond["payroll"] += amt
            source_detail.append(("Payroll", x["row"], "", amt, pd_,
                                   "beyond horizon", amt))
        else:
            weeks[wk]["payroll"] += amt
            source_detail.append(("Payroll", x["row"], "", amt, pd_,
                                   f"week {wk + 1}", amt))

    # Recurring: project occurrences across the 13-week window and beyond
    for x in recurring:
        amt = dec(x["amount"])
        direction = x["direction"].lower()
        if amt is None or direction not in ("inflow", "outflow"):
            errors.append(f"Recurring row {x['row']} ({x['description']}): amount "
                          f"and direction (inflow/outflow) are both required.")
            continue
        signed = amt if direction == "inflow" else -amt
        try:
            occs = project_recurring(x, start, weeks_end + timedelta(days=370))
        except ValueError as e:
            errors.append(str(e))
            continue
        for occ in occs:
            wk = week_index(occ, start)
            if wk is None:
                beyond["recurring"] += signed
                source_detail.append(("Recurring", x["row"], x["description"],
                                       signed, occ, "beyond horizon", signed))
            else:
                weeks[wk]["recurring"] += signed
                source_detail.append(("Recurring", x["row"], x["description"],
                                       signed, occ, f"week {wk + 1}", signed))

    # Roll the balance forward; AP is an outflow (subtract), recurring already signed
    bal = beginning_balance
    for i, w in enumerate(weeks):
        w["beginning"] = bal
        net = w["ar"] - w["ap"] - w["payroll"] + w["recurring"]
        w["net"] = net
        bal = bal + net
        w["ending"] = bal

    identity_rhs = beginning_balance + sum(w["net"] for w in weeks)
    if identity_rhs != weeks[-1]["ending"]:
        errors.append(f"Arithmetic identity failed: beginning balance + sum of "
                      f"net flows ({identity_rhs:,.2f}) does not equal week-13 "
                      f"ending balance ({weeks[-1]['ending']:,.2f}).")

    reforecast = None
    if actuals:
        reforecast = []
        for a in actuals:
            we = pdate(a["week_ending"])
            act = dec(a["actual_net_flow"])
            if we is None or act is None:
                errors.append(f"Actuals row {a['row']}: week_ending and "
                              f"actual_net_flow are both required.")
                continue
            wk = week_index(we - timedelta(days=6), start)
            if wk is None:
                errors.append(f"Actuals row {a['row']}: week_ending {we} does "
                              f"not fall within the 13-week forecast window.")
                continue
            forecast_net = weeks[wk]["net"]
            reforecast.append({"week": wk + 1, "week_ending": we,
                               "forecast": forecast_net, "actual": act,
                               "variance": act - forecast_net})

    return {"weeks": weeks, "beyond": beyond, "source_detail": source_detail,
            "assump_idx": assump_idx, "errors": errors, "reforecast": reforecast,
            "labels": labels}


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


def sheet_forecast(wb, meta, result, start, valid):
    ws = wb.active
    ws.title = "13-Week Forecast"
    widths(ws, {1: 10, 2: 18, 3: 16, 4: 16, 5: 16, 6: 16, 7: 16, 8: 18})
    cols = ["Week", "Week Ending", "Beginning Bal.", "AR Collections",
            "AP Disbursements", "Payroll", "Recurring (net)", "Ending Bal."]
    for c, v in enumerate(cols, start=1):
        ws.cell(row=1, column=c, value=v)
    hdr(ws, len(cols))
    r = 2
    for i, w in enumerate(result["weeks"]):
        we = start + timedelta(days=7 * i + 6)
        vals = [i + 1, we, w["beginning"], w["ar"], -w["ap"], -w["payroll"],
                w["recurring"], w["ending"]]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=float(v) if isinstance(v, Decimal) else v)
            if isinstance(v, Decimal):
                cell.number_format = MONEY
            if c == 2:
                cell.number_format = DATEF
        r += 1
    r += 1
    ws.cell(row=r, column=2, value="Beginning balance").font = Font(bold=True)
    ws.cell(row=r, column=3, value=float(meta["beginning_balance"])).number_format = MONEY
    r += 1
    v = ws.cell(row=r, column=2,
                value="GATE MET - identity ties exactly" if valid else
                "GATE NOT MET - see Source Detail / errors in chat")
    v.font = OK_FONT if valid else BAD_FONT
    if not valid:
        v.fill = BAD_FILL


def sheet_assumptions(wb, result):
    ws = wb.create_sheet("Assumptions")
    widths(ws, {1: 18, 2: 14, 3: 20})
    cols = ["Aging Bucket", "Week Offset", "Pct Collected In Week"]
    for c, v in enumerate(cols, start=1):
        ws.cell(row=1, column=c, value=v)
    hdr(ws, len(cols))
    r = 2
    for bucket, sched in result["assump_idx"].items():
        for wk, pct in sched:
            ws.cell(row=r, column=1, value=bucket)
            ws.cell(row=r, column=2, value=wk)
            ws.cell(row=r, column=3, value=float(pct))
            r += 1


def sheet_beyond(wb, result):
    ws = wb.create_sheet("Beyond Horizon")
    widths(ws, {1: 20, 2: 16})
    for c, v in enumerate(["Category", "Amount"], start=1):
        ws.cell(row=1, column=c, value=v)
    hdr(ws, 2)
    r = 2
    for k, v in result["beyond"].items():
        ws.cell(row=r, column=1, value=k.upper())
        cell = ws.cell(row=r, column=2, value=float(v))
        cell.number_format = MONEY
        r += 1


def sheet_source(wb, result):
    ws = wb.create_sheet("Source Detail")
    widths(ws, {1: 12, 2: 8, 3: 24, 4: 16, 5: 14, 6: 26, 7: 16})
    cols = ["Source", "Row", "Description", "Amount", "Date", "Placed In", "Placed Amount"]
    for c, v in enumerate(cols, start=1):
        ws.cell(row=1, column=c, value=v)
    hdr(ws, len(cols))
    r = 2
    counts = defaultdict(int)
    for src, row, desc, amt, dt, placed, placed_amt in result["source_detail"]:
        counts[src] += 1
        ws.cell(row=r, column=1, value=src)
        ws.cell(row=r, column=2, value=row)
        ws.cell(row=r, column=3, value=desc)
        c4 = ws.cell(row=r, column=4, value=float(amt)); c4.number_format = MONEY
        ws.cell(row=r, column=5, value=dt).number_format = DATEF
        ws.cell(row=r, column=6, value=placed)
        c7 = ws.cell(row=r, column=7, value=float(placed_amt)); c7.number_format = MONEY
        r += 1
    r += 1
    ws.cell(row=r, column=1, value="Row-count reconciliation:").font = Font(bold=True)
    r += 1
    for src, n in counts.items():
        ws.cell(row=r, column=1, value=f"  {src}: {n} placements recorded")
        r += 1


def sheet_reforecast(wb, result):
    if result["reforecast"] is None:
        return
    ws = wb.create_sheet("Reforecast Variance")
    widths(ws, {1: 8, 2: 16, 3: 16, 4: 16, 5: 16})
    cols = ["Week", "Week Ending", "Forecast Net", "Actual Net", "Variance"]
    for c, v in enumerate(cols, start=1):
        ws.cell(row=1, column=c, value=v)
    hdr(ws, len(cols))
    r = 2
    for rf in result["reforecast"]:
        ws.cell(row=r, column=1, value=rf["week"])
        ws.cell(row=r, column=2, value=rf["week_ending"]).number_format = DATEF
        for c, k in ((3, "forecast"), (4, "actual"), (5, "variance")):
            cell = ws.cell(row=r, column=c, value=float(rf[k]))
            cell.number_format = MONEY
            if k == "variance" and rf[k] != ZERO:
                cell.font = BAD_FONT
        r += 1


def write_template(kind: str, out: Path):
    cols = TEMPLATES[kind]
    with out.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(cols)
    print(f"Blank '{kind}' template written to {out}")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                  formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--ar"); ap.add_argument("--ap"); ap.add_argument("--payroll")
    ap.add_argument("--recurring"); ap.add_argument("--assumptions")
    ap.add_argument("--actuals")
    ap.add_argument("--beginning-balance")
    ap.add_argument("--forecast-start")
    ap.add_argument("--buckets", default="30,60,90",
                    help="aging bucket boundaries in days past due, comma-separated")
    ap.add_argument("--client", default="")
    ap.add_argument("--out", required=True)
    ap.add_argument("--template", choices=sorted(TEMPLATES))
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    out = Path(args.out)

    if args.template:
        write_template(args.template, out)
        return 0

    missing_required = [n for n, v in (("--ar", args.ar), ("--ap", args.ap),
                        ("--payroll", args.payroll), ("--recurring", args.recurring),
                        ("--assumptions", args.assumptions),
                        ("--beginning-balance", args.beginning_balance),
                        ("--forecast-start", args.forecast_start)) if not v]
    if missing_required:
        sys.exit(f"Missing required arguments: {', '.join(missing_required)}")

    start = pdate(args.forecast_start)
    beginning_balance = dec(args.beginning_balance)
    bounds = [int(b) for b in args.buckets.split(",")]

    ar = load_rows(Path(args.ar), TEMPLATES["ar"])
    ap_rows = load_rows(Path(args.ap), TEMPLATES["ap"])
    payroll = load_rows(Path(args.payroll), TEMPLATES["payroll"])
    recurring = load_rows(Path(args.recurring), TEMPLATES["recurring"])
    assumptions = load_rows(Path(args.assumptions), TEMPLATES["assumptions"])
    actuals = load_rows(Path(args.actuals), TEMPLATES["actuals"]) if args.actuals else None

    result = build(ar, ap_rows, payroll, recurring, assumptions, actuals,
                    beginning_balance, start, bounds)
    errors = result["errors"]
    valid = not errors

    meta = {"client": args.client or "(not stated)", "beginning_balance": beginning_balance,
            "forecast_start": start, "prepared": datetime.now().strftime("%Y-%m-%d %H:%M")}

    print("=" * 76)
    print("13-WEEK CASH FORECAST")
    print("=" * 76)
    print(f"Client: {meta['client']}    Forecast start: {start}")
    print(f"Beginning balance: {beginning_balance:,.2f}")
    print(f"Week-13 ending balance: {result['weeks'][-1]['ending']:,.2f}")
    print(f"Beyond horizon: AR {result['beyond']['ar']:,.2f}  "
          f"AP {result['beyond']['ap']:,.2f}  "
          f"Payroll {result['beyond']['payroll']:,.2f}  "
          f"Recurring {result['beyond']['recurring']:,.2f}")
    if errors:
        print(f"\nGATE FAILURES ({len(errors)}):")
        for e in errors[:25]:
            print(f"  ! {e}")

    if not valid and not args.force:
        print("\n" + "=" * 76)
        print("WORKBOOK NOT WRITTEN.")
        print("Every input row must land in a week or Beyond Horizon, and the")
        print("beginning balance plus net flows must tie exactly to the week-13")
        print("ending balance. Fix the exceptions above and re-run.")
        print("=" * 76)
        return 1

    wb = Workbook()
    sheet_forecast(wb, meta, result, start, valid)
    sheet_assumptions(wb, result)
    sheet_beyond(wb, result)
    sheet_source(wb, result)
    sheet_reforecast(wb, result)
    wb.active = 0

    if not valid:
        out = out.with_name(out.stem + " [GATE NOT MET]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\n{'GATE MET' if valid else 'NOT MET (forced)'}.  Workbook: {out}")
    return 0 if valid else 1


if __name__ == "__main__":
    sys.exit(main())
```

