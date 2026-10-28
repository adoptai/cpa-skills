#!/usr/bin/env python3
"""
Recompute sales commissions from bookings and the comp plan's own tiered
rate schedule, and tie the result to what payroll actually paid.

Conventions stated here, not guessed anywhere in the code:
  * `rate_pct` is a percentage (5 means 5%), applied as rate_pct/100 times
    the amount apportioned to that tier.
  * A blank `tier_max_amount` means "and above" - unbounded, not a guessed
    ceiling.
  * Each deal is mapped to a payroll period as the calendar month
    (YYYY-MM) of its close_date. The payroll register's own `period` column
    must use the same YYYY-MM format to tie out; a register using a
    different period convention will show as non-matching rather than
    being forced to align.
  * Commission tiers reset to zero at the start of each comp plan period
    (`plan_effective_date`) and accumulate chronologically by close_date
    within that period.

Gates:
  * HARD (blocks the workbook unless --force): no usable comp plan tiers
    supplied at all - there is nothing to recompute a rate from.
  * SOFT (routed to Exceptions, workbook still produced):
      - a deal whose close_date matches zero or more than one comp plan
        period for that rep - never defaulted to a flat or "standard" rate;
      - a plan period where any tier has a blank rate_pct - the ENTIRE plan
        period is treated as incomplete and every deal that falls in it
        becomes an exception, because cumulative-based tier placement for
        later deals cannot be trusted once a rate is undefined mid-schedule;
      - a deal whose amount span is not fully covered by contiguous tiers
        (a gap between tiers, or an amount this rep's highest tier doesn't
        reach) - the deal is never partially rated; it is excluded from the
        recomputed total entirely and reported with the uncovered range.
  * Recomputed vs. paid is never netted across reps or periods - every
    discrepancy is reported individually, by rep and period, with its own
    dollar effect.

Usage - emit a blank bookings worksheet:
    python3 commission_recalc.py --template --out bookings_template.csv

Usage - recompute:
    python3 commission_recalc.py --bookings bookings.csv --plan comp_plan.csv \
        --payroll payroll.csv --client "Northwind Systems Inc" \
        --out "Northwind - Commission Recalc.xlsx"

--bookings CSV:
    rep, deal_id, amount, close_date

--plan CSV (one row per tier; rows sharing rep + plan_effective_date +
  plan_expiry_date form one plan period):
    rep, plan_effective_date, plan_expiry_date, tier_min_amount,
    tier_max_amount, rate_pct

--payroll CSV:
    rep, period, commission_paid
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
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is required.  pip install openpyxl")

ZERO = Decimal("0.00")
INF = Decimal("Infinity")
MONEY = '#,##0.00;[Red](#,##0.00)'

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")

BOOKING_COLS = ["rep", "deal_id", "amount", "close_date"]
PLAN_COLS = ["rep", "plan_effective_date", "plan_expiry_date",
             "tier_min_amount", "tier_max_amount", "rate_pct"]
PAYROLL_COLS = ["rep", "period", "commission_paid"]


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


def load_csv(path: Path, cols: list[str]) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    out = []
    for i, r in enumerate(rows, start=1):
        x = {k: clean(r.get(k)) for k in cols}
        x["row"] = i
        out.append(x)
    return out


def write_template(out: Path):
    with out.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(BOOKING_COLS)
    print(f"Template written: {out}")


def style_header(ws, row, cols):
    for i, c in enumerate(cols, start=1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(c) + 2)


def fmt_amt(v):
    return "and above" if v == INF else f"{v:,.2f}"


# ------------------------------------------------------------------ plan load

def load_plans(path: Path):
    rows = load_csv(path, PLAN_COLS)
    plans = defaultdict(list)   # (rep, eff, exp) -> [tier, ...]
    load_errors = []
    for r in rows:
        try:
            eff = pdate(r["plan_effective_date"])
            exp = pdate(r["plan_expiry_date"])
            tmin = dec(r["tier_min_amount"])
            tmax_raw = r["tier_max_amount"]
            tmax = INF if tmax_raw == "" else dec(tmax_raw)
            rate = dec(r["rate_pct"]) if r["rate_pct"] != "" else None
        except ValueError as e:
            load_errors.append(f"Plan row {r['row']}: {e}")
            continue
        if eff is None or exp is None or tmin is None:
            load_errors.append(f"Plan row {r['row']}: missing required date/amount field")
            continue
        key = (r["rep"], eff, exp)
        plans[key].append({"min": tmin, "max": tmax, "rate": rate, "row": r["row"]})
    for key in plans:
        plans[key].sort(key=lambda t: t["min"])
    return plans, load_errors


def plan_invalid_reason(tiers):
    blanks = [t for t in tiers if t["rate"] is None]
    if blanks:
        rows = ", ".join(str(t["row"]) for t in blanks)
        return (f"tier(s) on row(s) {rows} have no rate_pct - commission not computed "
                f"for this plan period")
    return None


def split_across_tiers(tiers, span_start, span_end):
    """Return (breakdown, uncovered_ranges). breakdown: [(tier, portion), ...]."""
    breakdown = []
    uncovered = []
    pointer = span_start
    for t in tiers:
        if t["max"] <= pointer or t["min"] >= span_end:
            continue
        lo = max(pointer, t["min"])
        hi = min(span_end, t["max"])
        if lo >= hi:
            continue
        if lo > pointer:
            uncovered.append((pointer, lo))
        breakdown.append((t, hi - lo))
        pointer = hi
    if pointer < span_end:
        uncovered.append((pointer, span_end))
    return breakdown, uncovered


def find_plan_keys(plans, rep, close_date):
    hits = [key for key in plans if key[0] == rep and key[1] <= close_date <= key[2]]
    return hits


# ---------------------------------------------------------------------- main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                  formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--bookings")
    ap.add_argument("--plan")
    ap.add_argument("--payroll")
    ap.add_argument("--client", default="")
    ap.add_argument("--out", required=True)
    ap.add_argument("--template", action="store_true",
                     help="emit a blank bookings worksheet instead of recomputing")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    if args.template:
        write_template(Path(args.out))
        return 0

    if not args.bookings or not args.plan or not args.payroll:
        sys.exit("--bookings, --plan and --payroll are required (or pass --template)")

    hard_errors = []
    plans, plan_load_errors = load_plans(Path(args.plan))
    hard_errors.extend(plan_load_errors)
    if not plans:
        hard_errors.append("No usable comp plan tiers supplied - there is nothing to "
                            "recompute a rate from.")

    invalid_plans = {key: plan_invalid_reason(tiers) for key, tiers in plans.items()}
    invalid_plans = {k: v for k, v in invalid_plans.items() if v}

    bookings = load_csv(Path(args.bookings), BOOKING_COLS)
    if not bookings:
        sys.exit(f"No rows in bookings {args.bookings}")

    by_rep = defaultdict(list)
    for r in bookings:
        try:
            amount = dec(r["amount"])
            close = pdate(r["close_date"])
        except ValueError as e:
            by_rep[r["rep"]].append({**r, "amount": None, "close_date": None, "error": str(e)})
            continue
        by_rep[r["rep"]].append({**r, "amount": amount, "close_date": close, "error": None})

    schedule_rows = []
    exceptions = []
    recomputed_by_rep_period = defaultdict(lambda: ZERO)

    for rep, deals in by_rep.items():
        deals = [d for d in deals if d["error"] is None]
        for d in deals:
            if d["error"]:
                continue
        bad = [d for d in by_rep[rep] if d["error"]]
        for d in bad:
            exceptions.append({"rep": rep, "deal_id": d["deal_id"],
                                "reason": f"Unparseable row: {d['error']}"})

        deals.sort(key=lambda d: (d["close_date"], d["deal_id"]))
        cumulative = defaultdict(lambda: ZERO)   # plan_key -> running total

        for d in deals:
            keys = find_plan_keys(plans, rep, d["close_date"])
            if len(keys) == 0:
                exceptions.append({"rep": rep, "deal_id": d["deal_id"],
                                    "reason": f"No comp plan found covering close_date "
                                              f"{d['close_date']}"})
                continue
            if len(keys) > 1:
                desc = "; ".join(f"{k[1]} to {k[2]}" for k in keys)
                exceptions.append({"rep": rep, "deal_id": d["deal_id"],
                                    "reason": f"Overlapping plans found for close_date "
                                              f"{d['close_date']}: {desc}"})
                continue
            key = keys[0]
            if key in invalid_plans:
                exceptions.append({"rep": rep, "deal_id": d["deal_id"],
                                    "reason": f"Plan period {key[1]} to {key[2]}: "
                                              f"{invalid_plans[key]}"})
                continue

            tiers = plans[key]
            span_start = cumulative[key]
            span_end = span_start + d["amount"]
            breakdown, uncovered = split_across_tiers(tiers, span_start, span_end)
            cumulative[key] = span_end

            if uncovered:
                gaps = "; ".join(f"{lo:,.2f}-{('and above' if hi == INF else f'{hi:,.2f}')}"
                                  for lo, hi in uncovered)
                exceptions.append({"rep": rep, "deal_id": d["deal_id"],
                                    "reason": f"Amount span {span_start:,.2f}-{span_end:,.2f} "
                                              f"not fully covered by supplied tiers for plan "
                                              f"{key[1]} to {key[2]} (gap: {gaps}) - deal "
                                              f"excluded from recomputed commission"})
                continue

            commission = ZERO
            tier_desc = []
            for t, portion in breakdown:
                amt = (portion * t["rate"] / Decimal(100))
                commission += amt
                tier_desc.append(f"{fmt_amt(t['min'])}-{fmt_amt(t['max'])}@{t['rate']}%: "
                                  f"{portion:,.2f} -> {amt:,.2f}")

            period = d["close_date"].strftime("%Y-%m")
            recomputed_by_rep_period[(rep, period)] += commission
            schedule_rows.append({
                "rep": rep, "deal_id": d["deal_id"], "close_date": d["close_date"],
                "amount": d["amount"], "plan_period": f"{key[1]} to {key[2]}",
                "cumulative_before": span_start, "cumulative_after": span_end,
                "commission": commission, "tier_breakdown": "; ".join(tier_desc),
                "payroll_period": period,
            })

    payroll_rows = load_csv(Path(args.payroll), PAYROLL_COLS)
    paid_by_rep_period = defaultdict(lambda: ZERO)
    dup_payroll = defaultdict(int)
    for r in payroll_rows:
        try:
            paid = dec(r["commission_paid"]) or ZERO
        except ValueError as e:
            exceptions.append({"rep": r["rep"], "deal_id": "",
                                "reason": f"Payroll row {r['row']}: {e}"})
            continue
        key = (r["rep"], r["period"])
        dup_payroll[key] += 1
        paid_by_rep_period[key] += paid

    all_keys = sorted(set(recomputed_by_rep_period) | set(paid_by_rep_period))
    discrepancies = []
    for key in all_keys:
        recomputed = recomputed_by_rep_period.get(key, ZERO)
        paid = paid_by_rep_period.get(key, ZERO)
        if recomputed != paid:
            discrepancies.append({"rep": key[0], "period": key[1],
                                   "recomputed": recomputed, "paid": paid,
                                   "diff": recomputed - paid})

    valid = not hard_errors
    meta = {
        "Client": args.client or "(not stated)",
        "Reps": len(by_rep),
        "Deals recomputed": len(schedule_rows),
        "Exceptions": len(exceptions),
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    print("=" * 76)
    print("COMMISSION RECALCULATION")
    print("=" * 76)
    print(f"Client {meta['Client']}   Reps {meta['Reps']}   "
          f"Deals recomputed {meta['Deals recomputed']}")
    print(f"Exceptions {len(exceptions)}   Discrepancies vs. payroll {len(discrepancies)}")
    if hard_errors:
        print(f"\nHARD GATE FAILURES ({len(hard_errors)}):")
        for e in hard_errors:
            print(f"  ! {e}")

    if not valid and not args.force:
        print("\n" + "=" * 76)
        print("WORKBOOK NOT WRITTEN.")
        print("=" * 76)
        return 1

    wb = Workbook()
    ws = wb.active
    ws.title = "Recalculated Schedule"
    cols = ["rep", "deal_id", "close_date", "amount", "plan_period",
            "cumulative_before", "cumulative_after", "commission", "tier_breakdown",
            "payroll_period"]
    style_header(ws, 1, cols)
    r = 2
    for s in sorted(schedule_rows, key=lambda s: (s["rep"], s["close_date"])):
        ws.cell(row=r, column=1, value=s["rep"])
        ws.cell(row=r, column=2, value=s["deal_id"])
        ws.cell(row=r, column=3, value=str(s["close_date"]))
        ws.cell(row=r, column=4, value=float(s["amount"])).number_format = MONEY
        ws.cell(row=r, column=5, value=s["plan_period"])
        ws.cell(row=r, column=6, value=float(s["cumulative_before"])).number_format = MONEY
        ws.cell(row=r, column=7, value=float(s["cumulative_after"])).number_format = MONEY
        ws.cell(row=r, column=8, value=float(s["commission"])).number_format = MONEY
        ws.cell(row=r, column=9, value=s["tier_breakdown"])
        ws.cell(row=r, column=10, value=s["payroll_period"])
        r += 1

    ws2 = wb.create_sheet("Payroll Tie-Out")
    style_header(ws2, 1, ["rep", "period", "recomputed", "paid", "difference", "result"])
    r = 2
    for key in all_keys:
        recomputed = recomputed_by_rep_period.get(key, ZERO)
        paid = paid_by_rep_period.get(key, ZERO)
        diff = recomputed - paid
        ws2.cell(row=r, column=1, value=key[0])
        ws2.cell(row=r, column=2, value=key[1])
        ws2.cell(row=r, column=3, value=float(recomputed)).number_format = MONEY
        ws2.cell(row=r, column=4, value=float(paid)).number_format = MONEY
        ws2.cell(row=r, column=5, value=float(diff)).number_format = MONEY
        result = ws2.cell(row=r, column=6, value="TIE" if diff == ZERO else "BREAK")
        result.font = OK_FONT if diff == ZERO else BAD_FONT
        if diff != ZERO:
            ws2.cell(row=r, column=5).fill = BAD_FILL
        r += 1

    ws3 = wb.create_sheet("Discrepancies")
    style_header(ws3, 1, ["rep", "period", "recomputed", "paid", "dollar_effect"])
    r = 2
    for d in sorted(discrepancies, key=lambda d: abs(d["diff"]), reverse=True):
        ws3.cell(row=r, column=1, value=d["rep"])
        ws3.cell(row=r, column=2, value=d["period"])
        ws3.cell(row=r, column=3, value=float(d["recomputed"])).number_format = MONEY
        ws3.cell(row=r, column=4, value=float(d["paid"])).number_format = MONEY
        ws3.cell(row=r, column=5, value=float(d["diff"])).number_format = MONEY
        r += 1
    if r == 2:
        ws3.cell(row=2, column=1, value="(no discrepancies - recomputed ties to paid "
                                          "for every rep/period)")

    ws4 = wb.create_sheet("Exceptions")
    style_header(ws4, 1, ["rep", "deal_id", "reason"])
    r = 2
    for e in exceptions:
        ws4.cell(row=r, column=1, value=e["rep"])
        ws4.cell(row=r, column=2, value=e["deal_id"])
        ws4.cell(row=r, column=3, value=e["reason"])
        ws4.cell(row=r, column=1).fill = WARN_FILL
        r += 1
    if r == 2:
        ws4.cell(row=2, column=1, value="(no exceptions)")

    wb.active = 0
    out = Path(args.out)
    if not valid:
        out = out.with_name(out.stem + " [GATE NOT MET]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\n{'GATE PASSED' if valid else 'GATE NOT MET (forced)'}.  Workbook: {out}")
    return 0 if valid else 1


if __name__ == "__main__":
    sys.exit(main())
