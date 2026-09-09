#!/usr/bin/env python3
"""
Period-over-period flux analysis with dual materiality and a draft-then-confirm
explanation gate.

Flags a line as requiring explanation when it breaches the dollar OR the percent
threshold (percent ignored below an absolute floor), plus four always-flag
conditions that are qualitatively material at any magnitude:

    sign flip · new account · disappeared account ·
    zero variance where the underlying activity changed · round-number plug

If --prior-commentary is given, a flagged account with commentary on file from
the prior period gets a labeled DRAFT explanation carried into the worklist —
never a fabricated cause, only a reuse of what a person already wrote last
period, offered for a reviewer to confirm or replace. A draft never counts as
explained on its own. Only a status of "confirmed", with components that sum
to the variance and an evidence reference, closes a flagged line. Reports
explanation coverage and refuses to call the analysis complete otherwise.

Usage:
    python3 flux_compare.py --current cp.csv --prior pp.csv \
        --prior-commentary last_period_explanations.csv \
        --dollar-threshold 10000 --percent-threshold 10 --absolute-floor 2000 \
        --entity "Acme Holdings LLC" --cp-label "Sep 2026" --pp-label "Aug 2026" \
        --out "Acme Holdings - Sep 2026 Flux.xlsx"

Input CSVs (current and prior, same shape — a two-period trial balance export):
    account_id      stable key used to match periods           required
    account_name    human-readable line name                   required
    amount          signed                                      required
    account_type    asset|liability|equity|revenue|expense      optional
    activity_changed   yes|no|blank  (current-period file only)     optional

--prior-commentary CSV (read-only; last period's CONFIRMED explanations):
    account_id, cause, evidence, owner, notes
      Carried forward as a draft cause only. Components are never carried
      forward — a component tied to last period's dollar movement would
      misstate this period's, so the reviewer re-quantifies from scratch.

--explanations CSV (this period's worklist, read on a second run):
    account_id, cause, components, evidence, owner, notes, status
      components : semicolon-separated signed amounts that must sum to the
                   variance, e.g. "18000; -3200"
      evidence   : document reference, or the literal "client representation"
      status     : "confirmed" or blank/"draft" — only "confirmed" counts
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
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
DRAFT_FILL = PatternFill("solid", fgColor="DDEBF7")

# Amounts that are suspiciously round on a line that should be computed.
# Roundness has to scale with magnitude or the test is pure noise: a modulus
# has to be a meaningful fraction (>=5%) of the amount itself to mean anything.
ROUND_MODULI = (Decimal("1000000"), Decimal("500000"), Decimal("100000"),
                Decimal("50000"), Decimal("25000"), Decimal("10000"),
                Decimal("5000"), Decimal("1000"))
ROUND_MIN = Decimal("1000")
ROUND_SIGNIFICANCE = Decimal("20")   # modulus >= |amount| / 20


def dec(raw) -> Decimal | None:
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


def is_round(v: Decimal) -> str:
    a = abs(v)
    if a < ROUND_MIN:
        return ""
    need = a / ROUND_SIGNIFICANCE
    for m in ROUND_MODULI:
        if m >= need and a >= m and a % m == ZERO:
            return f"exact multiple of {m:,.0f}"
    return ""


def load_tb(path: Path, label: str) -> dict[str, dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No rows in {path}")
    out: dict[str, dict] = {}
    for i, r in enumerate(rows, start=1):
        aid = " ".join(str(r.get("account_id", "") or "").split())
        name = " ".join(str(r.get("account_name", "") or "").split())
        if not aid:
            aid = name
        if not aid:
            raise ValueError(f"{label} row {i}: needs an account_id or an account_name.")
        amt = dec(r.get("amount"))
        if amt is None:
            raise ValueError(f"{label} row {i} ({aid}): amount is required.")
        if aid in out:
            raise ValueError(
                f"{label}: duplicate account_id {aid!r}. Account IDs must be unique - "
                f"aggregate the duplicates or give them distinct IDs."
            )
        act = str(r.get("activity_changed", "") or "").strip().lower()
        out[aid] = {
            "account_id": aid, "account_name": name or aid, "amount": amt,
            "account_type": " ".join(str(r.get("account_type", "") or "").split()).lower(),
            "activity_changed": act in ("yes", "y", "true", "1"),
        }
    return out


def load_commentary(path: Path | None) -> dict[str, dict]:
    """Prior-period commentary: a read-only source for draft causes. No
    components are read from here — see the docstring on why."""
    if not path or not path.exists():
        return {}
    out = {}
    with path.open(newline="", encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            aid = " ".join(str(r.get("account_id", "") or "").split())
            cause = " ".join(str(r.get("cause", "") or "").split())
            if not aid or not cause:
                continue
            out[aid] = {
                "cause": cause,
                "evidence": " ".join(str(r.get("evidence", "") or "").split()),
                "owner": " ".join(str(r.get("owner", "") or "").split()),
            }
    return out


def load_explanations(path: Path | None) -> dict[str, dict]:
    if not path or not path.exists():
        return {}
    out = {}
    with path.open(newline="", encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            aid = " ".join(str(r.get("account_id", "") or "").split())
            if not aid:
                continue
            comps = []
            raw = str(r.get("components", "") or "")
            for part in re.split(r"[;|]", raw):
                part = part.strip()
                if part:
                    v = dec(part)
                    if v is not None:
                        comps.append(v)
            status = " ".join(str(r.get("status", "") or "").split()).lower()
            out[aid] = {
                "cause": " ".join(str(r.get("cause", "") or "").split()),
                "components_raw": raw.strip(),
                "components": comps,
                "evidence": " ".join(str(r.get("evidence", "") or "").split()),
                "owner": " ".join(str(r.get("owner", "") or "").split()),
                "notes": " ".join(str(r.get("notes", "") or "").split()),
                "confirmed": status == "confirmed",
                "status_raw": status or "draft",
            }
    return out


def compare(cp: dict, pp: dict, dollar: Decimal, pct: Decimal,
            floor: Decimal) -> list[dict]:
    rows = []
    for aid in sorted(set(cp) | set(pp)):
        c, p = cp.get(aid), pp.get(aid)
        cp_amt = c["amount"] if c else None
        pp_amt = p["amount"] if p else None
        src = c or p
        var = (cp_amt or ZERO) - (pp_amt or ZERO)

        pct_var = None
        if pp_amt not in (None, ZERO):
            pct_var = (var / abs(pp_amt)) * Decimal(100)

        flags: list[str] = []
        if c and not p:
            flags.append("NEW account this period")
        if p and not c:
            flags.append("DISAPPEARED - present last period, absent this period")
        if cp_amt is not None and pp_amt is not None:
            if (cp_amt > ZERO and pp_amt < ZERO) or (cp_amt < ZERO and pp_amt > ZERO):
                flags.append("SIGN FLIP - direction changed")
            if var == ZERO and c and c["activity_changed"]:
                flags.append("UNCHANGED but activity changed - stale rolled-forward figure?")
            if var == ZERO and cp_amt != ZERO and c and not c["activity_changed"]:
                flags.append("identical to prior period to the penny - confirm still correct")
        if cp_amt is not None:
            rnd = is_round(cp_amt)
            if rnd:
                flags.append(f"round number ({rnd}) - computed or plugged?")

        # dual materiality
        mat_d = abs(var) >= dollar
        mat_p = (pct_var is not None
                 and abs(pct_var) >= pct
                 and max(abs(cp_amt or ZERO), abs(pp_amt or ZERO)) >= floor)
        material = mat_d or mat_p
        why = []
        if mat_d:
            why.append(f"|variance| >= {dollar:,.0f}")
        if mat_p:
            why.append(f"|%| >= {pct}% above {floor:,.0f} floor")

        needs = material or bool(flags)
        rows.append({
            "account_id": aid, "account_name": src["account_name"],
            "account_type": src["account_type"],
            "pp": pp_amt, "cp": cp_amt, "variance": var, "pct": pct_var,
            "material": material, "materiality_basis": "; ".join(why),
            "flags": flags, "needs_explanation": needs,
        })
    rows.sort(key=lambda r: abs(r["variance"]), reverse=True)
    return rows


def attach(rows: list[dict], expl: dict, commentary: dict) -> tuple[list[dict], list[str]]:
    problems: list[str] = []
    for r in rows:
        e = expl.get(r["account_id"])
        prior = commentary.get(r["account_id"])
        r["is_draft"] = False

        if e:
            r["cause"] = e["cause"]
            r["components_raw"] = e["components_raw"]
            r["evidence"] = e["evidence"]
            r["owner"] = e["owner"]
            r["notes"] = e["notes"]
            r["status"] = e["status_raw"]
        elif r["needs_explanation"] and prior:
            # A carried-forward candidate, never a fabricated cause: the exact
            # words a person wrote last period, clearly labeled as unconfirmed
            # and with no dollar components attached to it.
            r["cause"] = f"DRAFT, confirm or replace: {prior['cause']}"
            r["components_raw"] = ""
            r["evidence"] = prior["evidence"]
            r["owner"] = prior["owner"]
            r["notes"] = "Carried forward from prior-period commentary. Re-quantify " \
                         "components for this period before confirming."
            r["status"] = "draft"
            r["is_draft"] = True
        else:
            r["cause"], r["components_raw"] = "", ""
            r["evidence"] = r["owner"] = r["notes"] = ""
            r["status"] = ""

        r["components_tie"] = ""
        r["explained"] = False

        if not r["needs_explanation"]:
            continue

        confirmed = bool(e and e["confirmed"])
        if not confirmed:
            # A draft, or an unconfirmed row, never counts — this is the gate
            # that makes "a reviewer edits it" a fact rather than a promise.
            continue

        comps = e["components"] if e else []
        if comps:
            total = sum(comps, ZERO)
            diff = total - r["variance"]
            if diff == ZERO:
                r["components_tie"] = "ties"
                r["explained"] = True
            else:
                r["components_tie"] = f"OFF BY {diff:,.2f}"
                problems.append(
                    f"{r['account_id']} ({r['account_name']}): components sum to "
                    f"{total:,.2f} but the variance is {r['variance']:,.2f} - "
                    f"{diff:,.2f} of the movement is unexplained. Do not absorb the "
                    f"remainder into 'other'."
                )
        else:
            r["components_tie"] = "no components given"
            problems.append(
                f"{r['account_id']} ({r['account_name']}): marked confirmed with no "
                f"quantified components. An explanation that cannot be tested is a "
                f"restatement of the variance."
            )
        if r["explained"] and not r["evidence"]:
            r["explained"] = False
            problems.append(
                f"{r['account_id']} ({r['account_name']}): confirmed with no evidence "
                f"reference. Cite a document, or state 'client representation' so the "
                f"weaker basis is visible to a reviewer."
            )
    for aid in sorted(set(expl) - {r["account_id"] for r in rows}):
        problems.append(
            f"Explanation supplied for account_id {aid!r}, which appears in neither "
            f"period. Likely a typo."
        )
    return rows, problems


# ------------------------------------------------------------------ workbook

def hdr(ws, n, row=1):
    for c in range(1, n + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def widths(ws, w):
    for c, v in w.items():
        ws.column_dimensions[get_column_letter(c)].width = v


SCHED_COLS = [
    ("account_name", "Account", 40), ("account_type", "Type", 12),
    ("pp", "Prior period", 16), ("cp", "Current period", 16),
    ("variance", "Variance $", 16), ("pct", "Variance %", 12),
    ("material", "Material", 10), ("flagstr", "Flag reason", 42),
    ("status", "Status", 12),
    ("cause", "Cause", 56), ("components_raw", "Components", 26),
    ("components_tie", "Tie", 14),
    ("evidence", "Evidence", 24), ("owner", "Owner", 14),
    ("account_id", "Account ID", 14),
]


def sheet_schedule(wb, rows, title, subset=None):
    ws = wb.create_sheet(title)
    ws.append([lbl for _, lbl, _ in SCHED_COLS])
    hdr(ws, len(SCHED_COLS))
    data = subset if subset is not None else rows
    for r in data:
        r["flagstr"] = "; ".join(r["flags"])
        out = []
        for key, _, _ in SCHED_COLS:
            v = r.get(key)
            if key in ("pp", "cp", "variance"):
                out.append(float(v) if v is not None else None)
            elif key == "pct":
                out.append(float(v) if v is not None else None)
            elif key == "material":
                out.append("YES" if v else "")
            else:
                out.append(v or None)
        ws.append(out)

    idx = {k: i for i, (k, _, _) in enumerate(SCHED_COLS)}
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for k in ("pp", "cp", "variance"):
            row[idx[k]].number_format = MONEY
        row[idx["pct"]].number_format = PCT
        if row[idx["material"]].value == "YES":
            row[idx["material"]].font = BAD_FONT
        if row[idx["flagstr"]].value:
            row[idx["flagstr"]].fill = WARN_FILL
        tie = row[idx["components_tie"]].value
        if tie and tie != "ties":
            row[idx["components_tie"]].fill = BAD_FILL
            row[idx["components_tie"]].font = BAD_FONT
        elif tie == "ties":
            row[idx["components_tie"]].font = OK_FONT
        if row[idx["status"]].value == "draft":
            row[idx["cause"]].fill = DRAFT_FILL
            row[idx["status"]].fill = DRAFT_FILL
        if (row[idx["material"]].value == "YES" or row[idx["flagstr"]].value) \
                and not row[idx["cause"]].value:
            row[idx["cause"]].fill = BAD_FILL
            row[idx["cause"]].value = "*** UNEXPLAINED ***"
        if str(row[idx["evidence"]].value or "").lower() == "client representation":
            row[idx["evidence"]].fill = DRAFT_FILL

    if ws.max_row == 1:
        ws.cell(row=2, column=1, value="None.")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(SCHED_COLS))}{max(ws.max_row, 2)}"
    widths(ws, {i: w for i, (_, _, w) in enumerate(SCHED_COLS, start=1)})


def sheet_summary(wb, rows, meta, stats, problems):
    ws = wb.active
    ws.title = "Summary"
    widths(ws, {1: 4, 2: 50, 3: 20, 4: 20, 5: 62})
    r = 1

    def line(txt, *, bold=False, size=11, fill=None, col=2):
        nonlocal r
        c = ws.cell(row=r, column=col, value=txt)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c.fill = fill
        r += 1

    line("PERIOD-OVER-PERIOD FLUX ANALYSIS", bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    complete = stats["needs"] > 0 and stats["explained"] == stats["needs"] and not problems
    if stats["needs"] == 0:
        line("No line breached materiality and no always-flag condition arose.",
             bold=True, fill=DRAFT_FILL)
    elif complete:
        line("ANALYSIS COMPLETE - every flagged line has a confirmed, evidenced, "
             "quantified cause.", bold=True)
        ws.cell(row=r - 1, column=2).font = OK_FONT
    else:
        line(f"ANALYSIS INCOMPLETE - {stats['needs'] - stats['explained']} flagged "
             f"line(s) not yet confirmed"
             + (f", {len(problems)} explanation problem(s)" if problems else "")
             + ". See Unexplained tab.", bold=True, fill=BAD_FILL)
    r += 1

    line("COVERAGE", bold=True, size=12)
    for label, val in [
        ("Accounts compared", stats["total"]),
        ("Requiring explanation", stats["needs"]),
        ("Drafted from prior-period commentary", stats["drafted"]),
        ("Confirmed", stats["explained"]),
        ("Confirmation rate (count)", f"{stats['rate_count']:.0f}%"),
        ("Confirmation rate (dollars)", f"{stats['rate_dollars']:.0f}%"),
    ]:
        ws.cell(row=r, column=2, value="  " + label)
        ws.cell(row=r, column=3, value=val)
        r += 1
    r += 1

    top = [x for x in rows if x["variance"] != ZERO][:10]
    if top:
        line("TOP 10 MOVEMENTS BY ABSOLUTE DOLLAR", bold=True, size=12)
        for h, col in zip(["Account", "Prior", "Current", "Variance", "Status"],
                          (2, 3, 4, 5, 6)):
            hc = ws.cell(row=r, column=col, value=h)
            hc.fill, hc.font = HDR_FILL, HDR_FONT
        r += 1
        for x in top:
            ws.cell(row=r, column=2, value=x["account_name"])
            for col, key in ((3, "pp"), (4, "cp"), (5, "variance")):
                v = x[key]
                c = ws.cell(row=r, column=col, value=float(v) if v is not None else None)
                c.number_format = MONEY
            label = "confirmed" if x["explained"] else (
                "draft, unconfirmed" if x.get("is_draft") else
                ("unexplained" if x["needs_explanation"] else "")
            )
            cc = ws.cell(row=r, column=6, value=label or "-")
            if x["needs_explanation"] and not x["explained"]:
                cc.fill, cc.font = BAD_FILL, BAD_FONT
            r += 1
        r += 1

    line("A cause is causal and quantified. 'Timing' and 'higher volume' restate "
         "the variance rather than explaining it. A draft carried forward from last "
         "period is a starting point, not an answer — it becomes an answer only once "
         "a reviewer confirms it against this period's actual detail.", bold=True)


def sheet_flags(wb, rows):
    subset = [r for r in rows if r["flags"]]
    sheet_schedule(wb, rows, "Always-Flag Exceptions", subset)


def sheet_drafts(wb, rows):
    subset = [r for r in rows if r.get("is_draft")]
    sheet_schedule(wb, rows, "Draft Queue", subset)


def sheet_unexplained(wb, rows, problems):
    subset = [r for r in rows if r["needs_explanation"] and not r["explained"]]
    sheet_schedule(wb, rows, "Unexplained", subset)
    ws = wb["Unexplained"]
    if problems:
        r = ws.max_row + 3
        c = ws.cell(row=r, column=1, value="EXPLANATION PROBLEMS")
        c.font = Font(bold=True, size=12)
        for p in problems:
            r += 1
            cell = ws.cell(row=r, column=1, value=p)
            cell.font = BAD_FONT
            cell.alignment = Alignment(wrap_text=True, vertical="top")


# ------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--current", required=True)
    ap.add_argument("--prior", required=True)
    ap.add_argument("--prior-commentary")
    ap.add_argument("--explanations")
    ap.add_argument("--dollar-threshold", default="10000")
    ap.add_argument("--percent-threshold", default="10")
    ap.add_argument("--absolute-floor", default="2000",
                    help="below this magnitude, percent variance is ignored as noise")
    ap.add_argument("--entity", default="")
    ap.add_argument("--cp-label", default="Current period")
    ap.add_argument("--pp-label", default="Prior period")
    ap.add_argument("--out", required=True)
    ap.add_argument("--worklist-out", default="flux_worklist.csv")
    args = ap.parse_args()

    dollar = dec(args.dollar_threshold) or ZERO
    pct = dec(args.percent_threshold) or ZERO
    floor = dec(args.absolute_floor) or ZERO

    cp = load_tb(Path(args.current), "current")
    pp = load_tb(Path(args.prior), "prior")
    rows = compare(cp, pp, dollar, pct, floor)
    commentary = load_commentary(Path(args.prior_commentary) if args.prior_commentary else None)
    expl = load_explanations(Path(args.explanations) if args.explanations else None)
    rows, problems = attach(rows, expl, commentary)

    needs = [r for r in rows if r["needs_explanation"]]
    explained = [r for r in needs if r["explained"]]
    drafted = [r for r in needs if r.get("is_draft")]
    flagged_dollars = sum((abs(r["variance"]) for r in needs), ZERO)
    explained_dollars = sum((abs(r["variance"]) for r in explained), ZERO)
    stats = {
        "total": len(rows), "needs": len(needs), "explained": len(explained),
        "drafted": len(drafted),
        "rate_count": (len(explained) / len(needs) * 100) if needs else 100.0,
        "rate_dollars": (float(explained_dollars / flagged_dollars) * 100)
                        if flagged_dollars else 100.0,
    }

    # worklist for the reviewer: every open item, drafts pre-filled, others blank
    open_items = [r for r in needs if not r["explained"]]
    if open_items:
        with Path(args.worklist_out).open("w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["account_id", "account_name", "prior", "current", "variance",
                        "pct", "flag_reason", "cause", "components", "evidence",
                        "owner", "notes", "status"])
            for r in open_items:
                w.writerow([
                    r["account_id"], r["account_name"],
                    f"{r['pp']:.2f}" if r["pp"] is not None else "",
                    f"{r['cp']:.2f}" if r["cp"] is not None else "",
                    f"{r['variance']:.2f}",
                    f"{r['pct']:.1f}" if r["pct"] is not None else "",
                    "; ".join(r["flags"]) or r["materiality_basis"],
                    r["cause"], "", r["evidence"], r["owner"], r["notes"],
                    "draft" if r.get("is_draft") else "",
                ])

    meta = {
        "Entity": args.entity or "(not stated)",
        "Current period": args.cp_label,
        "Prior period": args.pp_label,
        "Dollar threshold": float(dollar),
        "Percent threshold": f"{pct}%",
        "Absolute floor (percent test ignored below)": float(floor),
        "Prior-period commentary supplied": "yes" if commentary else "no",
    }

    wb = Workbook()
    sheet_summary(wb, rows, meta, stats, problems)
    sheet_schedule(wb, rows, "Flux Schedule")
    sheet_flags(wb, rows)
    sheet_drafts(wb, rows)
    sheet_unexplained(wb, rows, problems)
    wb.active = 0

    complete = (stats["needs"] == 0) or (stats["explained"] == stats["needs"] and not problems)
    out = Path(args.out)
    if not complete:
        out = out.with_name(out.stem + " [INCOMPLETE]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))

    # ---- console
    print("=" * 72)
    print(f"FLUX ANALYSIS:  {args.pp_label} -> {args.cp_label}")
    print("=" * 72)
    print(f"Materiality: |var| >= {dollar:,.0f}  OR  |%| >= {pct}%  "
          f"(percent ignored below {floor:,.0f})")
    print(f"Accounts compared     : {stats['total']}")
    print(f"Requiring explanation : {stats['needs']}")
    print(f"Drafted from prior commentary : {stats['drafted']}")
    print(f"Confirmed             : {stats['explained']}  "
          f"({stats['rate_count']:.0f}% of items, {stats['rate_dollars']:.0f}% of dollars)")

    flags = [r for r in rows if r["flags"]]
    if flags:
        print(f"\nAlways-flag exceptions ({len(flags)}):")
        for r in flags[:20]:
            print(f"  ! {r['account_name'][:40]:<40} {r['variance']:>14,.2f}  "
                  f"{'; '.join(r['flags'])[:60]}")
        if len(flags) > 20:
            print(f"  ... and {len(flags) - 20} more")

    if open_items:
        print(f"\nOPEN ({len(open_items)}), {len(drafted)} with a draft cause to review:")
        for r in open_items[:20]:
            tag = " [DRAFT]" if r.get("is_draft") else ""
            print(f"  ? {r['account_name'][:40]:<40} {r['variance']:>14,.2f}{tag}")
        if len(open_items) > 20:
            print(f"  ... and {len(open_items) - 20} more")
        print(f"\nWorklist written to {args.worklist_out}. Confirm or replace each "
              f"cause, quantify components, set status=confirmed, and re-run with "
              f"--explanations.")

    if problems:
        print(f"\nEXPLANATION PROBLEMS ({len(problems)}):")
        for p in problems[:20]:
            print(f"  ! {p}")

    print(f"\nStatus  : {'COMPLETE' if complete else 'INCOMPLETE'}")
    print(f"Workbook: {out}")
    return 0 if complete else 1


if __name__ == "__main__":
    sys.exit(main())
