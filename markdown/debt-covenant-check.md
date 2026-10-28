---
name: debt-covenant-check
description: Computes covenant ratios from the trial balance using the credit agreement's own component mappings and formulas, tracks headroom over time against prior compliance certificates, and flags approaching breaches — but only after confirming the trial balance itself balances. Use this whenever the user mentions debt covenants, covenant compliance, covenant headroom, leverage ratio, fixed charge coverage ratio, compliance certificate, loan covenant testing, or "are we in compliance with our credit agreement." Trigger on "check our covenants," "compute the leverage ratio," "are we close to breaching a covenant," "build the compliance certificate," or when a trial balance and covenant definitions are supplied together. Pairs with trial-balance-integrity, whose balance check this skill reuses before computing anything — if the TB doesn't balance yet, run that skill first.
---

# Debt Covenant Check

This is a quarterly ritual with real career risk attached if it's wrong, so it starts from the
same discipline as `trial-balance-integrity`: nothing is computed on a trial balance that doesn't
balance. Every covenant's formula, threshold, and account mapping comes from the client's actual
credit agreement — none of it is assumed or defaulted by the script.

## The gate

The script refuses to compute anything, or refuses a specific covenant, when:

1. **The trial balance doesn't balance.** Debits must equal credits before any covenant is
   touched. If it doesn't, the whole run is blocked — run `trial-balance-integrity` first.
2. **A component's mapped account doesn't exist in the TB.** That covenant becomes an exception,
   never computed with the missing account treated as zero.
3. **A covenant has no threshold or comparison operator supplied.** The script refuses that
   covenant by name rather than assuming a "typical" leverage or coverage threshold.
4. **A formula can't be evaluated** from the components actually supplied — divide-by-zero, an
   unrecognized component label, or unsupported syntax are all reported, never guessed around.

Early-warning headroom flags are opt-in: supply `--warning-band-pct` or the Approaching Breach tab
stays intentionally empty rather than silently applying an assumed band.

## Inputs

1. **Trial balance** — `account, debit, credit`.
2. **Covenant components** — `covenant_name, component_label, account, operator` (`add` or
   `subtract`). This is how the client's own credit agreement defines what rolls into, say,
   "EBITDA" or "Fixed Charges" for that specific facility — addbacks and exclusions vary by
   agreement and are never assumed.
3. **Covenant definitions** — `covenant_name, formula_description, formula, threshold,
   comparison`. `formula` is written using the component labels exactly as they appear in the
   components file (e.g. `Total Debt / EBITDA`). `comparison` is `>=`, `<=`, `>`, or `<`.
4. **Prior compliance certificates** (optional) — `covenant_name, period, actual_value`, for the
   headroom trend.
5. **`--warning-band-pct`** (optional) — the percentage-of-threshold band that counts as
   "approaching" a breach. No default; omitting it means no early-warning test runs.

## Step 1 — Get blank templates

```bash
python3 scripts/debt_covenant_check.py --template components --out components_template.csv
python3 scripts/debt_covenant_check.py --template definitions --out definitions_template.csv
```

## Step 2 — Run the check

```bash
python3 scripts/debt_covenant_check.py --tb trial_balance.csv \
  --components covenant_components.csv --definitions covenant_definitions.csv \
  --certificates prior_certificates.csv --warning-band-pct 15 \
  --client "Northwind Systems Inc" --period "FY2026 Q1" \
  --out "Northwind - Covenant Compliance.xlsx"
```

A covenant failing its test (e.g., leverage above the cap) is a real finding, not a gate failure —
the workbook still writes and reports the failure clearly. The gate is about whether the
computation itself is trustworthy, not whether the client happens to be in compliance.

## Deliver

**Workbook tabs:**

1. **Covenant Compliance** — each covenant's computed value, threshold, comparison, pass/fail,
   headroom, whether it's in the approaching-breach band, and the prior-period trend beneath it.
2. **Component Trace** — every account feeding every covenant, with its TB value and how it rolled
   into the component (add/subtract), so the computation can be audited line by line.
3. **Exceptions** — covenants that couldn't be computed and why (missing account, missing
   threshold, unevaluable formula).
4. **Approaching Breach** — only populated when `--warning-band-pct` is supplied; otherwise it
   states plainly that no early-warning test was run, rather than appearing as a clean empty tab.

**Then, in chat:** whether the TB balanced, the pass/fail count, and any covenant currently failing
or approaching its band — those are the only numbers that matter to the reader in the first
sentence.

## What to escalate

- **Any covenant that fails.** Report the computed value, the threshold, and the shortfall — do
  not characterize materiality or likely lender response.
- **A covenant approaching its warning band** for the first time, or trending toward it across the
  certificate history.
- **A component mapping that looks inconsistent with the credit agreement's actual definitions**
  discovered while building the components file — that's a conversation with whoever negotiated
  the facility, not a script decision.
- **A trial balance that doesn't balance.** This blocks the entire check and needs to be resolved
  before covenant compliance can be assessed at all.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry. Trial
balance and credit agreement terms never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*

---

# Appendix — bundled files

If you installed the `.skill` package these files are already in place and you can ignore this appendix. If you copied the skill as text, create the files below at the paths shown, alongside your `SKILL.md`. The skill will not run without them.

## `scripts/debt_covenant_check.py`

```python
#!/usr/bin/env python3
"""
Compute covenant ratios from the trial balance per the credit agreement's own
definitions, track headroom over time, and flag approaching breaches.

Gates:
  * the trial balance must balance (sum of debits = sum of credits) before any
    covenant is computed - same pattern as trial-balance-integrity. If it
    doesn't balance, everything is blocked.
  * every account referenced by a covenant's component mapping must exist in
    the trial balance with a non-blank value, or that covenant is an
    exception - never zero-filled or skipped silently.
  * a covenant with no threshold/comparison supplied is never computed - the
    script refuses that covenant specifically rather than inventing a
    "typical" bank covenant threshold.
  * an early-warning headroom band is only applied when --warning-band-pct is
    supplied - omitting it means no early-warning flag is raised, and the
    workbook says so plainly rather than silently skipping it.

No covenant threshold, ratio definition, or EBITDA addback is hardcoded.
Everything is sourced from the client's own credit agreement, supplied via
--components and --definitions.

Usage - emit blank templates:
    python3 debt_covenant_check.py --template components --out components_t.csv
    python3 debt_covenant_check.py --template definitions --out definitions_t.csv
    (choices: components, definitions, certificates)

Usage - run the check:
    python3 debt_covenant_check.py --tb trial_balance.csv \
        --components covenant_components.csv --definitions covenant_definitions.csv \
        --certificates prior_certificates.csv --warning-band-pct 15 \
        --client "Northwind Systems Inc" --period "FY2026 Q1" \
        --out "Northwind - Covenant Compliance.xlsx"

--tb CSV:              account, debit, credit
--components CSV:      covenant_name, component_label, account, operator
                       operator: add | subtract
--definitions CSV:     covenant_name, formula_description, formula,
                       threshold, comparison
                       formula is one of: "numerator / denominator" style
                       expressions written using component_label names
                       exactly as they appear in --components, e.g.
                       "EBITDA / Fixed Charges". comparison: >= | <= | > | <
--certificates CSV:    covenant_name, period, actual_value
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
MONEY = '#,##0.00;[Red](#,##0.00)'
RATIO = '0.00"x"'

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")

TEMPLATES = {
    "components": ["covenant_name", "component_label", "account", "operator"],
    "definitions": ["covenant_name", "formula_description", "formula", "threshold", "comparison"],
    "certificates": ["covenant_name", "period", "actual_value"],
}

COMPARISONS = {
    ">=": lambda v, t: v >= t, "<=": lambda v, t: v <= t,
    ">": lambda v, t: v > t, "<": lambda v, t: v < t,
}


def dec(raw):
    if raw is None:
        return None
    s = str(raw).strip()
    if s in ("", "-", "--", "n/a", "N/A", "None"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace(",", "").replace("$", "").strip()
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


def check_tb_balances(tb: list[dict]) -> tuple[Decimal, Decimal, dict[str, Decimal], list[str]]:
    errors = []
    total_debit = total_credit = ZERO
    tb_values: dict[str, Decimal] = {}
    for x in tb:
        d = dec(x["debit"]) or ZERO
        c = dec(x["credit"]) or ZERO
        total_debit += d
        total_credit += c
        acct = x["account"]
        if not acct:
            errors.append(f"TB row {x['row']}: account is blank.")
            continue
        tb_values[acct] = tb_values.get(acct, ZERO) + (d - c)
    if total_debit != total_credit:
        errors.append(f"Trial balance does not balance: debits {total_debit:,.2f} vs "
                      f"credits {total_credit:,.2f}, difference {total_debit - total_credit:,.2f}.")
    return total_debit, total_credit, tb_values, errors


def compute_components(components: list[dict], tb_values: dict[str, Decimal]):
    """Return {covenant_name: {component_label: value}} and a list of exceptions."""
    out = defaultdict(lambda: defaultdict(Decimal))
    exceptions = []
    trace = []
    for c in components:
        name, label, acct, op = c["covenant_name"], c["component_label"], c["account"], c["operator"].lower()
        if op not in ("add", "subtract"):
            exceptions.append(f"{name}/{label}: operator '{c['operator']}' must be "
                              f"'add' or 'subtract' (row {c['row']}).")
            continue
        if acct not in tb_values:
            exceptions.append(f"Covenant '{name}', component '{label}': account "
                              f"'{acct}' does not exist in the trial balance "
                              f"(row {c['row']}). Not computed.")
            continue
        val = tb_values[acct]
        signed = val if op == "add" else -val
        out[name][label] += signed
        trace.append({"covenant": name, "component": label, "account": acct,
                      "operator": op, "tb_value": val, "signed_value": signed})
    return out, exceptions, trace


def eval_formula(formula: str, components: dict[str, Decimal]) -> Decimal | None:
    """Evaluate a formula string of component labels combined with + - * / and
    parentheses. Labels are matched longest-first so multi-word labels don't
    get partially replaced."""
    expr = formula
    for label in sorted(components, key=len, reverse=True):
        token = "C{}".format(abs(hash(label)) % 1_000_000)
        expr = expr.replace(label, token)
    # map tokens back to values for eval in a restricted namespace
    ns = {}
    for label in components:
        token = "C{}".format(abs(hash(label)) % 1_000_000)
        ns[token] = float(components[label])
    if not re.fullmatch(r"[C0-9a-zA-Z_+\-*/(). ]+", expr):
        raise ValueError(f"Formula contains unsupported characters: {formula!r}")
    try:
        return Decimal(str(eval(expr, {"__builtins__": {}}, ns)))
    except ZeroDivisionError:
        return None
    except Exception as e:
        raise ValueError(f"Could not evaluate formula {formula!r}: {e}")


def build(tb, components, definitions, certificates, warning_band_pct):
    total_debit, total_credit, tb_values, tb_errors = check_tb_balances(tb)
    if tb_errors:
        return {"blocked": True, "errors": tb_errors, "covenants": [], "trace": [],
                "exceptions": [], "approaching": []}

    comp_values, comp_exceptions, trace = compute_components(components, tb_values)

    defs_by_name = {d["covenant_name"]: d for d in definitions}
    cert_idx = defaultdict(list)
    for c in certificates:
        cert_idx[c["covenant_name"]].append(c)

    errors = list(comp_exceptions)
    covenants = []
    approaching = []
    all_names = sorted(set(list(comp_values)) | set(defs_by_name))
    for name in all_names:
        d = defs_by_name.get(name)
        if not d or not d.get("threshold") or not d.get("comparison"):
            errors.append(f"Covenant '{name}': no threshold/comparison supplied in "
                          f"--definitions. Refusing to compute it rather than assume "
                          f"a typical bank covenant threshold.")
            continue
        comparison = d["comparison"].strip()
        if comparison not in COMPARISONS:
            errors.append(f"Covenant '{name}': comparison '{comparison}' must be one "
                          f"of {sorted(COMPARISONS)}.")
            continue
        threshold = dec(d["threshold"])
        if threshold is None:
            errors.append(f"Covenant '{name}': threshold is not a parseable number.")
            continue
        comps = comp_values.get(name, {})
        if not comps:
            errors.append(f"Covenant '{name}': has a definition but no component "
                          f"mappings in --components. Not computed.")
            continue
        try:
            value = eval_formula(d["formula"], comps)
        except ValueError as e:
            errors.append(f"Covenant '{name}': {e}")
            continue
        if value is None:
            errors.append(f"Covenant '{name}': formula divides by zero given current "
                          f"trial balance values. Not computed.")
            continue
        passed = COMPARISONS[comparison](value, threshold)
        headroom = value - threshold if comparison in (">=", ">") else threshold - value

        warn = False
        if warning_band_pct is not None and threshold != ZERO:
            band = abs(threshold) * warning_band_pct / 100
            warn = passed and abs(headroom) <= band
            if warn:
                approaching.append({"covenant": name, "value": value, "threshold": threshold,
                                    "headroom": headroom, "band": band})

        trend = sorted(cert_idx.get(name, []), key=lambda c: c["period"])
        covenants.append({"name": name, "description": d.get("formula_description", ""),
                          "value": value, "threshold": threshold, "comparison": comparison,
                          "passed": passed, "headroom": headroom, "warn": warn,
                          "trend": [(t["period"], dec(t["actual_value"])) for t in trend]})

    return {"blocked": False, "errors": errors, "covenants": covenants, "trace": trace,
            "exceptions": comp_exceptions, "approaching": approaching,
            "total_debit": total_debit, "total_credit": total_credit}


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


def sheet_compliance(wb, meta, result, valid):
    ws = wb.active
    ws.title = "Covenant Compliance"
    widths(ws, {1: 24, 2: 34, 3: 14, 4: 14, 5: 8, 6: 10, 7: 14, 8: 10})
    cols = ["Covenant", "Description", "Computed Value", "Threshold", "Test",
            "Pass/Fail", "Headroom", "Approaching"]
    for c, v in enumerate(cols, start=1):
        ws.cell(row=1, column=c, value=v)
    hdr(ws, len(cols))
    r = 2
    for cv in result["covenants"]:
        vals = [cv["name"], cv["description"], float(cv["value"]), float(cv["threshold"]),
                cv["comparison"], "PASS" if cv["passed"] else "FAIL",
                float(cv["headroom"]), "YES" if cv["warn"] else ""]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            if c in (3, 4, 7):
                cell.number_format = RATIO
            if c == 6:
                cell.font = OK_FONT if v == "PASS" else BAD_FONT
                if v == "FAIL":
                    cell.fill = BAD_FILL
            if c == 8 and v == "YES":
                cell.fill = WARN_FILL
        r += 1
        for period, actual in cv["trend"]:
            ws.cell(row=r, column=1, value=f"  prior: {period}")
            if actual is not None:
                ws.cell(row=r, column=3, value=float(actual)).number_format = RATIO
            r += 1
    r += 1
    v = ws.cell(row=r, column=1,
                value="ALL DEFINED COVENANTS COMPUTED" if valid else
                "NOT ALL COVENANTS COMPUTED - see Exceptions tab")
    v.font = OK_FONT if valid else BAD_FONT
    if not valid:
        v.fill = BAD_FILL


def sheet_trace(wb, result):
    ws = wb.create_sheet("Component Trace")
    widths(ws, {1: 20, 2: 24, 3: 20, 4: 10, 5: 16, 6: 16})
    cols = ["Covenant", "Component", "Account", "Operator", "TB Value", "Signed Value"]
    for c, v in enumerate(cols, start=1):
        ws.cell(row=1, column=c, value=v)
    hdr(ws, len(cols))
    r = 2
    for t in result["trace"]:
        ws.cell(row=r, column=1, value=t["covenant"])
        ws.cell(row=r, column=2, value=t["component"])
        ws.cell(row=r, column=3, value=t["account"])
        ws.cell(row=r, column=4, value=t["operator"])
        ws.cell(row=r, column=5, value=float(t["tb_value"])).number_format = MONEY
        ws.cell(row=r, column=6, value=float(t["signed_value"])).number_format = MONEY
        r += 1


def sheet_exceptions(wb, result):
    ws = wb.create_sheet("Exceptions")
    widths(ws, {1: 100})
    ws.cell(row=1, column=1, value="Exception")
    hdr(ws, 1)
    r = 2
    for e in result["errors"]:
        ws.cell(row=r, column=1, value=e).fill = BAD_FILL
        r += 1
    if not result["errors"]:
        ws.cell(row=2, column=1, value="None.").font = OK_FONT


def sheet_approaching(wb, result, warning_band_pct):
    ws = wb.create_sheet("Approaching Breach")
    widths(ws, {1: 24, 2: 14, 3: 14, 4: 14, 5: 14})
    if warning_band_pct is None:
        ws.cell(row=1, column=1,
                value="--warning-band-pct not supplied: no early-warning flags were "
                      "evaluated. This tab is intentionally empty, not a clean result.")
        ws.cell(row=1, column=1).fill = WARN_FILL
        return
    cols = ["Covenant", "Value", "Threshold", "Headroom", "Warning Band"]
    for c, v in enumerate(cols, start=1):
        ws.cell(row=1, column=c, value=v)
    hdr(ws, len(cols))
    r = 2
    for a in result["approaching"]:
        ws.cell(row=r, column=1, value=a["covenant"])
        ws.cell(row=r, column=2, value=float(a["value"])).number_format = RATIO
        ws.cell(row=r, column=3, value=float(a["threshold"])).number_format = RATIO
        ws.cell(row=r, column=4, value=float(a["headroom"])).number_format = RATIO
        ws.cell(row=r, column=5, value=float(a["band"])).number_format = RATIO
        for c in range(1, 6):
            ws.cell(row=r, column=c).fill = WARN_FILL
        r += 1
    if not result["approaching"]:
        ws.cell(row=2, column=1, value="None within the warning band.").font = OK_FONT


def write_template(kind: str, out: Path):
    cols = TEMPLATES[kind]
    with out.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(cols)
    print(f"Blank '{kind}' template written to {out}")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                  formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--tb")
    ap.add_argument("--components")
    ap.add_argument("--definitions")
    ap.add_argument("--certificates")
    ap.add_argument("--warning-band-pct", type=str, default=None)
    ap.add_argument("--client", default="")
    ap.add_argument("--period", default="")
    ap.add_argument("--out", required=True)
    ap.add_argument("--template", choices=sorted(TEMPLATES))
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    if args.template:
        write_template(args.template, Path(args.out))
        return 0

    for n, v in (("--tb", args.tb), ("--components", args.components),
                 ("--definitions", args.definitions)):
        if not v:
            sys.exit(f"Missing required argument: {n}")

    tb = load(Path(args.tb), ["account", "debit", "credit"])
    components = load(Path(args.components), TEMPLATES["components"])
    definitions = load(Path(args.definitions), TEMPLATES["definitions"])
    certificates = load(Path(args.certificates), TEMPLATES["certificates"]) if args.certificates else []
    warning_band_pct = dec(args.warning_band_pct) if args.warning_band_pct else None

    result = build(tb, components, definitions, certificates, warning_band_pct)
    errors = result["errors"]
    valid = not errors and not result["blocked"]

    print("=" * 76)
    print("DEBT COVENANT CHECK")
    print("=" * 76)
    print(f"Client: {args.client or '(not stated)'}    Period: {args.period or '(not stated)'}")
    if result["blocked"]:
        print("\nTRIAL BALANCE DOES NOT BALANCE - no covenant was computed.")
        for e in errors:
            print(f"  ! {e}")
    else:
        print(f"TB balances: debits {result['total_debit']:,.2f} = "
              f"credits {result['total_credit']:,.2f}")
        print(f"Covenants computed: {len(result['covenants'])}   "
              f"Pass: {sum(1 for c in result['covenants'] if c['passed'])}   "
              f"Fail: {sum(1 for c in result['covenants'] if not c['passed'])}")
        if result["approaching"]:
            print(f"Approaching breach: {len(result['approaching'])}")
        if errors:
            print(f"\nEXCEPTIONS ({len(errors)}):")
            for e in errors[:25]:
                print(f"  ! {e}")

    if not valid and not args.force:
        print("\n" + "=" * 76)
        print("WORKBOOK NOT WRITTEN.")
        print("The trial balance must balance, and every covenant needs its full")
        print("account mapping, threshold, and comparison before it can be computed.")
        print("=" * 76)
        return 1

    wb = Workbook()
    sheet_compliance(wb, {"client": args.client, "period": args.period}, result, valid)
    sheet_trace(wb, result)
    sheet_exceptions(wb, result)
    sheet_approaching(wb, result, warning_band_pct)
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

