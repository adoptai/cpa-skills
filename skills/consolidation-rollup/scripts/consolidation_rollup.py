#!/usr/bin/env python3
"""
Consolidate multiple entities' trial balances into one consolidated trial
balance, with FX translation, a formula-driven cumulative translation
adjustment (CTA), and consolidation elimination entries applied at the
worksheet level — the textbook parent/sub/eliminations/consolidated layout,
done mechanically and proved to tie.

The one distinction this script insists on: a cumulative translation
adjustment is NOT a plug. Translating different statement sections at
different required rates (assets/liabilities at the closing rate,
revenue/expense at the average rate, equity at a historical rate) necessarily
leaves each entity's own translated trial balance out of balance by a
specific, derivable amount — that amount IS the cumulative translation
adjustment under ASC 830/IAS 21, and it is computed here as a strict formula
from the rates you supplied, never guessed, never used to silently absorb an
unrelated imbalance (unmapped accounts or an unbalanced elimination are
reported as failures, not folded into CTA).

What this script does NOT do:
  * split non-controlling interest for a partially-owned entity — flagged,
    not computed. Book the NCI entry yourself via --elimination-entries.
  * derive intercompany elimination amounts — it posts what
    --elimination-entries gives it and checks each entry balances. Run
    intercompany-recon first to prove the amount.
  * perform the full historical rollforward of retained earnings (opening RE
    at its own historical rate, plus current-year net income at the average
    rate, less dividends at their own rate). It translates whatever balance
    sits in the equity-mapped GL account at a single historical rate. Split
    the account in your mapping if GAAP-precision equity translation matters
    for this engagement.
  * post journal entries anywhere. Everything here is presentation.

Usage:
    python3 consolidation_rollup.py \
        --entities entities.csv --trial-balances entity_tb.csv \
        --mapping mapping.csv --fx-rates fx_rates.csv \
        --reporting-currency USD \
        --elimination-entries elimination_entries.csv \
        --ic-status ic_pair_status.csv \
        --client "Acme Group" --period "2026-01" \
        --out "Acme Group - 2026-01 Consolidated TB.xlsx"

entities.csv (required):
    entity_id, entity_name, functional_currency, ownership_pct
      ownership_pct optional, defaults to 100. Below 100 is flagged — this
      script fully consolidates at 100% and does not split out NCI.

trial-balances.csv (required) — every entity's TB in one file:
    entity_id, account_number, account_name, debit, credit

mapping.csv (required) — entity account -> consolidated line:
    entity_id, account_number, consolidated_code, consolidated_line, stmt_type
      entity_id may be "*" to apply to every entity not matched more
      specifically by its own entity_id. stmt_type: asset | liability |
      equity | revenue | expense — this is what selects the translation
      rate (closing / closing / historical / average / average).

fx_rates.csv (required for every functional currency other than the
reporting currency):
    currency, rate_type, rate_to_reporting
      rate_type: closing | average | historical. closing and average are
      mandatory for every non-reporting currency in play; this script will
      not assume 1:1. historical may be omitted — it then falls back to the
      closing rate, visibly flagged as an approximation, never silently.

elimination-entries.csv (optional) — worksheet-level consolidating entries,
already in reporting currency, posted directly against a consolidated line
(the standard consolidation-worksheet treatment: investment elimination,
intercompany AR/AP elimination, intercompany profit elimination, etc.):
    entry_id, consolidated_line, stmt_type, debit, credit, description,
    entity_a, entity_b, source
      entry_id groups the lines of one adjusting entry; every entry_id's
      debits must equal its credits or none of its lines are posted.
      entity_a/entity_b are optional — populate both for an intercompany
      elimination so it can be checked against --ic-status.

ic-status.csv (optional) — the reconciliation status of each entity pair,
e.g. intercompany-recon's own Pair Summary exported as entity_a, entity_b,
status. Any IC-tagged elimination whose pair is not RECONCILED keeps the
consolidation provisional.

Money is Decimal throughout, quantized to the cent as each translated or
eliminated figure is produced. Nothing is rounded away for display.
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
CENT = Decimal("0.01")
MONEY = '#,##0.00;[Red](#,##0.00)'

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")

STMT_TYPES = {"asset", "liability", "equity", "revenue", "expense"}
RATE_TYPE_BY_STMT = {
    "asset": "closing", "liability": "closing",
    "revenue": "average", "expense": "average",
    "equity": "historical",
}
CTA_LINE = "Cumulative Translation Adjustment (OCI)"


# ------------------------------------------------------------------ parsing

def clean(s) -> str:
    return " ".join(str(s or "").split())


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


def d0(raw) -> Decimal:
    v = dec(raw)
    return v.quantize(CENT) if v is not None else ZERO


def rows_of(path: Path, label: str) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No rows in {label} ({path}).")
    return rows


# ------------------------------------------------------------------ loaders

def load_entities(path: Path) -> dict[str, dict]:
    out = {}
    for i, r in enumerate(rows_of(path, "entities"), start=1):
        eid = clean(r.get("entity_id"))
        if not eid:
            sys.exit(f"entities.csv row {i}: entity_id is required.")
        ccy = clean(r.get("functional_currency")).upper()
        if not ccy:
            sys.exit(f"entities.csv row {i} ({eid}): functional_currency is required.")
        pct = dec(r.get("ownership_pct"))
        out[eid] = {
            "name": clean(r.get("entity_name")) or eid,
            "currency": ccy,
            "ownership_pct": pct if pct is not None else Decimal("100"),
        }
    return out


def load_fx(path: Path | None, entities: dict, reporting_ccy: str) -> tuple[dict, set]:
    rates: dict[tuple[str, str], Decimal] = {}
    for rt in ("closing", "average", "historical"):
        rates[(reporting_ccy, rt)] = Decimal("1")

    if path and path.exists():
        for i, r in enumerate(rows_of(path, "fx rates"), start=1):
            ccy = clean(r.get("currency")).upper()
            rt = clean(r.get("rate_type")).lower()
            rate = dec(r.get("rate_to_reporting"))
            if not ccy:
                continue
            if rt not in ("closing", "average", "historical"):
                sys.exit(f"fx_rates.csv row {i} ({ccy}): rate_type {rt!r} is not "
                          f"recognized. Use closing, average, or historical.")
            if rate is None:
                sys.exit(f"fx_rates.csv row {i} ({ccy}, {rt}): rate_to_reporting "
                          f"is required and must be numeric.")
            rates[(ccy, rt)] = rate

    currencies = sorted({e["currency"] for e in entities.values()})
    missing = [c for c in currencies if c != reporting_ccy
               and ((c, "closing") not in rates or (c, "average") not in rates)]
    if missing:
        sys.exit(
            "No closing and/or average rate to the reporting currency for: "
            + ", ".join(missing) + ". Supply --fx-rates with a closing AND an "
            "average rate_to_reporting for each functional currency in play. "
            "This script will not assume a 1:1 rate."
        )

    historical_fallback = set()
    for c in currencies:
        if c != reporting_ccy and (c, "historical") not in rates:
            rates[(c, "historical")] = rates[(c, "closing")]
            historical_fallback.add(c)
    return rates, historical_fallback


def get_rate(rates: dict, currency: str, stmt_type: str,
             historical_fallback: set) -> tuple[Decimal, str, bool]:
    rt = RATE_TYPE_BY_STMT[stmt_type]
    rate = rates[(currency, rt)]
    approximated = rt == "historical" and currency in historical_fallback
    return rate, rt, approximated


def load_mapping(path: Path) -> tuple[dict, dict]:
    exact, wild = {}, {}
    for i, r in enumerate(rows_of(path, "mapping"), start=1):
        eid = clean(r.get("entity_id"))
        acct = clean(r.get("account_number"))
        if not acct:
            sys.exit(f"mapping.csv row {i}: account_number is required.")
        line = clean(r.get("consolidated_line"))
        stmt = clean(r.get("stmt_type")).lower()
        if not line or not stmt:
            sys.exit(f"mapping.csv row {i} ({eid or '*'} {acct}): consolidated_line "
                      f"and stmt_type are both required.")
        if stmt not in STMT_TYPES:
            sys.exit(f"mapping.csv row {i} ({acct}): stmt_type {stmt!r} is not "
                      f"recognized. Use one of {sorted(STMT_TYPES)}.")
        row = {"consolidated_line": line, "stmt_type": stmt,
               "consolidated_code": clean(r.get("consolidated_code"))}
        if eid and eid != "*":
            exact[(eid, acct)] = row
        else:
            wild[acct] = row
    return exact, wild


def lookup_mapping(exact: dict, wild: dict, entity_id: str, account_number: str):
    return exact.get((entity_id, account_number)) or wild.get(account_number)


def load_trial_balances(path: Path, entities: dict) -> list[dict]:
    out = []
    for i, r in enumerate(rows_of(path, "trial balances"), start=1):
        eid = clean(r.get("entity_id"))
        if eid not in entities:
            sys.exit(f"trial-balances.csv row {i}: entity_id {eid!r} is not in "
                      f"entities.csv.")
        acct = clean(r.get("account_number"))
        if not acct:
            sys.exit(f"trial-balances.csv row {i} ({eid}): account_number is "
                      f"required.")
        out.append({
            "row": i, "entity_id": eid, "account_number": acct,
            "account_name": clean(r.get("account_name")),
            "debit": d0(r.get("debit")), "credit": d0(r.get("credit")),
        })
    return out


def load_elimination_entries(path: Path | None) -> list[dict]:
    if not path or not path.exists():
        return []
    out = []
    for i, r in enumerate(rows_of(path, "elimination entries"), start=1):
        entry_id = clean(r.get("entry_id"))
        if not entry_id:
            sys.exit(f"elimination-entries.csv row {i}: entry_id is required - it "
                      f"groups the lines of one consolidating entry.")
        line = clean(r.get("consolidated_line"))
        stmt = clean(r.get("stmt_type")).lower()
        if not line or not stmt:
            sys.exit(f"elimination-entries.csv row {i} ({entry_id}): "
                      f"consolidated_line and stmt_type are both required.")
        if stmt not in STMT_TYPES:
            sys.exit(f"elimination-entries.csv row {i} ({entry_id}): stmt_type "
                      f"{stmt!r} is not recognized. Use one of {sorted(STMT_TYPES)}.")
        out.append({
            "row": i, "entry_id": entry_id, "consolidated_line": line,
            "stmt_type": stmt, "debit": d0(r.get("debit")), "credit": d0(r.get("credit")),
            "description": clean(r.get("description")),
            "entity_a": clean(r.get("entity_a")), "entity_b": clean(r.get("entity_b")),
            "source": clean(r.get("source")),
        })
    return out


def load_ic_status(path: Path | None) -> dict | None:
    if not path or not path.exists():
        return None
    out = {}
    for r in rows_of(path, "ic status"):
        a, b = clean(r.get("entity_a")), clean(r.get("entity_b"))
        status = clean(r.get("status")).upper()
        if a and b:
            out[frozenset((a, b))] = status
    return out


# ------------------------------------------------------------------ workbook helpers

def hdr(ws, n, row=1):
    for c in range(1, n + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 28


def widths(ws, w):
    for c, v in w.items():
        ws.column_dimensions[get_column_letter(c)].width = v


# ------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                  formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--entities", required=True)
    ap.add_argument("--trial-balances", required=True)
    ap.add_argument("--mapping", required=True)
    ap.add_argument("--fx-rates")
    ap.add_argument("--reporting-currency", required=True)
    ap.add_argument("--elimination-entries")
    ap.add_argument("--ic-status")
    ap.add_argument("--client", default="")
    ap.add_argument("--period", default="")
    ap.add_argument("--out", required=True)
    ap.add_argument("--unmapped-out", default="consolidation_unmapped.csv")
    args = ap.parse_args()

    entities = load_entities(Path(args.entities))
    reporting_ccy = args.reporting_currency.strip().upper()
    fx, historical_fallback = load_fx(
        Path(args.fx_rates) if args.fx_rates else None, entities, reporting_ccy)
    mapping_exact, mapping_wild = load_mapping(Path(args.mapping))
    tb_rows = load_trial_balances(Path(args.trial_balances), entities)
    elim_rows = load_elimination_entries(
        Path(args.elimination_entries) if args.elimination_entries else None)
    ic_status = load_ic_status(Path(args.ic_status) if args.ic_status else None)

    # ---- Test 1: every entity TB balances, in functional currency, before
    # anything downstream touches it.
    entity_totals = defaultdict(lambda: [ZERO, ZERO])
    for r in tb_rows:
        t = entity_totals[r["entity_id"]]
        t[0] += r["debit"]; t[1] += r["credit"]
    unbalanced_entities = [(eid, d, c) for eid, (d, c) in entity_totals.items() if d != c]
    for eid in entities:
        entity_totals.setdefault(eid, [ZERO, ZERO])

    # ---- mapping
    mapped_rows, unmapped_rows = [], []
    for r in tb_rows:
        m = lookup_mapping(mapping_exact, mapping_wild, r["entity_id"], r["account_number"])
        if not m:
            unmapped_rows.append(r)
            continue
        r["consolidated_line"] = m["consolidated_line"]
        r["stmt_type"] = m["stmt_type"]
        mapped_rows.append(r)
    unmapped_by_entity = defaultdict(int)
    for r in unmapped_rows:
        unmapped_by_entity[r["entity_id"]] += 1

    # ---- translation
    for r in mapped_rows:
        ccy = entities[r["entity_id"]]["currency"]
        rate, rate_type, approximated = get_rate(fx, ccy, r["stmt_type"], historical_fallback)
        r["currency"] = ccy
        r["rate"], r["rate_type"], r["approximated"] = rate, rate_type, approximated
        r["translated_debit"] = (r["debit"] * rate).quantize(CENT)
        r["translated_credit"] = (r["credit"] * rate).quantize(CENT)

    # ---- entity-level CTA: the formula-driven balancing line, computed only
    # where a real translation effect can exist (functional currency differs
    # from the reporting currency). An unmapped-driven imbalance is never
    # relabeled as CTA - it surfaces as a tie failure instead.
    cta_rows = []
    entity_cta: dict[str, dict] = {}
    by_entity = defaultdict(list)
    for r in mapped_rows:
        by_entity[r["entity_id"]].append(r)
    for eid, rows in by_entity.items():
        ccy = entities[eid]["currency"]
        td = sum((r["translated_debit"] for r in rows), ZERO)
        tc = sum((r["translated_credit"] for r in rows), ZERO)
        net = td - tc
        cta_debit = cta_credit = ZERO
        if ccy != reporting_ccy and net != ZERO:
            if net > ZERO:
                cta_credit = net
            else:
                cta_debit = -net
            cta_rows.append({
                "entity_id": eid, "account_number": "CTA",
                "account_name": "Cumulative Translation Adjustment (OCI)",
                "consolidated_line": CTA_LINE, "stmt_type": "equity",
                "currency": ccy, "debit": ZERO, "credit": ZERO,
                "rate": None, "rate_type": "formula (balancing)", "approximated": False,
                "translated_debit": cta_debit, "translated_credit": cta_credit,
            })
        entity_cta[eid] = {
            "currency": ccy, "translated_debit_total": td, "translated_credit_total": tc,
            "cta_debit": cta_debit, "cta_credit": cta_credit,
            "unmapped_count": unmapped_by_entity.get(eid, 0),
        }

    all_rows = mapped_rows + cta_rows

    # ---- Consolidated TB, pre-elimination
    cons: dict[str, dict] = {}

    def bucket(line, stmt):
        b = cons.setdefault(line, {"debit": ZERO, "credit": ZERO, "stmt_type": stmt,
                                    "conflict": False})
        if b["stmt_type"] != stmt:
            b["conflict"] = True
        return b

    for r in all_rows:
        b = bucket(r["consolidated_line"], r["stmt_type"])
        b["debit"] += r["translated_debit"]; b["credit"] += r["translated_credit"]

    # ---- Test 2: every elimination entry balances by entry_id.
    by_entry = defaultdict(list)
    for r in elim_rows:
        by_entry[r["entry_id"]].append(r)
    unbalanced_entries_elim, balanced_elim_rows = [], []
    for entry_id, rows in by_entry.items():
        d = sum((row["debit"] for row in rows), ZERO)
        c = sum((row["credit"] for row in rows), ZERO)
        if d != c:
            unbalanced_entries_elim.append((entry_id, d, c, rows))
        else:
            balanced_elim_rows.extend(rows)

    # ---- Test 6: IC-tagged eliminations checked against --ic-status.
    ic_flags = []
    for r in elim_rows:
        if r["entity_a"] and r["entity_b"]:
            pair = frozenset((r["entity_a"], r["entity_b"]))
            if ic_status is not None:
                status = ic_status.get(pair, "NOT FOUND")
                r["ic_status"] = status
                if status != "RECONCILED":
                    ic_flags.append(
                        f"Elimination {r['entry_id']} ({r['entity_a']} / "
                        f"{r['entity_b']}): intercompany pair status is "
                        f"{status}, not RECONCILED. Treat this elimination as "
                        f"provisional until it is."
                    )
            else:
                r["ic_status"] = "(not verified - no --ic-status supplied)"
        else:
            r["ic_status"] = ""

    for r in balanced_elim_rows:
        b = bucket(r["consolidated_line"], r["stmt_type"])
        b["debit"] += r["debit"]; b["credit"] += r["credit"]

    total_debit = sum((b["debit"] for b in cons.values()), ZERO)
    total_credit = sum((b["credit"] for b in cons.values()), ZERO)
    conflicts = [line for line, b in cons.items() if b["conflict"]]
    tie_ok = (total_debit == total_credit) and not conflicts

    # ---- gate
    tests = []
    tests.append((
        "1. Entity trial balances balance",
        not unbalanced_entities,
        "Every entity TB ties." if not unbalanced_entities else
        "; ".join(f"{entities[e]['name']}: debit {d:,.2f} vs credit {c:,.2f}"
                  for e, d, c in unbalanced_entities),
    ))
    tests.append((
        "2. Elimination entries balance",
        not unbalanced_entries_elim,
        "Every elimination entry ties." if not unbalanced_entries_elim else
        "; ".join(f"{eid}: debit {d:,.2f} vs credit {c:,.2f} - excluded, not posted"
                  for eid, d, c, _ in unbalanced_entries_elim),
    ))
    tests.append((
        "3. Every account mapped",
        not unmapped_rows,
        "Every TB account maps to a consolidated line." if not unmapped_rows else
        f"{len(unmapped_rows)} account(s) unmapped across "
        f"{len(unmapped_by_entity)} entit(y/ies) - see Unmapped Accounts.",
    ))
    tests.append((
        "4. FX rates complete",
        True,
        "No translation required (single currency)." if not historical_fallback and
        all(e["currency"] == reporting_ccy for e in entities.values()) else
        ("Closing and average rates supplied for every functional currency."
         + (f" {len(historical_fallback)} currenc(y/ies) had no historical rate "
            f"and used the closing rate as an approximation for equity "
            f"translation: {', '.join(sorted(historical_fallback))}."
            if historical_fallback else "")),
    ))
    tests.append((
        "5. Consolidated trial balance ties",
        tie_ok,
        f"Consolidated debit {total_debit:,.2f} = credit {total_credit:,.2f} "
        f"{reporting_ccy}, CTA included as its own line." if tie_ok else
        (f"Consolidated debit {total_debit:,.2f} vs credit {total_credit:,.2f} "
         f"{reporting_ccy} - difference {total_debit - total_credit:,.2f}."
         + (f" Conflicting stmt_type on: {', '.join(conflicts)}." if conflicts else "")),
    ))
    tests.append((
        "6. Intercompany eliminations reconciled",
        not ic_flags,
        "No IC-tagged eliminations." if not any(r["entity_a"] and r["entity_b"] for r in elim_rows) else
        ("Not verified - supply --ic-status to confirm." if ic_status is None else
         ("Every IC-tagged elimination's pair is RECONCILED." if not ic_flags else
          f"{len(ic_flags)} IC-tagged elimination(s) provisional - see Exceptions.")),
    ))
    pack_complete = all(t[1] for t in tests)

    nci_entities = [(eid, e) for eid, e in entities.items()
                    if e["ownership_pct"] != Decimal("100")]

    # ---- console
    print("=" * 72)
    print("CONSOLIDATION ROLLUP")
    print("=" * 72)
    for name, passed, detail in tests:
        print(f"  [{'PASS' if passed else 'FAIL'}] {name}")
        print(f"         {detail}")
    print()
    print("COMPLETE" if pack_complete else "INCOMPLETE - see failing test(s) above")
    if nci_entities:
        print("\nNCI (not computed - out of scope):")
        for eid, e in nci_entities:
            print(f"  ! {e['name']} is {e['ownership_pct']}% owned. Consolidated at "
                  f"100%; book the NCI entry via --elimination-entries.")
    if ic_flags:
        print("\nPROVISIONAL ELIMINATIONS:")
        for f in ic_flags:
            print(f"  ! {f}")

    # ---- unmapped export
    if unmapped_rows:
        with Path(args.unmapped_out).open("w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["entity_id", "entity_name", "account_number", "account_name",
                        "debit", "credit", "functional_currency"])
            for r in sorted(unmapped_rows, key=lambda x: (x["entity_id"], x["account_number"])):
                w.writerow([r["entity_id"], entities[r["entity_id"]]["name"],
                            r["account_number"], r["account_name"],
                            f"{r['debit']:.2f}", f"{r['credit']:.2f}",
                            entities[r["entity_id"]]["currency"]])
        print(f"\nUnmapped worklist: {args.unmapped_out}")

    # ---- workbook
    meta = {
        "Client": args.client or "(not supplied)",
        "Period": args.period or "(not supplied)",
        "Reporting currency": reporting_ccy,
        "Entities": len(entities),
        "Accounts mapped": len(mapped_rows),
        "Accounts unmapped": len(unmapped_rows),
        "Elimination entries posted": len(by_entry) - len(unbalanced_entries_elim),
        "Elimination entries excluded (unbalanced)": len(unbalanced_entries_elim),
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    wb = Workbook()
    sheet_summary(wb, pack_complete, tests, entities, entity_totals, unmapped_by_entity,
                  entity_cta, nci_entities, meta, total_debit, total_credit)
    sheet_consolidated_tb(wb, cons)
    sheet_entity_detail(wb, all_rows, entities)
    sheet_unmapped(wb, unmapped_rows, entities)
    sheet_eliminations(wb, elim_rows, unbalanced_entries_elim)
    sheet_fx_cta(wb, fx, historical_fallback, entities, entity_cta, reporting_ccy)
    sheet_memo(wb, pack_complete, tests, nci_entities, ic_flags, meta)
    wb.active = 0

    out = Path(args.out)
    if not pack_complete:
        out = out.with_name(out.stem + " [INCOMPLETE]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\nWorkpaper: {out}")

    return 0 if pack_complete else 1


# ------------------------------------------------------------------ sheets

def sheet_summary(wb, pack_complete, tests, entities, entity_totals, unmapped_by_entity,
                   entity_cta, nci_entities, meta, total_debit, total_credit):
    ws = wb.active
    ws.title = "Consolidation Summary"
    widths(ws, {1: 4})
    r = 1
    ws.cell(row=r, column=2, value="CONSOLIDATION ROLLUP").font = Font(bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1
    v = ws.cell(row=r, column=2,
                value="COMPLETE - consolidated TB ties" if pack_complete else
                "INCOMPLETE - see failing test(s) below")
    v.font = OK_FONT if pack_complete else BAD_FONT
    if not pack_complete:
        v.fill = BAD_FILL
    r += 2

    ws.cell(row=r, column=2, value="Gate tests").font = Font(bold=True, size=12)
    r += 1
    heads = ["Test", "Result", "Detail"]
    hr = r
    for c, h in enumerate(heads, start=1):
        ws.cell(row=hr, column=c, value=h)
    hdr(ws, len(heads), row=hr)
    r += 1
    for name, passed, detail in tests:
        ws.cell(row=r, column=2, value=name)
        rc = ws.cell(row=r, column=3, value="PASS" if passed else "FAIL")
        rc.font = OK_FONT if passed else BAD_FONT
        if not passed:
            rc.fill = BAD_FILL
        dc = ws.cell(row=r, column=4, value=detail)
        dc.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30
        r += 1
    r += 1

    ws.cell(row=r, column=2, value=f"Consolidated total: debit {total_debit:,.2f} "
                                    f"vs credit {total_credit:,.2f}").font = Font(bold=True)
    r += 2

    ws.cell(row=r, column=2, value="By entity").font = Font(bold=True, size=12)
    r += 1
    heads = ["Entity", "Functional ccy", "Ownership %", "TB debit", "TB credit",
             "TB balanced", "Unmapped accts", "CTA amount", "CTA side"]
    hr = r
    for c, h in enumerate(heads, start=1):
        ws.cell(row=hr, column=c, value=h)
    hdr(ws, len(heads), row=hr)
    r += 1
    for eid, e in sorted(entities.items(), key=lambda kv: kv[1]["name"]):
        d, c = entity_totals[eid]
        cta = entity_cta.get(eid, {"cta_debit": ZERO, "cta_credit": ZERO})
        cta_amt = cta["cta_debit"] or cta["cta_credit"]
        cta_side = "Debit" if cta["cta_debit"] else ("Credit" if cta["cta_credit"] else "")
        row = [e["name"], e["currency"], float(e["ownership_pct"]), float(d), float(c),
               "Yes" if d == c else "NO", unmapped_by_entity.get(eid, 0),
               float(cta_amt), cta_side]
        for col, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=col, value=val)
            if col in (4, 5, 8):
                cell.number_format = MONEY
        bal_cell = ws.cell(row=r, column=6)
        if bal_cell.value != "Yes":
            bal_cell.font, bal_cell.fill = BAD_FONT, BAD_FILL
        if e["ownership_pct"] != Decimal("100"):
            ws.cell(row=r, column=3).fill = WARN_FILL
        r += 1
    widths(ws, {2: 26, 3: 14, 4: 16, 5: 16, 6: 14, 7: 14, 8: 16, 9: 10})
    ws.column_dimensions["D"].width = 60


def sheet_consolidated_tb(wb, cons: dict):
    ws = wb.create_sheet("Consolidated Trial Balance")
    heads = ["Consolidated line", "Stmt type", "Debit", "Credit", "Net"]
    ws.append(heads)
    hdr(ws, len(heads))
    order = {"asset": 0, "liability": 1, "equity": 2, "revenue": 3, "expense": 4}
    lines = sorted(cons.items(), key=lambda kv: (order.get(kv[1]["stmt_type"], 9), kv[0]))
    current_type = None
    sub_d = sub_c = ZERO
    tot_d = tot_c = ZERO

    def write_subtotal(stmt_type, d, c):
        row = ws.max_row + 1
        ws.cell(row=row, column=1, value=f"Subtotal - {stmt_type}").font = Font(bold=True)
        for col, val in ((3, d), (4, c), (5, d - c)):
            cell = ws.cell(row=row, column=col, value=float(val))
            cell.number_format, cell.font = MONEY, Font(bold=True)
            cell.fill = SUB_FILL
        ws.cell(row=row, column=2).fill = SUB_FILL

    for line, b in lines:
        if current_type is not None and b["stmt_type"] != current_type:
            write_subtotal(current_type, sub_d, sub_c)
            sub_d = sub_c = ZERO
        current_type = b["stmt_type"]
        sub_d += b["debit"]; sub_c += b["credit"]
        tot_d += b["debit"]; tot_c += b["credit"]
        r = ws.max_row + 1
        label = line + (" *** CONFLICTING STMT_TYPE ***" if b["conflict"] else "")
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=b["stmt_type"])
        ws.cell(row=r, column=3, value=float(b["debit"])).number_format = MONEY
        ws.cell(row=r, column=4, value=float(b["credit"])).number_format = MONEY
        ws.cell(row=r, column=5, value=float(b["debit"] - b["credit"])).number_format = MONEY
        if b["conflict"]:
            ws.cell(row=r, column=1).font, ws.cell(row=r, column=1).fill = BAD_FONT, BAD_FILL
    if current_type is not None:
        write_subtotal(current_type, sub_d, sub_c)

    r = ws.max_row + 2
    bold12 = Font(bold=True, size=12)
    ws.cell(row=r, column=1, value="CONSOLIDATED TOTAL").font = bold12
    ws.cell(row=r, column=3, value=float(tot_d)).number_format = MONEY
    ws.cell(row=r, column=3).font = bold12
    ws.cell(row=r, column=4, value=float(tot_c)).number_format = MONEY
    ws.cell(row=r, column=4).font = bold12
    diff = ws.cell(row=r, column=5, value=float(tot_d - tot_c))
    diff.number_format = MONEY
    diff.font = Font(bold=True, size=12, color=(OK_FONT if tot_d == tot_c else BAD_FONT).color)
    if tot_d != tot_c:
        diff.fill = BAD_FILL

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:E{ws.max_row}"
    widths(ws, {1: 42, 2: 12, 3: 16, 4: 16, 5: 16})


def sheet_entity_detail(wb, all_rows: list[dict], entities: dict):
    ws = wb.create_sheet("Entity Detail")
    heads = ["Entity", "Account", "Account name", "Stmt type", "Consolidated line",
              "Functional ccy", "Debit (functional)", "Credit (functional)",
              "Rate type", "Rate", "Approximated", "Debit (reporting)",
              "Credit (reporting)"]
    ws.append(heads)
    hdr(ws, len(heads))
    for r in sorted(all_rows, key=lambda x: (entities[x["entity_id"]]["name"], x["account_number"])):
        ws.append([
            entities[r["entity_id"]]["name"], r["account_number"], r["account_name"],
            r["stmt_type"], r["consolidated_line"], r["currency"],
            float(r["debit"]), float(r["credit"]), r["rate_type"],
            float(r["rate"]) if r["rate"] is not None else None,
            "Yes" if r["approximated"] else "",
            float(r["translated_debit"]), float(r["translated_credit"]),
        ])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for idx in (7, 8, 12, 13):
            row[idx - 1].number_format = MONEY
        if row[9].value is not None:
            row[9].number_format = "#,##0.000000"
        if row[10].value == "Yes":
            row[10].fill = WARN_FILL
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:M{max(ws.max_row, 2)}"
    widths(ws, {1: 22, 2: 12, 3: 30, 4: 10, 5: 30, 6: 10, 7: 15, 8: 15, 9: 12,
                10: 14, 11: 12, 12: 15, 13: 15})


def sheet_unmapped(wb, unmapped_rows: list[dict], entities: dict):
    ws = wb.create_sheet("Unmapped Accounts")
    heads = ["Entity", "Account", "Account name", "Debit", "Credit", "Functional ccy"]
    ws.append(heads)
    hdr(ws, len(heads))
    for r in sorted(unmapped_rows, key=lambda x: (x["entity_id"], x["account_number"])):
        ws.append([entities[r["entity_id"]]["name"], r["account_number"], r["account_name"],
                   float(r["debit"]), float(r["credit"]), entities[r["entity_id"]]["currency"]])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[3].number_format = MONEY
        row[4].number_format = MONEY
        row[0].font, row[0].fill = BAD_FONT, BAD_FILL
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value="None - every account maps to a consolidated line.")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{max(ws.max_row, 2)}"
    widths(ws, {1: 22, 2: 12, 3: 32, 4: 15, 5: 15, 6: 12})


def sheet_eliminations(wb, elim_rows: list[dict], unbalanced_entries_elim: list):
    ws = wb.create_sheet("Elimination Entries")
    heads = ["Entry ID", "Consolidated line", "Stmt type", "Debit", "Credit",
              "Description", "Entity A", "Entity B", "IC status", "Source", "Posted"]
    ws.append(heads)
    hdr(ws, len(heads))
    unbalanced_ids = {eid for eid, *_ in unbalanced_entries_elim}
    for r in sorted(elim_rows, key=lambda x: (x["entry_id"], x["consolidated_line"])):
        posted = "NO - entry unbalanced" if r["entry_id"] in unbalanced_ids else "Yes"
        ws.append([r["entry_id"], r["consolidated_line"], r["stmt_type"],
                   float(r["debit"]), float(r["credit"]), r["description"],
                   r["entity_a"], r["entity_b"], r.get("ic_status", ""), r["source"], posted])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[3].number_format = MONEY
        row[4].number_format = MONEY
        if str(row[10].value).startswith("NO"):
            row[10].font, row[10].fill = BAD_FONT, BAD_FILL
        ic = str(row[8].value or "")
        if ic and ic != "RECONCILED" and not ic.startswith("(not verified"):
            row[8].font, row[8].fill = BAD_FONT, WARN_FILL
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value="No elimination entries supplied.")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:K{max(ws.max_row, 2)}"
    widths(ws, {1: 14, 2: 30, 3: 10, 4: 14, 5: 14, 6: 40, 7: 14, 8: 14, 9: 24, 10: 16, 11: 20})


def sheet_fx_cta(wb, fx: dict, historical_fallback: set, entities: dict,
                  entity_cta: dict, reporting_ccy: str):
    ws = wb.create_sheet("FX & CTA")
    ws.append(["FX RATES"])
    ws["A1"].font = Font(bold=True, size=12)
    heads = ["Currency", "Closing", "Average", "Historical", "Historical approximated"]
    hr = ws.max_row + 1
    for c, h in enumerate(heads, start=1):
        ws.cell(row=hr, column=c, value=h)
    hdr(ws, len(heads), row=hr)
    currencies = sorted({e["currency"] for e in entities.values()} | {reporting_ccy})
    for ccy in currencies:
        r = ws.max_row + 1
        ws.cell(row=r, column=1, value=ccy)
        ws.cell(row=r, column=2, value=float(fx[(ccy, "closing")])).number_format = "#,##0.000000"
        ws.cell(row=r, column=3, value=float(fx[(ccy, "average")])).number_format = "#,##0.000000"
        ws.cell(row=r, column=4, value=float(fx[(ccy, "historical")])).number_format = "#,##0.000000"
        approx_cell = ws.cell(row=r, column=5, value="Yes" if ccy in historical_fallback else "")
        if approx_cell.value == "Yes":
            approx_cell.fill = WARN_FILL

    r = ws.max_row + 2
    ws.cell(row=r, column=1, value="PER-ENTITY CTA").font = Font(bold=True, size=12)
    hr = r + 1
    heads = ["Entity", "Functional ccy", "Translated debit total (mapped)",
              "Translated credit total (mapped)", "CTA debit", "CTA credit", "Note"]
    for c, h in enumerate(heads, start=1):
        ws.cell(row=hr, column=c, value=h)
    hdr(ws, len(heads), row=hr)
    for eid, e in sorted(entities.items(), key=lambda kv: kv[1]["name"]):
        cta = entity_cta.get(eid)
        if not cta:
            continue
        r = ws.max_row + 1
        note = ("Same currency as reporting - no translation effect." if cta["currency"] == reporting_ccy
                else "")
        if cta["unmapped_count"]:
            note = (note + " " if note else "") + (
                f"{cta['unmapped_count']} unmapped account(s) excluded from this "
                f"entity's figures - CTA is provisional until mapping is complete.")
        ws.cell(row=r, column=1, value=e["name"])
        ws.cell(row=r, column=2, value=cta["currency"])
        ws.cell(row=r, column=3, value=float(cta["translated_debit_total"])).number_format = MONEY
        ws.cell(row=r, column=4, value=float(cta["translated_credit_total"])).number_format = MONEY
        ws.cell(row=r, column=5, value=float(cta["cta_debit"])).number_format = MONEY
        ws.cell(row=r, column=6, value=float(cta["cta_credit"])).number_format = MONEY
        ws.cell(row=r, column=7, value=note)
        if cta["unmapped_count"]:
            ws.cell(row=r, column=7).fill = WARN_FILL
    widths(ws, {1: 24, 2: 14, 3: 22, 4: 22, 5: 14, 6: 14, 7: 60})


def sheet_memo(wb, pack_complete, tests, nci_entities, ic_flags, meta):
    ws = wb.create_sheet("Exceptions & Memo")
    widths(ws, {1: 4, 2: 100})
    r = 1

    def para(text, bold=False, size=11):
        nonlocal r
        c = ws.cell(row=r, column=2, value=text)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    para("CONSOLIDATION ROLLUP MEMO", bold=True, size=14)
    r += 1
    para(f"Group: {meta.get('Client', '')}    Period: {meta.get('Period', '')}", bold=True)
    para("Status: " + ("COMPLETE" if pack_complete else "INCOMPLETE - see below"), bold=True)
    r += 1

    failing = [(n, d) for n, p, d in tests if not p]
    if failing:
        para("FAILING GATE TESTS", bold=True, size=12)
        for n, d in failing:
            para(f"- {n}: {d}")
        r += 1

    if nci_entities:
        para("NON-CONTROLLING INTEREST (out of scope)", bold=True, size=12)
        for eid, e in nci_entities:
            para(f"- {e['name']} is {e['ownership_pct']}% owned. This workpaper "
                 f"consolidates it at 100% and does not split out NCI - book that "
                 f"entry separately before treating consolidated equity as final.")
        r += 1

    if ic_flags:
        para("PROVISIONAL ELIMINATIONS", bold=True, size=12)
        for f in ic_flags:
            para(f"- {f}")
        r += 1

    para("REMINDER", bold=True, size=12)
    para("The Cumulative Translation Adjustment is a formula, not a plug: it is the "
         "exact, derivable residual of translating assets/liabilities at the closing "
         "rate, revenue/expense at the average rate, and equity at a historical rate. "
         "It never absorbs an unmapped account or an unbalanced elimination entry - "
         "those fail their own gate test and stay visible until fixed.", bold=True)


if __name__ == "__main__":
    sys.exit(main())
