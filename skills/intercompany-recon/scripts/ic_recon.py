#!/usr/bin/env python3
"""
Match intercompany AR/AP legs across a pair of entities and build an
intercompany reconciliation workpaper with proposed elimination entries.

Model: every leg is recorded from ONE entity's own books and is signed from
that entity's perspective — positive means "the counterparty owes me,"
negative means "I owe the counterparty." For two entities A and B to agree,
A's legs against B and B's legs against A must net to zero once B's legs are
mirrored into A's frame (negated). That mirrored-amount comparison is the
entire matching problem, and it is structurally identical to the bank/book
match in bank-rec-to-gl — this script reuses that cascade.

What this script does NOT do: decide whether an unmatched item is timing,
an FX difference, a missing entry, or a data error. It proposes a hypothesis
in the "suggested cause" column of unmatched.csv; the preparer confirms the
real classification and feeds it back with --classifications. The gate math
only uses confirmed classifications, never the hypothesis.

Usage:
    python3 ic_recon.py \
        --entities entities.csv --fx-rates fx_rates.csv \
        --reporting-currency USD \
        --transactions transactions.csv --balances balances.csv \
        --classifications classifications.csv \
        --as-of 2026-01-31 --period "2026-01" --client "Acme Group" \
        --out "Acme - 2026-01 Intercompany Recon.xlsx"

entities.csv (required):
    entity_id, entity_name, functional_currency

fx_rates.csv (required whenever a functional currency differs from
--reporting-currency): one rate per currency, applied uniformly for the
period. This is a single translation rate, not a transaction-date rate —
that gap is exactly what the fx_difference classification exists to surface.
    currency, rate_to_reporting          (units of reporting ccy per 1 unit)

transactions.csv (the IC subledger detail — optional if balances.csv alone
is supplied, but required for transaction-level matching):
    entity, counterparty, doc_id, date, description, amount
      entity, counterparty = entity_id values from entities.csv
      amount = signed, functional currency of `entity`
               (+ counterparty owes entity, - entity owes counterparty)

balances.csv (optional; the stated GL/TB intercompany balance per entity per
counterparty — used as the authoritative closing position and, when
transactions.csv is also given, to prove the subledger ties to the GL):
    entity, counterparty, amount, as_of
      amount = signed, same convention as transactions.csv

--classifications CSV (produced after reviewing unmatched.csv):
    key, classification, je_entity, je_required, expected_resolution, notes
      classification = timing_in_transit | fx_difference | missing_entry |
                        data_error | unidentified
      je_entity       = which entity should book the correcting JE
                         (defaults to the counterparty of the leg if blank)

--expected-pairs CSV (optional; coverage check):
    entity_a, entity_b

Money is Decimal throughout. Nothing is rounded except for display.
"""

from __future__ import annotations

import argparse
import csv
import itertools
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
DATEF = "yyyy-mm-dd"

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
TOP = Border(top=Side(style="thin"))
DBL = Border(top=Side(style="thin"), bottom=Side(style="double"))

STOPWORDS = {
    "the", "and", "for", "inc", "llc", "ltd", "co", "corp", "company", "payment",
    "pmt", "intercompany", "ic", "invoice", "inv", "charge", "allocation",
    "transfer", "settlement", "ref", "no", "id", "to", "from", "of", "on",
}

CLASSIFICATIONS = {
    "timing_in_transit": "Timing - recorded by one entity, not yet by the other",
    "fx_difference":     "FX - both recorded, translated amounts differ",
    "missing_entry":     "Missing entry - counterparty never recorded it",
    "data_error":        "Data error - wrong amount, period, or account",
    "unidentified":      "UNIDENTIFIED - not explained",
}


# ------------------------------------------------------------------ parsing

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


def pdate(raw) -> date:
    s = str(raw or "").strip()
    if not s:
        raise ValueError("Every row needs a date.")
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d-%b-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Unrecognized date {raw!r}. Prefer ISO YYYY-MM-DD.")


def tokens(text: str) -> set[str]:
    words = re.findall(r"[a-z0-9]+", (text or "").lower())
    return {w for w in words if len(w) > 2 and w not in STOPWORDS and not w.isdigit()}


def amounts_close(a: Decimal, b: Decimal, tol_pct: Decimal, tol_floor: Decimal) -> bool:
    if a == b:
        return True
    base = max(abs(a), abs(b))
    return abs(a - b) <= max(tol_floor, base * tol_pct)


def load_entities(path: Path) -> dict[str, dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No rows in {path}")
    out = {}
    for r in rows:
        eid = (r.get("entity_id") or "").strip()
        if not eid:
            raise ValueError("entities.csv has a row with no entity_id.")
        out[eid] = {
            "name": (r.get("entity_name") or eid).strip(),
            "currency": (r.get("functional_currency") or "").strip().upper(),
        }
    return out


def load_fx(path: Path | None, entities: dict, reporting_ccy: str) -> dict[str, Decimal]:
    fx = {reporting_ccy: Decimal("1")}
    if path and path.exists():
        with path.open(newline="", encoding="utf-8-sig") as fh:
            for r in csv.DictReader(fh):
                ccy = (r.get("currency") or "").strip().upper()
                rate = dec(r.get("rate_to_reporting"))
                if not ccy or rate is None:
                    continue
                fx[ccy] = rate
    missing = sorted({e["currency"] for e in entities.values()
                       if e["currency"] and e["currency"] not in fx})
    if missing:
        sys.exit(
            "No FX rate to the reporting currency for: " + ", ".join(missing) +
            ". Supply --fx-rates with a rate_to_reporting for each currency. "
            "This script will not assume a 1:1 rate."
        )
    return fx


def to_reporting(amount: Decimal, currency: str, fx: dict) -> Decimal | None:
    rate = fx.get(currency)
    if rate is None:
        return None
    return (amount * rate).quantize(Decimal("0.01"))


# ------------------------------------------------------------------ loading legs

_seq: dict[str, int] = defaultdict(int)


def _next_key(entity: str) -> str:
    _seq[entity] += 1
    return f"{entity}{_seq[entity]:04d}"


def load_legs(path: Path | None, entities: dict, fx: dict, source: str) -> list[dict]:
    if not path or not path.exists():
        return []
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    out = []
    for i, r in enumerate(rows, start=1):
        entity = (r.get("entity") or "").strip()
        counterparty = (r.get("counterparty") or "").strip()
        if entity not in entities:
            raise ValueError(f"{path.name} row {i}: entity {entity!r} not in entities.csv")
        if counterparty not in entities:
            raise ValueError(f"{path.name} row {i}: counterparty {counterparty!r} not in entities.csv")
        if entity == counterparty:
            raise ValueError(f"{path.name} row {i}: entity and counterparty are both {entity!r}")
        amt = dec(r.get("amount"))
        if amt is None:
            raise ValueError(f"{path.name} row {i}: amount is required and must be signed "
                              f"(+ counterparty owes entity, - entity owes counterparty).")
        if amt == ZERO:
            raise ValueError(f"{path.name} row {i}: zero-amount row. Remove it or explain it "
                              f"as an exception; it is not an intercompany balance.")
        currency = entities[entity]["currency"]
        desc = " ".join((r.get("description") or "").split())
        row_date = pdate(r.get("date") or r.get("as_of"))
        out.append({
            "key": _next_key(entity),
            "source": source,
            "entity": entity,
            "counterparty": counterparty,
            "doc_id": (r.get("doc_id") or "").strip(),
            "date": row_date,
            "description": desc,
            "amount": amt,
            "currency": currency,
            "amount_rpt": to_reporting(amt, currency, fx),
            "tokens": tokens(desc),
            "matched": False,
            "match_rule": "",
            "match_keys": [],
            "day_var": None,
        })
    return out


# ------------------------------------------------------------------ matching

def _link(a: dict, b: dict, rule: str) -> None:
    a["matched"] = b["matched"] = True
    a["match_rule"] = b["match_rule"] = rule
    a["match_keys"].append(b["key"])
    b["match_keys"].append(a["key"])
    var = (a["date"] - b["date"]).days
    a["day_var"] = -var
    b["day_var"] = var


def match_pair(side1: list[dict], side2: list[dict], tol_days: int,
               tol_pct: Decimal, tol_floor: Decimal, token_thresh: float,
               subset_cap: int) -> list[dict]:
    """side1/side2 each carry a 'cmp' field already expressed in the same
    frame (side2's cmp is the negated, reporting-currency mirror of its raw
    amount). Pairing on tolerant amounts is deliberately generous — it is
    there to find candidate pairs across independent FX translations, not to
    declare them equal. The caller computes a.cmp - b.cmp for every matched
    pair and carries any nonzero variance into the mismatch queue for
    classification; matching never silently absorbs it."""
    pairs: list[dict] = []

    def open_side(rows):
        return [r for r in rows if not r["matched"]]

    def record(a, b, rule):
        _link(a, b, rule)
        pairs.append({"rule": rule, "a": a, "b": b})

    # Pass 1 - doc_id match, regardless of amount (strongest evidence of
    # "same transaction"; a mismatch here is itself the finding).
    idx = defaultdict(list)
    for b in side2:
        if b["doc_id"]:
            idx[b["doc_id"]].append(b)
    for a in side1:
        if a["matched"] or not a["doc_id"]:
            continue
        for cand in idx.get(a["doc_id"], []):
            if cand["matched"]:
                continue
            record(a, cand, "1 doc_id match")
            break

    # Pass 2 - amount close (post-FX) + exact date
    idx = defaultdict(list)
    for b in open_side(side2):
        if b["cmp"] is not None:
            idx[b["date"]].append(b)
    for a in open_side(side1):
        if a["cmp"] is None:
            continue
        cands = [c for c in idx.get(a["date"], [])
                 if not c["matched"] and amounts_close(a["cmp"], c["cmp"], tol_pct, tol_floor)]
        if cands:
            record(a, cands[0], "2 amount + date")

    # Pass 3 - amount close within date tolerance (nearest date wins)
    for a in open_side(side1):
        if a["cmp"] is None:
            continue
        cands = [c for c in open_side(side2)
                 if c["cmp"] is not None and not c["matched"]
                 and abs((a["date"] - c["date"]).days) <= tol_days
                 and amounts_close(a["cmp"], c["cmp"], tol_pct, tol_floor)]
        if cands:
            best = min(cands, key=lambda c: abs((a["date"] - c["date"]).days))
            record(a, best, f"3 amount, date within {tol_days}d")

    # Pass 4 - amount close + description token overlap
    for a in open_side(side1):
        if a["cmp"] is None or not a["tokens"]:
            continue
        best, best_score = None, 0.0
        for b in open_side(side2):
            if b["cmp"] is None or not b["tokens"]:
                continue
            if not amounts_close(a["cmp"], b["cmp"], tol_pct, tol_floor):
                continue
            inter = len(a["tokens"] & b["tokens"])
            score = inter / min(len(a["tokens"]), len(b["tokens"]))
            if score > best_score:
                best, best_score = b, score
        if best and best_score >= token_thresh:
            record(a, best, f"4 amount + description ({best_score:.0%} overlap)")

    # Pass 5 - many side2 legs -> one side1 leg (batched settlement)
    for a in open_side(side1):
        if a["cmp"] is None:
            continue
        pool = [b for b in open_side(side2)
                if b["cmp"] is not None and (b["cmp"] > ZERO) == (a["cmp"] > ZERO)
                and abs((a["date"] - b["date"]).days) <= tol_days
                and abs(b["cmp"]) <= abs(a["cmp"])]
        pool.sort(key=lambda b: abs(b["cmp"]), reverse=True)
        pool = pool[:subset_cap]
        found = None
        for size in range(2, min(len(pool), 6) + 1):
            for combo in itertools.combinations(pool, size):
                if amounts_close(sum((c["cmp"] for c in combo), ZERO), a["cmp"], tol_pct, tol_floor):
                    found = combo
                    break
            if found:
                break
        if found:
            rule = f"5 many-to-one ({len(found)} legs batched)"
            for c in found:
                record(a, c, rule)
            a["matched"] = True

    # Pass 6 - one side2 leg -> many side1 legs (split settlement)
    for b in open_side(side2):
        if b["cmp"] is None:
            continue
        pool = [a for a in open_side(side1)
                if a["cmp"] is not None and (a["cmp"] > ZERO) == (b["cmp"] > ZERO)
                and abs((a["date"] - b["date"]).days) <= tol_days
                and abs(a["cmp"]) <= abs(b["cmp"])]
        pool.sort(key=lambda a: abs(a["cmp"]), reverse=True)
        pool = pool[:subset_cap]
        found = None
        for size in range(2, min(len(pool), 6) + 1):
            for combo in itertools.combinations(pool, size):
                if amounts_close(sum((c["cmp"] for c in combo), ZERO), b["cmp"], tol_pct, tol_floor):
                    found = combo
                    break
            if found:
                break
        if found:
            rule = f"6 one-to-many ({len(found)} legs)"
            for c in found:
                record(c, b, rule)
            b["matched"] = True

    return pairs


# ------------------------------------------------------------------ classification / gate

def load_classifications(path: Path | None) -> dict[str, dict]:
    if not path or not path.exists():
        return {}
    out = {}
    with path.open(newline="", encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            k = (r.get("key") or "").strip()
            if not k:
                continue
            cl = (r.get("classification") or "").strip().lower()
            if cl and cl not in CLASSIFICATIONS:
                raise ValueError(
                    f"Unknown classification {cl!r} for {k}. Valid values: "
                    + ", ".join(CLASSIFICATIONS)
                )
            out[k] = {
                "classification": cl,
                "je_entity": (r.get("je_entity") or "").strip(),
                "je_required": (r.get("je_required") or "").strip(),
                "expected_resolution": (r.get("expected_resolution") or "").strip(),
                "notes": (r.get("notes") or "").strip(),
            }
    return out


PENNY = Decimal("0.01")


def suggest_cause(disc: dict, tol_days: int, asof: date) -> str:
    if disc["kind"] == "variance":
        return "fx_difference? - paired (same doc id/amount window), translated amounts differ"
    if disc["amount"] is None:
        return f"cannot classify - no FX rate for {disc['currency']}; supply rate first"
    age = (asof - disc["date"]).days
    if age <= tol_days:
        return "timing_in_transit? - recent; may clear next period"
    return "missing_entry? - no corresponding record found within tolerance"


#: classifications that represent a leg one entity has not booked yet but
#: will (or must be corrected to) — these are the ones that move the
#: *adjusted* elimination basis once confirmed. fx_difference and
#: unidentified never do: an fx_difference is translation noise within an
#: already-matched pair (route it to CTA / the booking entity's own
#: correction, not into this pack's elimination entry), and unidentified
#: is, by definition, not yet explained.
EXPECTED_TO_POST = {"timing_in_transit", "missing_entry", "data_error"}


def build_pair_rec(side1: list[dict], side2: list[dict], matched_pairs: list[dict],
                    cls: dict, closing1: Decimal, closing2: Decimal) -> dict:
    """Builds the list of discrepancies that, together, account for the full
    gap between side1's and side2's recorded positions (diff = closing1 -
    closing2). An unmatched leg contributes its own cmp (signed +/- depending
    on which side it sits on); a matched-but-inexact pair contributes its
    variance (a.cmp - b.cmp). Every discrepancy needs a classification before
    its contribution is excluded from the unexplained total — nothing is
    netted away just because it was "close enough" to match on.

    Also computes adj1/adj2: the position each side's books will show once
    every confirmed timing/missing/data-error item is posted — the number
    the proposed elimination entry is actually based on."""
    discrepancies = []
    adj1, adj2 = closing1, closing2
    for leg in side1:
        if not leg["matched"]:
            discrepancies.append({
                "key": leg["key"], "amount": leg["cmp"], "kind": "unmatched", "side": 1,
                "entity": leg["entity"], "counterparty": leg["counterparty"],
                "date": leg["date"], "doc_id": leg["doc_id"],
                "description": leg["description"], "currency": leg["currency"],
                "leg": leg,
            })
    for leg in side2:
        if not leg["matched"]:
            discrepancies.append({
                "key": leg["key"], "amount": (-leg["cmp"] if leg["cmp"] is not None else None),
                "kind": "unmatched", "side": 2,
                "entity": leg["entity"], "counterparty": leg["counterparty"],
                "date": leg["date"], "doc_id": leg["doc_id"],
                "description": leg["description"], "currency": leg["currency"],
                "leg": leg,
            })
    for p in matched_pairs:
        a, b = p["a"], p["b"]
        if a["cmp"] is None or b["cmp"] is None:
            continue  # amount unverifiable - already flagged globally (missing FX rate)
        variance = a["cmp"] - b["cmp"]
        if abs(variance) <= PENNY:
            continue  # pure rounding dust, not a plug - never requires classification
        discrepancies.append({
            "key": f"VAR-{a['key']}-{b['key']}", "amount": variance, "kind": "variance",
            "side": None,
            "entity": a["entity"], "counterparty": a["counterparty"],
            "date": max(a["date"], b["date"]), "doc_id": a["doc_id"] or b["doc_id"],
            "description": f"Matched on {p['rule']} but amounts differ "
                            f"({a['entity']} {a['cmp']:,.2f} vs {b['entity']} {b['cmp']:,.2f})",
            "currency": "", "leg": None,
        })

    unexplained = ZERO
    unclassified, unidentified_items, variance_items = [], [], []
    for d in discrepancies:
        c = cls.get(d["key"], {}).get("classification", "")
        amt = d["amount"] if d["amount"] is not None else ZERO
        if not c or d["amount"] is None:
            unclassified.append(d)
            unexplained += amt
            continue
        if c == "unidentified":
            unidentified_items.append(d)
            unexplained += amt
            continue
        # confirmed and explained - excluded from unexplained.
        if d["kind"] == "variance":
            variance_items.append(d)
        elif c in EXPECTED_TO_POST:
            leg_cmp = d["leg"]["cmp"]
            if d["side"] == 1:
                adj2 += leg_cmp   # B will post the mirror; it lands at A's cmp.
            else:
                adj1 += leg_cmp   # A will post the mirror; it lands at B's cmp.

    return {
        "discrepancies": discrepancies, "unclassified": unclassified,
        "unidentified_items": unidentified_items, "variance_items": variance_items,
        "unidentified": sum((d["amount"] for d in unidentified_items), ZERO),
        "difference": unexplained,
        "reconciled": not unclassified and not unidentified_items,
        "adj1": adj1, "adj2": adj2,
    }


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


def sheet_summary(wb, pack_complete: bool, pair_rows: list[dict], meta: dict):
    ws = wb.active
    ws.title = "Pair Summary"
    widths(ws, {1: 4})
    r = 1
    ws.cell(row=r, column=2, value="INTERCOMPANY RECONCILIATION").font = Font(bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1
    v = ws.cell(row=r, column=2,
                value="COMPLETE - every pair reconciled" if pack_complete else
                "INCOMPLETE - see Chart Status / Exceptions below")
    v.font = OK_FONT if pack_complete else BAD_FONT
    if not pack_complete:
        v.fill = BAD_FILL
    r += 2

    heads = ["Entity A", "Entity B", "Per A: B owes A (rpt ccy)",
             "Per B: B owes A (rpt ccy)", "Raw difference", "Unexplained difference",
             "Unidentified", "Discrepancy count", "Status"]
    hr = r
    for c, h in enumerate(heads, start=1):
        ws.cell(row=hr, column=c, value=h)
    hdr(ws, len(heads), row=hr)
    r += 1
    for row in pair_rows:
        cells = [row["a_name"], row["b_name"], float(row["closing1"]), float(row["closing2"]),
                  float(row["raw_diff"]), float(row["difference"]), float(row["unidentified"]),
                  row["discrepancy_count"], row["status"]]
        for c, val in enumerate(cells, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            if c in (3, 4, 5, 6, 7):
                cell.number_format = MONEY
        status_cell = ws.cell(row=r, column=9)
        if row["status"] == "RECONCILED":
            status_cell.font = OK_FONT
        elif row["status"] == "NOT RECEIVED":
            status_cell.font, status_cell.fill = BAD_FONT, WARN_FILL
        else:
            status_cell.font, status_cell.fill = BAD_FONT, BAD_FILL
        r += 1
    widths(ws, {1: 20, 2: 20, 3: 22, 4: 22, 5: 16, 6: 18, 7: 14, 8: 14, 9: 18})
    ws.freeze_panes = f"A{hr + 1}"
    ws.auto_filter.ref = f"A{hr}:I{max(r - 1, hr)}"


def sheet_unmatched(wb, all_discrepancies: list[dict], cls: dict, tol_days: int, asof: date):
    ws = wb.create_sheet("Mismatch Queue")
    heads = ["Key", "Kind", "Entity", "Counterparty", "Date", "Age (d)", "Doc ID",
             "Description", "Amount (reporting ccy)", "Suggested cause",
             "Classification", "JE entity", "JE req'd", "Expected resolution", "Notes"]
    ws.append(heads)
    hdr(ws, len(heads))
    for d in sorted(all_discrepancies, key=lambda x: (x["entity"], x["date"])):
        info = cls.get(d["key"], {})
        c = info.get("classification", "")
        ws.append([
            d["key"], d["kind"], d["entity"], d["counterparty"], d["date"],
            (asof - d["date"]).days, d["doc_id"], d["description"],
            float(d["amount"]) if d["amount"] is not None else "NO FX RATE",
            suggest_cause(d, tol_days, asof),
            c or "*** UNCLASSIFIED ***", info.get("je_entity", ""),
            info.get("je_required", ""), info.get("expected_resolution", ""),
            info.get("notes", ""),
        ])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[4].number_format = DATEF
        if isinstance(row[8].value, (int, float)):
            row[8].number_format = MONEY
        if row[10].value in ("*** UNCLASSIFIED ***", "unidentified"):
            row[10].font, row[10].fill = BAD_FONT, BAD_FILL
        if isinstance(row[5].value, int) and row[5].value > 90:
            row[5].fill = WARN_FILL
    if ws.max_row == 1:
        ws.cell(row=2, column=2, value="None - every leg and matched pair ties within rounding.")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:O{max(ws.max_row, 2)}"
    widths(ws, {1: 20, 2: 10, 3: 10, 4: 12, 5: 12, 6: 8, 7: 12, 8: 44,
                9: 16, 10: 36, 11: 20, 12: 20, 13: 10, 14: 22, 15: 32})


def sheet_matched(wb, all_pairs: list[dict]):
    ws = wb.create_sheet("Matched Detail")
    heads = ["Rule", "Entity A key", "A date", "A description", "A amount (rpt)",
             "Entity B key", "B date", "B description", "B amount (rpt)",
             "Day var", "Variance (rpt)"]
    ws.append(heads)
    hdr(ws, len(heads))
    for p in sorted(all_pairs, key=lambda x: x["rule"]):
        a, b = p["a"], p["b"]
        var = (a["cmp"] - b["cmp"]) if a["cmp"] is not None and b["cmp"] is not None else None
        ws.append([p["rule"], a["key"], a["date"], a["description"],
                   float(a["amount_rpt"]) if a["amount_rpt"] is not None else None,
                   b["key"], b["date"], b["description"],
                   float(b["amount_rpt"]) if b["amount_rpt"] is not None else None,
                   (a["date"] - b["date"]).days, float(var) if var is not None else None])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[2].number_format = DATEF
        row[6].number_format = DATEF
        row[4].number_format = MONEY
        row[8].number_format = MONEY
        if row[10].value:
            row[10].number_format = MONEY
            row[10].fill = WARN_FILL
        if isinstance(row[9].value, int) and abs(row[9].value) > 5:
            row[9].fill = WARN_FILL
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value="No matches.")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:K{max(ws.max_row, 2)}"
    widths(ws, {1: 26, 2: 9, 3: 11, 4: 34, 5: 14, 6: 9, 7: 11, 8: 34, 9: 14, 10: 8, 11: 14})


def sheet_eliminations(wb, pair_rows: list[dict]):
    ws = wb.create_sheet("Proposed Eliminations")
    heads = ["Entity A", "Entity B", "Elimination basis (rpt ccy)",
             "Dr - IC Payable at", "Cr - IC Receivable at", "Residual unresolved",
             "Status", "Notes"]
    ws.append(heads)
    hdr(ws, len(heads))
    n = 0
    for row in pair_rows:
        basis = row["elim_basis"]
        if basis == ZERO and row["residual"] == ZERO:
            continue
        n += 1
        if row["elim_direction"] == "a_owed":
            # A owes B: A holds the payable, B holds the receivable.
            dr_entity, cr_entity = row["a_name"], row["b_name"]
        elif row["elim_direction"] == "b_owed":
            # B owes A: B holds the payable, A holds the receivable.
            dr_entity, cr_entity = row["b_name"], row["a_name"]
        else:
            dr_entity = cr_entity = "(direction unclear - see notes)"
        note = row.get("elim_note", "")
        ws.append([row["a_name"], row["b_name"], float(basis), dr_entity, cr_entity,
                   float(row["residual"]), row["status"], note])
    for r in ws.iter_rows(min_row=2, max_row=ws.max_row):
        r[2].number_format = MONEY
        r[5].number_format = MONEY
        if r[5].value:
            r[5].font, r[5].fill = BAD_FONT, WARN_FILL
    if n == 0:
        ws.cell(row=2, column=1, value="No intercompany balances to eliminate.")
    widths(ws, {1: 20, 2: 20, 3: 20, 4: 26, 5: 26, 6: 16, 7: 16, 8: 50})


def sheet_memo(wb, pack_complete: bool, pair_rows: list[dict], meta: dict, flags: list[str]):
    ws = wb.create_sheet("Exceptions & Memo")
    widths(ws, {1: 4, 2: 100})
    r = 1

    def para(text, bold=False, size=11):
        nonlocal r
        c = ws.cell(row=r, column=2, value=text)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    para("INTERCOMPANY RECONCILIATION MEMO", bold=True, size=14)
    r += 1
    para(f"Group: {meta.get('Client', '')}    Period: {meta.get('Period', '')}", bold=True)
    para(f"Entity pairs: {len(pair_rows)}   "
         f"Reconciled: {sum(1 for p in pair_rows if p['status'] == 'RECONCILED')}   "
         f"Exceptions: {sum(1 for p in pair_rows if p['status'] not in ('RECONCILED',))}")
    para("Status: " + ("COMPLETE" if pack_complete else "INCOMPLETE - see below"), bold=True)
    r += 1

    if flags:
        para("EXCEPTIONS AND ESCALATIONS", bold=True, size=12)
        for f in flags:
            para("- " + f)
        r += 1

    para("REMINDER", bold=True, size=12)
    para("Never plug. An unmatched intercompany leg with no confirmed classification "
         "stays visible as unexplained; it is never absorbed into the elimination entry. "
         "A pair that ties with no confirmed classifications behind it is coincidence, "
         "not proof.", bold=True)


# ------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                  formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--entities", required=True)
    ap.add_argument("--fx-rates")
    ap.add_argument("--reporting-currency", required=True)
    ap.add_argument("--transactions")
    ap.add_argument("--balances")
    ap.add_argument("--classifications")
    ap.add_argument("--expected-pairs")
    ap.add_argument("--as-of", help="period end YYYY-MM-DD (default: latest row date)")
    ap.add_argument("--date-tolerance", type=int, default=5)
    ap.add_argument("--fx-tolerance-pct", type=float, default=0.5,
                     help="percent tolerance on reporting-currency amount match (default 0.5)")
    ap.add_argument("--fx-tolerance-floor", type=float, default=1.00,
                     help="minimum absolute tolerance in reporting currency (default 1.00)")
    ap.add_argument("--token-threshold", type=float, default=0.6)
    ap.add_argument("--subset-cap", type=int, default=14)
    ap.add_argument("--client", default="")
    ap.add_argument("--period", default="")
    ap.add_argument("--out", required=True)
    ap.add_argument("--unmatched-out", default="ic_unmatched.csv")
    args = ap.parse_args()

    if not args.transactions and not args.balances:
        sys.exit("Supply --transactions, --balances, or both.")

    entities = load_entities(Path(args.entities))
    reporting_ccy = args.reporting_currency.strip().upper()
    fx = load_fx(Path(args.fx_rates) if args.fx_rates else None, entities, reporting_ccy)

    txn_legs = load_legs(Path(args.transactions) if args.transactions else None,
                          entities, fx, "subledger")
    bal_legs = load_legs(Path(args.balances) if args.balances else None,
                          entities, fx, "balance")

    flags: list[str] = []

    # integrity: where both a subledger and a stated balance exist for the
    # same (entity, counterparty), the subledger must sum to the balance.
    txn_sum = defaultdict(lambda: ZERO)
    for leg in txn_legs:
        txn_sum[(leg["entity"], leg["counterparty"])] += leg["amount"]
    bal_by_key = {(leg["entity"], leg["counterparty"]): leg for leg in bal_legs}
    for key, total in txn_sum.items():
        bal = bal_by_key.get(key)
        if bal and bal["amount"] != total:
            flags.append(
                f"{entities[key[0]]['name']}'s subledger vs {entities[key[1]]['name']} "
                f"sums to {total:,.2f} {entities[key[0]]['currency']} but the stated "
                f"balance is {bal['amount']:,.2f}. The subledger is incomplete or the "
                f"stated balance is wrong - resolve before relying on this reconciliation."
            )

    # Working legs per (entity, counterparty): subledger detail when present,
    # else the single stated balance stands in for the whole position.
    legs_by_key: dict[tuple[str, str], list[dict]] = defaultdict(list)
    has_txn = {(e, c) for e, c in txn_sum}
    for leg in txn_legs:
        legs_by_key[(leg["entity"], leg["counterparty"])].append(leg)
    for leg in bal_legs:
        k = (leg["entity"], leg["counterparty"])
        if k not in has_txn:
            legs_by_key[k].append(leg)

    all_pair_ids = sorted({tuple(sorted((e, c))) for (e, c) in legs_by_key})

    if args.expected_pairs and Path(args.expected_pairs).exists():
        with Path(args.expected_pairs).open(newline="", encoding="utf-8-sig") as fh:
            for r in csv.DictReader(fh):
                ea, eb = (r.get("entity_a") or "").strip(), (r.get("entity_b") or "").strip()
                if ea and eb:
                    all_pair_ids.append(tuple(sorted((ea, eb))))
    all_pair_ids = sorted(set(all_pair_ids))

    cls = load_classifications(Path(args.classifications) if args.classifications else None)

    everything = txn_legs + bal_legs
    asof = pdate(args.as_of) if args.as_of else (
        max((leg["date"] for leg in everything), default=date.today()))

    tol_pct = Decimal(str(args.fx_tolerance_pct)) / 100
    tol_floor = Decimal(str(args.fx_tolerance_floor))

    pair_rows, all_matched_pairs, all_discrepancies = [], [], []

    for a, b in all_pair_ids:
        side1 = legs_by_key.get((a, b), [])
        side2_raw = legs_by_key.get((b, a), [])
        if not side1 and not side2_raw:
            pair_rows.append({
                "a_name": entities.get(a, {}).get("name", a),
                "b_name": entities.get(b, {}).get("name", b),
                "closing1": ZERO, "closing2": ZERO, "raw_diff": ZERO, "difference": ZERO,
                "unidentified": ZERO, "discrepancy_count": 0, "status": "NOT RECEIVED",
                "elim_basis": ZERO, "elim_direction": "", "residual": ZERO,
                "elim_note": "No intercompany activity submitted by either entity.",
            })
            continue

        for leg in side1:
            leg["cmp"] = leg["amount_rpt"]
        side2 = [dict(leg) for leg in side2_raw]
        for leg in side2:
            leg["cmp"] = -leg["amount_rpt"] if leg["amount_rpt"] is not None else None

        # closing1/closing2 are both expressed in the SAME frame after the
        # mirror: "amount B owes A," per each entity's own books. A tie
        # means closing1 == closing2, exactly like bank vs. book in
        # bank-rec-to-gl.
        closing1 = sum((leg["cmp"] for leg in side1 if leg["cmp"] is not None), ZERO)
        closing2 = sum((leg["cmp"] for leg in side2 if leg["cmp"] is not None), ZERO)

        matched_pairs = match_pair(
            side1, side2, args.date_tolerance, tol_pct, tol_floor,
            args.token_threshold, args.subset_cap)
        all_matched_pairs.extend(matched_pairs)

        # write back match state onto the *original* side2 legs, since side2
        # is a working copy (cmp differs in sign from the stored legs).
        by_key = {leg["key"]: leg for leg in side2_raw}
        for leg in side2:
            orig = by_key[leg["key"]]
            orig["matched"], orig["match_rule"] = leg["matched"], leg["match_rule"]

        rec = build_pair_rec(side1, side2, matched_pairs, cls, closing1, closing2)
        all_discrepancies.extend(rec["discrepancies"])

        raw_diff = closing1 - closing2

        if rec["reconciled"]:
            status = "RECONCILED"
        elif rec["unidentified_items"]:
            status = "UNIDENTIFIED"
        elif rec["unclassified"]:
            status = "UNCLASSIFIED ITEMS"
        else:
            status = "EXCEPTION"

        # Elimination is based on the adjusted (post-expected-posting)
        # position, not the raw closing balances — otherwise a pair with a
        # one-sided missing entry looks like it has "no basis" to eliminate
        # even once fully explained.
        adj1, adj2 = rec["adj1"], rec["adj2"]
        residual = abs(rec["difference"])
        if residual == ZERO:
            elim_basis = abs(adj1)
            if adj1 > ZERO:
                elim_direction = "b_owed"
            elif adj1 < ZERO:
                elim_direction = "a_owed"
            else:
                elim_direction = ""
        else:
            # not fully explained yet - eliminate only the smaller, proven
            # balance; the gap stays visible, never folded into the entry.
            elim_basis = min(abs(adj1), abs(adj2))
            if adj1 > ZERO and adj2 > ZERO:
                elim_direction = "b_owed"
            elif adj1 < ZERO and adj2 < ZERO:
                elim_direction = "a_owed"
            else:
                elim_direction = ""
        elim_note = ""
        if not elim_direction and elim_basis != ZERO:
            elim_note = "Net positions disagree in sign - resolve exceptions before eliminating."
        if residual != ZERO:
            note2 = (f"Residual {residual:,.2f} unresolved - do not eliminate until "
                     f"classified; see Mismatch Queue.")
            elim_note = (elim_note + " " + note2).strip()
        if rec["variance_items"]:
            var_total = sum((d["amount"] for d in rec["variance_items"]), ZERO)
            elim_note = (elim_note + f" Matched-item FX translation variance of "
                         f"{var_total:,.2f} classified and excluded from this entry - "
                         f"route to the booking entity's correction or CTA, as "
                         f"applicable.").strip()

        pair_rows.append({
            "a_name": entities[a]["name"], "b_name": entities[b]["name"],
            "closing1": closing1, "closing2": closing2, "raw_diff": raw_diff,
            "difference": rec["difference"], "unidentified": rec["unidentified"],
            "discrepancy_count": len(rec["discrepancies"]), "status": status,
            "elim_basis": elim_basis, "elim_direction": elim_direction,
            "residual": residual, "elim_note": elim_note,
        })

        if status != "RECONCILED":
            flags.append(
                f"{entities[a]['name']} vs {entities[b]['name']}: {status} - "
                f"{len(rec['discrepancies'])} discrepancy(ies), unexplained difference "
                f"{rec['difference']:,.2f} {reporting_ccy}."
            )
        for d in rec["unidentified_items"]:
            flags.append(f"{d['key']}: UNIDENTIFIED item ({d['amount']:,.2f} "
                          f"{reporting_ccy}) - escalate regardless of size.")

    # FX rate gaps are a hard-visible exception, not a silent 1:1 assumption.
    no_rate = [leg for leg in txn_legs + bal_legs if leg["amount_rpt"] is None]
    if no_rate:
        missing_ccy = sorted({leg["currency"] for leg in no_rate})
        flags.append(
            f"{len(no_rate)} leg(s) could not be converted to {reporting_ccy} - "
            f"no FX rate supplied for: {', '.join(missing_ccy)}."
        )

    pack_complete = all(p["status"] == "RECONCILED" for p in pair_rows)

    if all_discrepancies:
        with Path(args.unmatched_out).open("w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["key", "kind", "entity", "counterparty", "date", "doc_id",
                        "description", "amount_rpt", "suggested_cause", "classification",
                        "je_entity", "je_required", "expected_resolution", "notes"])
            for d in sorted(all_discrepancies, key=lambda x: (x["entity"], x["date"])):
                info = cls.get(d["key"], {})
                w.writerow([
                    d["key"], d["kind"], d["entity"], d["counterparty"], d["date"].isoformat(),
                    d["doc_id"], d["description"],
                    f"{d['amount']:.2f}" if d["amount"] is not None else "",
                    suggest_cause(d, args.date_tolerance, asof),
                    info.get("classification", ""), info.get("je_entity", ""),
                    info.get("je_required", ""), info.get("expected_resolution", ""),
                    info.get("notes", ""),
                ])

    meta = {
        "Client": args.client or "(not supplied)",
        "Period": args.period or "(not supplied)",
        "Reporting currency": reporting_ccy,
        "As-of date": asof.isoformat(),
        "Entities": len(entities),
        "Entity pairs": len(pair_rows),
        "Date tolerance (days)": args.date_tolerance,
        "FX amount tolerance": f"{args.fx_tolerance_pct}% (floor {args.fx_tolerance_floor:.2f})",
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    wb = Workbook()
    sheet_summary(wb, pack_complete, pair_rows, meta)
    sheet_unmatched(wb, all_discrepancies, cls, args.date_tolerance, asof)
    sheet_matched(wb, all_matched_pairs)
    sheet_eliminations(wb, pair_rows)
    sheet_memo(wb, pack_complete, pair_rows, meta, flags)
    wb.active = 0

    out = Path(args.out)
    if not pack_complete:
        out = out.with_name(out.stem + " [INCOMPLETE]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))

    # ---- console
    print("=" * 70)
    print("INTERCOMPANY RECONCILIATION")
    print("=" * 70)
    for row in pair_rows:
        print(f"  {row['a_name']:<20} vs {row['b_name']:<20} {row['status']:<18} "
              f"diff {row['difference']:>12,.2f} {reporting_ccy}")
    print()
    print("COMPLETE" if pack_complete else "INCOMPLETE - not every pair is reconciled")
    if flags:
        print("\nEXCEPTIONS / ESCALATIONS:")
        for f in flags:
            print(f"  ! {f}")
    print(f"\nWorkpaper: {out}")
    if all_discrepancies:
        print(f"Unmatched worklist: {args.unmatched_out}")
    return 0 if pack_complete else 1


if __name__ == "__main__":
    sys.exit(main())
