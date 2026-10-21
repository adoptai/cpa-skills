#!/usr/bin/env python3
"""
Match bank credits (deposits) to open AR invoices and build a cash
application workpaper with an exception queue for everything that does not
apply cleanly.

Model: a deposit is cash that landed in the bank; an invoice is an open
receivable on the books. Applying cash is the act of saying which invoice(s)
a deposit pays. The strongest evidence is an explicit remittance advice
telling you exactly which invoices a payment covers; next is an invoice
number embedded in the bank reference/description; next is an amount that
ties exactly (or in combination) to open invoice balances for an identified
customer. Anything short of that is not applied automatically - it is
surfaced in the Exception Queue with candidate invoices, and the preparer
decides. This script never guesses which invoice a short-pay belongs to and
never assumes a payer's identity from a fuzzy name match.

What this script does NOT do: decide whether a short-pay is a discount, a
disputed deduction, or an error, or whether an overpayment should sit as a
customer credit or be refunded. It proposes a hypothesis in the "suggested
cause" column; the preparer confirms the real disposition and feeds it back
with --manual-applications. The gate only clears on confirmed classifications
and "unidentified_payer" never clears it - see CLASSIFICATIONS below.

Usage:
    python3 cash_apply.py \
        --deposits deposits.csv --invoices invoices.csv \
        --remittances remittances.csv --aliases aliases.csv \
        --manual-applications manual_applications.csv \
        --control-total 184250.00 --as-of 2026-10-20 \
        --client "Acme Inc" --period "2026-10" \
        --out "Acme - 2026-10 Cash Application.xlsx"

deposits.csv (required) - one row per bank credit to apply:
    deposit_id, date, amount, reference, payer_name, description
      deposit_id  = unique id; auto-assigned D0001... if blank
      amount      = positive, the credit amount
      reference   = ACH trace / check number / wire ref, if present
      payer_name  = name printed on the bank side (often a lockbox name,
                    not the legal customer name)
      description = raw bank description text

invoices.csv (required) - the open AR subledger as of --as-of:
    customer_id, customer_name, invoice_number, invoice_date, due_date,
    original_amount, open_balance
      open_balance must be > 0; a zero/paid invoice in this extract is a
      data-quality note, not a blocking error.

remittances.csv (optional, the detail transcribed from remittance advices
or emails - this script does not parse PDFs; see SKILL.md Step 1):
    deposit_id, deposit_reference, customer_id, invoice_number,
    amount_applied, deduction_code, deduction_reason
      Link a line to a deposit via deposit_id (preferred) or, if blank, via
      deposit_reference matched exactly against deposits.reference.
      customer_id may be left blank; it is inferred from invoice_number.

aliases.csv (optional) - maps how a customer's name appears on the bank
side to the AR customer master. Exact, light-normalized match only; no
fuzzy guessing of identity:
    payer_alias, customer_id

manual-applications.csv (optional, the feedback loop after reviewing the
Exception Queue):
    deposit_id, invoice_number, amount_applied, classification, notes
      invoice_number and amount_applied may be blank for a classification
      that does not apply cash to a specific invoice (e.g. a pure refund).
      classification = short_pay_discount | short_pay_deduction |
                        short_pay_dispute | overpayment_credit_balance |
                        overpayment_refund_due | unidentified_payer |
                        data_error | misc_resolved

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

CLASSIFICATIONS = {
    "short_pay_discount":          "Short-pay - authorized early-payment/trade discount taken",
    "short_pay_deduction":         "Short-pay - deduction for damages, returns, or chargeback",
    "short_pay_dispute":           "Short-pay - customer disputes the invoice, unresolved",
    "overpayment_credit_balance":  "Overpayment - hold as a credit balance on the account",
    "overpayment_refund_due":      "Overpayment - refund owed to the customer",
    "unidentified_payer":          "UNIDENTIFIED - payer/customer could not be determined",
    "data_error":                  "Data error - wrong invoice, amount, or duplicate",
    "misc_resolved":               "Resolved by preparer judgment - see notes",
}

#: classifications that represent cash genuinely disposed of (even though it
#: never ties to a specific invoice dollar) - these clear the gate once
#: confirmed. unidentified_payer never does: saying "unidentified" states the
#: finding, it does not resolve it. See AGENTS.md rule 4.
RESOLVING = {
    "short_pay_discount", "short_pay_deduction", "short_pay_dispute",
    "overpayment_credit_balance", "overpayment_refund_due",
    "data_error", "misc_resolved",
}

PENNY = Decimal("0.01")


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


def norm_name(s: str) -> str:
    s = re.sub(r"[.,]", "", (s or "")).strip().lower()
    return " ".join(s.split())


def norm_num(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def amounts_close(a: Decimal, b: Decimal, tol: Decimal) -> bool:
    return abs(a - b) <= tol


def find_subset(pool: list[dict], target: Decimal, tol: Decimal, cap: int) -> tuple[dict, ...] | None:
    """Smallest-first combination search for a set of open invoices whose
    remaining balances sum exactly (within tol) to target. Capped so a large
    candidate pool cannot blow up combinatorially; a miss here is not a
    failure, it just means the amount goes to the Exception Queue instead of
    being guessed at."""
    cands = [c for c in pool if ZERO < c["remaining"] <= target + tol]
    cands.sort(key=lambda c: c["remaining"], reverse=True)
    cands = cands[:cap]
    for size in range(2, min(len(cands), 6) + 1):
        for combo in itertools.combinations(cands, size):
            if amounts_close(sum((c["remaining"] for c in combo), ZERO), target, tol):
                return combo
    return None


# ------------------------------------------------------------------ loading

def load_deposits(path: Path) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No rows in {path}")
    out = []
    seen_ids = set()
    for i, r in enumerate(rows, start=1):
        did = (r.get("deposit_id") or "").strip() or f"D{i:04d}"
        if did in seen_ids:
            raise ValueError(f"{path.name} row {i}: duplicate deposit_id {did!r}")
        seen_ids.add(did)
        amt = dec(r.get("amount"))
        if amt is None or amt <= ZERO:
            raise ValueError(f"{path.name} row {i} ({did}): amount must be a positive number.")
        out.append({
            "deposit_id": did,
            "date": pdate(r.get("date")),
            "amount": amt,
            "reference": (r.get("reference") or "").strip(),
            "payer_name": (r.get("payer_name") or "").strip(),
            "description": (r.get("description") or "").strip(),
            "applied": ZERO,
            "applications": [],
            "flags": [],
        })
    return out


def load_invoices(path: Path) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No rows in {path}")
    out = []
    flags = []
    seen = set()
    for i, r in enumerate(rows, start=1):
        cust = (r.get("customer_id") or "").strip()
        inv = (r.get("invoice_number") or "").strip()
        if not cust or not inv:
            raise ValueError(f"{path.name} row {i}: customer_id and invoice_number are required.")
        key = (cust, inv)
        if key in seen:
            raise ValueError(f"{path.name} row {i}: duplicate invoice {inv!r} for customer {cust!r}.")
        seen.add(key)
        bal = dec(r.get("open_balance"))
        if bal is None:
            raise ValueError(f"{path.name} row {i}: open_balance is required for {inv!r}.")
        if bal <= ZERO:
            flags.append(f"Invoice {inv} ({cust}) has open_balance {bal:,.2f} - "
                          f"not open; remove it from the extract or confirm the data.")
            continue
        out.append({
            "customer_id": cust,
            "customer_name": (r.get("customer_name") or cust).strip(),
            "invoice_number": inv,
            "invoice_date": pdate(r.get("invoice_date")) if r.get("invoice_date") else None,
            "due_date": pdate(r.get("due_date")) if r.get("due_date") else None,
            "original_amount": dec(r.get("original_amount")) or bal,
            "open_balance": bal,
            "remaining": bal,
            "applications": [],
        })
    return out, flags


def load_aliases(path: Path | None) -> dict[str, str]:
    if not path or not path.exists():
        return {}
    out = {}
    with path.open(newline="", encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            alias = norm_name(r.get("payer_alias") or "")
            cid = (r.get("customer_id") or "").strip()
            if alias and cid:
                out[alias] = cid
    return out


def load_remittances(path: Path | None, deposits_by_id: dict, deposits_by_ref: dict,
                      invoices_by_key: dict) -> tuple[list[dict], list[str]]:
    if not path or not path.exists():
        return [], []
    flags = []
    out = []
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    for i, r in enumerate(rows, start=1):
        did = (r.get("deposit_id") or "").strip()
        ref = (r.get("deposit_reference") or "").strip()
        resolved_id = did if did in deposits_by_id else None
        if resolved_id is None and ref:
            cands = deposits_by_ref.get(ref, [])
            if len(cands) == 1:
                resolved_id = cands[0]
            elif len(cands) > 1:
                flags.append(f"remittances.csv row {i}: deposit_reference {ref!r} matches "
                              f"{len(cands)} deposits - ambiguous, not auto-linked.")
        inv = (r.get("invoice_number") or "").strip()
        amt = dec(r.get("amount_applied"))
        if resolved_id is None:
            flags.append(f"remittances.csv row {i}: could not link to a deposit "
                          f"(deposit_id {did!r}, deposit_reference {ref!r}) - payment advice "
                          f"received but no matching bank credit found. Check for a timing "
                          f"difference or a missing deposit in the bank export.")
            continue
        if not inv or amt is None:
            flags.append(f"remittances.csv row {i}: missing invoice_number or amount_applied - skipped.")
            continue
        cid = (r.get("customer_id") or "").strip()
        if not cid:
            # infer from the invoice, scanning all customers for this invoice_number
            owners = [k[0] for k in invoices_by_key if k[1] == inv]
            if len(owners) == 1:
                cid = owners[0]
        out.append({
            "deposit_id": resolved_id, "customer_id": cid, "invoice_number": inv,
            "amount_applied": amt,
            "deduction_code": (r.get("deduction_code") or "").strip(),
            "deduction_reason": (r.get("deduction_reason") or "").strip(),
        })
    return out, flags


def load_manual(path: Path | None) -> list[dict]:
    if not path or not path.exists():
        return []
    out = []
    with path.open(newline="", encoding="utf-8-sig") as fh:
        for i, r in enumerate(csv.DictReader(fh), start=1):
            did = (r.get("deposit_id") or "").strip()
            if not did:
                continue
            cl = (r.get("classification") or "").strip().lower()
            if cl and cl not in CLASSIFICATIONS:
                raise ValueError(f"manual-applications.csv row {i}: unknown classification "
                                  f"{cl!r}. Valid values: {', '.join(CLASSIFICATIONS)}")
            out.append({
                "deposit_id": did,
                "invoice_number": (r.get("invoice_number") or "").strip(),
                "amount_applied": dec(r.get("amount_applied")),
                "classification": cl,
                "notes": (r.get("notes") or "").strip(),
            })
    return out


# ------------------------------------------------------------------ applying

def apply_amount(deposit: dict, invoice: dict, amount: Decimal, rule: str) -> None:
    deposit["applied"] += amount
    deposit["applications"].append({
        "invoice": invoice, "amount": amount, "rule": rule,
    })
    invoice["remaining"] -= amount
    invoice["applications"].append({"deposit": deposit, "amount": amount, "rule": rule})


def find_reference_invoice_numbers(text: str, pool: list[dict]) -> list[dict]:
    text_up = (text or "").upper()
    hits = []
    for inv in pool:
        num = inv["invoice_number"]
        if len(norm_num(num)) < 3:
            continue
        pattern = r"(?<![A-Z0-9])" + re.escape(num.upper()) + r"(?![A-Z0-9])"
        if re.search(pattern, text_up):
            hits.append(inv)
    return hits


def resolve_customers(deposit: dict, remit_rows: list[dict], aliases: dict,
                       invoices: list[dict], names_by_customer: dict) -> list[str]:
    if remit_rows:
        return sorted({r["customer_id"] for r in remit_rows if r["customer_id"]})
    payer_key = norm_name(deposit["payer_name"])
    alias_hit = aliases.get(payer_key)
    if alias_hit:
        return [alias_hit]
    # An exact (normalized) match against the AR customer master is not a
    # fuzzy guess - it's the same identity stated two ways - so it does not
    # require an aliases.csv entry. A near-miss still goes through aliases.
    name_hit = names_by_customer.get(payer_key)
    if name_hit:
        return [name_hit]
    text = f"{deposit['reference']} {deposit['description']}"
    hits = find_reference_invoice_numbers(text, invoices)
    owners = sorted({inv["customer_id"] for inv in hits})
    if len(owners) == 1:
        return owners
    return []


def process_deposit(deposit: dict, remit_rows: list[dict], aliases: dict,
                     invoices_by_customer: dict, invoices_by_key: dict,
                     names_by_customer: dict, tol: Decimal, subset_cap: int) -> list[dict]:
    """Returns the list of discrepancies left over for this deposit after
    every automatic pass has run. Structural data problems found along the
    way (a remittance line naming an invoice that doesn't exist, or applying
    more than an invoice has open) are recorded on deposit["flags"] instead -
    they are data to fix and re-run, not a leftover-cash disposition, and
    they must stay visible even if the deposit's dollars net to zero through
    some other pass. Folding them into the leftover amount would let a data
    error hide behind an unrelated match."""
    discrepancies = []

    # Phase 1 - remittance-driven application (strongest evidence).
    for r in remit_rows:
        inv = invoices_by_key.get((r["customer_id"], r["invoice_number"]))
        if inv is None:
            deposit["flags"].append(
                f"Remittance names invoice {r['invoice_number']!r} for customer "
                f"{r['customer_id']!r}, not found in the open AR extract.")
            continue
        amt = r["amount_applied"]
        if inv["remaining"] - amt < -tol:
            over = amt - inv["remaining"]
            deposit["flags"].append(
                f"Remittance applies {amt:,.2f} to invoice {inv['invoice_number']} which "
                f"only has {inv['remaining']:,.2f} open ({over:,.2f} over-applied).")
            amt = inv["remaining"]
        if amt > ZERO:
            apply_amount(deposit, inv, amt, "remittance detail")

    leftover = deposit["amount"] - deposit["applied"]
    if leftover <= tol:
        return discrepancies

    # Phase 2 - automatic amount-based matching for whatever remains.
    customer_ids = resolve_customers(deposit, remit_rows, aliases,
                                      [inv for invs in invoices_by_customer.values() for inv in invs],
                                      names_by_customer)
    if not customer_ids:
        discrepancies.append({
            "kind": "unidentified_payer", "deposit": deposit, "invoice": None,
            "amount": leftover,
            "detail": "Payer could not be identified - no exact customer-name match, alias, "
                      "remittance, or invoice number match in the reference/description.",
        })
        return discrepancies

    pool = [inv for cid in customer_ids for inv in invoices_by_customer.get(cid, [])
            if inv["remaining"] > tol]

    # Pass A - reference-embedded invoice number, unique and exact.
    ref_hits = [inv for inv in find_reference_invoice_numbers(
        f"{deposit['reference']} {deposit['description']}", pool)]
    if len(ref_hits) == 1 and amounts_close(ref_hits[0]["remaining"], leftover, tol):
        apply_amount(deposit, ref_hits[0], leftover, "reference-embedded invoice number")
        return discrepancies

    # Pass B - unique exact amount match.
    exact = [inv for inv in pool if amounts_close(inv["remaining"], leftover, tol)]
    if len(exact) == 1:
        apply_amount(deposit, exact[0], leftover, "exact amount match")
        return discrepancies
    if len(exact) > 1:
        discrepancies.append({
            "kind": "ambiguous_amount", "deposit": deposit, "invoice": None,
            "amount": leftover,
            "detail": "Amount ties exactly to " + ", ".join(
                f"{i['invoice_number']} ({i['remaining']:,.2f})" for i in exact) +
                " - more than one invoice matches; not auto-applied.",
            "candidates": exact,
        })
        return discrepancies

    # Pass C - lump sum across several open invoices.
    combo = find_subset(pool, leftover, tol, subset_cap)
    if combo:
        total = sum((c["remaining"] for c in combo), ZERO)
        for c in combo:
            share = c["remaining"]
            if c is combo[-1]:
                share += leftover - total  # absorb rounding dust on the last line
            apply_amount(deposit, c, share, "lump sum (subset match)")
        return discrepancies

    # No automatic rule applies - leave it for the preparer, with candidates.
    nearby = sorted(pool, key=lambda i: abs(i["remaining"] - leftover))[:8]
    total_open = sum((i["remaining"] for i in pool), ZERO)
    if not pool:
        if invoices_by_customer.get(customer_ids[0] if len(customer_ids) == 1 else "", []):
            cause = ("customer identified but every open invoice is already fully applied "
                     "this run - this leftover does not correspond to a specific invoice")
        else:
            cause = "customer identified but has no open invoices on this extract"
    elif leftover < min((i["remaining"] for i in pool), default=ZERO):
        cause = "short_pay? - less than any single open invoice for this customer"
    elif leftover > total_open:
        cause = "overpayment? - exceeds total open balance for this customer"
    else:
        cause = "ambiguous? - no exact or combination match among open invoices"
    discrepancies.append({
        "kind": "unapplied_remainder", "deposit": deposit, "invoice": None,
        "amount": leftover, "detail": cause, "candidates": nearby,
        "customer_ids": customer_ids,
    })
    return discrepancies


def merge_manual(discrepancies: list[dict], manual: list[dict], invoices_by_key: dict,
                  deposits_by_id: dict, tol: Decimal) -> tuple[dict[str, dict], list[str]]:
    """Applies confirmed manual applications and returns {key: classification
    info} for every discrepancy key, plus any flags (bad references)."""
    flags = []
    cls_by_key: dict[str, dict] = {}
    by_dep_inv = defaultdict(list)
    for d in discrepancies:
        key = d["key"]
        by_dep_inv[d["deposit"]["deposit_id"]].append(d)

    for m in manual:
        dep = deposits_by_id.get(m["deposit_id"])
        if dep is None:
            flags.append(f"manual-applications.csv: deposit_id {m['deposit_id']!r} not found.")
            continue
        # Apply cash if an invoice and amount were given.
        if m["invoice_number"] and m["amount_applied"]:
            inv = None
            for d in by_dep_inv.get(m["deposit_id"], []):
                if d["kind"] == "unapplied_remainder":
                    inv = next((c for c in d.get("candidates", [])
                                if c["invoice_number"] == m["invoice_number"]), None)
                    if inv is None:
                        inv = invoices_by_key.get((d.get("customer_ids", [""])[0], m["invoice_number"]))
            if inv is None:
                # fall back to a global lookup by invoice number alone.
                owners = [k for k in invoices_by_key if k[1] == m["invoice_number"]]
                if len(owners) == 1:
                    inv = invoices_by_key[owners[0]]
            if inv is None:
                flags.append(f"manual-applications.csv: invoice {m['invoice_number']!r} for "
                              f"deposit {m['deposit_id']} not found in open AR.")
            else:
                amt = m["amount_applied"]
                if inv["remaining"] - amt < -tol:
                    flags.append(f"manual-applications.csv: applying {amt:,.2f} to "
                                 f"{inv['invoice_number']} exceeds its open balance "
                                 f"({inv['remaining']:,.2f}) - not applied.")
                else:
                    apply_amount(dep, inv, amt, "manual (preparer confirmed)")
        # Record the classification against the whole deposit's leftover key.
        key = f"{m['deposit_id']}"
        if m["classification"]:
            cls_by_key[key] = {"classification": m["classification"], "notes": m["notes"]}
    return cls_by_key, flags


# ------------------------------------------------------------------ workbook

def hdr(ws, n, row=1):
    for c in range(1, n + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 28


def widths(ws, w):
    for c, v in w.items():
        ws.column_dimensions[get_column_letter(c)].width = v


def sheet_summary(wb, complete: bool, meta: dict, totals: dict, by_rule: dict):
    ws = wb.active
    ws.title = "Application Summary"
    widths(ws, {1: 4, 2: 30, 3: 20})
    r = 1
    ws.cell(row=r, column=2, value="CASH APPLICATION").font = Font(bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1
    v = ws.cell(row=r, column=2,
                value="COMPLETE - every deposit applied or resolved" if complete else
                "INCOMPLETE - see Exception Queue / Exceptions & Memo")
    v.font = OK_FONT if complete else BAD_FONT
    if not complete:
        v.fill = BAD_FILL
    r += 2

    rows = [
        ("Total deposits", totals["deposits"]),
        ("Applied to invoices", totals["applied"]),
        ("Resolved (classified, not applied to an invoice)", totals["resolved"]),
        ("Unidentified / unclassified (exceptions)", totals["exception"]),
    ]
    for label, val in rows:
        ws.cell(row=r, column=2, value=label)
        c = ws.cell(row=r, column=3, value=float(val))
        c.number_format = MONEY
        r += 1
    r += 1
    ws.cell(row=r, column=2, value="Applied by rule").font = Font(bold=True)
    r += 1
    for rule, amt in sorted(by_rule.items(), key=lambda kv: -kv[1]):
        ws.cell(row=r, column=2, value=rule)
        c = ws.cell(row=r, column=3, value=float(amt))
        c.number_format = MONEY
        r += 1


def sheet_population(wb, deposits: list[dict], control_total: Decimal | None):
    ws = wb.create_sheet("Population Proof")
    heads = ["Deposit ID", "Date", "Payer (bank side)", "Amount", "Applied",
             "Resolved (non-cash)", "Exception amount", "Status"]
    ws.append(heads)
    hdr(ws, len(heads))
    total_amt = total_applied = total_resolved = total_exc = ZERO
    for d in deposits:
        applied = d["applied"]
        resolved = d.get("resolved_amount", ZERO)
        exc = d["amount"] - applied - resolved
        total_amt += d["amount"]; total_applied += applied
        total_resolved += resolved; total_exc += exc
        status = d.get("status", "")
        ws.append([d["deposit_id"], d["date"], d["payer_name"], float(d["amount"]),
                   float(applied), float(resolved), float(exc), status])
        row = ws[ws.max_row]
        row[1].number_format = DATEF
        for i in (3, 4, 5, 6):
            row[i].number_format = MONEY
        if status in ("UNCLASSIFIED", "UNIDENTIFIED"):
            row[7].font, row[7].fill = BAD_FONT, BAD_FILL
        elif status == "RESOLVED":
            row[7].fill = WARN_FILL
        elif status == "APPLIED":
            row[7].font = OK_FONT
    r = ws.max_row + 1
    ws.cell(row=r, column=1, value="TOTAL").font = Font(bold=True)
    for col, val in ((4, total_amt), (5, total_applied), (6, total_resolved), (7, total_exc)):
        c = ws.cell(row=r, column=col, value=float(val))
        c.number_format = MONEY
        c.font = Font(bold=True)
    if control_total is not None:
        r += 2
        ws.cell(row=r, column=1, value="Control total (period bank credits)").font = Font(bold=True)
        ws.cell(row=r, column=4, value=float(control_total)).number_format = MONEY
        r += 1
        diff = total_amt - control_total
        c = ws.cell(row=r, column=1, value="Difference (deposits population vs. control)")
        c.font = Font(bold=True)
        cv = ws.cell(row=r, column=4, value=float(diff))
        cv.number_format = MONEY
        if abs(diff) > PENNY:
            cv.font, cv.fill = BAD_FONT, BAD_FILL
        else:
            cv.font = OK_FONT
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{max(ws.max_row, 2)}"
    widths(ws, {1: 14, 2: 12, 3: 26, 4: 14, 5: 14, 6: 16, 7: 16, 8: 16})


def sheet_applied(wb, deposits: list[dict]):
    ws = wb.create_sheet("Applied Detail")
    heads = ["Deposit ID", "Date", "Customer", "Invoice", "Invoice date",
             "Amount applied", "Rule", "Invoice remaining after"]
    ws.append(heads)
    hdr(ws, len(heads))
    for d in deposits:
        for a in d["applications"]:
            inv = a["invoice"]
            ws.append([d["deposit_id"], d["date"], inv["customer_name"], inv["invoice_number"],
                       inv["invoice_date"], float(a["amount"]), a["rule"], float(inv["remaining"])])
            row = ws[ws.max_row]
            row[1].number_format = DATEF
            if row[4].value:
                row[4].number_format = DATEF
            row[5].number_format = MONEY
            row[7].number_format = MONEY
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value="No applications.")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{max(ws.max_row, 2)}"
    widths(ws, {1: 14, 2: 12, 3: 24, 4: 16, 5: 12, 6: 14, 7: 26, 8: 18})


def sheet_exceptions(wb, discrepancies: list[dict], cls_by_key: dict):
    ws = wb.create_sheet("Exception Queue")
    heads = ["Deposit ID", "Date", "Payer (bank side)", "Kind", "Amount",
             "Candidate invoices", "Suggested cause", "Classification", "Notes"]
    ws.append(heads)
    hdr(ws, len(heads))
    for d in sorted(discrepancies, key=lambda x: (x["deposit"]["deposit_id"],)):
        dep = d["deposit"]
        info = cls_by_key.get(dep["deposit_id"], {})
        cl = info.get("classification", "")
        cands = ""
        if d.get("candidates"):
            cands = ", ".join(f"{c['invoice_number']} ({c['remaining']:,.2f})"
                               for c in d["candidates"])
        ws.append([dep["deposit_id"], dep["date"], dep["payer_name"], d["kind"],
                   float(d["amount"]), cands, d.get("detail", ""),
                   cl or "*** UNCLASSIFIED ***", info.get("notes", "")])
        row = ws[ws.max_row]
        row[1].number_format = DATEF
        row[4].number_format = MONEY
        if row[7].value in ("*** UNCLASSIFIED ***", "unidentified_payer"):
            row[7].font, row[7].fill = BAD_FONT, BAD_FILL
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value="None - every deposit applied in full.")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{max(ws.max_row, 2)}"
    widths(ws, {1: 14, 2: 12, 3: 24, 4: 20, 5: 14, 6: 34, 7: 44, 8: 22, 9: 28})


def sheet_open_ar(wb, invoices: list[dict]):
    ws = wb.create_sheet("Remaining Open AR")
    heads = ["Customer", "Invoice", "Invoice date", "Due date", "Original amount",
             "Open balance (start)", "Applied this run", "Open balance (end)"]
    ws.append(heads)
    hdr(ws, len(heads))
    for inv in sorted(invoices, key=lambda i: (i["customer_name"], i["invoice_number"])):
        applied = inv["open_balance"] - inv["remaining"]
        ws.append([inv["customer_name"], inv["invoice_number"], inv["invoice_date"],
                   inv["due_date"], float(inv["original_amount"]), float(inv["open_balance"]),
                   float(applied), float(inv["remaining"])])
        row = ws[ws.max_row]
        if row[2].value:
            row[2].number_format = DATEF
        if row[3].value:
            row[3].number_format = DATEF
        for i in (4, 5, 6, 7):
            row[i].number_format = MONEY
        if inv["remaining"] <= PENNY:
            row[7].font = OK_FONT
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{max(ws.max_row, 2)}"
    widths(ws, {1: 24, 2: 16, 3: 12, 4: 12, 5: 14, 6: 16, 7: 14, 8: 16})


def sheet_memo(wb, complete: bool, meta: dict, totals: dict, flags: list[str]):
    ws = wb.create_sheet("Exceptions & Memo")
    widths(ws, {1: 4, 2: 100})
    r = 1

    def para(text, bold=False, size=11):
        nonlocal r
        c = ws.cell(row=r, column=2, value=text)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    para("CASH APPLICATION MEMO", bold=True, size=14)
    r += 1
    para(f"Client: {meta.get('Client', '')}    Period: {meta.get('Period', '')}", bold=True)
    para(f"Deposits: total {totals['deposits']:,.2f}   applied {totals['applied']:,.2f}   "
         f"resolved {totals['resolved']:,.2f}   exceptions {totals['exception']:,.2f}")
    para("Status: " + ("COMPLETE" if complete else "INCOMPLETE - see Exception Queue"), bold=True)
    r += 1

    if flags:
        para("EXCEPTIONS AND ESCALATIONS", bold=True, size=12)
        for f in flags:
            para("- " + f)
        r += 1

    para("REMINDER", bold=True, size=12)
    para("Never plug. Cash that does not tie exactly to an invoice, or a combination of "
         "invoices, is never force-applied to the nearest one. It stays visible in the "
         "Exception Queue until the preparer confirms what it is.", bold=True)


# ------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                  formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--deposits", required=True)
    ap.add_argument("--invoices", required=True)
    ap.add_argument("--remittances")
    ap.add_argument("--aliases")
    ap.add_argument("--manual-applications")
    ap.add_argument("--control-total", type=str, default=None)
    ap.add_argument("--as-of", help="period end YYYY-MM-DD (default: latest deposit date)")
    ap.add_argument("--rounding-tolerance", type=str, default="0.01",
                     help="max cents treated as rounding dust, not a plug (default 0.01)")
    ap.add_argument("--subset-cap", type=int, default=12)
    ap.add_argument("--client", default="")
    ap.add_argument("--period", default="")
    ap.add_argument("--out", required=True)
    ap.add_argument("--exceptions-out", default="cash_app_exceptions.csv")
    args = ap.parse_args()

    tol = dec(args.rounding_tolerance) or ZERO
    control_total = dec(args.control_total) if args.control_total else None

    deposits = load_deposits(Path(args.deposits))
    invoices, inv_flags = load_invoices(Path(args.invoices))
    aliases = load_aliases(Path(args.aliases) if args.aliases else None)

    deposits_by_id = {d["deposit_id"]: d for d in deposits}
    deposits_by_ref = defaultdict(list)
    for d in deposits:
        if d["reference"]:
            deposits_by_ref[d["reference"]].append(d["deposit_id"])

    invoices_by_key = {(i["customer_id"], i["invoice_number"]): i for i in invoices}
    invoices_by_customer = defaultdict(list)
    names_multi = defaultdict(set)
    for i in invoices:
        invoices_by_customer[i["customer_id"]].append(i)
        names_multi[norm_name(i["customer_name"])].add(i["customer_id"])
    # Only a name that maps to exactly one customer_id is safe to treat as an
    # implicit alias - two customers sharing a display name must go through
    # an explicit aliases.csv entry instead of being resolved here.
    names_by_customer = {name: next(iter(ids)) for name, ids in names_multi.items() if len(ids) == 1}

    remit_rows, remit_flags = load_remittances(
        Path(args.remittances) if args.remittances else None,
        deposits_by_id, deposits_by_ref, invoices_by_key)
    remit_by_deposit = defaultdict(list)
    for r in remit_rows:
        remit_by_deposit[r["deposit_id"]].append(r)

    manual = load_manual(Path(args.manual_applications) if args.manual_applications else None)

    all_discrepancies = []
    for i, key in enumerate(invoices_by_key):
        invoices_by_key[key]["key"] = key  # harmless no-op guard for clarity
    for d in deposits:
        discs = process_deposit(d, remit_by_deposit.get(d["deposit_id"], []), aliases,
                                 invoices_by_customer, invoices_by_key, names_by_customer,
                                 tol, args.subset_cap)
        for disc in discs:
            disc["key"] = d["deposit_id"]
        all_discrepancies.extend(discs)

    cls_by_key, manual_flags = merge_manual(all_discrepancies, manual, invoices_by_key,
                                             deposits_by_id, tol)

    # Re-derive discrepancy amounts after manual applications may have
    # reduced a deposit's leftover - recompute per deposit leftover and keep
    # only what is still genuinely outstanding, carrying classification along.
    remaining_discs = []
    flags = list(inv_flags) + list(remit_flags) + list(manual_flags)
    for d in deposits:
        for f in d["flags"]:
            flags.append(f"{d['deposit_id']}: {f}")
        leftover = d["amount"] - d["applied"]
        info = cls_by_key.get(d["deposit_id"])
        d["resolved_amount"] = ZERO
        if leftover <= tol:
            d["status"] = "APPLIED"
            continue
        if info and info["classification"]:
            d["resolved_amount"] = leftover
            d["status"] = "UNIDENTIFIED" if info["classification"] == "unidentified_payer" else "RESOLVED"
        else:
            d["status"] = "UNCLASSIFIED"
        # keep one representative discrepancy row per deposit for the queue,
        # using the richest one already computed (first match for this id).
        own = [x for x in all_discrepancies if x["deposit"]["deposit_id"] == d["deposit_id"]]
        if own:
            # amount on the row should reflect the CURRENT leftover, not the
            # pre-manual-application amount.
            own[-1]["amount"] = leftover
            remaining_discs.append(own[-1])
        else:
            remaining_discs.append({
                "kind": "unapplied_remainder", "deposit": d, "invoice": None,
                "amount": leftover, "detail": "Leftover after manual application.",
                "key": d["deposit_id"],
            })

    total_deposits = sum((d["amount"] for d in deposits), ZERO)
    total_applied = sum((d["applied"] for d in deposits), ZERO)
    total_resolved = sum((d["resolved_amount"] for d in deposits), ZERO)
    total_exception = total_deposits - total_applied - total_resolved

    by_rule = defaultdict(lambda: ZERO)
    for d in deposits:
        for a in d["applications"]:
            by_rule[a["rule"]] += a["amount"]

    if control_total is not None and abs(total_deposits - control_total) > PENNY:
        flags.append(f"Deposits population ({total_deposits:,.2f}) does not tie to the "
                     f"supplied control total ({control_total:,.2f}) - the deposits extract "
                     f"may be incomplete or filtered.")

    blocking = any(d["status"] in ("UNCLASSIFIED", "UNIDENTIFIED") for d in deposits)
    data_errors = any(d["flags"] for d in deposits) or bool(remit_flags) or bool(manual_flags) or bool(inv_flags)
    control_ok = control_total is None or abs(total_deposits - control_total) <= PENNY
    complete = (not blocking) and control_ok and not data_errors

    meta = {
        "Client": args.client or "(not supplied)",
        "Period": args.period or "(not supplied)",
        "As-of date": (pdate(args.as_of) if args.as_of
                       else max((d["date"] for d in deposits), default=date.today())).isoformat(),
        "Deposits": len(deposits),
        "Open invoices": len(invoices),
        "Rounding tolerance": f"{tol}",
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    totals = {"deposits": total_deposits, "applied": total_applied,
              "resolved": total_resolved, "exception": total_exception}

    wb = Workbook()
    sheet_summary(wb, complete, meta, totals, by_rule)
    sheet_population(wb, deposits, control_total)
    sheet_applied(wb, deposits)
    sheet_exceptions(wb, remaining_discs, cls_by_key)
    sheet_open_ar(wb, invoices)
    sheet_memo(wb, complete, meta, totals, flags)
    wb.active = 0

    out = Path(args.out)
    if not complete:
        out = out.with_name(out.stem + " [INCOMPLETE]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))

    if remaining_discs:
        with Path(args.exceptions_out).open("w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["deposit_id", "date", "payer_name", "kind", "amount",
                        "candidate_invoices", "suggested_cause", "classification", "notes"])
            for d in sorted(remaining_discs, key=lambda x: x["deposit"]["deposit_id"]):
                dep = d["deposit"]
                info = cls_by_key.get(dep["deposit_id"], {})
                cands = ";".join(c["invoice_number"] for c in d.get("candidates", []))
                w.writerow([dep["deposit_id"], dep["date"].isoformat(), dep["payer_name"],
                            d["kind"], f"{d['amount']:.2f}", cands, d.get("detail", ""),
                            info.get("classification", ""), info.get("notes", "")])

    print("=" * 70)
    print("CASH APPLICATION")
    print("=" * 70)
    print(f"  Deposits:   {total_deposits:>14,.2f}")
    print(f"  Applied:    {total_applied:>14,.2f}")
    print(f"  Resolved:   {total_resolved:>14,.2f}")
    print(f"  Exceptions: {total_exception:>14,.2f}")
    print()
    print("COMPLETE" if complete else "INCOMPLETE - see Exception Queue / Exceptions & Memo")
    if flags:
        print("\nEXCEPTIONS / ESCALATIONS:")
        for f in flags:
            print(f"  ! {f}")
    print(f"\nWorkpaper: {out}")
    if remaining_discs:
        print(f"Exception worklist: {args.exceptions_out}")
    return 0 if complete else 1


if __name__ == "__main__":
    sys.exit(main())
