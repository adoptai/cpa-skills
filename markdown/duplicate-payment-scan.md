---
name: duplicate-payment-scan
description: Scans the full AP invoice and payment population — never a sample — for duplicate-payment candidates across four risk-ranked tiers (exact duplicate invoice numbers, same vendor/amount near in time, vendor-name variants, and reused payment references), with every finding traceable to the source rows it's based on. Use this whenever the user mentions duplicate payments, double payments, duplicate invoices, paying a vendor twice, AP duplicate scan, finding overpayments, or recovering cash from AP. Also trigger on "scan AP for duplicates," "did we pay this vendor twice," "find duplicate invoices," "duplicate payment risk," or when 12+ months of AP invoice and payment registers are provided for review. Reuses the full-population scan pattern from journal-entry-anomaly-scan, pointed at AP instead of the GL. Runs fully local — no vendor or payment data leaves the machine.
---

# Duplicate Payment Scan

Findings here are self-evidently valuable — every confirmed duplicate is recoverable cash. The
discipline is in proving each finding with the actual matched rows rather than asserting
"duplicate" and moving on.

## The gate

The scan refuses to produce a clean workbook when:

1. **Coverage cannot be proven.** Every row read must end up either scanned or in the
   unscanned/unparsed list with a stated reason. Scanned count plus unscanned count must equal
   total rows read — checked explicitly, not assumed.
2. **A `record_id` is duplicated** in the source file. Row identity must be unambiguous before
   any finding can reliably point back to a specific transaction.

A row with an unparseable amount or date is not itself a gate failure — it is excluded from
matching and listed, with its reason, on the Unscanned/Unparsed Rows tab, so the scope limitation
is visible rather than silently absorbed into a clean-looking result.

## Inputs

**AP invoice and payment register (12+ months)** — one file, one row per invoice or payment
event: `record_id, vendor, amount, invoice_date, invoice_no, payment_date, payment_ref`.
`invoice_no`, `payment_date`, and `payment_ref` may be blank on any given row.

## Step 1 — Run the scan

```bash
python3 scripts/duplicate_payment_scan.py --register ap_register.csv --window-days 5 \
  --client "Northwind Systems Inc" \
  --out "Northwind - Duplicate Payment Scan.xlsx"
```

`--window-days` has no built-in default — state the proximity window you actually want checked
(a tighter window for weekly-run AP, a wider one for monthly-batch AP) rather than relying on a
number this skill would otherwise have to invent.

## Step 2 — What each tier catches, and why it's ranked there

| Tier | Match | Why it's ranked there |
|---|---|---|
| **1** | Exact vendor + amount + `invoice_no` duplicated | The same invoice number paid twice for the same vendor and amount is close to unambiguous — the highest-confidence finding in the set. |
| **2** | Exact vendor + amount + date within the window | Same vendor, same amount, close in time, different invoice numbers — classic re-submission or re-entry duplicate. |
| **3** | Same amount + date within the window, vendor name a normalized variant of another (e.g. "Acme Supply Co" vs. "Acme Supply") | Catches duplicates hidden by vendor-master inconsistency — often the hardest tier to find by eye and the one worth the normalization log. |
| **4** | Same `payment_ref` reused across different invoice numbers | A check or EFT reference applied to more than one invoice is either a batching convention or a posting error — always worth a look. |

Each finding carries the matched records' row references and fields, never a bare assertion.

## Deliver

**Workbook tabs:**

1. **Risk Summary** — population rows read, scanned vs. unscanned, population dollar total, and
   finding counts and dollars at risk by tier.
2. **Findings** — every match with tier, basis, vendor(s), amount, matched record references, and
   dollars at risk.
3. **Unscanned-Unparsed Rows** — every row excluded from scanning, with the reason.
4. **Vendor Normalization Log** — every raw vendor string mapped to its normalized form, so a
   Tier 3 match can be audited back to the exact spelling variants involved.

**Then, in chat:** the population scanned, total dollars at risk by tier, and the highest-tier
findings first. State plainly if the unscanned count is material — a scan with a large unparsed
population proves less than it appears to.

## What to escalate

- **Any Tier 1 finding** — confirm with the bank or payment processor whether both payments
  actually cleared before assuming it's a false positive.
- **A payment reference reused across invoices to different vendors**, if it occurs — a
  reference collision this specific is unusual and worth confirming isn't a data-entry artifact
  versus a real misapplied payment.
- **A vendor-name variant cluster (Tier 3) spanning many distinct normalized groups** — may
  indicate the vendor master itself needs cleanup, a separate project from the duplicates found
  this pass.
- **Concentration of findings in one AP clerk's or approver's batch**, if that data is
  available separately — the scan does not have approver data and cannot make this link itself,
  but a reviewer with access to it should check.

Report each as a fact with the record references. Do not characterize intent — the scan reports
pattern and evidence, never motive.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.
Vendor and payment data never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*

---

# Appendix — bundled files

If you installed the `.skill` package these files are already in place and you can ignore this appendix. If you copied the skill as text, create the files below at the paths shown, alongside your `SKILL.md`. The skill will not run without them.

## `scripts/duplicate_payment_scan.py`

```python
#!/usr/bin/env python3
"""
Scan the FULL AP invoice/payment population for duplicate-payment candidates,
risk-ranked across four tiers, with every finding traceable to source rows.

Gates:
  * every row read must end up either scanned or in the unscanned/unparsed
    list with a stated reason - parsed_count + unparsed_count must equal
    total rows read, and that identity is printed and checked
  * a duplicate record_id makes the population ambiguous and blocks the
    workbook (a scan cannot trust row identity it cannot establish)
  * every finding must carry the record refs and fields it is based on -
    never a bare "duplicate" assertion
  * the population total ($ and row count) is always stated on the Risk
    Summary tab so scope is explicit

Risk tiers (1 = highest risk):
  1. exact vendor + amount + invoice_no duplicated
  2. exact vendor + amount + date within --window-days
  3. same amount + date within --window-days across vendor-name variants
     (normalized vendor matches, raw vendor string does not)
  4. same payment_ref reused across different invoice numbers

Usage - emit a blank register template:
    python3 duplicate_payment_scan.py --template --out ap_register_template.csv

Usage - scan:
    python3 duplicate_payment_scan.py --register ap_register.csv --window-days 5 \
        --client "Northwind Systems Inc" \
        --out "Northwind - Duplicate Payment Scan.xlsx"

--register CSV (12+ months of AP invoices and payments):
    record_id, vendor, amount, invoice_date, invoice_no, payment_date, payment_ref
    (invoice_no, payment_date, payment_ref may be blank per row)
"""

from __future__ import annotations

import argparse
import re
import sys
import csv
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
DATEF = "yyyy-mm-dd"

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")

COLS = ["record_id", "vendor", "amount", "invoice_date", "invoice_no", "payment_date", "payment_ref"]
SUFFIXES = {"inc", "llc", "co", "corp", "ltd", "company", "incorporated", "limited"}


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


def norm_vendor(s: str) -> str:
    s = clean(s).lower()
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    tokens = [t for t in s.split() if t and t not in SUFFIXES]
    return " ".join(tokens)


def write_template(path: Path) -> None:
    with path.open("w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=COLS)
        w.writeheader()
        w.writerow({
            "record_id": "AP-00001", "vendor": "Acme Supply Co", "amount": "1250.00",
            "invoice_date": "2026-01-15", "invoice_no": "9981",
            "payment_date": "2026-01-29", "payment_ref": "CHK-4471",
        })
    print(f"Template written: {path}")


def load(path: Path):
    with path.open(newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        missing = [c for c in COLS if c not in (reader.fieldnames or [])]
        if missing:
            sys.exit(f"{path}: missing required column(s) {missing}")
        raw_rows = [{k: clean(r.get(k)) for k in COLS} for r in reader]
    if not raw_rows:
        sys.exit(f"No rows in {path}")

    scanned, unscanned = [], []
    seen_ids = {}
    dup_id_errors = []
    for i, r in enumerate(raw_rows, start=2):  # row 1 is header
        r["row"] = i
        if not r["record_id"]:
            unscanned.append({**r, "reason": "blank record_id"})
            continue
        if r["record_id"] in seen_ids:
            dup_id_errors.append(
                f"record_id {r['record_id']!r} duplicated at rows "
                f"{seen_ids[r['record_id']]} and {i} - population identity is ambiguous"
            )
            unscanned.append({**r, "reason": f"duplicate record_id (also at row {seen_ids[r['record_id']]})"})
            continue
        seen_ids[r["record_id"]] = i
        try:
            amt = dec(r["amount"])
            if amt is None:
                raise ValueError("blank amount")
            r["amount_dec"] = amt
        except ValueError as e:
            unscanned.append({**r, "reason": str(e)})
            continue
        try:
            r["invoice_date_dec"] = pdate(r["invoice_date"]) if r["invoice_date"] else None
        except ValueError as e:
            unscanned.append({**r, "reason": str(e)})
            continue
        r["vendor_norm"] = norm_vendor(r["vendor"])
        scanned.append(r)
    return scanned, unscanned, len(raw_rows), dup_id_errors


def pairs_within_window(records, window_days):
    """Yield (a, b) pairs from records (sorted by date) within window_days."""
    recs = sorted((r for r in records if r["invoice_date_dec"]), key=lambda r: r["invoice_date_dec"])
    for i, a in enumerate(recs):
        for b in recs[i + 1:]:
            delta = (b["invoice_date_dec"] - a["invoice_date_dec"]).days
            if delta > window_days:
                break
            yield a, b


def scan(records, window_days):
    findings = []
    tier1_pairs, tier2_pairs = set(), set()

    # Tier 1: exact vendor + amount + invoice_no duplicated
    by_vendor_amt_inv = defaultdict(list)
    for r in records:
        if r["invoice_no"]:
            by_vendor_amt_inv[(r["vendor"], r["amount_dec"], r["invoice_no"])].append(r)
    for key, group in by_vendor_amt_inv.items():
        if len(group) > 1:
            for i, a in enumerate(group):
                for b in group[i + 1:]:
                    tier1_pairs.add(frozenset((a["row"], b["row"])))
            findings.append({
                "tier": 1,
                "basis": "exact vendor + amount + invoice_no",
                "vendor": key[0], "amount": key[1],
                "records": [r["record_id"] for r in group],
                "row_refs": [r["row"] for r in group],
                "at_risk": key[1] * (len(group) - 1),
            })

    # Tier 2: exact vendor + amount + date within window, excluding tier-1 pairs
    by_vendor_amt = defaultdict(list)
    for r in records:
        by_vendor_amt[(r["vendor"], r["amount_dec"])].append(r)
    for key, group in by_vendor_amt.items():
        reported = set()
        for a, b in pairs_within_window(group, window_days):
            pair_key = frozenset((a["row"], b["row"]))
            if pair_key in tier1_pairs or pair_key in reported:
                continue
            reported.add(pair_key)
            tier2_pairs.add(pair_key)
            findings.append({
                "tier": 2,
                "basis": f"exact vendor + amount + date within {window_days}d",
                "vendor": key[0], "amount": key[1],
                "records": [a["record_id"], b["record_id"]],
                "row_refs": [a["row"], b["row"]],
                "at_risk": key[1],
            })

    # Tier 3: same amount + date within window, vendor-name variants (normalized match, raw differs)
    by_norm_amt = defaultdict(list)
    for r in records:
        by_norm_amt[(r["vendor_norm"], r["amount_dec"])].append(r)
    variant_log = []
    for key, group in by_norm_amt.items():
        raws = {r["vendor"] for r in group}
        if len(raws) < 2:
            continue
        for a, b in pairs_within_window(group, window_days):
            if a["vendor"] == b["vendor"]:
                continue
            pair_key = frozenset((a["row"], b["row"]))
            if pair_key in tier1_pairs or pair_key in tier2_pairs:
                continue
            findings.append({
                "tier": 3,
                "basis": f"normalized-vendor match ('{key[0]}') + amount + date within {window_days}d",
                "vendor": f"{a['vendor']} / {b['vendor']}", "amount": key[1],
                "records": [a["record_id"], b["record_id"]],
                "row_refs": [a["row"], b["row"]],
                "at_risk": key[1],
            })
        for raw in raws:
            variant_log.append({"normalized": key[0], "raw_vendor": raw})

    # Tier 4: same payment_ref reused across different invoice numbers
    by_ref = defaultdict(list)
    for r in records:
        if r["payment_ref"]:
            by_ref[r["payment_ref"]].append(r)
    for ref, group in by_ref.items():
        invoice_nos = {r["invoice_no"] for r in group if r["invoice_no"]}
        if len(invoice_nos) > 1:
            amounts = [r["amount_dec"] for r in group]
            at_risk = sum(amounts, ZERO) - max(amounts)
            findings.append({
                "tier": 4,
                "basis": "payment_ref reused across different invoice_no values",
                "vendor": "/".join(sorted({r["vendor"] for r in group})),
                "amount": sum(amounts, ZERO),
                "records": [r["record_id"] for r in group],
                "row_refs": [r["row"] for r in group],
                "at_risk": at_risk,
                "ref": ref,
            })

    findings.sort(key=lambda f: (f["tier"], -f["at_risk"]))
    return findings, variant_log


def style_header(ws, cols):
    for c, name in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=c, value=name)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"


def autosize(ws, cols):
    for c, name in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(c)].width = max(len(name) + 2, 14)


def sheet_summary(wb, meta, findings, total_rows, scanned_n, unscanned_n, pop_total, valid):
    ws = wb.active
    ws.title = "Risk Summary"
    ws["A1"] = "DUPLICATE PAYMENT SCAN"
    ws["A1"].font = Font(bold=True, size=14)
    row = 3
    for k, v in meta.items():
        ws.cell(row=row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=row, column=2, value=v)
        row += 1
    row += 1
    ws.cell(row=row, column=1, value="Population rows read").font = Font(bold=True)
    ws.cell(row=row, column=2, value=total_rows); row += 1
    ws.cell(row=row, column=1, value="Scanned").font = Font(bold=True)
    ws.cell(row=row, column=2, value=scanned_n); row += 1
    ws.cell(row=row, column=1, value="Unscanned/unparsed").font = Font(bold=True)
    ws.cell(row=row, column=2, value=unscanned_n); row += 1
    ws.cell(row=row, column=1, value="Population $ total (scanned)").font = Font(bold=True)
    c = ws.cell(row=row, column=2, value=float(pop_total)); c.number_format = MONEY; row += 2

    ws.cell(row=row, column=1, value="Tier").font = HDR_FONT
    ws.cell(row=row, column=1).fill = HDR_FILL
    ws.cell(row=row, column=2, value="Findings").font = HDR_FONT
    ws.cell(row=row, column=2).fill = HDR_FILL
    ws.cell(row=row, column=3, value="$ at risk").font = HDR_FONT
    ws.cell(row=row, column=3).fill = HDR_FILL
    row += 1
    for tier in (1, 2, 3, 4):
        tf = [f for f in findings if f["tier"] == tier]
        ws.cell(row=row, column=1, value=f"Tier {tier}")
        ws.cell(row=row, column=2, value=len(tf))
        c = ws.cell(row=row, column=3, value=float(sum((f["at_risk"] for f in tf), ZERO)))
        c.number_format = MONEY
        row += 1
    row += 1
    status = ws.cell(row=row, column=1, value="GATE STATUS: " + ("CLEAN" if valid else "NOT MET"))
    status.font = OK_FONT if valid else BAD_FONT
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 18


def sheet_findings(wb, findings):
    ws = wb.create_sheet("Findings")
    cols = ["tier", "basis", "vendor", "amount", "records", "row_refs", "at_risk"]
    style_header(ws, cols)
    tier_fill = {1: "FFC7CE", 2: "FFEB9C", 3: "FFEB9C", 4: "FFC7CE"}
    for r, f in enumerate(findings, start=2):
        ws.cell(row=r, column=1, value=f["tier"])
        ws.cell(row=r, column=2, value=f["basis"])
        ws.cell(row=r, column=3, value=f["vendor"])
        ac = ws.cell(row=r, column=4, value=float(f["amount"])); ac.number_format = MONEY
        ws.cell(row=r, column=5, value=", ".join(f["records"]))
        ws.cell(row=r, column=6, value=", ".join(str(x) for x in f["row_refs"]))
        rc = ws.cell(row=r, column=7, value=float(f["at_risk"])); rc.number_format = MONEY
        for c in range(1, 8):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=tier_fill[f["tier"]])
    autosize(ws, cols)


def sheet_unscanned(wb, unscanned):
    ws = wb.create_sheet("Unscanned-Unparsed Rows")
    cols = COLS + ["reason"]
    style_header(ws, cols)
    for r, u in enumerate(unscanned, start=2):
        for c, name in enumerate(COLS, start=1):
            ws.cell(row=r, column=c, value=u.get(name, ""))
        rc = ws.cell(row=r, column=len(COLS) + 1, value=u["reason"])
        for c in range(1, len(COLS) + 2):
            ws.cell(row=r, column=c).fill = WARN_FILL
    autosize(ws, cols)


def sheet_variants(wb, variant_log):
    ws = wb.create_sheet("Vendor Normalization Log")
    cols = ["normalized", "raw_vendor"]
    style_header(ws, cols)
    seen = set()
    r = 2
    for v in variant_log:
        key = (v["normalized"], v["raw_vendor"])
        if key in seen:
            continue
        seen.add(key)
        ws.cell(row=r, column=1, value=v["normalized"])
        ws.cell(row=r, column=2, value=v["raw_vendor"])
        r += 1
    autosize(ws, cols)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--register")
    ap.add_argument("--out", required=True)
    ap.add_argument("--window-days", type=int,
                    help="day window for date-proximity matching (tiers 2 and 3). Required unless --template.")
    ap.add_argument("--client", default="")
    ap.add_argument("--template", action="store_true")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    if args.template:
        write_template(Path(args.out))
        return 0

    if not args.register:
        sys.exit("--register is required (unless --template)")
    if args.window_days is None:
        sys.exit("--window-days is required - there is no built-in default window")

    scanned, unscanned, total_rows, dup_id_errors = load(Path(args.register))
    errors = list(dup_id_errors)
    if len(scanned) + len(unscanned) != total_rows:
        errors.append(
            f"internal coverage failure: {len(scanned)} scanned + {len(unscanned)} unscanned "
            f"!= {total_rows} rows read"
        )
    valid = not errors

    findings, variant_log = scan(scanned, args.window_days) if valid or args.force else ([], [])
    pop_total = sum((r["amount_dec"] for r in scanned), ZERO)

    meta = {
        "Client": args.client or "(not stated)",
        "Window (days)": args.window_days,
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    print("=" * 76)
    print("DUPLICATE PAYMENT SCAN - full population, no sampling")
    print("=" * 76)
    print(f"Client          : {meta['Client']}")
    print(f"Rows read       : {total_rows}   scanned {len(scanned)}   unscanned {len(unscanned)}")
    print(f"Population $    : {pop_total:,.2f}")
    print(f"Window          : {args.window_days} days")
    for tier in (1, 2, 3, 4):
        tf = [f for f in findings if f["tier"] == tier]
        print(f"  Tier {tier}: {len(tf)} finding(s), ${sum((f['at_risk'] for f in tf), ZERO):,.2f} at risk")
    if errors:
        print(f"\nGATE FAILURES ({len(errors)}):")
        for e in errors:
            print(f"  ! {e}")

    if not valid and not args.force:
        print("\n" + "=" * 76)
        print("WORKBOOK NOT WRITTEN.")
        print("Row identity in this population is ambiguous (duplicate record_id) or the")
        print("coverage count does not reconcile. Fix the source extract and re-run.")
        print("=" * 76)
        return 1

    wb = Workbook()
    sheet_summary(wb, meta, findings, total_rows, len(scanned), len(unscanned), pop_total, valid)
    sheet_findings(wb, findings)
    sheet_unscanned(wb, unscanned)
    sheet_variants(wb, variant_log)
    wb.active = 0
    out = Path(args.out)
    if not valid:
        out = out.with_name(out.stem + " [GATE NOT MET]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\n{'GATE CLEAN' if valid else 'NOT MET (forced)'}.  Workbook: {out}")
    return 0 if valid else 1


if __name__ == "__main__":
    sys.exit(main())
```

