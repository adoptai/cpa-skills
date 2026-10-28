#!/usr/bin/env python3
"""
Draft collection emails from an AR aging and a client-supplied escalation
policy. Nothing is ever sent - every draft is status "drafted - pending
review" and a human sends it, or doesn't.

What this does, mechanically, and no further:
  * Matches each overdue, non-disputed invoice to exactly one escalation
    bucket from the policy file. Bucket day ranges and wording are entirely
    client-supplied - this script never assumes 30/60/90 or any other
    standard aging convention.
  * Groups a customer's matched invoices into one draft, using the tier of
    the most-escalated bucket that applies to any of their open invoices.
  * Renders that tier's template with the customer's name, invoice list,
    total, and oldest due date merged in.

Gates:
  * HARD (blocks the workbook unless --force): no escalation policy supplied
    (nothing to drive escalation from); or the dollar total appearing across
    all drafted emails does not exactly equal the sum of the non-disputed,
    bucket-matched invoices used to build them - an internal consistency
    failure that must never reach a reviewer.
  * SOFT (routed to Review Queue / Exceptions, workbook still produced):
    an invoice matching zero or more than one policy bucket; a customer
    with no contact email; any invoice flagged disputed (excluded from that
    customer's draft entirely - a customer whose every open invoice is
    disputed gets no draft at all); a template referencing a field this
    script does not provide.

Usage - emit a blank aging worksheet:
    python3 dunning_drafts.py --template --out aging_template.csv

Usage - build the queue:
    python3 dunning_drafts.py --aging aging.csv --contacts contacts.csv \
        --policy policy.csv --disputes disputes.csv \
        --client "Northwind Systems Inc" --as-of 2025-12-31 \
        --out "Northwind - Collections Queue.xlsx"

--aging CSV:
    customer, invoice_no, amount, due_date, days_past_due

--contacts CSV:
    customer, contact_name, email

--policy CSV (bucket ranges and wording are entirely client-supplied):
    bucket_min_days, bucket_max_days, tier_name, tone, template_text

  template_text may use: {customer} {contact_name} {total} {invoice_count}
  {invoice_list} {oldest_due_date} {tier_name} {tone}

--disputes CSV (optional):
    customer, invoice_no, disputed (y/n)
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

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")

AGING_COLS = ["customer", "invoice_no", "amount", "due_date", "days_past_due"]
CONTACT_COLS = ["customer", "contact_name", "email"]
POLICY_COLS = ["bucket_min_days", "bucket_max_days", "tier_name", "tone", "template_text"]
DISPUTE_COLS = ["customer", "invoice_no", "disputed"]


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
        w.writerow(AGING_COLS)
    print(f"Template written: {out}")


def match_buckets(buckets, days):
    hits = []
    for b in buckets:
        if b["min"] <= days <= b["max"]:
            hits.append(b)
    return hits


def style_header(ws, row, cols):
    for i, c in enumerate(cols, start=1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(c) + 2)


# ---------------------------------------------------------------------- main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                  formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--aging")
    ap.add_argument("--contacts")
    ap.add_argument("--policy")
    ap.add_argument("--disputes")
    ap.add_argument("--client", default="")
    ap.add_argument("--as-of", default="")
    ap.add_argument("--out", required=True)
    ap.add_argument("--template", action="store_true",
                     help="emit a blank aging worksheet instead of building")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    if args.template:
        write_template(Path(args.out))
        return 0

    if not args.aging or not args.contacts:
        sys.exit("--aging and --contacts are required (or pass --template)")

    hard_errors = []

    policy_rows = load_csv(Path(args.policy), POLICY_COLS) if args.policy else []
    buckets = []
    for r in policy_rows:
        try:
            lo, hi = int(clean(r["bucket_min_days"])), int(clean(r["bucket_max_days"]))
        except ValueError:
            hard_errors.append(f"Policy row {r['row']}: bucket_min_days/bucket_max_days "
                                f"must be integers")
            continue
        buckets.append({"min": lo, "max": hi, "tier_name": r["tier_name"],
                         "tone": r["tone"], "template_text": r["template_text"]})
    if not buckets:
        hard_errors.append("No usable escalation policy buckets supplied - there is "
                            "nothing to drive escalation from.")
    buckets.sort(key=lambda b: b["min"])

    contacts = {r["customer"]: {"contact_name": r["contact_name"], "email": r["email"]}
                for r in load_csv(Path(args.contacts), CONTACT_COLS)}

    disputes = set()
    if args.disputes:
        for r in load_csv(Path(args.disputes), DISPUTE_COLS):
            if r["disputed"].strip().lower() in ("y", "yes", "true", "1"):
                disputes.add((r["customer"], r["invoice_no"]))

    aging_rows = load_csv(Path(args.aging), AGING_COLS)
    if not aging_rows:
        sys.exit(f"No rows in aging {args.aging}")

    by_customer = defaultdict(list)
    exceptions = []
    review_queue = []

    for r in aging_rows:
        try:
            amount = dec(r["amount"])
            days = int(clean(r["days_past_due"]))
            due = pdate(r["due_date"])
        except ValueError as e:
            exceptions.append({"category": "Unparseable row", "customer": r["customer"],
                                "invoice_no": r["invoice_no"], "detail": str(e)})
            continue
        if amount is None or days is None or days <= 0:
            continue  # not overdue - not in scope for dunning

        key = (r["customer"], r["invoice_no"])
        item = {"customer": r["customer"], "invoice_no": r["invoice_no"],
                "amount": amount, "due_date": due, "days_past_due": days}

        if key in disputes:
            review_queue.append({**item, "reason": "disputed - excluded from draft"})
            continue

        hits = match_buckets(buckets, days)
        if len(hits) != 1:
            why = "no bucket covers this day count" if not hits else \
                "multiple buckets match: " + ", ".join(h["tier_name"] for h in hits)
            exceptions.append({"category": "Unmatched bucket", "customer": r["customer"],
                                "invoice_no": r["invoice_no"], "detail": why})
            continue

        item["bucket"] = hits[0]
        by_customer[r["customer"]].append(item)

    drafts = []
    total_drafted = ZERO
    total_consumed = ZERO
    for customer, items in by_customer.items():
        contact = contacts.get(customer)
        if not contact or not contact.get("email"):
            exceptions.append({"category": "Missing contact", "customer": customer,
                                "invoice_no": "", "detail": "no contact email on file - "
                                "no draft generated"})
            continue

        tier = max((i["bucket"] for i in items), key=lambda b: b["min"])
        total = sum((i["amount"] for i in items), ZERO)
        oldest = min(i["due_date"] for i in items if i["due_date"])
        invoice_list = "\n".join(
            f"  {i['invoice_no']}: {i['amount']:,.2f} due {i['due_date']}"
            for i in sorted(items, key=lambda i: i["due_date"] or datetime.min.date()))

        fields = {
            "customer": customer, "contact_name": contact["contact_name"],
            "total": f"{total:,.2f}", "invoice_count": len(items),
            "invoice_list": invoice_list, "oldest_due_date": str(oldest),
            "tier_name": tier["tier_name"], "tone": tier["tone"],
        }
        try:
            body = tier["template_text"].format(**fields)
        except (KeyError, IndexError) as e:
            exceptions.append({"category": "Template error", "customer": customer,
                                "invoice_no": "", "detail": f"template references unknown "
                                f"field {e} - no draft generated"})
            continue

        drafts.append({"customer": customer, "email": contact["email"],
                        "tier_name": tier["tier_name"], "total": total,
                        "oldest_due_date": oldest, "invoice_count": len(items),
                        "body": body})
        total_drafted += total
        total_consumed += sum((i["amount"] for i in items), ZERO)

    if total_drafted != total_consumed:
        hard_errors.append(f"Tie-out failure: drafted total {total_drafted:,.2f} does not "
                            f"equal consumed invoice total {total_consumed:,.2f}")

    valid = not hard_errors

    meta = {
        "Client": args.client or "(not stated)",
        "As of": args.as_of or "(not stated)",
        "Overdue invoices in aging": len(aging_rows),
        "Customers drafted": len(drafts),
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    print("=" * 76)
    print("COLLECTIONS DUNNING DRAFTS")
    print("=" * 76)
    print(f"Client {meta['Client']}   As of {meta['As of']}")
    print(f"Drafted {len(drafts)} customer emails, total {total_drafted:,.2f}")
    print(f"Exceptions {len(exceptions)}   Review queue (disputed) {len(review_queue)}")
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
    ws.title = "Collection Queue"
    cols = ["customer", "tier_name", "total_overdue", "oldest_due_date", "invoice_count"]
    style_header(ws, 1, cols)
    r = 2
    for d in sorted(drafts, key=lambda d: d["total"], reverse=True):
        ws.cell(row=r, column=1, value=d["customer"])
        ws.cell(row=r, column=2, value=d["tier_name"])
        ws.cell(row=r, column=3, value=float(d["total"])).number_format = MONEY
        ws.cell(row=r, column=4, value=str(d["oldest_due_date"]))
        ws.cell(row=r, column=5, value=d["invoice_count"])
        r += 1

    ws2 = wb.create_sheet("Drafted Emails")
    style_header(ws2, 1, ["customer", "email", "status", "body"])
    r = 2
    for d in drafts:
        ws2.cell(row=r, column=1, value=d["customer"])
        ws2.cell(row=r, column=2, value=d["email"])
        ws2.cell(row=r, column=3, value="drafted - pending review")
        cell = ws2.cell(row=r, column=4, value=d["body"])
        cell.alignment = Alignment(wrap_text=True)
        r += 1
    ws2.column_dimensions["D"].width = 90

    ws3 = wb.create_sheet("Review Queue - Exceptions")
    style_header(ws3, 1, ["category", "customer", "invoice_no", "detail"])
    r = 2
    for item in review_queue:
        ws3.cell(row=r, column=1, value="Disputed")
        ws3.cell(row=r, column=2, value=item["customer"])
        ws3.cell(row=r, column=3, value=item["invoice_no"])
        ws3.cell(row=r, column=4, value=item["reason"])
        ws3.cell(row=r, column=1).fill = WARN_FILL
        r += 1
    for e in exceptions:
        ws3.cell(row=r, column=1, value=e["category"])
        ws3.cell(row=r, column=2, value=e["customer"])
        ws3.cell(row=r, column=3, value=e["invoice_no"])
        ws3.cell(row=r, column=4, value=e["detail"])
        ws3.cell(row=r, column=1).fill = WARN_FILL
        r += 1
    if r == 2:
        ws3.cell(row=2, column=1, value="(no exceptions)")

    ws4 = wb.create_sheet("Tie-Out")
    style_header(ws4, 1, ["check", "value"])
    tie_ok = total_drafted == total_consumed
    ws4.append(["Total drafted across all emails", float(total_drafted)])
    ws4.append(["Total non-disputed, bucket-matched invoices consumed", float(total_consumed)])
    ws4.append(["Difference", float(total_drafted - total_consumed)])
    ws4.append(["Result", "TIE" if tie_ok else "BREAK"])
    for row in range(2, 5):
        ws4.cell(row=row, column=2).number_format = MONEY if row < 5 else None
    ws4.cell(row=5, column=2).font = OK_FONT if tie_ok else BAD_FONT

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
