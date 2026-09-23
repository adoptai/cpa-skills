#!/usr/bin/env python3
"""
Synthetic AR aging test dataset generator (ADO-72).

Deterministic (seeded) generator that produces a raw AR subledger --
customers, transactions (invoices / payments / credit memos / write-offs) --
plus a hand-verifiable "answer key" aging report computed independently from
the same transaction data. Ten deliberate edge cases are hardcoded onto
customers CUST-031 through CUST-040; customers CUST-001 through CUST-030
carry randomized (but seeded, reproducible) "normal" activity for volume.

As-of date: 2026-06-30
Bucket boundaries (days past due): 30, 60, 90 -- current, 1-30, 31-60,
61-90, 90+. Boundary convention: a day count *equal to* a boundary falls in
the LOWER bucket (days_past_due <= 30 -> "1-30", not "31-60"). This matches
the bucket_of() convention already used by skills/ar-aging-tie-out in the
cpa-skills repo (`if days <= b: return labels[i + 1]`), so the two datasets
don't disagree on where a boundary invoice lands.

Aging basis: DAYS PAST DUE DATE, not days since invoice date (the edge case
list explicitly says "invoices due exactly on a bucket boundary").

Run: python3 generate_synthetic_data.py
Produces customers.csv, ar_transactions.csv, expected_aging.csv in this
directory (overwrites in place).
"""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

random.seed(20260630)  # deterministic -- rerun reproduces byte-identical output

OUT = Path(__file__).parent
AS_OF = date(2026, 6, 30)
RANGE_START = date(2025, 10, 1)

BUCKET_BOUNDS = [30, 60, 90]
BUCKET_LABELS = ["current", "1-30", "31-60", "61-90", "90+"]

TERMS_DAYS = {"Net 15": 15, "Net 30": 30, "Net 45": 45, "Net 60": 60}
TERMS_CHOICES = ["Net 15", "Net 30", "Net 45", "Net 60"]

NORMAL_NAMES = [
    "Harborview Fabrication", "Bristlecone Logistics", "Meridian Office Supply",
    "Falcon Ridge Industrial", "Cobalt Point Distributors", "Amberlane Freight",
    "Northwind Packaging", "Silverton Hardware", "Cedar Hollow Electrical",
    "Bluepeak Machine Works", "Ironvale Sheet Metal", "Windmere Textiles",
    "Granite Bay Wholesale", "Copper Creek Supply", "Fairview Print & Signage",
    "Redstone Contract Furnishings", "Larkspur Foodservice Equipment",
    "Thornbury Plastics", "Ashgrove Building Materials", "Wrenfield Tooling",
    "Maple Ledge Distribution", "Stonebriar Industrial Parts",
    "Quarrystone Logistics", "Hollowbrook Packaging", "Elmsworth Wholesale",
    "Pinecrest Supply Chain", "Rustwood Fasteners", "Briarcliff Textiles",
    "Oakmere Distributors", "Sablefield Equipment Rental",
]
EDGE_NAMES = [
    "Boundary Thirty", "Boundary Sixty", "Boundary Ninety",
    "Partial Pay Industries", "Multi-Invoice Payers", "Overpaid Credit",
    "Unapplied Cash Ventures", "Fully Settled Supply", "Future Dated Freight",
    "Net Credit Holdings",
]
assert len(NORMAL_NAMES) == 30
assert len(EDGE_NAMES) == 10

SUFFIXES = ["LLC", "Inc.", "Co.", "Corp."]

txn_seq = 0
inv_seq = 100000


def next_txn_id():
    global txn_seq
    txn_seq += 1
    return f"TXN-{txn_seq:06d}"


def next_invoice_number():
    global inv_seq
    inv_seq += 1
    return f"INV-{inv_seq}"


def bucket_of(days_past_due):
    if days_past_due <= 0:
        return "current"
    for bound, label in zip(BUCKET_BOUNDS, BUCKET_LABELS[1:]):
        if days_past_due <= bound:
            return label
    return BUCKET_LABELS[-1]


def rand_date(start, end):
    span = (end - start).days
    return start + timedelta(days=random.randint(0, span))


def rand_amount(lo=500, hi=85000):
    # weighted toward the lower/mid range, occasional large invoice
    if random.random() < 0.08:
        return round(random.uniform(30000, hi), 2)
    return round(random.uniform(lo, 22000), 2)


# ---------------------------------------------------------------------------
# customers.csv
# ---------------------------------------------------------------------------
customers = []
INACTIVE = {"CUST-020", "CUST-025", "CUST-030"}
for i, name in enumerate(NORMAL_NAMES + EDGE_NAMES, start=1):
    cid = f"CUST-{i:03d}"
    terms = random.choice(TERMS_CHOICES)
    credit_limit = round(random.uniform(5000, 150000) / 500) * 500
    customers.append({
        "customer_id": cid,
        "customer_name": f"{name} {random.choice(SUFFIXES)}",
        "payment_terms": terms,
        "credit_limit": credit_limit,
        "is_active": "false" if cid in INACTIVE else "true",
    })

CUST_BY_ID = {c["customer_id"]: c for c in customers}

# ---------------------------------------------------------------------------
# ar_transactions.csv -- built as a list of dicts, one per row
# ---------------------------------------------------------------------------
transactions = []


def add_invoice(customer_id, txn_date, amount, invoice_number=None, status="open"):
    terms = TERMS_DAYS[CUST_BY_ID[customer_id]["payment_terms"]]
    inv_no = invoice_number or next_invoice_number()
    transactions.append({
        "txn_id": next_txn_id(),
        "customer_id": customer_id,
        "txn_type": "invoice",
        "invoice_number": inv_no,
        "txn_date": txn_date,
        "due_date": txn_date + timedelta(days=terms),
        "amount": round(amount, 2),
        "applied_to_invoice": "",
        "currency": "USD",
        "status": status,
    })
    return inv_no


def add_application(customer_id, txn_type, txn_date, amount, applied_to_invoice,
                     status="applied"):
    # amount is stored negative -- it reduces the AR balance
    transactions.append({
        "txn_id": next_txn_id(),
        "customer_id": customer_id,
        "txn_type": txn_type,
        "invoice_number": "",
        "txn_date": txn_date,
        "due_date": "",
        "amount": -round(abs(amount), 2),
        "applied_to_invoice": applied_to_invoice,
        "currency": "USD",
        "status": status,
    })


# --- 30 "normal" customers: randomized but seeded invoice/payment activity --
for c in customers[:30]:
    cid = c["customer_id"]
    n_invoices = random.randint(6, 11)
    for _ in range(n_invoices):
        inv_date = rand_date(RANGE_START, AS_OF - timedelta(days=1))
        amount = rand_amount()
        inv_no = add_invoice(cid, inv_date, amount)
        terms = TERMS_DAYS[c["payment_terms"]]
        due = inv_date + timedelta(days=terms)

        behavior = random.random()
        if behavior < 0.38:
            # paid in full shortly after issuance -- nets to zero, excluded
            pay_date = inv_date + timedelta(days=random.randint(5, terms + 10))
            if pay_date <= AS_OF:
                add_application(cid, "payment", pay_date, amount, inv_no)
        elif behavior < 0.56:
            # partial payment -- residual ages by the invoice's own due date
            pay_date = inv_date + timedelta(days=random.randint(5, terms + 15))
            if pay_date <= AS_OF:
                partial = round(amount * random.uniform(0.25, 0.7), 2)
                add_application(cid, "payment", pay_date, partial, inv_no)
        elif behavior < 0.63:
            # small credit memo (return/discount) -- residual ages normally
            cm_date = inv_date + timedelta(days=random.randint(5, terms + 15))
            if cm_date <= AS_OF:
                credit = round(amount * random.uniform(0.05, 0.25), 2)
                add_application(cid, "credit_memo", cm_date, credit, inv_no)
        elif behavior < 0.93:
            # left fully open -- ages in full
            pass
        else:
            # written off
            wo_date = due + timedelta(days=random.randint(30, 90))
            if wo_date <= AS_OF:
                add_application(cid, "writeoff", wo_date, amount, inv_no,
                                 status="written_off")

    for row in transactions:
        if row["customer_id"] == cid and row["txn_type"] == "invoice":
            pass  # status left as "open"; the aging skill recomputes it anyway

# statuses on the invoice rows for the 30 normal customers are informational
# only ("open"/"written_off" set above); the dataset intentionally does NOT
# maintain a trustworthy running "paid"/"partial" status field on invoices --
# an aging skill must net the applications itself rather than read a status
# column, exactly as skills/ar-aging-tie-out already does for its own input.

# ---------------------------------------------------------------------------
# The 10 hardcoded edge cases (CUST-031..CUST-040)
# ---------------------------------------------------------------------------

# 1) CUST-031 -- boundary at EXACTLY 30 days past due, due_date = AS_OF - 30
c = "CUST-031"
d31 = AS_OF - timedelta(days=30) - timedelta(days=TERMS_DAYS[CUST_BY_ID[c]["payment_terms"]])
inv31 = add_invoice(c, d31, 12000.00, invoice_number="INV-970100")  # shared w/ CUST-034, see #10 dup

# 2) CUST-032 -- boundary at EXACTLY 60 days past due
c = "CUST-032"
d32 = AS_OF - timedelta(days=60) - timedelta(days=TERMS_DAYS[CUST_BY_ID[c]["payment_terms"]])
inv32 = add_invoice(c, d32, 8400.00)

# 3) CUST-033 -- boundary at EXACTLY 90 days past due
c = "CUST-033"
d33 = AS_OF - timedelta(days=90) - timedelta(days=TERMS_DAYS[CUST_BY_ID[c]["payment_terms"]])
inv33 = add_invoice(c, d33, 15250.00)

# 4) CUST-034 -- partial payment leaving a residual balance
c = "CUST-034"
d34 = AS_OF - timedelta(days=45) - timedelta(days=TERMS_DAYS[CUST_BY_ID[c]["payment_terms"]])
inv34 = add_invoice(c, d34, 10000.00, invoice_number="INV-970100", status="partial")  # DUPLICATE of CUST-031's number
add_application(c, "payment", AS_OF - timedelta(days=20), 6500.00, inv34)

# 5) CUST-035 -- one payment event applied across two different invoices
c = "CUST-035"
termsdays = TERMS_DAYS[CUST_BY_ID[c]["payment_terms"]]
invE_date = AS_OF - timedelta(days=50) - timedelta(days=termsdays)
invF_date = AS_OF - timedelta(days=40) - timedelta(days=termsdays)
invE = add_invoice(c, invE_date, 6000.00)
invF = add_invoice(c, invF_date, 9000.00)
pay_date_35 = AS_OF - timedelta(days=10)
add_application(c, "payment", pay_date_35, 4000.00, invE)  # same wire, split
add_application(c, "payment", pay_date_35, 6000.00, invF)  # across 2 invoices

# 6) CUST-036 -- credit memo that overpays an invoice (negative balance)
c = "CUST-036"
d36 = AS_OF - timedelta(days=20) - timedelta(days=TERMS_DAYS[CUST_BY_ID[c]["payment_terms"]])
inv36 = add_invoice(c, d36, 5000.00)
add_application(c, "credit_memo", AS_OF - timedelta(days=5), 6000.00, inv36)

# 7) CUST-037 -- unapplied cash (payment with no applied_to_invoice)
c = "CUST-037"
d37 = AS_OF - timedelta(days=100) - timedelta(days=TERMS_DAYS[CUST_BY_ID[c]["payment_terms"]])
inv37 = add_invoice(c, d37, 15000.00)  # left open, ages past 90
unapplied_date = AS_OF - timedelta(days=15)
add_application(c, "payment", unapplied_date, 2500.00, "", status="unapplied")

# 8) CUST-038 -- a fully paid invoice (must not appear in aging at all)
c = "CUST-038"
d38 = AS_OF - timedelta(days=70) - timedelta(days=TERMS_DAYS[CUST_BY_ID[c]["payment_terms"]])
inv38 = add_invoice(c, d38, 7500.00, status="paid")
add_application(c, "payment", AS_OF - timedelta(days=30), 7500.00, inv38, status="applied")

# 9) CUST-039 -- an invoice dated AFTER the as-of date (must be excluded)
c = "CUST-039"
d39_normal = AS_OF - timedelta(days=10) - timedelta(days=TERMS_DAYS[CUST_BY_ID[c]["payment_terms"]])
inv39_normal = add_invoice(c, d39_normal, 4200.00)  # this one IS in range/open
future_date = AS_OF + timedelta(days=10)  # 2026-07-10, after the as-of date
inv39_future = add_invoice(c, future_date, 9900.00)

# 10) CUST-040 -- a customer with a net credit balance overall
c = "CUST-040"
d40 = AS_OF - timedelta(days=5) - timedelta(days=TERMS_DAYS[CUST_BY_ID[c]["payment_terms"]])
inv40 = add_invoice(c, d40, 3000.00)  # small open invoice, lands in "current"
add_application(c, "credit_memo", AS_OF - timedelta(days=2), 8000.00, "",
                 status="unapplied")  # unapplied credit, net customer balance < 0

random.shuffle(transactions)  # so file order doesn't mirror generation order
transactions.sort(key=lambda r: r["txn_date"])  # then a stable, realistic sort

# ---------------------------------------------------------------------------
# expected_aging.csv -- computed independently from the transaction rows
# ---------------------------------------------------------------------------

def compute_aging():
    # Open balance per (customer_id, invoice_number), keyed together so a
    # duplicate invoice_number across two customers can never cross-apply.
    invoice_rows = {}  # (customer_id, invoice_number) -> dict(due_date, balance)
    unapplied = []  # rows with no applied_to_invoice: (customer_id, date, amount)

    for r in transactions:
        cid = r["customer_id"]
        if r["txn_type"] == "invoice":
            if r["txn_date"] > AS_OF:
                continue  # cutoff: not yet invoiced as of the as-of date
            key = (cid, r["invoice_number"])
            invoice_rows[key] = {"due_date": r["due_date"], "balance": r["amount"]}
        else:
            if r["txn_date"] > AS_OF:
                continue  # a future-dated application shouldn't occur, but guard anyway
            if r["applied_to_invoice"]:
                key = (cid, r["applied_to_invoice"])
                if key in invoice_rows:
                    invoice_rows[key]["balance"] += r["amount"]
                # if the target invoice doesn't exist (or is post-cutoff), the
                # application has nothing to net against -- not exercised by
                # this dataset, so it is intentionally left unhandled here.
            else:
                unapplied.append((cid, r["txn_date"], r["amount"]))

    per_customer = {c["customer_id"]: {l: 0.0 for l in BUCKET_LABELS} for c in customers}

    for (cid, _inv_no), row in invoice_rows.items():
        bal = round(row["balance"], 2)
        if bal == 0:
            continue  # fully paid / written off -- does not appear in aging
        days_past_due = (AS_OF - row["due_date"]).days
        b = bucket_of(days_past_due)
        per_customer[cid][b] += bal

    for cid, adate, amount in unapplied:
        # No invoice to peg unapplied cash/credits to -- aged from its own
        # transaction date instead of a due date (documented in README.md).
        days_past_due = (AS_OF - adate).days
        b = bucket_of(days_past_due)
        per_customer[cid][b] += amount

    out = []
    for c in customers:
        cid = c["customer_id"]
        buckets = per_customer[cid]
        total = round(sum(buckets.values()), 2)
        out.append({
            "customer_id": cid,
            "customer_name": c["customer_name"],
            "current": round(buckets["current"], 2),
            "1-30": round(buckets["1-30"], 2),
            "31-60": round(buckets["31-60"], 2),
            "61-90": round(buckets["61-90"], 2),
            "90+": round(buckets["90+"], 2),
            "total": total,
        })
    return out


expected_aging = compute_aging()

# ---------------------------------------------------------------------------
# write files
# ---------------------------------------------------------------------------

with (OUT / "customers.csv").open("w", newline="", encoding="utf-8") as fh:
    w = csv.DictWriter(fh, fieldnames=[
        "customer_id", "customer_name", "payment_terms", "credit_limit", "is_active"])
    w.writeheader()
    w.writerows(customers)

with (OUT / "ar_transactions.csv").open("w", newline="", encoding="utf-8") as fh:
    w = csv.DictWriter(fh, fieldnames=[
        "txn_id", "customer_id", "txn_type", "invoice_number", "txn_date", "due_date",
        "amount", "applied_to_invoice", "currency", "status"])
    w.writeheader()
    for r in transactions:
        row = dict(r)
        row["txn_date"] = r["txn_date"].isoformat()
        row["due_date"] = r["due_date"].isoformat() if r["due_date"] else ""
        w.writerow(row)

with (OUT / "expected_aging.csv").open("w", newline="", encoding="utf-8") as fh:
    w = csv.DictWriter(fh, fieldnames=[
        "customer_id", "customer_name", "current", "1-30", "31-60", "61-90", "90+", "total"])
    w.writeheader()
    w.writerows(expected_aging)

print(f"customers: {len(customers)}")
print(f"transactions: {len(transactions)}")
print(f"expected_aging rows: {len(expected_aging)}")

grand_total = round(sum(row["total"] for row in expected_aging), 2)
print(f"grand total across all customers: {grand_total:,.2f}")
