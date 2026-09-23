# AR Aging synthetic test dataset (ADO-72)

Synthetic accounts-receivable subledger for testing an AR aging skill/tool against a
hand-verifiable ground truth. Entirely fictional — no real client, customer, or invoice
data. Generated deterministically by `generate_synthetic_data.py` (seed `20260630`);
rerunning it reproduces these files byte-for-byte.

Originally generated and pushed to [adoptai/skills](https://github.com/adoptai/skills)'s
top-level `synthetic_data/ar-aging/` (ADO-72), since that was the only existing shared
fixture location in the org at the time. Forked into this repo, `synthetic_data/ar-aging/`,
so it lives alongside the skill it actually exercises, `skills/ar-aging-tie-out`
(`ar_aging.py`). Column and bucket-boundary conventions (invoice-level detail, netting by
`(customer_id, invoice_number)`, a day count landing exactly on 30/60/90 falling in the
lower bucket) already followed `ar-aging-tie-out`'s own `bucket_of()` — see "Aging
methodology" below.

## Parameters

| | |
|---|---|
| As-of date | **2026-06-30** |
| Date range | 2025-10-01 through 2026-06-30 (one deliberate exception, see edge case 9) |
| Customers | 40 (30 randomized "normal" + 10 dedicated edge-case customers) |
| Transactions | 439 |
| Invoice amounts | $577.21 – $80,906.27 |
| Currency | USD only (single-currency, documented simplification) |
| Bucket boundaries | 30 / 60 / 90 days **past due** |

## Files

### `customers.csv`

`customer_id, customer_name, payment_terms, credit_limit, is_active`

- `payment_terms` is one of `Net 15` / `Net 30` / `Net 45` / `Net 60`.
- `is_active` is `true`/`false`; three customers (`CUST-020`, `CUST-025`, `CUST-030`) are
  inactive. This is cosmetic realism only — it has no bearing on the aging computation.

### `ar_transactions.csv`

`txn_id, customer_id, txn_type, invoice_number, txn_date, due_date, amount, applied_to_invoice, currency, status`

- `txn_type` is one of `invoice | payment | credit_memo | writeoff`.
- `invoice_number` and `due_date` are populated **only** on `invoice` rows — the other
  three types reduce an existing invoice's balance and carry no invoice number of their
  own.
- `amount` is signed: positive on `invoice` rows, negative on `payment` /
  `credit_memo` / `writeoff` rows (they reduce the AR balance).
- `applied_to_invoice` holds the `invoice_number` a payment/credit/writeoff reduces. It
  is blank for invoices, and also blank for **unapplied cash/credit** (a payment or
  credit memo with nothing to net against — see edge case 7).
- `status` is informational only, exactly as `cpa-skills/skills/ar-aging-tie-out`
  treats a reported aging bucket: **don't trust it, recompute the balance from the
  transactions.** It is not guaranteed to be updated consistently (e.g. it is not the
  signal that tells you an invoice is fully paid — the netted balance is).
- Invoice numbers are **not globally unique** — see edge case 10 below. Always net a
  payment/credit/writeoff against an invoice scoped to the same `customer_id`, never by
  `invoice_number` alone.

### `expected_aging.csv` — ground truth

`customer_id, customer_name, current, 1-30, 31-60, 61-90, 90+, total`

One row per customer (all 40, including customers with an all-zero row), aggregated by
recomputed bucket as of 2026-06-30. `total` is the sum of the five bucket columns, and it
foots for every row.

## Aging methodology (how `expected_aging.csv` was computed)

1. **Basis: days past due date**, not days since invoice date. `days_past_due = as_of -
   due_date`.
2. **Bucket assignment** (`bucket_of()` in the generator):
   - `days_past_due <= 0` → **current**
   - `1 <= days_past_due <= 30` → **1-30**
   - `31 <= days_past_due <= 60` → **31-60**
   - `61 <= days_past_due <= 90` → **61-90**
   - `days_past_due > 90` → **90+**

   A day count that lands **exactly on** a boundary (30 / 60 / 90) falls in the
   **lower** bucket. This mirrors `cpa-skills/skills/ar-aging-tie-out`'s own
   `bucket_of()` (`if days <= b: return labels[i + 1]`), so a boundary invoice doesn't
   get bucketed differently depending which of the two datasets/tools you're looking
   at.
3. **Netting**: every invoice's balance starts at its `amount` and is reduced by every
   `payment`/`credit_memo`/`writeoff` row whose `applied_to_invoice` + `customer_id`
   match it. The **net balance** is what gets aged and bucketed by the invoice's own
   `due_date` — a partially paid or over-credited invoice still ages from when it was
   originally due, not from the date of the last application against it.
4. **Zero-balance invoices are excluded** — a fully paid or fully written-off invoice
   contributes `$0` to every bucket, i.e. it doesn't appear.

   **Write-off convention, pinned:** a `writeoff` row nets against its invoice exactly
   like a payment or credit memo — it is not a separate lifecycle state. The moment a
   write-off transaction is dated on or before the as-of date, the invoice's net balance
   goes to zero and it drops out of the aging **immediately**, the same run in which it's
   computed. There is no interim state where a written-off invoice stays visible "until
   period close" — this generator has no such concept, and neither should a skill reading
   this fixture. (A write-off dated *after* the as-of date wouldn't apply yet, by the same
   cutoff rule as any other application — see point 6 — but this dataset doesn't exercise
   that combination.) Concretely: `TXN-000285` writes off the full `3,474.53` balance of
   `INV-100174` (`CUST-020`) on `2026-01-13`; that invoice contributes nothing to
   `CUST-020`'s row in `expected_aging.csv`. 14 such write-offs are scattered through the
   bulk data (`grep writeoff ar_transactions.csv`); there is no dedicated `CUST-03x` case
   for this one; instead the bulk data carries it since it doesn't need isolation to be
   unambiguous — netting to zero is the same rule already exercised by edge case 8.
5. **Unapplied cash/credit** (blank `applied_to_invoice`) has no invoice to inherit a
   due date from, so it is aged from **its own `txn_date`** instead
   (`days_past_due = as_of - txn_date`). This is a deliberate modeling choice — the
   generator's `compute_aging()` documents it inline — since there's no other date to
   peg it to.
6. **Cutoff**: any `invoice` row with `txn_date > as_of` is excluded entirely — not
   counted in any bucket, not even as a $0 line. It hasn't happened yet as of the
   as-of date.

## The ten edge cases

Ten customers, `CUST-031` through `CUST-040`, are hand-authored so each edge case is
isolated and traceable. (Bulk/"normal" activity lives on `CUST-001`–`CUST-030`.)

### 1. Invoice due exactly 30 days past due — `CUST-031` / Boundary Thirty
Invoice `INV-970100`, due 2026-05-31.
`days_past_due = 2026-06-30 − 2026-05-31 = 30` → boundary, falls in **1-30** (not
31-60). `1-30 = 12,000.00`.

### 2. Invoice due exactly 60 days past due — `CUST-032` / Boundary Sixty
Invoice `INV-100258`, due 2026-05-01.
`days_past_due = 2026-06-30 − 2026-05-01 = 60` → boundary, falls in **31-60**.
`31-60 = 8,400.00`.

### 3. Invoice due exactly 90 days past due — `CUST-033` / Boundary Ninety
Invoice `INV-100259`, due 2026-04-01.
`days_past_due = 2026-06-30 − 2026-04-01 = 90` → boundary, falls in **61-90**.
`61-90 = 15,250.00`.

### 4. Partial payment leaving a residual balance — `CUST-034` / Partial Pay Industries
Invoice `INV-970100` (see edge case 10 — shared number, different customer) for
`10,000.00`, due 2026-05-16. Payment of `6,500.00` applied against it.
Residual balance `= 10,000.00 − 6,500.00 = 3,500.00`.
`days_past_due = 2026-06-30 − 2026-05-16 = 45` → **31-60**. `31-60 = 3,500.00`.

### 5. Payment applied across multiple invoices — `CUST-035` / Multi-Invoice Payers
One remittance, same date (2026-06-20), split into two ledger rows:
- Invoice `INV-100260` for `6,000.00`, due 2026-05-11, paid `4,000.00` → residual
  `2,000.00`. `days_past_due = 2026-06-30 − 2026-05-11 = 50` → **31-60**.
- Invoice `INV-100261` for `9,000.00`, due 2026-05-21, paid `6,000.00` → residual
  `3,000.00`. `days_past_due = 2026-06-30 − 2026-05-21 = 40` → **31-60**.

Both residuals land in the same bucket: `31-60 = 2,000.00 + 3,000.00 = 5,000.00`.

### 6. Credit memo overpaying an invoice (negative balance) — `CUST-036` / Overpaid Credit
Invoice `INV-100262` for `5,000.00`, due 2026-06-10. Credit memo of `6,000.00` applied
against it. Balance `= 5,000.00 − 6,000.00 = −1,000.00` (a credit).
`days_past_due = 2026-06-30 − 2026-06-10 = 20` → **1-30**, still aged by the invoice's
own due date even though the net balance is negative. `1-30 = −1,000.00`.

### 7. Unapplied cash — `CUST-037` / Unapplied Cash Ventures
Invoice `INV-100263` for `15,000.00`, due 2026-03-22, left open in full:
`days_past_due = 2026-06-30 − 2026-03-22 = 100` → **90+** (`90+ = 15,000.00`).
Separately, a payment of `2,500.00` dated 2026-06-15 with no `applied_to_invoice`.
Aged from its own date: `days_past_due = 2026-06-30 − 2026-06-15 = 15` → **1-30**
(`1-30 = −2,500.00`). Customer total `= 15,000.00 − 2,500.00 = 12,500.00`.

### 8. Fully paid invoice (must not appear) — `CUST-038` / Fully Settled Supply
Invoice `INV-100264` for `7,500.00`, paid `7,500.00` in full on 2026-05-31.
Balance `= 0.00` → excluded from every bucket. `expected_aging.csv` still carries a row
for `CUST-038` (every customer gets one), but every column, including `total`, is
`0.00`.

### 9. Invoice dated after the as-of date (excluded) — `CUST-039` / Future Dated Freight
Invoice `INV-100265` for `4,200.00`, due 2026-06-20 (in range):
`days_past_due = 10` → **1-30** (`1-30 = 4,200.00`).
A second invoice, `INV-100266` for `9,900.00`, is dated **2026-07-10** — after the
2026-06-30 as-of date — and is excluded from the aging entirely (not counted as $0,
simply omitted from the netting/bucketing pass). Customer total is `4,200.00`, not
`14,100.00`.

### 10. Customer with a net credit balance overall — `CUST-040` / Net Credit Holdings
Invoice `INV-100267` for `3,000.00`, due 2026-06-25: `days_past_due = 5` → **1-30**
(`+3,000.00`). Unapplied credit memo for `8,000.00` dated 2026-06-28: aged from its own
date, `days_past_due = 2` → **1-30** (`−8,000.00`). Combined: `1-30 = 3,000.00 −
8,000.00 = −5,000.00`, and it is the customer's only nonzero bucket, so `total =
−5,000.00` — a net credit balance across the whole customer, not just one invoice.

### Bonus: duplicate invoice numbers across two different customers
`INV-970100` is used **twice** — once for `CUST-031` (edge case 1, `$12,000.00`, no
applications) and once for `CUST-034` (edge case 4, `$10,000.00`, partially paid). They
are netted independently because the netting key is `(customer_id, invoice_number)`,
not `invoice_number` alone; a payment on `CUST-034`'s copy of `INV-970100` never
touches `CUST-031`'s balance and vice versa. Confirmed correct in
`expected_aging.csv`: `CUST-031` = `12,000.00` in `1-30`, `CUST-034` = `3,500.00` in
`31-60` — the two invoices don't collide.

## Regenerating

```bash
cd synthetic_data/ar-aging
python3 generate_synthetic_data.py
```

Note on shape: `skills/ar-aging-tie-out/scripts/ar_aging.py --aging ...` expects a
pre-existing aging-detail CSV (one row per open invoice, with the reported bucket and open
balance already assigned) to re-age and tie out. `ar_transactions.csv` here is the raw
subledger it would be *derived from* — invoices, payments, credit memos, and writeoffs as
separate rows, netted by this dataset's own `expected_aging.csv`/`generate_synthetic_data.py`
logic rather than the script's. They're companion fixtures for the same skill, not
interchangeable inputs; adapting one into the other's shape is left to whoever wires up the
test.

Deterministic — the script seeds `random` with `20260630` and writes all three CSVs in
place. No external dependencies (standard library only).
