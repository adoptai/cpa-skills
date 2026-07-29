---
name: revenue-trace-to-source
description: Trace a revenue sample from the general ledger through to the invoice, the contract, and the cash receipt — proving each recorded amount agrees across the whole chain, and refusing to record an item as tested unless every required evidence link is present. Use this whenever the user mentions revenue testing, tracing revenue, vouching revenue, revenue substantive procedures, agreeing revenue to invoices or contracts or cash, revenue recognition testing, cutoff testing on revenue, or revenue existence and occurrence. Also trigger on "trace these revenue transactions," "vouch revenue to source," "test revenue recognition," "agree revenue to cash," "did this revenue actually happen," "revenue cutoff testing," or when a revenue population and supporting documents are provided together. Runs fully local — no client revenue or contract data leaves the machine.
---

# Revenue Trace to Source

Before anything else, be clear about which way you are walking, because the two directions test
different assertions and are constantly confused:

- **GL → source** (this skill) tests **occurrence and existence**. Does recorded revenue represent
  a real transaction, at the right amount, in the right period? It finds fictitious and
  overstated revenue.
- **Source → GL** tests **completeness**. Is every real transaction recorded? It finds unrecorded
  revenue and skimming.

**Tracing from the ledger cannot detect unrecorded revenue**, because unrecorded revenue is not in
the population you sampled from. If completeness is the concern, this is the wrong procedure and
the workpaper must say so rather than implying coverage it does not have. The script states the
direction and the assertion on the face of the workpaper for exactly this reason.

## The gate

An item is not "tested" until the evidence exists. The script refuses to produce a clean
workpaper when:

1. **A required evidence link is missing.** Each selection needs the invoice, the contract or
   order, and the cash receipt — or an explicit, documented reason why one does not apply.
   Concluding "agreed" with a blank evidence field is rejected.
2. **Amounts disagree across the chain.** GL amount, invoice amount, contract-derived amount
   (price × quantity), and cash received must agree or carry a named difference.
3. **The population does not tie to GL revenue.** You supply the total; testing a filtered
   extract proves nothing about the account.
4. **A conclusion contradicts its own figures** — marked `agreed` while a difference exists.

## Inputs

1. **The revenue population** and the GL revenue total it should equal. Use the
   `audit-sampling` skill to select from it — that gives you a reproducible seed and a documented
   basis, which this skill assumes rather than reinvents.
2. **The selections to test**, with the GL amount and date for each.
3. **Supporting documents** per selection: invoice, contract or purchase order, shipping or
   delivery evidence, and the cash receipt or bank credit.
4. **The period end date**, for cutoff testing.
5. **Subsequent-period credit memos**, if available. See below — these are the highest-yield
   document in a revenue test and are almost never requested.

## Step 1 — Build the trace worksheet

```bash
python3 scripts/revenue_trace.py --selections selections.csv \
  --population-total 18422900.00 --period-end 2025-12-31 \
  --client "Northwind Systems Inc" \
  --assertion "Occurrence of recorded revenue for FY2025" \
  --template --out trace_worksheet.csv
```

That emits a worksheet with the evidence columns to fill in. Then run the test:

```bash
python3 scripts/revenue_trace.py --selections trace_worksheet.csv \
  --population-total 18422900.00 --period-end 2025-12-31 \
  --credit-memos subsequent_credits.csv \
  --client "Northwind Systems Inc" \
  --assertion "Occurrence of recorded revenue for FY2025" \
  --out "Northwind - FY2025 Revenue Trace.xlsx"
```

## Step 2 — What each link proves, and what it doesn't

Each link in the chain answers a different question. Treating them as interchangeable is how a
revenue test looks complete while proving little.

| Link | What it proves | What it does not prove |
|---|---|---|
| **Invoice** | An amount was billed | That anything was delivered, or that the customer accepted it |
| **Contract or order** | The customer agreed to buy, at a price | That the obligation was satisfied this period |
| **Delivery or shipping** | Something was transferred | That the customer will pay |
| **Cash receipt** | The customer paid | Nothing about period — cash can arrive in any period |

**Cash is the strongest single link**, because a customer paying real money for a real thing is
hard to fabricate. But it is timing-neutral: a year-end receivable subsequently collected is
perfectly normal. What matters is the combination — *no cash, no receivable, and no delivery
evidence* is the signature of fictitious revenue, and no single link would have surfaced it.

The script computes a **completeness-of-evidence score** per item and reports which links are
missing, so a reviewer can see at a glance whether the population was tested on invoices alone.

## Step 3 — Cutoff and recognition

Cutoff is where legitimate businesses make errors and where aggressive ones push. The script
tests:

- **Revenue recorded in the period with delivery after period end** — recognised before the
  performance obligation was satisfied.
- **Revenue recorded in the period with an invoice dated after period end.**
- **Concentration in the final days of the period.** Reports revenue by day for the last stretch
  of the period and flags a spike. A spike is not itself wrong — many businesses have real
  period-end seasonality — but it is where to look.
- **Subsequent-period credit memos reversing tested revenue.** This is the single highest-yield
  test in the set. Revenue recorded at period end and credited shortly afterward was often never
  a valid sale: goods shipped early, a side agreement permitting return, or a bill-and-hold
  arrangement. Supply `--credit-memos` and the script matches them against the selections.

Recognition timing beyond cutoff — multiple performance obligations, variable consideration,
principal versus agent, licences versus services — is judgment. The script flags contracts whose
terms suggest such a question exists; it does not conclude on the treatment. Route those to the
person signing.

## Step 4 — Deliver

**Workbook tabs:**

1. **Workpaper Summary** — the direction and assertion stated explicitly, the population tie,
   coverage, evidence completeness across the sample, exception counts, and the conclusion block.
2. **Trace Detail** — every selection with GL, invoice, contract, delivery, and cash amounts and
   dates, each difference, the evidence score, and the disposition.
3. **Exceptions** — every difference and every missing link, with the dollar effect.
4. **Cutoff Testing** — items failing cutoff, revenue by day near period end, and the
   subsequent-credit-memo matches.
5. **Evidence Coverage** — a matrix of which links were obtained across the sample. If the sample
   was tested on invoices alone, this tab makes it obvious.
6. **Judgment Items** — recognition questions identified but deliberately not concluded, with
   what would need to be determined and by whom.

**Then, in chat:** the direction and assertion first — one sentence, because it bounds everything
else — then the population tie, the exceptions with dollar effect, cutoff findings, and what needs
a judgment call. If completeness is the client's actual concern, say plainly that this procedure
does not address it.

## What to escalate

- **Revenue with no cash, no receivable, and no delivery evidence.** The fictitious-revenue
  pattern. Escalate regardless of amount.
- **Subsequent credit memos reversing period-end revenue**, particularly clustered by customer or
  by salesperson.
- **Contract terms discovered during testing that the accounting does not reflect** — return
  rights, acceptance clauses, contingent fees, bill-and-hold. A side agreement changes the
  accounting and by design does not appear in the accounting records.
- **Invoices with no contract or order** on a population where contracts are the norm.
- **Related-party revenue** within third-party revenue.
- **A customer whose payments consistently come from a different entity** than the one invoiced.
- **Revenue recognised on a contract signed after period end**, which is unambiguous.
- **Round-dollar revenue transactions at period end** to a single customer.

Report each as a fact with the document reference. Do not characterize intent — describe what the
documents show and let the engagement decide.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry.
Contract terms and customer data never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*

---

# Appendix — bundled files

If you installed the `.skill` package these files are already in place and you can ignore this appendix. If you copied the skill as text, create the files below at the paths shown, alongside your `SKILL.md`. The skill will not run without them.

## `scripts/revenue_trace.py`

```python
#!/usr/bin/env python3
"""
Trace revenue from the GL to invoice, contract, delivery, and cash receipt.

Direction matters and is stated on the workpaper: GL -> source tests OCCURRENCE
and EXISTENCE. It cannot detect unrecorded revenue, because unrecorded revenue is
not in the population sampled from. If completeness is the concern, this is the
wrong procedure.

Gates:
  * an item is not "tested" until the evidence exists - a conclusion of 'agreed'
    with a missing required link or a blank evidence reference is rejected
  * amounts must agree across GL, invoice, contract and cash, or carry a named
    difference
  * the population must tie to the GL revenue total
  * a conclusion may not contradict its own figures

Usage - emit a blank worksheet:
    python3 revenue_trace.py --selections selections.csv \
        --population-total 18422900.00 --period-end 2025-12-31 \
        --template --out trace_worksheet.csv

Usage - test:
    python3 revenue_trace.py --selections trace_worksheet.csv \
        --population-total 18422900.00 --period-end 2025-12-31 \
        --credit-memos subsequent_credits.csv \
        --client "Northwind Systems Inc" \
        --assertion "Occurrence of recorded revenue for FY2025" \
        --out "Northwind - FY2025 Revenue Trace.xlsx"

--selections CSV:
    item_id, customer, gl_date, gl_amount, gl_reference,
    invoice_no, invoice_date, invoice_amount,
    contract_ref, contract_date, contract_price, contract_quantity,
    delivery_ref, delivery_date,
    cash_receipt_ref, cash_receipt_date, cash_amount,
    conclusion, disposition, notes,
    na_reason_contract (optional), na_reason_cash (optional),
    na_reason_delivery (optional), related_party (optional),
    contract_terms_note (optional)

  conclusion: agreed | exception | open_item | n/a_explained

--credit-memos CSV:
    customer, invoice_no, memo_date, amount, reason
"""

from __future__ import annotations

import argparse
import csv
import sys
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
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
PCT = '0.0"%"'
DATEF = "yyyy-mm-dd"

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
TOP = Border(top=Side(style="thin"))

CONCLUSIONS = {"agreed", "exception", "open_item", "n/a_explained"}

COLS = ["item_id", "customer", "gl_date", "gl_amount", "gl_reference",
        "invoice_no", "invoice_date", "invoice_amount",
        "contract_ref", "contract_date", "contract_price", "contract_quantity",
        "delivery_ref", "delivery_date",
        "cash_receipt_ref", "cash_receipt_date", "cash_amount",
        "conclusion", "disposition", "notes",
        "na_reason_contract", "na_reason_cash", "na_reason_delivery",
        "related_party", "contract_terms_note"]

LINKS = ("invoice", "contract", "delivery", "cash")

JUDGMENT_HINTS = ("return right", "right of return", "acceptance", "contingent",
                  "bill and hold", "bill-and-hold", "milestone", "multiple element",
                  "performance obligation", "variable", "rebate", "consignment",
                  "agent", "principal", "licence", "license", "sla", "renewal",
                  "side agreement", "side letter")


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


def load(path: Path) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No selections in {path}")
    out = []
    for i, r in enumerate(rows, start=1):
        x = {k: clean(r.get(k)) for k in COLS}
        x["row"] = i
        x["item_id"] = x["item_id"] or f"R{i:04d}"
        x["conclusion"] = x["conclusion"].lower()
        for f in ("gl_amount", "invoice_amount", "contract_price",
                  "contract_quantity", "cash_amount"):
            x[f] = dec(r.get(f))
        for f in ("gl_date", "invoice_date", "contract_date", "delivery_date",
                  "cash_receipt_date"):
            x[f] = pdate(r.get(f))
        if x["gl_amount"] is None:
            raise ValueError(f"Row {i} ({x['item_id']}): gl_amount is required.")
        out.append(x)
    return out


def load_memos(path: Path | None) -> list[dict]:
    if not path or not path.exists():
        return []
    with path.open(newline="", encoding="utf-8-sig") as fh:
        return [{"customer": clean(r.get("customer")),
                 "invoice_no": clean(r.get("invoice_no")),
                 "memo_date": pdate(r.get("memo_date")),
                 "amount": dec(r.get("amount")) or ZERO,
                 "reason": clean(r.get("reason"))} for r in csv.DictReader(fh)]


# ----------------------------------------------------------------------- tests

def evaluate(rows, period_end: date, memos) -> tuple[list[str], list[str], list[dict]]:
    errors, escal, judgment = [], [], []
    memo_idx = defaultdict(list)
    for m in memos:
        memo_idx[(m["customer"].lower(), m["invoice_no"].lower())].append(m)

    for x in rows:
        n, iid = x["row"], x["item_id"]
        present, missing = [], []

        if x["invoice_no"] or x["invoice_amount"] is not None:
            present.append("invoice")
        else:
            missing.append("invoice")
        if x["contract_ref"] or x["contract_price"] is not None:
            present.append("contract")
        elif x["na_reason_contract"]:
            present.append("contract (n/a documented)")
        else:
            missing.append("contract")
        if x["delivery_ref"] or x["delivery_date"]:
            present.append("delivery")
        elif x["na_reason_delivery"]:
            present.append("delivery (n/a documented)")
        else:
            missing.append("delivery")
        if x["cash_receipt_ref"] or x["cash_amount"] is not None:
            present.append("cash")
        elif x["na_reason_cash"]:
            present.append("cash (n/a documented)")
        else:
            missing.append("cash")

        x["links_present"] = present
        x["links_missing"] = missing
        x["evidence_score"] = len(present)

        # amount agreement
        gl = x["gl_amount"]
        x["diff_invoice"] = (gl - x["invoice_amount"]) if x["invoice_amount"] is not None else None
        contract_amt = None
        if x["contract_price"] is not None:
            q = x["contract_quantity"] if x["contract_quantity"] is not None else Decimal(1)
            contract_amt = x["contract_price"] * q
        x["contract_amount"] = contract_amt
        x["diff_contract"] = (gl - contract_amt) if contract_amt is not None else None
        x["diff_cash"] = (gl - x["cash_amount"]) if x["cash_amount"] is not None else None
        diffs = [d for d in (x["diff_invoice"], x["diff_contract"], x["diff_cash"])
                 if d is not None]
        x["max_diff"] = max((abs(d) for d in diffs), default=ZERO)

        # --- gate checks
        if x["conclusion"] not in CONCLUSIONS:
            errors.append(f"Row {n} ({iid}): conclusion '{x['conclusion']}' is not one of "
                          f"{sorted(CONCLUSIONS)}.")
        if x["conclusion"] == "agreed":
            if missing:
                errors.append(
                    f"Row {n} ({iid}): concluded 'agreed' but these evidence links are "
                    f"missing with no documented reason: {', '.join(missing)}. An item is "
                    f"not tested until the evidence exists - supply the reference, add an "
                    f"na_reason_*, or change the conclusion to 'open_item'.")
            if x["max_diff"] != ZERO:
                errors.append(
                    f"Row {n} ({iid}): concluded 'agreed' but amounts differ by up to "
                    f"{x['max_diff']:,.2f} across the chain. That is an exception, not an "
                    f"agreement.")
        if x["conclusion"] == "exception" and not x["disposition"]:
            errors.append(f"Row {n} ({iid}): exception with no disposition. State what "
                          f"should be done about it.")
        if x["conclusion"] == "n/a_explained" and not x["notes"]:
            errors.append(f"Row {n} ({iid}): concluded 'n/a_explained' with no explanation "
                          f"in notes.")

        # --- cutoff
        x["cutoff_flags"] = []
        if x["gl_date"] and x["gl_date"] <= period_end:
            if x["delivery_date"] and x["delivery_date"] > period_end:
                x["cutoff_flags"].append(
                    f"delivery {x['delivery_date']} is after period end {period_end} - "
                    f"revenue recognised before the performance obligation was satisfied")
            if x["invoice_date"] and x["invoice_date"] > period_end:
                x["cutoff_flags"].append(
                    f"invoice dated {x['invoice_date']}, after period end")
            if x["contract_date"] and x["contract_date"] > period_end:
                x["cutoff_flags"].append(
                    f"CONTRACT SIGNED {x['contract_date']}, after period end - revenue "
                    f"recognised on a contract that did not exist at the reporting date")
                escal.append(f"{iid} ({x['customer']}): contract signed "
                             f"{x['contract_date']}, after period end, with revenue of "
                             f"{gl:,.2f} recorded in the period. Unambiguous.")

        # --- fictitious-revenue pattern
        no_cash = x["cash_amount"] is None and not x["cash_receipt_ref"]
        no_delivery = not x["delivery_ref"] and not x["delivery_date"]
        no_ar_note = "receivable" not in (x["notes"] or "").lower()
        if no_cash and no_delivery and no_ar_note:
            escal.append(f"{iid} ({x['customer']}): revenue of {gl:,.2f} with no cash "
                         f"receipt, no delivery evidence, and no receivable noted. This is "
                         f"the fictitious-revenue pattern - escalate regardless of amount.")

        # --- subsequent credit memos
        hits = memo_idx.get((x["customer"].lower(), x["invoice_no"].lower()), [])
        after = [m for m in hits if m["memo_date"] and m["memo_date"] > period_end]
        x["credit_memos"] = after
        x["credit_memo_total"] = sum((m["amount"] for m in after), ZERO)
        if after:
            escal.append(
                f"{iid} ({x['customer']}): {len(after)} subsequent-period credit memo(s) "
                f"totalling {x['credit_memo_total']:,.2f} against tested revenue of "
                f"{gl:,.2f}. Revenue recorded at period end and credited shortly after was "
                f"often never a valid sale - early shipment, a return right, or "
                f"bill-and-hold.")

        # --- judgment items
        note = (x["contract_terms_note"] or "").lower()
        hit = [h for h in JUDGMENT_HINTS if h in note]
        if hit:
            judgment.append({"item_id": iid, "customer": x["customer"],
                             "amount": gl, "terms": x["contract_terms_note"],
                             "hints": ", ".join(hit)})
        if clean(x["related_party"]).lower() in ("yes", "y", "true", "1"):
            escal.append(f"{iid} ({x['customer']}): related-party revenue of {gl:,.2f} "
                         f"within third-party revenue - classification and disclosure "
                         f"matter.")

    return errors, escal, judgment


# -------------------------------------------------------------------- workbook

def hdr(ws, n, row=1):
    for c in range(1, n + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def widths(ws, w):
    for c, v in w.items():
        ws.column_dimensions[get_column_letter(c)].width = v


def sheet_summary(wb, meta, rows, errors, escal, judgment, pop_total, valid):
    ws = wb.active
    ws.title = "Workpaper Summary"
    widths(ws, {1: 4, 2: 50, 3: 20, 4: 16, 5: 64})
    r = 1

    def line(label, a=None, b=None, *, bold=False, size=11, fill=None, note=""):
        nonlocal r
        c = ws.cell(row=r, column=2, value=label)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c.fill = fill
        for col, v in ((3, a), (4, b)):
            if v is None:
                continue
            cc = ws.cell(row=r, column=col,
                         value=float(v) if isinstance(v, Decimal) else v)
            if isinstance(v, Decimal):
                cc.number_format = MONEY
            cc.font = Font(bold=bold)
        if note:
            n = ws.cell(row=r, column=5, value=note)
            n.font = Font(italic=True, size=9)
            n.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    line("REVENUE TRACE TO SOURCE", bold=True, size=14)
    r += 1

    line("DIRECTION AND ASSERTION", bold=True, size=12, fill=SUB_FILL)
    line("  Direction", "GL -> source", bold=True)
    line("  Assertion tested", meta.get("Assertion") or "(NOT STATED)", bold=True)
    line("  This procedure tests OCCURRENCE and EXISTENCE.", fill=WARN_FILL,
         note="It cannot detect unrecorded revenue, because unrecorded revenue is not in "
              "the population that was sampled from. If completeness is the concern, this "
              "is the wrong procedure and a source-to-GL trace is required.")
    r += 1

    for k, v in meta.items():
        if k == "Assertion":
            continue
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    v = ws.cell(row=r, column=2,
                value="EVIDENCE STANDARD MET" if valid else
                "EVIDENCE STANDARD NOT MET - see the Exceptions tab. An item concluded "
                "without evidence is not a tested item.")
    v.font = OK_FONT if valid else BAD_FONT
    if not valid:
        v.fill = BAD_FILL
    r += 2

    sel_total = sum((x["gl_amount"] for x in rows), ZERO)
    line("POPULATION AND COVERAGE", bold=True, size=12, fill=SUB_FILL)
    line("  Items traced", len(rows))
    line("  Value traced", sel_total)
    if pop_total is not None:
        line("  GL revenue total", pop_total)
        cov = (sel_total / pop_total * 100) if pop_total else ZERO
        line("  Coverage of revenue", None, float(cov))
        ws.cell(row=r - 1, column=4).number_format = PCT
    else:
        line("  GL revenue total", "NOT SUPPLIED", fill=BAD_FILL,
             note="Without the population total, coverage is unknown and the sample "
                  "supports no conclusion about the account.")
    r += 1

    con = Counter(x["conclusion"] for x in rows)
    line("RESULTS", bold=True, size=12, fill=SUB_FILL)
    for k in ("agreed", "exception", "open_item", "n/a_explained"):
        line(f"  {k}", con.get(k, 0))
    r += 1

    line("EVIDENCE COMPLETENESS", bold=True, size=12, fill=SUB_FILL)
    for link in LINKS:
        got = sum(1 for x in rows
                  if any(p.startswith(link) for p in x["links_present"]))
        pct = (got / len(rows) * 100) if rows else 0
        ws.cell(row=r, column=2, value=f"  {link} obtained")
        ws.cell(row=r, column=3, value=f"{got} of {len(rows)}")
        c = ws.cell(row=r, column=4, value=pct)
        c.number_format = PCT
        if pct < 100:
            c.fill = WARN_FILL
        r += 1
    line("  Cash is the strongest single link but is timing-neutral. No cash, no "
         "receivable and no delivery together is the fictitious-revenue signature - and "
         "no single link would surface it.", bold=True)
    r += 1

    cut = [x for x in rows if x["cutoff_flags"]]
    if cut:
        line("CUTOFF FINDINGS", bold=True, size=12, fill=BAD_FILL)
        for x in cut:
            line(f"  {x['item_id']} ({x['customer']}): {'; '.join(x['cutoff_flags'])}")
        r += 1

    if escal:
        line("ESCALATE", bold=True, size=12, fill=BAD_FILL)
        for e in escal:
            line("  " + e)
        r += 1

    if judgment:
        line("JUDGMENT ITEMS - NOT CONCLUDED HERE", bold=True, size=12, fill=WARN_FILL)
        for j in judgment:
            line(f"  {j['item_id']} ({j['customer']}, {j['amount']:,.2f}): "
                 f"{j['hints']} - route the recognition treatment to the signer")
        r += 1

    r += 1
    line("CONCLUSION", bold=True, size=12)
    line("  [Complete after review. State whether recorded revenue tested is supported as "
         "to occurrence, what exceptions were identified, and whether any matter requires "
         "a judgement on recognition. Note explicitly that completeness was not tested by "
         "this procedure.]")
    r += 2
    line("Prepared by / date:  __________________  ____________", bold=True)
    line("Reviewed by / date:  __________________  ____________", bold=True)


def sheet_detail(wb, rows):
    ws = wb.create_sheet("Trace Detail")
    heads = ["Item", "Customer", "GL date", "GL amount", "GL ref",
             "Invoice", "Inv date", "Inv amount", "Diff",
             "Contract", "Contract date", "Contract amount", "Diff",
             "Delivery", "Delivery date",
             "Cash ref", "Cash date", "Cash amount", "Diff",
             "Links", "Evidence score", "Conclusion", "Disposition", "Notes"]
    ws.append(heads)
    hdr(ws, len(heads))
    for x in rows:
        ws.append([x["item_id"], x["customer"], x["gl_date"], float(x["gl_amount"]),
                   x["gl_reference"] or None,
                   x["invoice_no"] or None, x["invoice_date"],
                   float(x["invoice_amount"]) if x["invoice_amount"] is not None else None,
                   float(x["diff_invoice"]) if x["diff_invoice"] is not None else None,
                   x["contract_ref"] or None, x["contract_date"],
                   float(x["contract_amount"]) if x["contract_amount"] is not None else None,
                   float(x["diff_contract"]) if x["diff_contract"] is not None else None,
                   x["delivery_ref"] or None, x["delivery_date"],
                   x["cash_receipt_ref"] or None, x["cash_receipt_date"],
                   float(x["cash_amount"]) if x["cash_amount"] is not None else None,
                   float(x["diff_cash"]) if x["diff_cash"] is not None else None,
                   ", ".join(x["links_present"]), x["evidence_score"],
                   x["conclusion"], x["disposition"] or None, x["notes"] or None])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in (2, 6, 10, 14, 16):
            row[i].number_format = DATEF
        for i in (3, 7, 8, 11, 12, 17, 18):
            row[i].number_format = MONEY
        for i in (8, 12, 18):
            if isinstance(row[i].value, (int, float)) and row[i].value:
                row[i].font, row[i].fill = BAD_FONT, BAD_FILL
        if isinstance(row[20].value, int) and row[20].value < 4:
            row[20].fill = WARN_FILL
        if row[21].value == "exception":
            row[21].font, row[21].fill = BAD_FONT, BAD_FILL
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:X{max(ws.max_row, 2)}"
    widths(ws, {1: 9, 2: 26, 3: 11, 4: 15, 5: 14, 6: 13, 7: 11, 8: 14, 9: 12,
                10: 14, 11: 12, 12: 16, 13: 12, 14: 13, 15: 12, 16: 14, 17: 11,
                18: 14, 19: 12, 20: 42, 21: 9, 22: 15, 23: 28, 24: 46})


def sheet_exceptions(wb, rows, errors):
    ws = wb.create_sheet("Exceptions")
    heads = ["#", "Type", "Item", "Detail", "Dollar effect"]
    ws.append(heads)
    hdr(ws, len(heads))
    n = 0
    for e in errors:
        n += 1
        ws.append([n, "evidence standard", "", e, None])
    for x in rows:
        for label, d in (("invoice", x["diff_invoice"]), ("contract", x["diff_contract"]),
                         ("cash", x["diff_cash"])):
            if d is not None and d != ZERO:
                n += 1
                ws.append([n, f"amount disagreement ({label})", x["item_id"],
                           f"GL {x['gl_amount']:,.2f} vs {label} - difference "
                           f"{d:,.2f}", float(d)])
        for m in x["links_missing"]:
            n += 1
            ws.append([n, "missing evidence", x["item_id"],
                       f"no {m} obtained and no documented reason", None])
        for cf in x["cutoff_flags"]:
            n += 1
            ws.append([n, "cutoff", x["item_id"], cf, float(x["gl_amount"])])
    if n == 0:
        ws.cell(row=2, column=4, value="None - every item carries complete evidence and "
                                       "all amounts agree.")
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[4].number_format = MONEY
        row[3].alignment = Alignment(wrap_text=True, vertical="top")
        if row[1].value == "evidence standard":
            row[1].font, row[1].fill = BAD_FONT, BAD_FILL
    widths(ws, {1: 5, 2: 28, 3: 12, 4: 94, 5: 16})


def sheet_cutoff(wb, rows, period_end, memos):
    ws = wb.create_sheet("Cutoff Testing")
    widths(ws, {1: 4, 2: 30, 3: 16, 4: 14, 5: 70})
    r = 1
    ws.cell(row=r, column=2, value="CUTOFF TESTING").font = Font(bold=True, size=14)
    r += 2
    ws.cell(row=r, column=2, value=f"Period end: {period_end}").font = Font(bold=True)
    r += 2

    ws.cell(row=r, column=2, value="Revenue by day, final 15 days of the period"
            ).font = Font(bold=True, size=12)
    r += 1
    for h, col in zip(["Date", "Items", "Revenue"], (2, 3, 4)):
        c = ws.cell(row=r, column=col, value=h)
        c.fill, c.font = HDR_FILL, HDR_FONT
    r += 1
    per_day = defaultdict(lambda: {"n": 0, "amt": ZERO})
    for x in rows:
        if x["gl_date"] and period_end - timedelta(days=14) <= x["gl_date"] <= period_end:
            per_day[x["gl_date"]]["n"] += 1
            per_day[x["gl_date"]]["amt"] += x["gl_amount"]
    total_window = sum((a["amt"] for a in per_day.values()), ZERO)
    for d in sorted(per_day):
        a = per_day[d]
        ws.cell(row=r, column=2, value=d).number_format = DATEF
        ws.cell(row=r, column=3, value=a["n"])
        c = ws.cell(row=r, column=4, value=float(a["amt"]))
        c.number_format = MONEY
        if total_window and a["amt"] / total_window > Decimal("0.35"):
            c.fill = WARN_FILL
            ws.cell(row=r, column=5,
                    value="This day is more than a third of revenue in the window. Not "
                          "wrong by itself - many businesses have real period-end "
                          "seasonality - but it is where to look."
                    ).alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    if not per_day:
        ws.cell(row=r, column=2, value="No traced items dated in the final 15 days.")
        r += 1
    r += 1

    ws.cell(row=r, column=2, value="Items failing cutoff").font = Font(bold=True, size=12)
    r += 1
    cut = [x for x in rows if x["cutoff_flags"]]
    if not cut:
        ws.cell(row=r, column=2, value="None.").font = OK_FONT
        r += 1
    else:
        for h, col in zip(["Item", "Customer", "GL amount", "Finding"], (2, 3, 4, 5)):
            c = ws.cell(row=r, column=col, value=h)
            c.fill, c.font = HDR_FILL, HDR_FONT
        r += 1
        for x in cut:
            ws.cell(row=r, column=2, value=x["item_id"])
            ws.cell(row=r, column=3, value=x["customer"])
            c = ws.cell(row=r, column=4, value=float(x["gl_amount"]))
            c.number_format = MONEY
            cc = ws.cell(row=r, column=5, value="; ".join(x["cutoff_flags"]))
            cc.alignment = Alignment(wrap_text=True, vertical="top")
            cc.fill = BAD_FILL
            r += 1
    r += 1

    ws.cell(row=r, column=2,
            value="Subsequent-period credit memos against tested revenue"
            ).font = Font(bold=True, size=12)
    r += 1
    hits = [x for x in rows if x["credit_memos"]]
    if not memos:
        ws.cell(row=r, column=2,
                value="NOT PERFORMED - no subsequent credit memos supplied. This is the "
                      "highest-yield test in a revenue programme and is almost never "
                      "requested. Ask for them.").fill = WARN_FILL
        r += 1
    elif not hits:
        ws.cell(row=r, column=2,
                value="No subsequent credit memos matched any tested item.").font = OK_FONT
        r += 1
    else:
        for h, col in zip(["Item", "Customer", "Revenue tested", "Credited after period end"],
                          (2, 3, 4, 5)):
            c = ws.cell(row=r, column=col, value=h)
            c.fill, c.font = HDR_FILL, HDR_FONT
        r += 1
        for x in hits:
            ws.cell(row=r, column=2, value=x["item_id"])
            ws.cell(row=r, column=3, value=x["customer"])
            c1 = ws.cell(row=r, column=4, value=float(x["gl_amount"]))
            c2 = ws.cell(row=r, column=5, value=float(x["credit_memo_total"]))
            c1.number_format = c2.number_format = MONEY
            c2.font, c2.fill = BAD_FONT, BAD_FILL
            r += 1
            for m in x["credit_memos"]:
                ws.cell(row=r, column=3, value=f"    memo {m['memo_date']}: {m['reason']}")
                r += 1
        r += 1
        ws.cell(row=r, column=2,
                value="Revenue recorded at period end and credited shortly afterwards was "
                      "often never a valid sale - goods shipped early, a side agreement "
                      "permitting return, or bill-and-hold. Look for clustering by "
                      "customer or salesperson.").alignment = Alignment(wrap_text=True)


def sheet_coverage(wb, rows):
    ws = wb.create_sheet("Evidence Coverage")
    heads = ["Item", "Customer", "GL amount", "Invoice", "Contract", "Delivery",
             "Cash", "Score", "Missing"]
    ws.append(heads)
    hdr(ws, len(heads))
    for x in sorted(rows, key=lambda y: y["evidence_score"]):
        got = lambda link: ("Y" if any(p == link for p in x["links_present"])
                            else "n/a" if any(p.startswith(link) for p in x["links_present"])
                            else "")
        ws.append([x["item_id"], x["customer"], float(x["gl_amount"]),
                   got("invoice"), got("contract"), got("delivery"), got("cash"),
                   x["evidence_score"], ", ".join(x["links_missing"])])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[2].number_format = MONEY
        for i in (3, 4, 5, 6):
            if not row[i].value:
                row[i].fill = BAD_FILL
            elif row[i].value == "n/a":
                row[i].fill = WARN_FILL
        if isinstance(row[7].value, int) and row[7].value < 4:
            row[7].fill = WARN_FILL
    r = ws.max_row + 2
    ws.cell(row=r, column=1,
            value="If this tab shows invoices obtained and little else, the sample was "
                  "vouched to billing rather than traced to source, and it supports a much "
                  "narrower conclusion than the workpaper may imply."
            ).font = Font(italic=True)
    ws.freeze_panes = "C2"
    widths(ws, {1: 10, 2: 28, 3: 16, 4: 10, 5: 10, 6: 10, 7: 8, 8: 8, 9: 34})


def sheet_judgment(wb, judgment):
    ws = wb.create_sheet("Judgment Items")
    heads = ["Item", "Customer", "Amount", "Terms identified", "Question raised",
             "To be determined by"]
    ws.append(heads)
    hdr(ws, len(heads))
    for j in judgment:
        ws.append([j["item_id"], j["customer"], float(j["amount"]), j["terms"],
                   j["hints"], None])
    if ws.max_row == 1:
        ws.cell(row=2, column=4,
                value="No contract terms were flagged as raising a recognition question.")
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[2].number_format = MONEY
        row[3].alignment = Alignment(wrap_text=True, vertical="top")
    r = ws.max_row + 2
    ws.cell(row=r, column=1,
            value="These are identified, not concluded. Recognition timing beyond cutoff - "
                  "multiple performance obligations, variable consideration, principal "
                  "versus agent, licences versus services - is judgement and belongs with "
                  "the person signing."
            ).font = Font(italic=True)
    widths(ws, {1: 10, 2: 28, 3: 16, 4: 56, 5: 34, 6: 24})


# ------------------------------------------------------------------------ main

def write_template(rows, out: Path) -> None:
    with out.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(COLS)
        for x in rows:
            w.writerow([x.get(c) if not isinstance(x.get(c), (date, Decimal))
                        else (x[c].isoformat() if isinstance(x[c], date)
                              else f"{x[c]:.2f}")
                        for c in COLS])
    print(f"Worksheet template written to {out}")
    print("Fill in the invoice, contract, delivery and cash columns, set a conclusion "
          "for each\nitem, then re-run without --template.")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--selections", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--population-total")
    ap.add_argument("--period-end", required=True)
    ap.add_argument("--credit-memos")
    ap.add_argument("--client", default="")
    ap.add_argument("--assertion", default="")
    ap.add_argument("--template", action="store_true",
                    help="emit a blank trace worksheet instead of testing")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    period_end = pdate(args.period_end)
    rows = load(Path(args.selections))

    if args.template:
        write_template(rows, Path(args.out))
        return 0

    memos = load_memos(Path(args.credit_memos) if args.credit_memos else None)
    errors, escal, judgment = evaluate(rows, period_end, memos)
    pop_total = dec(args.population_total)
    if pop_total is None:
        errors.append("--population-total not supplied. Without the GL revenue total, "
                      "coverage is unknown and the sample supports no conclusion about "
                      "the account.")
    valid = not errors

    sel_total = sum((x["gl_amount"] for x in rows), ZERO)
    con = Counter(x["conclusion"] for x in rows)

    meta = {
        "Client": args.client or "(not stated)",
        "Assertion": args.assertion or "",
        "Period end": str(period_end),
        "Items traced": len(rows),
        "Subsequent credit memos supplied": "yes" if memos else "no",
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    # ---- console
    print("=" * 76)
    print("REVENUE TRACE TO SOURCE")
    print("=" * 76)
    print("Direction : GL -> source.  Tests OCCURRENCE and EXISTENCE.")
    print("            This procedure CANNOT detect unrecorded revenue - unrecorded")
    print("            revenue is not in the population sampled from. If completeness is")
    print("            the concern, a source-to-GL trace is required instead.")
    print(f"Assertion : {args.assertion or '*** NOT STATED - record what this supports ***'}")
    print()
    print(f"Client        : {meta['Client']}    Period end: {period_end}")
    print(f"Items traced  : {len(rows)}   value {sel_total:,.2f}")
    if pop_total:
        print(f"GL revenue    : {pop_total:,.2f}   coverage "
              f"{sel_total / pop_total * 100:.1f}%")
    print(f"  agreed {con['agreed']}   exception {con['exception']}   "
          f"open_item {con['open_item']}   n/a_explained {con['n/a_explained']}")
    print()
    print("Evidence obtained across the sample:")
    for link in LINKS:
        got = sum(1 for x in rows if any(p.startswith(link) for p in x["links_present"]))
        print(f"  {link:<10} {got:>4} of {len(rows)}  "
              f"({got / len(rows) * 100:.0f}%)" if rows else "")
    cut = [x for x in rows if x["cutoff_flags"]]
    if cut:
        print(f"\nCUTOFF FINDINGS ({len(cut)}):")
        for x in cut[:12]:
            print(f"  ! {x['item_id']} ({x['customer'][:22]}): "
                  f"{'; '.join(x['cutoff_flags'])}")
    if errors:
        print(f"\nEVIDENCE STANDARD FAILURES ({len(errors)}):")
        for e in errors[:25]:
            print(f"  ! {e}")
        if len(errors) > 25:
            print(f"  ... and {len(errors) - 25} more")
    if escal:
        print(f"\nESCALATE ({len(escal)}):")
        for e in escal[:15]:
            print(f"  ! {e}")
    if judgment:
        print(f"\nJUDGMENT ITEMS - identified, not concluded ({len(judgment)}):")
        for j in judgment[:10]:
            print(f"  ? {j['item_id']} ({j['customer'][:22]}): {j['hints']}")

    if not valid and not args.force:
        print("\n" + "=" * 76)
        print("WORKBOOK NOT WRITTEN.")
        print("An item concluded without its evidence is not a tested item. Obtain the")
        print("missing documents, or change the conclusion to 'open_item', and re-run.")
        print("=" * 76)
        return 1

    wb = Workbook()
    sheet_summary(wb, meta, rows, errors, escal, judgment, pop_total, valid)
    sheet_detail(wb, rows)
    sheet_exceptions(wb, rows, errors)
    sheet_cutoff(wb, rows, period_end, memos)
    sheet_coverage(wb, rows)
    sheet_judgment(wb, judgment)
    wb.active = 0
    out = Path(args.out)
    if not valid:
        out = out.with_name(out.stem + " [EVIDENCE STANDARD NOT MET]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"\n{'EVIDENCE STANDARD MET' if valid else 'NOT MET (forced)'}.  Workbook: {out}")
    return 0 if valid else 1


if __name__ == "__main__":
    sys.exit(main())
```

