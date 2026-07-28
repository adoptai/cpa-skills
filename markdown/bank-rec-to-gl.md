---
name: bank-rec-to-gl
description: Reconcile a bank statement to the general ledger cash account — match transactions, identify deposits in transit, outstanding checks, and unrecorded bank items, then produce a signable four-column reconciliation workpaper that proves adjusted bank balance equals adjusted book balance. Use this whenever the user mentions bank reconciliation, reconciling cash, "tie the bank to the GL," unreconciled items, outstanding checks, deposits in transit, an unexplained cash difference, a cash account that won't balance, month-end close of cash, or asks why the bank and the books disagree. Also trigger on "reconcile these," "match the bank to QuickBooks/NetSuite/Sage," "our cash is off by," "clear the old outstanding checks," or when a bank statement and a GL detail export are provided together. Runs fully local — no client cash data leaves the machine.
---

# Bank Reconciliation → General Ledger

A reconciliation is not a matching exercise. It is a **proof that two independently
maintained records of the same cash, which legitimately differ, differ only for reasons you
can name.** Matching is the mechanical part. The reconciliation is the explanation.

The output is a workpaper someone signs. That means every reconciling item has a cause, a
dollar amount, and an expected resolution date — and the unexplained difference line reads
zero. Not "close to zero." Zero.

## The one rule that matters most

**Never plug.** If a difference cannot be explained, it stays visible as an unexplained
difference and the reconciliation is reported as incomplete. A plugged reconciliation is how
misappropriation survives twelve consecutive months of close. The most common form of a plug
is a "reconciling item" with a vague label — `Timing`, `Bank adj`, `Misc difference`,
`Unidentified deposit` — that never resolves in a subsequent month. Treat any such item
inherited from a prior period as an open finding, not as an accepted balance.

## Inputs

You need three things. Ask for whatever is missing before starting.

1. **Bank side** — statement transactions with date, description, amount, and direction.
   If only a PDF statement is available, extract it first (the `bank-statement-to-excel`
   skill does this with a completeness proof; do not reconcile off an unproven extract).
   You also need the statement's opening and closing balance.
2. **Book side** — GL detail for the cash account for the same period: date, memo/payee,
   reference/check number, debit, credit, and the account's opening and closing balance per
   the trial balance.
3. **Prior-period reconciliation**, if one exists — specifically the list of items that were
   outstanding at the prior close. Without it you cannot tell a genuinely new outstanding
   check from one that has been outstanding for nine months.

Confirm before matching: same account, same period, and the GL closing balance agrees to the
trial balance. Reconciling to a GL balance that doesn't agree to the TB wastes the entire
exercise.

## Step 1 — Match

```bash
python3 scripts/reconcile.py \
  --bank bank.csv --book book.csv \
  --opening-bank 42150.22 --closing-bank 38904.71 \
  --opening-book 41980.22 --closing-book 39310.71 \
  --prior-outstanding prior_items.csv \
  --date-tolerance 5 \
  --out "Acme - Cash x4471 - 2026-01 Bank Rec.xlsx"
```

Matching runs as a cascade, most reliable first. Once a transaction is matched it is removed
from the pool, so a weaker rule can never steal a transaction from a stronger one. Every
match records **which rule matched it** — a reviewer needs to know a $4,000 item was matched
on exact amount and date, not on amount alone within a five-day window.

| Pass | Rule | Notes |
|---|---|---|
| 1 | Check number + amount | Strongest evidence available. Run before anything date-based. |
| 2 | Exact amount + exact date | |
| 3 | Exact amount within date tolerance | Default ±5 days for float/posting lag. Ask before widening. |
| 4 | Amount + description similarity | Requires a token overlap threshold, not a loose fuzzy score. |
| 5 | Many-to-one | Several GL receipts deposited as one bank deposit — very common. Subset-sum, capped to keep it honest. |
| 6 | One-to-many | One GL entry settling as several bank items. |

Everything unmatched after pass 6 is a reconciling item and must be classified by hand in
Step 2. **The script does not classify.** Classification is judgment and it is where the
value of a CPA actually sits.

## Step 2 — Classify every unmatched item

This is the part that cannot be automated, and the part a reviewer reads.

**Unmatched on the book side** (in the GL, not at the bank):

- *Deposit in transit* — receipt recorded before period end, deposited/cleared after. Legitimate
  timing. Verify it actually cleared in the subsequent period; note the clearing date. A
  deposit in transit that has not cleared by the time you're reconciling is not timing — it is
  an exception, and it is one of the classic fictitious-revenue and lapping patterns.
- *Outstanding check* — issued and recorded, not yet presented. Legitimate timing. Record the
  issue date, because age is the whole story: see stale-check handling below.
- *Book error* — wrong amount, wrong period, duplicate posting, wrong account. Requires a
  correcting journal entry, not a reconciling item. Draft the JE.

**Unmatched on the bank side** (at the bank, not in the GL):

- *Unrecorded bank item* — service charges, analysis fees, wire fees, interest earned, NSF
  returns and their fees, credit-card processor holdbacks, sweep transfers, direct debits set
  up outside AP. These are real economic events the books simply do not know about yet. They
  require a journal entry, and that JE is a deliverable of this engagement.
- *Bank error* — rare and worth naming loudly when found. Notify the bank; most agreements
  carry a 30–60 day notification deadline, so timing is substantive.
- *Unidentified* — the one you must never bury. An unidentified deposit or, far more
  seriously, an unidentified withdrawal is escalated in the memo regardless of size. Amount
  is not the test for escalation here; the absence of an explanation is.

## Step 3 — Build the reconciliation in proper form

Both sides converge, and the two adjusted balances must be equal:

```
  Balance per bank statement                              38,904.71
  Add:  Deposits in transit                                4,200.00
  Less: Outstanding checks                               ( 3,794.00)
  Add/(less): Bank errors                                      0.00
  = ADJUSTED BANK BALANCE                                 39,310.71

  Balance per general ledger                              39,310.71
  Add:  Unrecorded credits (interest, bank credits)            0.00
  Less: Unrecorded debits (service charges, NSF, fees)         0.00
  Add/(less): Book errors                                      0.00
  = ADJUSTED BOOK BALANCE                                 39,310.71

  UNEXPLAINED DIFFERENCE                                       0.00   <-- must be 0.00
```

The script computes both columns and the difference. If the difference is not zero it exits
non-zero and does not produce a clean workpaper. Note the reconciliation is *to the adjusted
book balance* — the adjusted book balance is what the GL should say once the required journal
entries are posted, which is why those JEs are part of the deliverable and not an
afterthought.

## Step 4 — Age the outstanding items

Aging is where a reconciliation stops being clerical.

- **Over 90 days outstanding** — investigate. Was it delivered? Is the payee waiting? Is it
  lost?
- **Over 6 months** — presumptively stale. Most bank agreements make a check stale-dated at
  six months. Consider void-and-reissue, and record the escheatment question.
- **Unclaimed property** — uncashed checks are reportable to the state after a dormancy
  period (commonly 3–5 years, varying by state and property type, generally the payee's
  state of last known address). Flag candidates. Do not assert the specific dormancy period
  or filing deadline for a given state from memory — name the item as a reportable candidate
  and tell the user to confirm against the applicable state's current rules. Being wrong here
  creates real penalty exposure.
- **Recurring identical reconciling items** — the same amount appearing as a reconciling item
  in consecutive months is almost never timing. It is a systemic posting error or a duplicate
  that has been rolling forward unexamined. Call it out by name.

## Step 5 — Deliver

**Workbook tabs:**

1. **Reconciliation** — the four-column form above. Signature and date block for preparer
   and reviewer. This is the workpaper.
2. **Reconciling Items** — every item: side, date, description, amount, classification,
   age in days, expected resolution, whether a JE is required, and who owns it.
3. **Proposed Journal Entries** — debit/credit ready to post, for the unrecorded bank items
   and book errors. Each with account, amount, and a one-line explanation. This is usually
   the single most-used tab.
4. **Matched Detail** — every matched pair with the rule that matched it and the day
   variance. Lets a reviewer sample and re-perform.
5. **Aging** — outstanding checks and deposits in transit bucketed 0–30 / 31–60 / 61–90 /
   91–180 / 180+, with stale and unclaimed-property candidates flagged.
6. **Exceptions & Memo** — unidentified items, bank errors, suspected duplicates, recurring
   items, control observations, and anything a reviewer must see before signing.

**Then write a short memo in chat**: period and account, bank and book closing balances,
count and total of reconciling items by classification, the required journal entries, the
unexplained difference (zero, or the reconciliation is incomplete), and any escalation. Lead
with escalations if there are any. Three sentences before the detail.

## Control observations worth raising unprompted

You will see these while reconciling, and they are worth more to the client than the
reconciliation itself:

- Checks clearing out of sequence, or gaps in the check sequence
- Round-dollar disbursements to unfamiliar payees
- Transfers between company accounts that appear on only one side (kiting pattern)
- Bank items with no corresponding AP record — payments set up outside the AP process
- The same person preparing the reconciliation and having disbursement authority
- Reconciliations prepared more than 30 days after period end, or unsigned by a reviewer

Raise them in the memo as observations, framed as control matters for management, not as
audit conclusions.

## Security posture

Fully local: `pandas`/`openpyxl` only. No network calls, no uploads, no telemetry. Inputs are
read-only. Intermediate files stay in the working directory and can be shredded after
delivery.

## Dependencies

```bash
pip install pandas openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*

---

# Appendix — bundled files

If you installed the `.skill` package these files are already in place and you can ignore this appendix. If you copied the skill as text, create the files below at the paths shown, alongside your `SKILL.md`. The skill will not run without them.

## `scripts/reconcile.py`

```python
#!/usr/bin/env python3
"""
Match bank transactions to GL cash detail and build a four-column bank
reconciliation workpaper.

The matcher is a cascade: strongest evidence first, and a matched transaction
leaves the pool so a weaker rule can never steal it from a stronger one. Every
match records the rule that produced it.

What this script does NOT do: classify unmatched items. Deposit-in-transit vs.
book error vs. unidentified withdrawal is judgment. Pass your classifications
back in with --classifications to get a complete workpaper.

Usage:
    python3 reconcile.py --bank bank.csv --book book.csv \
        --opening-bank 42150.22 --closing-bank 38904.71 \
        --opening-book 41980.22 --closing-book 39310.71 \
        --out "Acme - Cash x4471 - 2026-01 Bank Rec.xlsx"

Input CSVs (both sides, same shape):
    date        YYYY-MM-DD                    required
    description free text                     required
    amount      signed: + increases cash, - decreases cash   required
    check_no    optional
    reference   optional
    doc_id      optional, your own row key

--classifications CSV (produced after you review unmatched.csv):
    side, key, classification, expected_resolution, je_required, notes
      side           = bank | book
      key            = the Key value from the unmatched listing
      classification = deposit_in_transit | outstanding_check | book_error |
                       unrecorded_bank_debit | unrecorded_bank_credit |
                       bank_error | unidentified

--prior-outstanding CSV (optional, for aging):
    date, description, amount, check_no

Money is Decimal throughout. Nothing is rounded.
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
    "pmt", "ach", "deposit", "withdrawal", "transfer", "debit", "credit", "check",
    "chk", "ref", "no", "id", "to", "from", "of", "on", "trn", "des", "ppd", "ccd",
}

CLASSIFICATIONS = {
    "deposit_in_transit":     ("bank_add",  "Deposits in transit"),
    "outstanding_check":      ("bank_less", "Outstanding checks"),
    "bank_error":             ("bank_adj",  "Bank errors"),
    "unrecorded_bank_credit": ("book_add",  "Unrecorded credits (interest, bank credits)"),
    "unrecorded_bank_debit":  ("book_less", "Unrecorded debits (service charges, NSF, fees)"),
    "book_error":             ("book_adj",  "Book errors"),
    "unidentified":           ("unexplained", "UNIDENTIFIED - not explained"),
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
        raise ValueError("Every transaction needs a date.")
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d-%b-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Unrecognized date {raw!r}. Prefer ISO YYYY-MM-DD.")


def tokens(text: str) -> set[str]:
    words = re.findall(r"[a-z0-9]+", (text or "").lower())
    return {w for w in words if len(w) > 2 and w not in STOPWORDS and not w.isdigit()}


def load_side(path: Path, side: str) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        raw = list(csv.DictReader(fh))
    if not raw:
        sys.exit(f"No rows in {path}")
    prefix = "BNK" if side == "bank" else "GL"
    out = []
    for i, r in enumerate(raw, start=1):
        amt = dec(r.get("amount"))
        if amt is None:
            raise ValueError(f"{side} row {i}: amount is required and must be signed "
                             f"(+ increases cash, - decreases cash).")
        if amt == ZERO:
            raise ValueError(f"{side} row {i}: zero-amount row. Remove it or explain it "
                             f"as an exception; it is not a cash transaction.")
        desc = " ".join((r.get("description") or "").split())
        out.append({
            "side": side,
            "key": f"{prefix}{i:04d}",
            "doc_id": (r.get("doc_id") or "").strip(),
            "date": pdate(r.get("date")),
            "description": desc,
            "check_no": re.sub(r"\D", "", str(r.get("check_no") or "")) or "",
            "reference": (r.get("reference") or "").strip(),
            "amount": amt,
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


def match_all(bank: list[dict], book: list[dict], tol: int,
              token_thresh: float, subset_cap: int) -> list[dict]:
    pairs: list[dict] = []

    def open_side(rows):
        return [r for r in rows if not r["matched"]]

    def record(a, b, rule):
        _link(a, b, rule)
        pairs.append({"rule": rule, "bank": a, "book": b})

    # Pass 1 - check number + amount
    idx = defaultdict(list)
    for bk in book:
        if bk["check_no"]:
            idx[(bk["check_no"], bk["amount"])].append(bk)
    for bn in bank:
        if bn["matched"] or not bn["check_no"]:
            continue
        for cand in idx.get((bn["check_no"], bn["amount"]), []):
            if not cand["matched"]:
                record(bn, cand, "1 check# + amount")
                break

    # Pass 2 - exact amount + exact date
    idx = defaultdict(list)
    for bk in open_side(book):
        idx[(bk["amount"], bk["date"])].append(bk)
    for bn in open_side(bank):
        for cand in idx.get((bn["amount"], bn["date"]), []):
            if not cand["matched"]:
                record(bn, cand, "2 amount + date")
                break

    # Pass 3 - exact amount within date tolerance (nearest date wins)
    idx = defaultdict(list)
    for bk in open_side(book):
        idx[bk["amount"]].append(bk)
    for bn in open_side(bank):
        cands = [c for c in idx.get(bn["amount"], [])
                 if not c["matched"] and abs((bn["date"] - c["date"]).days) <= tol]
        if cands:
            best = min(cands, key=lambda c: abs((bn["date"] - c["date"]).days))
            record(bn, best, f"3 amount, date within {tol}d")

    # Pass 4 - amount + description token overlap
    for bn in open_side(bank):
        best, best_score = None, 0.0
        for bk in open_side(book):
            if bk["amount"] != bn["amount"]:
                continue
            if not bn["tokens"] or not bk["tokens"]:
                continue
            inter = len(bn["tokens"] & bk["tokens"])
            score = inter / min(len(bn["tokens"]), len(bk["tokens"]))
            if score > best_score:
                best, best_score = bk, score
        if best and best_score >= token_thresh:
            record(bn, best, f"4 amount + description ({best_score:.0%} overlap)")

    # Pass 5 - many book rows -> one bank row (batched deposits)
    for bn in open_side(bank):
        pool = [b for b in open_side(book)
                if (b["amount"] > ZERO) == (bn["amount"] > ZERO)
                and abs((bn["date"] - b["date"]).days) <= tol
                and abs(b["amount"]) <= abs(bn["amount"])]
        pool.sort(key=lambda b: abs(b["amount"]), reverse=True)
        pool = pool[:subset_cap]
        found = None
        for size in range(2, min(len(pool), 6) + 1):
            for combo in itertools.combinations(pool, size):
                if sum((c["amount"] for c in combo), ZERO) == bn["amount"]:
                    found = combo
                    break
            if found:
                break
        if found:
            rule = f"5 many-to-one ({len(found)} GL items batched)"
            for c in found:
                record(bn, c, rule)
            bn["matched"] = True

    # Pass 6 - one book row -> many bank rows (split settlement)
    for bk in open_side(book):
        pool = [b for b in open_side(bank)
                if (b["amount"] > ZERO) == (bk["amount"] > ZERO)
                and abs((b["date"] - bk["date"]).days) <= tol
                and abs(b["amount"]) <= abs(bk["amount"])]
        pool.sort(key=lambda b: abs(b["amount"]), reverse=True)
        pool = pool[:subset_cap]
        found = None
        for size in range(2, min(len(pool), 6) + 1):
            for combo in itertools.combinations(pool, size):
                if sum((c["amount"] for c in combo), ZERO) == bk["amount"]:
                    found = combo
                    break
            if found:
                break
        if found:
            rule = f"6 one-to-many ({len(found)} bank items)"
            for c in found:
                record(c, bk, rule)
            bk["matched"] = True

    return pairs


# ------------------------------------------------------------------ workpaper

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
                "expected_resolution": (r.get("expected_resolution") or "").strip(),
                "je_required": (r.get("je_required") or "").strip(),
                "notes": (r.get("notes") or "").strip(),
            }
    return out


def build_rec(unmatched: list[dict], cls: dict, closing_bank: Decimal,
              closing_book: Decimal) -> dict:
    buckets = defaultdict(lambda: {"total": ZERO, "items": []})
    unclassified = []
    for r in unmatched:
        c = cls.get(r["key"], {}).get("classification", "")
        if not c:
            unclassified.append(r)
            continue
        group, _label = CLASSIFICATIONS[c]
        buckets[c]["total"] += r["amount"]
        buckets[c]["items"].append(r)

    def tot(name):
        return buckets[name]["total"] if name in buckets else ZERO

    dit = tot("deposit_in_transit")
    osc = tot("outstanding_check")           # negative amounts
    bnk_err = tot("bank_error")
    unrec_cr = tot("unrecorded_bank_credit")
    unrec_dr = tot("unrecorded_bank_debit")  # negative amounts
    bk_err = tot("book_error")
    unident = tot("unidentified")
    unclass_total = sum((r["amount"] for r in unclassified), ZERO)

    adj_bank = closing_bank + dit + osc + bnk_err
    adj_book = closing_book + unrec_cr + unrec_dr + bk_err
    diff = adj_bank - adj_book

    return {
        "buckets": buckets, "unclassified": unclassified,
        "dit": dit, "osc": osc, "bnk_err": bnk_err,
        "unrec_cr": unrec_cr, "unrec_dr": unrec_dr, "bk_err": bk_err,
        "unidentified": unident, "unclassified_total": unclass_total,
        "adj_bank": adj_bank, "adj_book": adj_book, "difference": diff,
        "reconciled": diff == ZERO and unident == ZERO and not unclassified,
    }


def hdr(ws, n, row=1):
    for c in range(1, n + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 28


def widths(ws, w):
    for c, v in w.items():
        ws.column_dimensions[get_column_letter(c)].width = v


def sheet_reconciliation(wb, rec, meta, closing_bank, closing_book):
    ws = wb.active
    ws.title = "Reconciliation"
    widths(ws, {1: 4, 2: 52, 3: 16, 4: 16, 5: 40})
    r = 1

    def line(label, val=None, *, bold=False, indent=0, border=None, note=""):
        nonlocal r
        c = ws.cell(row=r, column=2, value=("    " * indent) + label)
        c.font = Font(bold=bold)
        if val is not None:
            m = ws.cell(row=r, column=4 if bold else 3, value=float(val))
            m.number_format = MONEY
            m.font = Font(bold=bold)
            if border:
                m.border = border
        if note:
            ws.cell(row=r, column=5, value=note).font = Font(italic=True, size=9)
        r += 1

    ws.cell(row=r, column=2, value="BANK RECONCILIATION").font = Font(bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    v = ws.cell(row=r, column=2,
                value="RECONCILED - adjusted bank equals adjusted book"
                if rec["reconciled"] else
                "NOT RECONCILED - see unexplained difference below")
    v.font = OK_FONT if rec["reconciled"] else BAD_FONT
    if not rec["reconciled"]:
        v.fill = BAD_FILL
    r += 2

    # --- bank column
    c = ws.cell(row=r, column=2, value="BANK SIDE")
    c.font, c.fill = Font(bold=True, size=12), SUB_FILL
    r += 1
    line("Balance per bank statement", closing_bank)
    line("Add: Deposits in transit", rec["dit"], indent=1,
         note="receipts recorded before period end, cleared after")
    line("Less: Outstanding checks", rec["osc"], indent=1,
         note="issued and recorded, not yet presented")
    line("Add/(less): Bank errors", rec["bnk_err"], indent=1,
         note="notify the bank - agreements carry notification deadlines" if rec["bnk_err"] else "")
    line("ADJUSTED BANK BALANCE", rec["adj_bank"], bold=True, border=TOP)
    r += 1

    # --- book column
    c = ws.cell(row=r, column=2, value="BOOK SIDE")
    c.font, c.fill = Font(bold=True, size=12), SUB_FILL
    r += 1
    line("Balance per general ledger", closing_book)
    line("Add: Unrecorded credits (interest, bank credits)", rec["unrec_cr"], indent=1,
         note="requires a journal entry - see Proposed Journal Entries")
    line("Less: Unrecorded debits (service charges, NSF, fees)", rec["unrec_dr"], indent=1,
         note="requires a journal entry - see Proposed Journal Entries")
    line("Add/(less): Book errors", rec["bk_err"], indent=1,
         note="requires a correcting journal entry")
    line("ADJUSTED BOOK BALANCE", rec["adj_book"], bold=True, border=TOP)
    r += 1

    line("UNEXPLAINED DIFFERENCE", rec["difference"], bold=True, border=DBL)
    dc = ws.cell(row=r - 1, column=4)
    dc.font = OK_FONT if rec["difference"] == ZERO else BAD_FONT
    if rec["difference"] != ZERO:
        dc.fill = BAD_FILL
    r += 1

    if rec["unidentified"] != ZERO:
        c = ws.cell(row=r, column=2,
                    value=f"UNIDENTIFIED ITEMS INCLUDED ABOVE: "
                          f"{rec['unidentified']:,.2f} - ESCALATE. An unidentified "
                          f"withdrawal is escalated regardless of amount.")
        c.font, c.fill = BAD_FONT, BAD_FILL
        r += 1
    if rec["unclassified"]:
        c = ws.cell(row=r, column=2,
                    value=f"{len(rec['unclassified'])} unmatched item(s) not yet "
                          f"classified ({rec['unclassified_total']:,.2f}). The "
                          f"reconciliation is incomplete until each is classified. "
                          f"Do not plug.")
        c.font, c.fill = BAD_FONT, WARN_FILL
        r += 1
    r += 2

    ws.cell(row=r, column=2, value="Prepared by / date:").font = Font(bold=True)
    ws.cell(row=r, column=3, value="__________________  ____________")
    r += 1
    ws.cell(row=r, column=2, value="Reviewed by / date:").font = Font(bold=True)
    ws.cell(row=r, column=3, value="__________________  ____________")


def sheet_items(wb, unmatched, cls, asof: date):
    ws = wb.create_sheet("Reconciling Items")
    heads = ["Key", "Side", "Date", "Age (d)", "Description", "Check No",
             "Amount", "Classification", "Expected resolution", "JE req'd", "Notes"]
    ws.append(heads)
    hdr(ws, len(heads))
    for r in sorted(unmatched, key=lambda x: (x["side"], x["date"])):
        info = cls.get(r["key"], {})
        c = info.get("classification", "")
        ws.append([
            r["key"], r["side"], r["date"], (asof - r["date"]).days,
            r["description"], r["check_no"], float(r["amount"]),
            c or "*** UNCLASSIFIED ***",
            info.get("expected_resolution", ""), info.get("je_required", ""),
            info.get("notes", ""),
        ])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[2].number_format = DATEF
        row[6].number_format = MONEY
        if row[7].value in ("*** UNCLASSIFIED ***", "unidentified"):
            row[7].font, row[7].fill = BAD_FONT, BAD_FILL
        if isinstance(row[3].value, int) and row[3].value > 180:
            row[3].fill = WARN_FILL
    if ws.max_row == 1:
        ws.cell(row=2, column=5, value="None - every transaction matched.")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:K{max(ws.max_row, 2)}"
    widths(ws, {1: 8, 2: 7, 3: 12, 4: 9, 5: 50, 6: 10, 7: 14, 8: 24, 9: 22, 10: 10, 11: 40})


def sheet_jes(wb, unmatched, cls):
    ws = wb.create_sheet("Proposed Journal Entries")
    heads = ["#", "Date", "Account (fill in)", "Debit", "Credit",
             "Memo", "Source key", "Basis"]
    ws.append(heads)
    hdr(ws, len(heads))

    je_kinds = {"unrecorded_bank_debit", "unrecorded_bank_credit", "book_error", "bank_error"}
    n = 0
    for r in sorted(unmatched, key=lambda x: x["date"]):
        c = cls.get(r["key"], {}).get("classification", "")
        if c not in je_kinds:
            continue
        n += 1
        amt, memo = r["amount"], r["description"][:70]
        basis = {
            "unrecorded_bank_debit": "Bank charge/debit not on the books",
            "unrecorded_bank_credit": "Bank credit/interest not on the books",
            "book_error": "Correcting entry - GL posted incorrectly",
            "bank_error": "Bank error - notify bank; book only if agreed",
        }[c]
        if amt < ZERO:
            ws.append([n, r["date"], "<expense / clearing account>", float(-amt), None,
                       memo, r["key"], basis])
            ws.append(["", "", "  Cash", None, float(-amt), "", "", ""])
        else:
            ws.append([n, r["date"], "Cash", float(amt), None, memo, r["key"], basis])
            ws.append(["", "", "  <income / clearing account>", None, float(amt), "", "", ""])

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[1].number_format = DATEF
        row[3].number_format = MONEY
        row[4].number_format = MONEY
    if n == 0:
        ws.cell(row=2, column=3,
                value="No journal entries required - no unrecorded bank items or "
                      "book errors identified.")
    else:
        r = ws.max_row + 2
        ws.cell(row=r, column=3, value="TOTALS (must be equal)").font = Font(bold=True)
        for col, letter in ((4, "D"), (5, "E")):
            c = ws.cell(row=r, column=col,
                        value=f"=SUM({letter}2:{letter}{ws.max_row - 1})")
            c.number_format, c.font = MONEY, Font(bold=True)
            c.border = TOP
    widths(ws, {1: 5, 2: 12, 3: 36, 4: 14, 5: 14, 6: 50, 7: 10, 8: 40})


def sheet_matched(wb, pairs):
    ws = wb.create_sheet("Matched Detail")
    heads = ["Rule", "Bank key", "Bank date", "Bank description", "Bank amount",
             "Book key", "Book date", "Book description", "Book amount", "Day var"]
    ws.append(heads)
    hdr(ws, len(heads))
    for p in sorted(pairs, key=lambda x: x["rule"]):
        bn, bk = p["bank"], p["book"]
        ws.append([p["rule"], bn["key"], bn["date"], bn["description"], float(bn["amount"]),
                   bk["key"], bk["date"], bk["description"], float(bk["amount"]),
                   (bn["date"] - bk["date"]).days])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[2].number_format = DATEF
        row[6].number_format = DATEF
        row[4].number_format = MONEY
        row[8].number_format = MONEY
        if isinstance(row[9].value, int) and abs(row[9].value) > 5:
            row[9].fill = WARN_FILL
    if ws.max_row == 1:
        ws.cell(row=2, column=1, value="No matches.")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:J{max(ws.max_row, 2)}"
    widths(ws, {1: 30, 2: 9, 3: 12, 4: 42, 5: 14, 6: 9, 7: 12, 8: 42, 9: 14, 10: 9})


def sheet_aging(wb, unmatched, cls, prior: list[dict], asof: date):
    ws = wb.create_sheet("Aging")
    heads = ["Key", "Date", "Age (days)", "Bucket", "Description", "Check No",
             "Amount", "Classification", "Flag"]
    ws.append(heads)
    hdr(ws, len(heads))

    def bucket(d):
        for lim, name in ((30, "0-30"), (60, "31-60"), (90, "61-90"), (180, "91-180")):
            if d <= lim:
                return name
        return "180+"

    aged = [(r, cls.get(r["key"], {}).get("classification", "")) for r in unmatched]
    aged = [(r, c) for r, c in aged if c in ("outstanding_check", "deposit_in_transit")]
    for r in prior:
        aged.append((r, "outstanding_check (prior period)"))

    for r, c in sorted(aged, key=lambda x: (asof - x[0]["date"]).days, reverse=True):
        age = (asof - r["date"]).days
        flags = []
        if c.startswith("outstanding_check"):
            if age > 180:
                flags.append("STALE (>6 months) - consider void/reissue")
                flags.append("unclaimed-property candidate - confirm state rules")
            elif age > 90:
                flags.append("investigate - delivered? lost?")
        if c == "deposit_in_transit" and age > 10:
            flags.append("DIT outstanding >10 days - not timing; investigate")
        ws.append([r.get("key", ""), r["date"], age, bucket(age), r["description"],
                   r.get("check_no", ""), float(r["amount"]), c, "; ".join(flags)])

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[1].number_format = DATEF
        row[6].number_format = MONEY
        if row[8].value:
            row[8].fill = WARN_FILL
    if ws.max_row == 1:
        ws.cell(row=2, column=5,
                value="No outstanding checks or deposits in transit to age.")
    ws.freeze_panes = "A2"
    widths(ws, {1: 8, 2: 12, 3: 11, 4: 10, 5: 44, 6: 10, 7: 14, 8: 30, 9: 56})


def sheet_memo(wb, rec, unmatched, cls, pairs, meta, flags: list[str]):
    ws = wb.create_sheet("Exceptions & Memo")
    widths(ws, {1: 4, 2: 100})
    r = 1

    def para(text, bold=False, size=11):
        nonlocal r
        c = ws.cell(row=r, column=2, value=text)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    para("RECONCILIATION MEMO", bold=True, size=14)
    r += 1
    para(f"Account: {meta.get('Account', '')}    Period: {meta.get('Period', '')}", bold=True)
    para(f"Transactions matched: {len(pairs)}   Reconciling items: {len(unmatched)}")
    para("Status: " + ("RECONCILED" if rec["reconciled"]
                       else "NOT RECONCILED - see below"), bold=True)
    r += 1

    if flags:
        para("EXCEPTIONS AND ESCALATIONS", bold=True, size=12)
        for f in flags:
            para("- " + f)
        r += 1

    para("CONTROL OBSERVATIONS TO CONSIDER", bold=True, size=12)
    for line in [
        "Checks clearing out of sequence, or gaps in the check sequence",
        "Round-dollar disbursements to unfamiliar payees",
        "Inter-company transfers appearing on only one side (kiting pattern)",
        "Bank debits with no corresponding AP record - payments outside the AP process",
        "The reconciliation preparer also holding disbursement authority",
        "Reconciliations prepared >30 days after period end, or unsigned by a reviewer",
    ]:
        para("- " + line)
    r += 1
    para("Frame these to management as control observations, not audit conclusions.",
         bold=False)
    r += 1
    para("Reminder: never plug. A reconciling item with a vague label (Timing, Bank "
         "adj, Misc) that does not resolve in the following period is an open finding, "
         "not an accepted balance.", bold=True)


# ------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--bank", required=True)
    ap.add_argument("--book", required=True)
    ap.add_argument("--closing-bank", required=True)
    ap.add_argument("--closing-book", required=True)
    ap.add_argument("--opening-bank")
    ap.add_argument("--opening-book")
    ap.add_argument("--classifications")
    ap.add_argument("--prior-outstanding")
    ap.add_argument("--as-of", help="period end YYYY-MM-DD, for aging (default: latest txn)")
    ap.add_argument("--date-tolerance", type=int, default=5)
    ap.add_argument("--token-threshold", type=float, default=0.6)
    ap.add_argument("--subset-cap", type=int, default=14,
                    help="max candidates considered in many-to-one matching")
    ap.add_argument("--account-label", default="")
    ap.add_argument("--period", default="")
    ap.add_argument("--out", required=True)
    ap.add_argument("--unmatched-out", default="unmatched.csv")
    args = ap.parse_args()

    bank = load_side(Path(args.bank), "bank")
    book = load_side(Path(args.book), "book")
    closing_bank, closing_book = dec(args.closing_bank), dec(args.closing_book)

    # integrity: does each side's activity roll to its stated closing balance?
    flags: list[str] = []
    for label, rows, opening, closing in (
        ("Bank", bank, dec(args.opening_bank), closing_bank),
        ("Book", book, dec(args.opening_book), closing_book),
    ):
        if opening is None:
            continue
        computed = opening + sum((r["amount"] for r in rows), ZERO)
        if computed != closing:
            flags.append(
                f"{label} side does not roll forward: opening {opening:,.2f} + activity "
                f"= {computed:,.2f}, but stated closing is {closing:,.2f} "
                f"(difference {computed - closing:,.2f}). The {label.lower()} data is "
                f"incomplete. Fix this before relying on the reconciliation."
            )

    pairs = match_all(bank, book, args.date_tolerance,
                      args.token_threshold, args.subset_cap)
    unmatched = [r for r in bank + book if not r["matched"]]

    asof = pdate(args.as_of) if args.as_of else max(r["date"] for r in bank + book)

    prior = []
    if args.prior_outstanding and Path(args.prior_outstanding).exists():
        with Path(args.prior_outstanding).open(newline="", encoding="utf-8-sig") as fh:
            for i, r in enumerate(csv.DictReader(fh), start=1):
                prior.append({
                    "key": f"P{i:04d}", "date": pdate(r.get("date")),
                    "description": " ".join((r.get("description") or "").split()),
                    "check_no": (r.get("check_no") or "").strip(),
                    "amount": dec(r.get("amount")) or ZERO,
                })

    cls = load_classifications(Path(args.classifications) if args.classifications else None)

    # A classification key that matches nothing is almost always a typo, and it
    # would otherwise silently leave a real item unclassified.
    live_keys = {r["key"] for r in unmatched}
    orphans = sorted(set(cls) - live_keys)
    if orphans:
        flags.append(
            "Classification keys that match no unmatched item (typo, or the item "
            "matched on a later run): " + ", ".join(orphans) +
            ". Valid keys are BNKnnnn (bank side) and GLnnnn (book side)."
        )

    rec = build_rec(unmatched, cls, closing_bank, closing_book)

    # write unmatched worklist so the preparer can classify
    if unmatched:
        with Path(args.unmatched_out).open("w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["side", "key", "date", "description", "check_no", "amount",
                        "classification", "expected_resolution", "je_required", "notes"])
            for r in sorted(unmatched, key=lambda x: (x["side"], x["date"])):
                info = cls.get(r["key"], {})
                w.writerow([r["side"], r["key"], r["date"].isoformat(), r["description"],
                            r["check_no"], f"{r['amount']:.2f}",
                            info.get("classification", ""),
                            info.get("expected_resolution", ""),
                            info.get("je_required", ""), info.get("notes", "")])

    if rec["unidentified"] != ZERO:
        flags.append(
            f"Unidentified items totalling {rec['unidentified']:,.2f}. An unidentified "
            f"withdrawal is escalated regardless of amount - absence of an explanation "
            f"is the test, not materiality."
        )
    for r in unmatched:
        c = cls.get(r["key"], {}).get("classification", "")
        age = (asof - r["date"]).days
        if c == "outstanding_check" and age > 180:
            flags.append(f"{r['key']} check {r['check_no'] or '(no #)'} outstanding "
                         f"{age} days ({r['amount']:,.2f}) - stale; consider "
                         f"void/reissue and assess unclaimed-property reporting.")
        if c == "deposit_in_transit" and age > 10:
            flags.append(f"{r['key']} deposit in transit {age} days "
                         f"({r['amount']:,.2f}) - a DIT that has not cleared is not "
                         f"timing. Investigate.")

    meta = {
        "Account": args.account_label or "(not supplied)",
        "Period": args.period or "(not supplied)",
        "As-of date for aging": asof.isoformat(),
        "Bank transactions": len(bank),
        "GL transactions": len(book),
        "Matched pairs": len(pairs),
        "Date tolerance (days)": args.date_tolerance,
        "Prepared": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    wb = Workbook()
    sheet_reconciliation(wb, rec, meta, closing_bank, closing_book)
    sheet_items(wb, unmatched, cls, asof)
    sheet_jes(wb, unmatched, cls)
    sheet_matched(wb, pairs)
    sheet_aging(wb, unmatched, cls, prior, asof)
    sheet_memo(wb, rec, unmatched, cls, pairs, meta, flags)
    wb.active = 0

    out = Path(args.out)
    if not rec["reconciled"]:
        out = out.with_name(out.stem + " [NOT RECONCILED]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))

    # ---- console
    print("=" * 70)
    print("BANK RECONCILIATION")
    print("=" * 70)
    print(f"Bank txns {len(bank)}   GL txns {len(book)}   matched pairs {len(pairs)}")
    print(f"Unmatched (reconciling items): {len(unmatched)}")
    print()
    print(f"  Balance per bank                {closing_bank:>14,.2f}")
    print(f"  Add: deposits in transit        {rec['dit']:>14,.2f}")
    print(f"  Less: outstanding checks        {rec['osc']:>14,.2f}")
    print(f"  Add/(less): bank errors         {rec['bnk_err']:>14,.2f}")
    print(f"  = ADJUSTED BANK                 {rec['adj_bank']:>14,.2f}")
    print()
    print(f"  Balance per GL                  {closing_book:>14,.2f}")
    print(f"  Add: unrecorded credits         {rec['unrec_cr']:>14,.2f}")
    print(f"  Less: unrecorded debits         {rec['unrec_dr']:>14,.2f}")
    print(f"  Add/(less): book errors         {rec['bk_err']:>14,.2f}")
    print(f"  = ADJUSTED BOOK                 {rec['adj_book']:>14,.2f}")
    print()
    print(f"  UNEXPLAINED DIFFERENCE          {rec['difference']:>14,.2f}   "
          f"{'OK' if rec['difference'] == ZERO else '*** MUST BE ZERO ***'}")
    if rec["unclassified"]:
        print(f"\n{len(rec['unclassified'])} unmatched item(s) are NOT YET CLASSIFIED.")
        print(f"Classify them in {args.unmatched_out}, then re-run with")
        print(f"  --classifications {args.unmatched_out}")
        print("The reconciliation is incomplete until then. Do not plug.")
    if flags:
        print("\nEXCEPTIONS / ESCALATIONS:")
        for f in flags:
            print(f"  ! {f}")
    print(f"\nWorkpaper: {out}")
    if unmatched:
        print(f"Unmatched worklist: {args.unmatched_out}")
    return 0 if rec["reconciled"] else 1


if __name__ == "__main__":
    sys.exit(main())
```

