---
name: flux-narrative
description: Compute period-over-period movement between two trial balances, apply dual dollar-and-percent materiality, and draft variance explanations for a reviewer to confirm rather than write from scratch. Use this whenever the user mentions flux analysis, flux report, month-over-month or period-over-period variance, explaining why an account moved, a flux narrative or flux commentary, preparing the flux tab for a close package, or asks why a balance changed from last month. Also trigger on "flux this close," "explain the movement on this account," "draft the variance commentary," "two trial balances, what changed," or when two periods of trial balance data are supplied together. Runs fully local — no client or company data leaves the machine.
---

# Flux Narrative Builder

The numbers side of a flux report takes seconds: subtract one trial balance from another. The narrative side is what actually takes the week, because every material movement needs a specific, sourced reason, not a restatement of the number that moved.

This skill computes the movement and drafts the narrative's starting point. It does not invent a cause from nothing. It flags what changed, and where a person already explained the same account last period, it offers that explanation back as a labeled draft for a reviewer to confirm or replace. **A draft never counts as an explanation on its own.** Nothing closes until a reviewer confirms it, re-quantified for this period, with a source attached.

## What counts as an explanation

An explanation is causal and quantified. Direction is not an explanation.

- ✗ "Revenue increased." — restates the variance
- ✗ "Higher volume." — no magnitude, no source, not testable
- ✗ "Timing." — the single most common non-explanation in practice; it hides errors
- ✓ "Revenue up $84K on the January renewal cohort, per the deferred revenue rollforward." Ties to $84K and can be tested against the schedule.

If the components don't sum to the variance, the explanation is incomplete and the remainder stays visible as unexplained. Do not let the last 15% of a movement disappear into "other."

## Inputs

1. **Current period trial balance** — the period just closed
2. **Prior period trial balance, as reported** — not a since-adjusted version of the prior file. If the prior period was restated, use the restated figures and say so.
3. **Prior-period commentary (optional)** — the confirmed cause, evidence, and owner from last period's flux review, keyed by account. This is what makes drafting possible: a recurring driver (a SaaS renewal cadence, a known seasonal pattern, an ongoing ramp) gets recognized as recurring instead of re-investigated cold every month.
4. **This period's explanations (optional, second run)** — the worklist this skill writes, filled in by a reviewer and fed back in to close the loop.

Confirm both periods are on a comparable basis before comparing: same entity, same accounting method, same chart of accounts. A restated prior period or a chart-of-accounts restructuring makes raw variances meaningless — normalize first and disclose the normalization, or the schedule misleads.

## Step 1 — Map the accounts

Align current-period accounts to prior-period accounts by account ID. Three outcomes, and the last two matter most:

- **Matched** — same account both periods. Compute the variance.
- **New this period** — an account with a current balance and no prior one. Every one needs a cause. High-signal: a new vendor, a new revenue stream, a new expense category. Also the classic tell for a misposting into a freshly created account.
- **Disappeared** — an account with a prior balance and nothing this period. Higher signal still. Either the activity genuinely stopped (name the event) or something was dropped.

## Step 2 — Run the comparison

```bash
python3 scripts/flux_compare.py \
  --current cp.csv --prior pp.csv \
  --prior-commentary last_period_explanations.csv \
  --dollar-threshold 10000 --percent-threshold 10 --absolute-floor 2000 \
  --entity "Acme Holdings LLC" --cp-label "Sep 2026" --pp-label "Aug 2026" \
  --out "Acme Holdings - Sep 2026 Flux.xlsx"
```

Input CSVs: `account_id`, `account_name`, `amount`, and optional `account_type` and `activity_changed` (`yes`/`no`/blank, current-period file only — see the unchanged test).

**Materiality is dual, and both tests must be satisfied to be immaterial.** A movement is material if it exceeds the dollar threshold **or** the percent threshold, subject to an absolute floor below which percent is ignored. A 400% swing on a $200 account is noise; a 3% move on an $8M revenue account is not. Set thresholds by engagement judgment and record them.

Regardless of threshold, these are **always** flagged:

- **Sign flips** — a balance that crossed zero. Qualitatively material at any magnitude.
- **New and disappeared accounts** — per Step 1.
- **Zero-variance accounts where activity changed** — mark `activity_changed = yes` on any account whose underlying facts moved even though the reported figure didn't. Depreciation identical to last month after a capital addition is not a coincidence.
- **Round-number amounts** on accounts that should be computed — an accrual of exactly $50,000 suggests a plug, not a calculation.
- **Accounts equal to the prior period to the penny** where there's ongoing activity.

## Step 3 — Review the drafts, not just the blanks

This is where this skill differs from a flat variance calculator. When `--prior-commentary` is supplied and a flagged account was explained last period, the worklist arrives with that explanation already sitting in the `cause` field, prefixed `DRAFT, confirm or replace:`. It carries no dollar components — a component tied to last month's movement would misstate this month's, so those always start blank.

Three things can happen to a draft:

- **It still applies.** Re-quantify the components against this period's actual detail, attach evidence, set `status = confirmed`.
- **It's close but not exact.** Edit the cause, then quantify and confirm.
- **It no longer applies.** Replace it. A stale carried-forward reason marked confirmed anyway is worse than an honest blank, because it looks resolved to the next reviewer.

Everything without a prior-period match starts blank, exactly like a flat variance tool, because there is nothing honest to draft from.

## Step 4 — Require confirmation

The script writes `flux_worklist.csv` listing every open item: unconfirmed drafts and true blanks together. Fill in `cause` (or accept the draft), `components` (which must sum to the variance), `evidence`, `owner`, and **`status = confirmed`**, then re-run with `--explanations`. A row left as `draft` — even with a perfectly good cause sitting in it — does not count. That is the entire point: the skill offers a starting point, a person decides it's actually true this period.

The script reports the **confirmation rate**, both by count and by dollar, and refuses to mark the analysis complete below 100% of flagged items. Where a cause rests on a **client- or business-unit-provided reason you have not corroborated**, mark `evidence = client representation` — a legitimate but weaker basis, and the workbook flags it so a reviewer knows which is which.

## Step 5 — Deliver

**Workbook tabs:**

1. **Summary** — thresholds used, accounts compared, confirmation rate by count and by dollar, top ten movements, and status.
2. **Flux Schedule** — every account: type, prior period, current period, dollar and percent variance, materiality verdict, flag reason, status, cause, components, tie check, evidence, owner. Sorted by absolute variance descending.
3. **Always-Flag Exceptions** — sign flips, new accounts, disappeared accounts, unchanged-but-activity-changed, and round-number suspects.
4. **Draft Queue** — every account carrying an unconfirmed draft, isolated so a reviewer can work through exactly this list first.
5. **Unexplained** — what remains open, drafts and true blanks together. An empty tab is the goal.

**Then, in chat:** the handful of movements that actually matter, in plain language, with their confirmed causes, followed by what's still in the draft queue or genuinely unexplained and who owns it. Lead with anything still open. A controller reading this wants to know what to review, not to re-read the schedule.

## When a movement suggests something worse

Some patterns are worth escalating rather than explaining:

- A confirmed cause that repeats verbatim for the same account for several consecutive periods — the signature of a systemic error rolling forward under a label that says "recurring," not a genuine recurring driver
- A round-dollar movement that exactly offsets an equal and opposite movement in another account
- An account marked "activity changed, amount unchanged" more than once
- A prior-period figure that doesn't agree to the prior period's own reported balance — someone changed a closed period, and it needs to be reconciled before any variance is meaningful

Raise these as observations with the facts you have. Do not characterize intent.

## Security posture

Fully local. No client or company financial data leaves the machine. Inputs read-only.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*

---

# Appendix — bundled files

If you installed the `.skill` package these files are already in place and you can ignore this appendix. If you copied the skill as text, create the files below at the paths shown, alongside your `SKILL.md`. The skill will not run without them.

## `scripts/flux_compare.py`

```python
#!/usr/bin/env python3
"""
Period-over-period flux analysis with dual materiality and a draft-then-confirm
explanation gate.

Flags a line as requiring explanation when it breaches the dollar OR the percent
threshold (percent ignored below an absolute floor), plus four always-flag
conditions that are qualitatively material at any magnitude:

    sign flip · new account · disappeared account ·
    zero variance where the underlying activity changed · round-number plug

If --prior-commentary is given, a flagged account with commentary on file from
the prior period gets a labeled DRAFT explanation carried into the worklist —
never a fabricated cause, only a reuse of what a person already wrote last
period, offered for a reviewer to confirm or replace. A draft never counts as
explained on its own. Only a status of "confirmed", with components that sum
to the variance and an evidence reference, closes a flagged line. Reports
explanation coverage and refuses to call the analysis complete otherwise.

Usage:
    python3 flux_compare.py --current cp.csv --prior pp.csv \
        --prior-commentary last_period_explanations.csv \
        --dollar-threshold 10000 --percent-threshold 10 --absolute-floor 2000 \
        --entity "Acme Holdings LLC" --cp-label "Sep 2026" --pp-label "Aug 2026" \
        --out "Acme Holdings - Sep 2026 Flux.xlsx"

Input CSVs (current and prior, same shape — a two-period trial balance export):
    account_id      stable key used to match periods           required
    account_name    human-readable line name                   required
    amount          signed                                      required
    account_type    asset|liability|equity|revenue|expense      optional
    activity_changed   yes|no|blank  (current-period file only)     optional

--prior-commentary CSV (read-only; last period's CONFIRMED explanations):
    account_id, cause, evidence, owner, notes
      Carried forward as a draft cause only. Components are never carried
      forward — a component tied to last period's dollar movement would
      misstate this period's, so the reviewer re-quantifies from scratch.

--explanations CSV (this period's worklist, read on a second run):
    account_id, cause, components, evidence, owner, notes, status
      components : semicolon-separated signed amounts that must sum to the
                   variance, e.g. "18000; -3200"
      evidence   : document reference, or the literal "client representation"
      status     : "confirmed" or blank/"draft" — only "confirmed" counts
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
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
PCT = '0.0"%"'

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
OK_FONT = Font(bold=True, color="006100")
BAD_FONT = Font(bold=True, color="9C0006")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
DRAFT_FILL = PatternFill("solid", fgColor="DDEBF7")

# Amounts that are suspiciously round on a line that should be computed.
# Roundness has to scale with magnitude or the test is pure noise: a modulus
# has to be a meaningful fraction (>=5%) of the amount itself to mean anything.
ROUND_MODULI = (Decimal("1000000"), Decimal("500000"), Decimal("100000"),
                Decimal("50000"), Decimal("25000"), Decimal("10000"),
                Decimal("5000"), Decimal("1000"))
ROUND_MIN = Decimal("1000")
ROUND_SIGNIFICANCE = Decimal("20")   # modulus >= |amount| / 20


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


def is_round(v: Decimal) -> str:
    a = abs(v)
    if a < ROUND_MIN:
        return ""
    need = a / ROUND_SIGNIFICANCE
    for m in ROUND_MODULI:
        if m >= need and a >= m and a % m == ZERO:
            return f"exact multiple of {m:,.0f}"
    return ""


def load_tb(path: Path, label: str) -> dict[str, dict]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        sys.exit(f"No rows in {path}")
    out: dict[str, dict] = {}
    for i, r in enumerate(rows, start=1):
        aid = " ".join(str(r.get("account_id", "") or "").split())
        name = " ".join(str(r.get("account_name", "") or "").split())
        if not aid:
            aid = name
        if not aid:
            raise ValueError(f"{label} row {i}: needs an account_id or an account_name.")
        amt = dec(r.get("amount"))
        if amt is None:
            raise ValueError(f"{label} row {i} ({aid}): amount is required.")
        if aid in out:
            raise ValueError(
                f"{label}: duplicate account_id {aid!r}. Account IDs must be unique - "
                f"aggregate the duplicates or give them distinct IDs."
            )
        act = str(r.get("activity_changed", "") or "").strip().lower()
        out[aid] = {
            "account_id": aid, "account_name": name or aid, "amount": amt,
            "account_type": " ".join(str(r.get("account_type", "") or "").split()).lower(),
            "activity_changed": act in ("yes", "y", "true", "1"),
        }
    return out


def load_commentary(path: Path | None) -> dict[str, dict]:
    """Prior-period commentary: a read-only source for draft causes. No
    components are read from here — see the docstring on why."""
    if not path or not path.exists():
        return {}
    out = {}
    with path.open(newline="", encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            aid = " ".join(str(r.get("account_id", "") or "").split())
            cause = " ".join(str(r.get("cause", "") or "").split())
            if not aid or not cause:
                continue
            out[aid] = {
                "cause": cause,
                "evidence": " ".join(str(r.get("evidence", "") or "").split()),
                "owner": " ".join(str(r.get("owner", "") or "").split()),
            }
    return out


def load_explanations(path: Path | None) -> dict[str, dict]:
    if not path or not path.exists():
        return {}
    out = {}
    with path.open(newline="", encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            aid = " ".join(str(r.get("account_id", "") or "").split())
            if not aid:
                continue
            comps = []
            raw = str(r.get("components", "") or "")
            for part in re.split(r"[;|]", raw):
                part = part.strip()
                if part:
                    v = dec(part)
                    if v is not None:
                        comps.append(v)
            status = " ".join(str(r.get("status", "") or "").split()).lower()
            out[aid] = {
                "cause": " ".join(str(r.get("cause", "") or "").split()),
                "components_raw": raw.strip(),
                "components": comps,
                "evidence": " ".join(str(r.get("evidence", "") or "").split()),
                "owner": " ".join(str(r.get("owner", "") or "").split()),
                "notes": " ".join(str(r.get("notes", "") or "").split()),
                "confirmed": status == "confirmed",
                "status_raw": status or "draft",
            }
    return out


def compare(cp: dict, pp: dict, dollar: Decimal, pct: Decimal,
            floor: Decimal) -> list[dict]:
    rows = []
    for aid in sorted(set(cp) | set(pp)):
        c, p = cp.get(aid), pp.get(aid)
        cp_amt = c["amount"] if c else None
        pp_amt = p["amount"] if p else None
        src = c or p
        var = (cp_amt or ZERO) - (pp_amt or ZERO)

        pct_var = None
        if pp_amt not in (None, ZERO):
            pct_var = (var / abs(pp_amt)) * Decimal(100)

        flags: list[str] = []
        if c and not p:
            flags.append("NEW account this period")
        if p and not c:
            flags.append("DISAPPEARED - present last period, absent this period")
        if cp_amt is not None and pp_amt is not None:
            if (cp_amt > ZERO and pp_amt < ZERO) or (cp_amt < ZERO and pp_amt > ZERO):
                flags.append("SIGN FLIP - direction changed")
            if var == ZERO and c and c["activity_changed"]:
                flags.append("UNCHANGED but activity changed - stale rolled-forward figure?")
            if var == ZERO and cp_amt != ZERO and c and not c["activity_changed"]:
                flags.append("identical to prior period to the penny - confirm still correct")
        if cp_amt is not None:
            rnd = is_round(cp_amt)
            if rnd:
                flags.append(f"round number ({rnd}) - computed or plugged?")

        # dual materiality
        mat_d = abs(var) >= dollar
        mat_p = (pct_var is not None
                 and abs(pct_var) >= pct
                 and max(abs(cp_amt or ZERO), abs(pp_amt or ZERO)) >= floor)
        material = mat_d or mat_p
        why = []
        if mat_d:
            why.append(f"|variance| >= {dollar:,.0f}")
        if mat_p:
            why.append(f"|%| >= {pct}% above {floor:,.0f} floor")

        needs = material or bool(flags)
        rows.append({
            "account_id": aid, "account_name": src["account_name"],
            "account_type": src["account_type"],
            "pp": pp_amt, "cp": cp_amt, "variance": var, "pct": pct_var,
            "material": material, "materiality_basis": "; ".join(why),
            "flags": flags, "needs_explanation": needs,
        })
    rows.sort(key=lambda r: abs(r["variance"]), reverse=True)
    return rows


def attach(rows: list[dict], expl: dict, commentary: dict) -> tuple[list[dict], list[str]]:
    problems: list[str] = []
    for r in rows:
        e = expl.get(r["account_id"])
        prior = commentary.get(r["account_id"])
        r["is_draft"] = False

        if e:
            r["cause"] = e["cause"]
            r["components_raw"] = e["components_raw"]
            r["evidence"] = e["evidence"]
            r["owner"] = e["owner"]
            r["notes"] = e["notes"]
            r["status"] = e["status_raw"]
        elif r["needs_explanation"] and prior:
            # A carried-forward candidate, never a fabricated cause: the exact
            # words a person wrote last period, clearly labeled as unconfirmed
            # and with no dollar components attached to it.
            r["cause"] = f"DRAFT, confirm or replace: {prior['cause']}"
            r["components_raw"] = ""
            r["evidence"] = prior["evidence"]
            r["owner"] = prior["owner"]
            r["notes"] = "Carried forward from prior-period commentary. Re-quantify " \
                         "components for this period before confirming."
            r["status"] = "draft"
            r["is_draft"] = True
        else:
            r["cause"], r["components_raw"] = "", ""
            r["evidence"] = r["owner"] = r["notes"] = ""
            r["status"] = ""

        r["components_tie"] = ""
        r["explained"] = False

        if not r["needs_explanation"]:
            continue

        confirmed = bool(e and e["confirmed"])
        if not confirmed:
            # A draft, or an unconfirmed row, never counts — this is the gate
            # that makes "a reviewer edits it" a fact rather than a promise.
            continue

        comps = e["components"] if e else []
        if comps:
            total = sum(comps, ZERO)
            diff = total - r["variance"]
            if diff == ZERO:
                r["components_tie"] = "ties"
                r["explained"] = True
            else:
                r["components_tie"] = f"OFF BY {diff:,.2f}"
                problems.append(
                    f"{r['account_id']} ({r['account_name']}): components sum to "
                    f"{total:,.2f} but the variance is {r['variance']:,.2f} - "
                    f"{diff:,.2f} of the movement is unexplained. Do not absorb the "
                    f"remainder into 'other'."
                )
        else:
            r["components_tie"] = "no components given"
            problems.append(
                f"{r['account_id']} ({r['account_name']}): marked confirmed with no "
                f"quantified components. An explanation that cannot be tested is a "
                f"restatement of the variance."
            )
        if r["explained"] and not r["evidence"]:
            r["explained"] = False
            problems.append(
                f"{r['account_id']} ({r['account_name']}): confirmed with no evidence "
                f"reference. Cite a document, or state 'client representation' so the "
                f"weaker basis is visible to a reviewer."
            )
    for aid in sorted(set(expl) - {r["account_id"] for r in rows}):
        problems.append(
            f"Explanation supplied for account_id {aid!r}, which appears in neither "
            f"period. Likely a typo."
        )
    return rows, problems


# ------------------------------------------------------------------ workbook

def hdr(ws, n, row=1):
    for c in range(1, n + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def widths(ws, w):
    for c, v in w.items():
        ws.column_dimensions[get_column_letter(c)].width = v


SCHED_COLS = [
    ("account_name", "Account", 40), ("account_type", "Type", 12),
    ("pp", "Prior period", 16), ("cp", "Current period", 16),
    ("variance", "Variance $", 16), ("pct", "Variance %", 12),
    ("material", "Material", 10), ("flagstr", "Flag reason", 42),
    ("status", "Status", 12),
    ("cause", "Cause", 56), ("components_raw", "Components", 26),
    ("components_tie", "Tie", 14),
    ("evidence", "Evidence", 24), ("owner", "Owner", 14),
    ("account_id", "Account ID", 14),
]


def sheet_schedule(wb, rows, title, subset=None):
    ws = wb.create_sheet(title)
    ws.append([lbl for _, lbl, _ in SCHED_COLS])
    hdr(ws, len(SCHED_COLS))
    data = subset if subset is not None else rows
    for r in data:
        r["flagstr"] = "; ".join(r["flags"])
        out = []
        for key, _, _ in SCHED_COLS:
            v = r.get(key)
            if key in ("pp", "cp", "variance"):
                out.append(float(v) if v is not None else None)
            elif key == "pct":
                out.append(float(v) if v is not None else None)
            elif key == "material":
                out.append("YES" if v else "")
            else:
                out.append(v or None)
        ws.append(out)

    idx = {k: i for i, (k, _, _) in enumerate(SCHED_COLS)}
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for k in ("pp", "cp", "variance"):
            row[idx[k]].number_format = MONEY
        row[idx["pct"]].number_format = PCT
        if row[idx["material"]].value == "YES":
            row[idx["material"]].font = BAD_FONT
        if row[idx["flagstr"]].value:
            row[idx["flagstr"]].fill = WARN_FILL
        tie = row[idx["components_tie"]].value
        if tie and tie != "ties":
            row[idx["components_tie"]].fill = BAD_FILL
            row[idx["components_tie"]].font = BAD_FONT
        elif tie == "ties":
            row[idx["components_tie"]].font = OK_FONT
        if row[idx["status"]].value == "draft":
            row[idx["cause"]].fill = DRAFT_FILL
            row[idx["status"]].fill = DRAFT_FILL
        if (row[idx["material"]].value == "YES" or row[idx["flagstr"]].value) \
                and not row[idx["cause"]].value:
            row[idx["cause"]].fill = BAD_FILL
            row[idx["cause"]].value = "*** UNEXPLAINED ***"
        if str(row[idx["evidence"]].value or "").lower() == "client representation":
            row[idx["evidence"]].fill = DRAFT_FILL

    if ws.max_row == 1:
        ws.cell(row=2, column=1, value="None.")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(SCHED_COLS))}{max(ws.max_row, 2)}"
    widths(ws, {i: w for i, (_, _, w) in enumerate(SCHED_COLS, start=1)})


def sheet_summary(wb, rows, meta, stats, problems):
    ws = wb.active
    ws.title = "Summary"
    widths(ws, {1: 4, 2: 50, 3: 20, 4: 20, 5: 62})
    r = 1

    def line(txt, *, bold=False, size=11, fill=None, col=2):
        nonlocal r
        c = ws.cell(row=r, column=col, value=txt)
        c.font = Font(bold=bold, size=size)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c.fill = fill
        r += 1

    line("PERIOD-OVER-PERIOD FLUX ANALYSIS", bold=True, size=14)
    r += 1
    for k, v in meta.items():
        ws.cell(row=r, column=2, value=k).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1
    r += 1

    complete = stats["needs"] > 0 and stats["explained"] == stats["needs"] and not problems
    if stats["needs"] == 0:
        line("No line breached materiality and no always-flag condition arose.",
             bold=True, fill=DRAFT_FILL)
    elif complete:
        line("ANALYSIS COMPLETE - every flagged line has a confirmed, evidenced, "
             "quantified cause.", bold=True)
        ws.cell(row=r - 1, column=2).font = OK_FONT
    else:
        line(f"ANALYSIS INCOMPLETE - {stats['needs'] - stats['explained']} flagged "
             f"line(s) not yet confirmed"
             + (f", {len(problems)} explanation problem(s)" if problems else "")
             + ". See Unexplained tab.", bold=True, fill=BAD_FILL)
    r += 1

    line("COVERAGE", bold=True, size=12)
    for label, val in [
        ("Accounts compared", stats["total"]),
        ("Requiring explanation", stats["needs"]),
        ("Drafted from prior-period commentary", stats["drafted"]),
        ("Confirmed", stats["explained"]),
        ("Confirmation rate (count)", f"{stats['rate_count']:.0f}%"),
        ("Confirmation rate (dollars)", f"{stats['rate_dollars']:.0f}%"),
    ]:
        ws.cell(row=r, column=2, value="  " + label)
        ws.cell(row=r, column=3, value=val)
        r += 1
    r += 1

    top = [x for x in rows if x["variance"] != ZERO][:10]
    if top:
        line("TOP 10 MOVEMENTS BY ABSOLUTE DOLLAR", bold=True, size=12)
        for h, col in zip(["Account", "Prior", "Current", "Variance", "Status"],
                          (2, 3, 4, 5, 6)):
            hc = ws.cell(row=r, column=col, value=h)
            hc.fill, hc.font = HDR_FILL, HDR_FONT
        r += 1
        for x in top:
            ws.cell(row=r, column=2, value=x["account_name"])
            for col, key in ((3, "pp"), (4, "cp"), (5, "variance")):
                v = x[key]
                c = ws.cell(row=r, column=col, value=float(v) if v is not None else None)
                c.number_format = MONEY
            label = "confirmed" if x["explained"] else (
                "draft, unconfirmed" if x.get("is_draft") else
                ("unexplained" if x["needs_explanation"] else "")
            )
            cc = ws.cell(row=r, column=6, value=label or "-")
            if x["needs_explanation"] and not x["explained"]:
                cc.fill, cc.font = BAD_FILL, BAD_FONT
            r += 1
        r += 1

    line("A cause is causal and quantified. 'Timing' and 'higher volume' restate "
         "the variance rather than explaining it. A draft carried forward from last "
         "period is a starting point, not an answer — it becomes an answer only once "
         "a reviewer confirms it against this period's actual detail.", bold=True)


def sheet_flags(wb, rows):
    subset = [r for r in rows if r["flags"]]
    sheet_schedule(wb, rows, "Always-Flag Exceptions", subset)


def sheet_drafts(wb, rows):
    subset = [r for r in rows if r.get("is_draft")]
    sheet_schedule(wb, rows, "Draft Queue", subset)


def sheet_unexplained(wb, rows, problems):
    subset = [r for r in rows if r["needs_explanation"] and not r["explained"]]
    sheet_schedule(wb, rows, "Unexplained", subset)
    ws = wb["Unexplained"]
    if problems:
        r = ws.max_row + 3
        c = ws.cell(row=r, column=1, value="EXPLANATION PROBLEMS")
        c.font = Font(bold=True, size=12)
        for p in problems:
            r += 1
            cell = ws.cell(row=r, column=1, value=p)
            cell.font = BAD_FONT
            cell.alignment = Alignment(wrap_text=True, vertical="top")


# ------------------------------------------------------------------ main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--current", required=True)
    ap.add_argument("--prior", required=True)
    ap.add_argument("--prior-commentary")
    ap.add_argument("--explanations")
    ap.add_argument("--dollar-threshold", default="10000")
    ap.add_argument("--percent-threshold", default="10")
    ap.add_argument("--absolute-floor", default="2000",
                    help="below this magnitude, percent variance is ignored as noise")
    ap.add_argument("--entity", default="")
    ap.add_argument("--cp-label", default="Current period")
    ap.add_argument("--pp-label", default="Prior period")
    ap.add_argument("--out", required=True)
    ap.add_argument("--worklist-out", default="flux_worklist.csv")
    args = ap.parse_args()

    dollar = dec(args.dollar_threshold) or ZERO
    pct = dec(args.percent_threshold) or ZERO
    floor = dec(args.absolute_floor) or ZERO

    cp = load_tb(Path(args.current), "current")
    pp = load_tb(Path(args.prior), "prior")
    rows = compare(cp, pp, dollar, pct, floor)
    commentary = load_commentary(Path(args.prior_commentary) if args.prior_commentary else None)
    expl = load_explanations(Path(args.explanations) if args.explanations else None)
    rows, problems = attach(rows, expl, commentary)

    needs = [r for r in rows if r["needs_explanation"]]
    explained = [r for r in needs if r["explained"]]
    drafted = [r for r in needs if r.get("is_draft")]
    flagged_dollars = sum((abs(r["variance"]) for r in needs), ZERO)
    explained_dollars = sum((abs(r["variance"]) for r in explained), ZERO)
    stats = {
        "total": len(rows), "needs": len(needs), "explained": len(explained),
        "drafted": len(drafted),
        "rate_count": (len(explained) / len(needs) * 100) if needs else 100.0,
        "rate_dollars": (float(explained_dollars / flagged_dollars) * 100)
                        if flagged_dollars else 100.0,
    }

    # worklist for the reviewer: every open item, drafts pre-filled, others blank
    open_items = [r for r in needs if not r["explained"]]
    if open_items:
        with Path(args.worklist_out).open("w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["account_id", "account_name", "prior", "current", "variance",
                        "pct", "flag_reason", "cause", "components", "evidence",
                        "owner", "notes", "status"])
            for r in open_items:
                w.writerow([
                    r["account_id"], r["account_name"],
                    f"{r['pp']:.2f}" if r["pp"] is not None else "",
                    f"{r['cp']:.2f}" if r["cp"] is not None else "",
                    f"{r['variance']:.2f}",
                    f"{r['pct']:.1f}" if r["pct"] is not None else "",
                    "; ".join(r["flags"]) or r["materiality_basis"],
                    r["cause"], "", r["evidence"], r["owner"], r["notes"],
                    "draft" if r.get("is_draft") else "",
                ])

    meta = {
        "Entity": args.entity or "(not stated)",
        "Current period": args.cp_label,
        "Prior period": args.pp_label,
        "Dollar threshold": float(dollar),
        "Percent threshold": f"{pct}%",
        "Absolute floor (percent test ignored below)": float(floor),
        "Prior-period commentary supplied": "yes" if commentary else "no",
    }

    wb = Workbook()
    sheet_summary(wb, rows, meta, stats, problems)
    sheet_schedule(wb, rows, "Flux Schedule")
    sheet_flags(wb, rows)
    sheet_drafts(wb, rows)
    sheet_unexplained(wb, rows, problems)
    wb.active = 0

    complete = (stats["needs"] == 0) or (stats["explained"] == stats["needs"] and not problems)
    out = Path(args.out)
    if not complete:
        out = out.with_name(out.stem + " [INCOMPLETE]" + out.suffix)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))

    # ---- console
    print("=" * 72)
    print(f"FLUX ANALYSIS:  {args.pp_label} -> {args.cp_label}")
    print("=" * 72)
    print(f"Materiality: |var| >= {dollar:,.0f}  OR  |%| >= {pct}%  "
          f"(percent ignored below {floor:,.0f})")
    print(f"Accounts compared     : {stats['total']}")
    print(f"Requiring explanation : {stats['needs']}")
    print(f"Drafted from prior commentary : {stats['drafted']}")
    print(f"Confirmed             : {stats['explained']}  "
          f"({stats['rate_count']:.0f}% of items, {stats['rate_dollars']:.0f}% of dollars)")

    flags = [r for r in rows if r["flags"]]
    if flags:
        print(f"\nAlways-flag exceptions ({len(flags)}):")
        for r in flags[:20]:
            print(f"  ! {r['account_name'][:40]:<40} {r['variance']:>14,.2f}  "
                  f"{'; '.join(r['flags'])[:60]}")
        if len(flags) > 20:
            print(f"  ... and {len(flags) - 20} more")

    if open_items:
        print(f"\nOPEN ({len(open_items)}), {len(drafted)} with a draft cause to review:")
        for r in open_items[:20]:
            tag = " [DRAFT]" if r.get("is_draft") else ""
            print(f"  ? {r['account_name'][:40]:<40} {r['variance']:>14,.2f}{tag}")
        if len(open_items) > 20:
            print(f"  ... and {len(open_items) - 20} more")
        print(f"\nWorklist written to {args.worklist_out}. Confirm or replace each "
              f"cause, quantify components, set status=confirmed, and re-run with "
              f"--explanations.")

    if problems:
        print(f"\nEXPLANATION PROBLEMS ({len(problems)}):")
        for p in problems[:20]:
            print(f"  ! {p}")

    print(f"\nStatus  : {'COMPLETE' if complete else 'INCOMPLETE'}")
    print(f"Workbook: {out}")
    return 0 if complete else 1


if __name__ == "__main__":
    sys.exit(main())
```

