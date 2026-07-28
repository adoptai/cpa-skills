# Free Skills for Accountants, Tax Professionals, and CPAs

Five production-grade skills for accounting and tax work. Each one runs entirely on your own
machine — no client data is transmitted, no cloud OCR, no API keys, no telemetry.

Every skill was written the way a controller or engagement leader would actually want the work
done, which means each one **refuses to produce a clean deliverable until it proves itself.**
The bank statement extract will not hand you a workbook unless it ties to the statement's own
opening and closing balance. The reconciliation will not close unless the unexplained
difference is zero. The review workbook rejects any tested item that lacks a citation. That
gate is the point — an unproven workpaper is worse than none, because it gets relied on.

---

## The skills

### 1. Bank Statement → Excel
`bank-statement-to-excel`

Converts PDF bank, credit card, and merchant-processor statements into a clean workbook with
dates, descriptions, debits, credits, and running balance — then proves the extraction is
complete against the statement's own control totals.

Three independent completeness tests run before delivery: balance roll-forward, agreement to
printed statement totals, and line-by-line running-balance continuity. That third test is the
one that earns its keep — it names the exact row where an extract diverges, so a transposed
digit takes seconds to find instead of an afternoon. Handles scanned statements through local
OCR, multi-account PDFs, and multi-month sets with a continuity check that catches a missing
statement month.

**Use it for:** lender packages, cash-flow analysis, reconciliation prep, clean-up engagements,
any time transaction detail is trapped in a PDF.

---

### 2. Bank Reconciliation → General Ledger
`bank-rec-to-gl`

Matches bank transactions to GL cash detail through a six-pass cascade — check number, exact
date, date tolerance, description, batched deposits, split settlements — then builds the
four-column reconciliation workpaper in proper form, proving adjusted bank equals adjusted
book.

Produces the proposed journal entries for unrecorded bank items as a deliverable rather than
an afterthought, ages outstanding checks with stale-dating and unclaimed-property flags, and
verifies both sides roll forward from their own opening balances before matching begins. The
governing rule is that it will not plug: an unexplained difference stays visible and the
reconciliation reports as incomplete.

**Use it for:** monthly close, clean-up work, audit prep, any cash account that won't tie.

---

### 3. Tax Return Review
`tax-return-review`

Reviews a prepared return against a structured checklist and produces evidence for every item
tested — form and line on one side, source document and page on the other.

Runs six passes, starting with **what isn't there**, because the errors that survive review
are the invisible ones: a missing form, a dropped carryforward, an unfiled information return.
Nothing on the face of the return reveals them. The validator rejects the workbook if any item
concluded "agreed" lacks a citation, if a blocking item has no owner, or if a stated difference
doesn't foot.

Includes checklists for 1040, 1120-S, 1065, 1120, depreciation, and amended returns. Those
checklists deliberately contain **no dollar thresholds or rates** — they change annually, so
the skill cites the form and line to verify rather than a figure recalled from memory.

**Use it for:** second review, signature-ready sign-off, new-client returns, staff review.

---

### 4. Year-over-Year Variance Analysis
`return-yoy-variance`

Compares two years line by line under dual materiality — dollar *or* percent, with an absolute
floor so a 400% swing on a $200 account doesn't crowd out a 3% move on $8M of revenue.

Also flags what a normal variance report misses: **lines that didn't move when they should
have.** Depreciation identical to last year after a year of capital additions is not a
coincidence, it's a stale schedule nobody updated. Plus sign flips, new and disappeared lines,
and round-number plugs.

Then it holds the line on explanation quality. "Timing" and "higher volume" are rejected as
restatements of the variance. An explanation needs quantified components that sum to the
movement, and the tool reports coverage as a percentage of flagged dollars rather than letting
the last 15% vanish into "other."

**Use it for:** return review, close packages, analytical procedures, board reporting.

---

### 5. Journal Entry Anomaly Scan
`journal-entry-anomaly-scan`

Full-population scan across 26 tests — duplicates, round-dollar, weekend and holiday,
after-hours, posting lag, post-close entries, self-approval, unauthorized users, suspense
accounts, reserve and revenue manual entries, unreversed accruals, description quality, and
Benford — with risk scores that accumulate per entry.

The design principle: a single flag is weak, converging flags are strong. The value isn't the
weekend list, it's the entry that lands on six lists at once. Among the highest-signal tests is
**threshold circumvention** — an entry of $9,987 under a $10,000 approval limit implies
knowledge of the control, and the companion test detects splitting across same-day entries.

Proves the population before testing it: every entry must balance, document numbering must be
continuous, and empty fields are reported as scope limitations rather than silently disabling
tests. Reports attributes and facts only — never conclusions about intent.

**Use it for:** required audit journal entry procedures, fraud risk work, forensic engagements,
internal audit, pre-audit self-review.

---

## Installing

**Option A — one-click (Claude desktop, Cowork, Claude Code)**

Download the `.skill` file and open it. It installs as a named skill you can invoke by
describing what you need in plain language — "convert these statements to Excel," "reconcile
January," "scan the GL for unusual entries."

**Option B — copy and paste**

Download the `.md` file. It contains the full skill plus every bundled script inlined, so it
works standalone in any assistant that accepts pasted instructions, or as a written procedure
for your team. Create the script files at the paths shown in the appendix.

**Dependencies** — Python 3.9+ and, depending on the skill:

```bash
pip install openpyxl pdfplumber

# only for scanned statements (optional, local OCR):
#   macOS:   brew install ocrmypdf
#   Debian:  sudo apt install ocrmypdf tesseract-ocr
```

---

## Security posture

Everything runs locally. No network calls, no uploads, no cloud OCR, no telemetry, no API keys.
Source documents are opened read-only and never modified. Intermediate working files stay in a
local directory you control and can shred under your document-retention policy. Output
filenames use client name and period only — never a full account number or TIN.

If you are subject to a written information security program, these skills are designed to sit
inside it rather than around it. Nothing leaves the machine, so there is no vendor to add to
your data map.

---

## A word on judgment

These skills automate the mechanical layer — matching, footing, extraction, pattern detection,
formatting — and deliberately stop where judgment starts.

The reconciliation matches transactions but will not classify a reconciling item as a deposit
in transit versus an unidentified withdrawal, because that distinction is the professional
work. The JE scan produces attributes but will not conclude on intent. The review tool enforces
that evidence exists but cannot tell you whether a position is correct. The variance tool
demands that explanations be quantified but cannot supply the cause.

That boundary is intentional. Everything on the near side of it should be fast, complete, and
provable. Everything on the far side is why the client hired a professional.

---

## Repository structure

```
skills/                     the skills themselves — edit these
  bank-statement-to-excel/
    SKILL.md
    scripts/
  bank-rec-to-gl/
  tax-return-review/
    reference/checklists.md  1040, 1120-S, 1065, 1120, depreciation, amended
  return-yoy-variance/
  journal-entry-anomaly-scan/

dist/                       .skill packages — download and open to install
markdown/                   self-contained single files — copy and paste
build.py                    regenerates dist/ and markdown/ from skills/
```

Edit `skills/`, then run `python3 build.py` to regenerate the two distribution formats.

---

## Modifying these

Each skill is a plain markdown file plus a Python script. Both are meant to be edited —
materiality thresholds, test weights, checklist items, date tolerances, and firm-specific
procedures are all parameters or plain text you can change. If your firm has its own review
checklist, use it and keep the evidence validation.

To rebuild the distribution files after editing:

```bash
python3 build.py
```

---

## Professional use

These are tools, not a substitute for professional judgment, and they are built on that
assumption throughout — each one automates the mechanical layer and stops where judgment
begins.

A few things to be explicit about:

- **You remain responsible for the work product.** Output must be reviewed by a qualified
  professional before it is relied on, filed, or delivered to a client. A passing proof means
  an extract ties to its source; it does not mean a return is correct or a reconciliation is
  complete in substance.
- **Nothing here is tax, accounting, audit, or legal advice.** The checklists are a starting
  framework, not an assurance program, and they are not a substitute for the applicable
  professional standards or your firm's own methodology.
- **Statutory figures are deliberately absent.** Thresholds, rates, limits, and deadlines
  change annually. Where a step depends on one, the skill tells you to verify it against
  current authority instead of asserting a number. Confirm those figures yourself.
- **Verify before you trust a run.** Sample-check output against the source on first use with
  any new document format, ERP export, or client. The proofs catch incompleteness, not every
  form of misclassification.

If your firm has a written information security program or an AI use policy, these are designed
to sit inside it: everything runs locally, so there is no vendor to add to your data map and no
third party processing client data.

---

## Contributing

Issues and pull requests are welcome, particularly:

- Bank and ERP export formats that trip up the parsers
- Additional review checklist items, especially state-specific ones
- New anomaly tests, or better weights for the existing ones
- False positives worth suppressing

Please do not include real client data in an issue, a test fixture, or a pull request. Synthetic
examples only.

---

## License

MIT — see [LICENSE](LICENSE). Free to use, modify, and redistribute, including commercially and
inside your firm. Attribution appreciated but not required.

---

*Built by [Adopt AI](https://adopt.ai).*
