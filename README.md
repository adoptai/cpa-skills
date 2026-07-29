# CPA Skills for AI Agents

Accounting and tax skills for AI agents. Built for accountants, tax professionals, controllers,
and auditors who want an agent to handle the mechanical layer of the work — extraction,
matching, footing, tie-outs, anomaly detection — without ever sending client data anywhere.

Works with Claude Code, Cowork, OpenAI Codex, Cursor, Windsurf, and any agent that supports the
[Agent Skills spec](https://agentskills.io).

Built by [Adopt AI](https://adopt.ai). Everything runs locally: no cloud, no API keys, no
uploads, no telemetry.

**Contributions welcome** — especially bank and ERP export formats that break the parsers. See
[Contributing](#contributing).

---

## What makes these different

Every skill in this repository **proves its own work, and refuses to deliver a workpaper that
doesn't tie.**

That's the whole design. An unproven workpaper is worse than no workpaper, because it gets
relied on and signed.

| Skill | It will not deliver unless… |
|---|---|
| `bank-statement-to-excel` | the extract ties to the statement's own opening and closing balance |
| `bank-rec-to-gl` | the unexplained difference is exactly `0.00` |
| `tax-return-review` | every tested item carries a form/line reference **and** a source document |
| `return-yoy-variance` | each explanation's components sum to the variance |
| `journal-entry-anomaly-scan` | every journal entry balances before a single anomaly test runs |
| `k1-extract-summarize` | aggregate K-1s foot to the entity return and ownership totals 100.000% |
| `payroll-tax-reconciliation` | register, 941s, W-2/W-3 and GL agree, with no unexplained difference |
| `audit-sampling` | the population ties to a control total and the selection is re-performable from a seed |
| `three-way-match` | matched invoices plus exceptions equal the whole population |
| `depreciation-tie-out` | cost and accumulated depreciation rollforwards foot, beginning equals prior-year ending |
| `sales-tax-reconciliation` | returns tie to the ledger, the liability rollforward foots, and every jurisdiction with sales has a return or a documented nexus conclusion |
| `ar-aging-tie-out` | the aging ties to the GL control account **and** every bucket recomputes from the invoice date |
| `revenue-trace-to-source` | every selection links to invoice, contract and cash — or it is an exception, not a test |

Two consequences worth knowing up front:

**No statutory figures.** You won't find a dollar threshold, rate, phase-out, standard
deduction, contribution limit, mileage rate, or deadline anywhere in these skills. They change
annually and a stale figure in a tool gets relied on. Instead the skills recompute from the
figure shown on the return and tell you which form and line to verify against current
authority.

**They stop where judgment starts.** The reconciliation matches transactions but won't decide
whether an item is a deposit in transit or an unidentified withdrawal. The journal entry scan
reports that an entry was posted at 2 a.m. by its own approver — and never characterizes intent.
That boundary is deliberate.

---

## What are Skills?

Skills are markdown files that give AI agents specialized knowledge and workflows for specific
tasks. Once installed, your agent recognizes when you're working on one of these tasks and
applies the right procedure — including the parts practitioners skip under time pressure.

## How these skills work together

```
        ┌───────────────────────────────┐
        │   bank-statement-to-excel     │   PDF → proven transaction detail
        └───────────────┬───────────────┘
                        │  feeds a proven extract into
                        ▼
        ┌───────────────────────────────┐
        │        bank-rec-to-gl         │   bank ↔ GL, four-column rec + JEs
        └───────────────────────────────┘

        ┌───────────────────────────────┐
        │     return-yoy-variance       │   PY vs CY, explained variances
        └───────────────┬───────────────┘
                        │  supplies Pass 5 (reasonableness) of
                        ▼
        ┌───────────────────────────────┐
        │      tax-return-review        │   checklist + evidence for every item
        └───────────────────────────────┘

        ┌───────────────────────────────┐
        │  journal-entry-anomaly-scan   │   full-population GL scan, risk-ranked
        └───────────────────────────────┘
                        ▲
                        │  its exceptions often become
                        │  reconciling items or review notes

        ┌───────────────────────────────┐
        │    k1-extract-summarize       │   K-1 packages → footed summary
        └───────────────┬───────────────┘
                        │  supplies flow-through detail to
                        ▼
        ┌───────────────────────────────┐
        │      tax-return-review        │
        └───────────────────────────────┘

        ┌───────────────────────────────┐
        │    depreciation-tie-out       │   FA register → schedule → return
        └───────────────┬───────────────┘
                        │  supports the depreciation pass of
                        ▼
        ┌───────────────────────────────┐
        │      tax-return-review        │
        └───────────────────────────────┘

        ┌───────────────────────────────┐      ┌──────────────────────────┐
        │  payroll-tax-reconciliation   │      │      audit-sampling      │
        │  register ↔ 941s ↔ W-3 ↔ GL   │      │  any population above    │
        └───────────────────────────────┘      └──────────────────────────┘

        ┌───────────────────────────────┐
        │       three-way-match         │   PO ↔ invoice ↔ receiving, + GRNI accrual
        └───────────────────────────────┘

        ┌───────────────────────────────┐
        │       audit-sampling          │   selects a reproducible sample
        └───────────────┬───────────────┘
                        │  feeds selections into
                        ▼
        ┌───────────────────────────────┐
        │   revenue-trace-to-source     │   GL → invoice → contract → cash
        └───────────────────────────────┘

        ┌───────────────────────────────┐      ┌──────────────────────────────┐
        │   sales-tax-reconciliation    │      │      ar-aging-tie-out        │
        │  returns ↔ ledger ↔ GL,       │      │  aging ↔ GL, every bucket    │
        │  + nexus screening            │      │  recomputed                  │
        └───────────────────────────────┘      └──────────────────────────────┘
```

Never reconcile off an unproven extract, and never analyse variances against a prior year that
doesn't agree to the return as filed. The skills cross-reference each other where this matters.

---

## Available skills

| Skill | Description |
|---|---|
| [bank-statement-to-excel](skills/bank-statement-to-excel) | Convert PDF bank, credit card, and merchant-processor statements into a clean workbook with dates, descriptions, debits, credits, and running balance — and prove the extraction is complete against the statement's own control totals. Local OCR for scans. |
| [bank-rec-to-gl](skills/bank-rec-to-gl) | Reconcile a bank statement to the GL cash account through a six-pass matching cascade, then produce the four-column reconciliation, the journal entries for unrecorded bank items, and aged outstanding-item schedules. |
| [tax-return-review](skills/tax-return-review) | Review a prepared return against a structured checklist with evidence for every item tested. Starts with what's *missing* — dropped carryforwards, absent forms, unfiled information returns. Checklists for 1040, 1120-S, 1065, 1120, depreciation, amended. |
| [return-yoy-variance](skills/return-yoy-variance) | Compare two years line by line under dual materiality, flag the movements *and* the suspicious non-movements, and enforce that every explanation is causal and quantified. |
| [journal-entry-anomaly-scan](skills/journal-entry-anomaly-scan) | Full-population journal entry scan across 26 tests — duplicates, round-dollar, weekend, after-hours, self-approval, threshold circumvention, unreversed accruals, Benford — with accumulating risk scores. |
| [k1-extract-summarize](skills/k1-extract-summarize) | Extract Schedule K-1 data box by box with the code preserved, then foot the aggregate K-1s to the entity return. Detects a missing K-1 through ownership percentages and a two-way recipient reconciliation. Full TINs rejected on input. |
| [payroll-tax-reconciliation](skills/payroll-tax-reconciliation) | Four-way tie across the payroll register, the four Forms 941, W-2/W-3, and the GL. Infers the Social Security wage base from the data rather than asserting one, and isolates undeposited trust-fund tax. |
| [audit-sampling](skills/audit-sampling) | MUS/PPS, stratified, and attribute selection with a mandatory seed so the sample can be re-performed exactly. Population must tie to a control total; negative balances need an explicit decision. Projects with tainting. |
| [three-way-match](skills/three-way-match) | PO to invoice to receiving, with four duplicate-detection patterns, quantity and price tolerances, vendor-level pattern rollup, and a quantified goods-received-not-invoiced accrual. |
| [depreciation-tie-out](skills/depreciation-tie-out) | Fixed asset register to depreciation schedule to return. Recomputes straight-line exactly; tests accelerated methods for consistency without asserting any rate. Catches disposed assets still depreciating and beginning balances that do not agree to the prior year. |
| [sales-tax-reconciliation](skills/sales-tax-reconciliation) | Filed returns to the sales ledger to the GL, with a liability rollforward that isolates collected-but-unremitted tax. Derives rates from the returns and produces a nexus *screening* schedule rather than a conclusion. Separates marketplace-facilitated sales throughout. |
| [ar-aging-tie-out](skills/ar-aging-tie-out) | Ties the aging to the GL control account and then recomputes every bucket from the invoice date — because the total can agree while the aging is wrong. Grosses up netted credits, tests cutoff, and uses subsequent receipts to separate collection timing from valuation. |
| [revenue-trace-to-source](skills/revenue-trace-to-source) | Traces revenue from the GL through invoice, contract, delivery and cash. States direction and assertion explicitly, since this tests occurrence and cannot detect unrecorded revenue. Matches subsequent credit memos against tested revenue. |

---

## Installation

### Option 1: CLI install (recommended)

Use [npx skills](https://github.com/vercel-labs/skills):

```bash
# Install all skills
npx skills add adoptai/cpa-skills

# Install specific skills
npx skills add adoptai/cpa-skills --skill bank-rec-to-gl audit-sampling

# List available skills
npx skills add adoptai/cpa-skills --list
```

The CLI detects which agents you have and asks where to install. Claude Code uses
`.claude/skills/`; universal agents share `.agents/skills/`.

> **Tip:** if you run this from *inside* an agent session, the CLI runs non-interactively and
> may only install to `.agents/skills/`, which Claude Code doesn't read. Pass the agent
> explicitly:
> ```bash
> npx skills add adoptai/cpa-skills -a claude-code
> ```

### Option 2: Claude Code plugin

```bash
/plugin marketplace add adoptai/cpa-skills
/plugin install cpa-skills
```

### Option 3: One-click `.skill` file (no terminal)

Download a file from [`dist/`](dist) and open it. It installs into Claude desktop or Cowork
with no command line involved.

| Skill | Download |
|---|---|
| Bank statement → Excel | [`bank-statement-to-excel.skill`](dist/bank-statement-to-excel.skill) |
| Bank reconciliation | [`bank-rec-to-gl.skill`](dist/bank-rec-to-gl.skill) |
| Tax return review | [`tax-return-review.skill`](dist/tax-return-review.skill) |
| YoY variance | [`return-yoy-variance.skill`](dist/return-yoy-variance.skill) |
| JE anomaly scan | [`journal-entry-anomaly-scan.skill`](dist/journal-entry-anomaly-scan.skill) |
| K-1 extract & summarize | [`k1-extract-summarize.skill`](dist/k1-extract-summarize.skill) |
| Payroll tax reconciliation | [`payroll-tax-reconciliation.skill`](dist/payroll-tax-reconciliation.skill) |
| Audit sampling | [`audit-sampling.skill`](dist/audit-sampling.skill) |
| Three-way match | [`three-way-match.skill`](dist/three-way-match.skill) |
| Depreciation tie-out | [`depreciation-tie-out.skill`](dist/depreciation-tie-out.skill) |
| Sales tax reconciliation | [`sales-tax-reconciliation.skill`](dist/sales-tax-reconciliation.skill) |
| AR aging tie-out | [`ar-aging-tie-out.skill`](dist/ar-aging-tie-out.skill) |
| Revenue trace to source | [`revenue-trace-to-source.skill`](dist/revenue-trace-to-source.skill) |

### Option 4: Copy and paste

Each file in [`markdown/`](markdown) is the complete skill with every script inlined. Paste it
into any assistant, or hand it to your team as a written procedure. No install, no terminal.

### Option 5: Clone and copy

```bash
git clone https://github.com/adoptai/cpa-skills.git
cp -r cpa-skills/skills/* .agents/skills/
```

### Option 6: Git submodule

```bash
git submodule add https://github.com/adoptai/cpa-skills.git .agents/cpa-skills
```

### Option 7: Fork and customize

Fork it, set your firm's materiality thresholds and review checklist, clone your fork into your
engagements. Thresholds and checklist items are plain parameters and plain text — see
[Modifying these](#modifying-these).

### Dependencies

Python 3.9+ and:

```bash
pip install openpyxl pdfplumber

# only for scanned statements (optional, local OCR):
#   macOS:   brew install ocrmypdf
#   Debian:  sudo apt install ocrmypdf tesseract-ocr
```

---

## Usage

Once installed, just describe the task:

```
"Convert these bank statement PDFs to Excel"
→ bank-statement-to-excel

"Reconcile January for the operating account"
→ bank-rec-to-gl

"Review this 1065 against my checklist before I sign it"
→ tax-return-review

"Why did our numbers move versus last year?"
→ return-yoy-variance

"Scan the GL for unusual journal entries"
→ journal-entry-anomaly-scan

"Summarize these K-1s and tell me if any are missing"
→ k1-extract-summarize

"Do our 941s tie to the W-3?"
→ payroll-tax-reconciliation

"Select 40 receivables for confirmation"
→ audit-sampling

"Did we pay any of these invoices twice?"
→ three-way-match

"Does our depreciation schedule tie to the return?"
→ depreciation-tie-out

"Do our sales tax returns tie to revenue, and which states are we missing?"
→ sales-tax-reconciliation

"Does the AR aging tie, and which receivables are really old?"
→ ar-aging-tie-out

"Trace these revenue transactions to invoices and cash"
→ revenue-trace-to-source
```

Or invoke directly:

```
/bank-rec-to-gl
/tax-return-review
```

---

## Skill categories

### Document extraction
- `bank-statement-to-excel` — PDF statements to proven transaction detail
- `k1-extract-summarize` — K-1 packages to a footed, box-by-box summary

### Reconciliation and close
- `bank-rec-to-gl` — bank to general ledger, with proposed journal entries
- `payroll-tax-reconciliation` — register, 941s, W-2/W-3 and GL in one tie
- `depreciation-tie-out` — fixed asset rollforward and schedule agreement
- `three-way-match` — AP completeness and the GRNI accrual
- `ar-aging-tie-out` — receivables agreement and the allowance
- `return-yoy-variance` — period-over-period variance with enforced explanations

### Tax
- `tax-return-review` — checklist review with evidence for every item tested
- `k1-extract-summarize` — flow-through detail for return input
- `depreciation-tie-out` — fixed assets and Form 4562 support
- `sales-tax-reconciliation` — multi-jurisdiction returns and nexus screening
- `return-yoy-variance` — prior-year comparison and analytical procedures

### Audit and assurance
- `audit-sampling` — reproducible MUS, stratified, and attribute selection
- `revenue-trace-to-source` — revenue occurrence, with the direction stated
- `journal-entry-anomaly-scan` — full-population JE testing with risk ranking
- `three-way-match` — purchasing and payables control testing
- `ar-aging-tie-out` — receivable existence, valuation, and cutoff
- `bank-rec-to-gl` — cash existence and completeness support

### Payroll
- `payroll-tax-reconciliation` — four-way tie, trust-fund exposure, employee-level integrity

---

## Security posture

Everything runs locally: no network calls, no uploads, no cloud OCR, no telemetry, no API keys.
Source documents open read-only and are never modified. Working files stay in a local directory
you control and can shred under your retention policy. Output filenames use client name and
period only — never a full account number or TIN.

If you're subject to a written information security program, these are designed to sit inside
it rather than around it. Nothing leaves the machine, so there's no vendor to add to your data
map and no third party processing client data.

`./validate-skills.sh` also fails the build if a spreadsheet, PDF, accounting export, or
TIN-shaped string is ever committed to this repository.

---

## Professional use

These are tools, not a substitute for professional judgment.

- **You remain responsible for the work product.** Output must be reviewed by a qualified
  professional before it is relied on, filed, or delivered. A passing proof means an extract
  ties to its source — not that a return is correct or a reconciliation is complete in
  substance.
- **Nothing here is tax, accounting, audit, or legal advice.** The checklists are a starting
  framework, not an assurance program, and not a substitute for applicable professional
  standards or your firm's methodology.
- **Statutory figures are deliberately absent.** Confirm every threshold, rate, limit, and
  deadline against current authority. The skills tell you which form and line to check.
- **Verify before you trust a run.** Sample-check output against the source the first time you
  use a skill with any new document format, ERP export, or client. The proofs catch
  incompleteness, not every form of misclassification.

---

## Modifying these

Each skill is a markdown file plus a Python script, both meant to be edited. Materiality
thresholds, test weights, date tolerances, approval limits, and checklist items are all
parameters or plain prose. If your firm has its own review checklist, use it — and keep the
evidence validation.

```bash
# edit skills/, then:
./validate-skills.sh     # spec compliance, manifest sync, client-data check
python3 build.py         # regenerate dist/ and markdown/
```

See [AGENTS.md](AGENTS.md) for the repository conventions and the domain rules any change has
to respect.

---

## Repository structure

```
.claude-plugin/     marketplace.json + plugin.json
.github/            issue and PR templates
skills/             the skills — edit these
dist/               generated .skill packages
markdown/           generated single-file copies
build.py            regenerates dist/ and markdown/
validate-skills.sh  spec + repo compliance checks
AGENTS.md           conventions and domain rules for agents
CLAUDE.md           Claude Code specifics
VERSIONS.md         version history and update checking
```

---

## Contributing

PRs and issues welcome. Most useful, in order: **bank or ERP export formats that break a
parser**, false positives worth suppressing, state-specific checklist items, and new anomaly
tests.

**Never include real client data** — not in an issue, a fixture, or a PR. Synthetic examples
only.

Any new skill has to name its output gate. See [CONTRIBUTING.md](CONTRIBUTING.md).

---

## License

[MIT](LICENSE) — use these however you want, including commercially and inside your firm.
Attribution appreciated, not required.
