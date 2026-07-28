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

---

## Installation

### Option 1: CLI install (recommended)

Use [npx skills](https://github.com/vercel-labs/skills):

```bash
# Install all skills
npx skills add adoptai/cpa-skills

# Install specific skills
npx skills add adoptai/cpa-skills --skill bank-rec-to-gl tax-return-review

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

### Reconciliation and close
- `bank-rec-to-gl` — bank to general ledger, with proposed journal entries
- `return-yoy-variance` — period-over-period variance with enforced explanations

### Tax
- `tax-return-review` — checklist review with evidence for every item tested
- `return-yoy-variance` — prior-year comparison and analytical procedures

### Audit and assurance
- `journal-entry-anomaly-scan` — full-population JE testing with risk ranking
- `bank-rec-to-gl` — cash existence and completeness support

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
