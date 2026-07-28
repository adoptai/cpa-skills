# Contributing

Contributions are welcome. These skills get better when people who do the work daily tell us
where they break.

## Before anything else: never include client data

These skills read bank statements, general ledgers, and tax returns. That makes an ordinary
bug report a potential confidentiality incident.

**Do not paste, attach, or commit:**

- Real bank statements, GL extracts, trial balances, returns, or K-1s
- Client names, account numbers, TINs, EINs, SSNs, addresses
- Spreadsheets or PDFs of any kind
- Screenshots containing any of the above

**Instead:** describe the shape of the problem and supply a synthetic example. Replace amounts
with made-up figures, names with `Acme Holdings LLC`, accounts with `x4471`. A three-row CSV
that reproduces the issue is more useful than a real file, and safe to discuss in public.

`./validate-skills.sh` scans for spreadsheets, PDFs, exports, and TIN-shaped strings and will
fail the build if any are committed. Please don't rely on it as your only safeguard.

## Most useful contributions

Ranked by how much they help:

1. **A bank or ERP export format that breaks a parser.** The highest-value contribution.
   Statement layouts vary enormously and we can only harden against formats we have seen.
   Include a synthetic file in the same layout.
2. **False positives worth suppressing.** If a test flags something innocuous every month in
   your environment, that is a real defect — noise trains people to ignore exception reports.
   Tell us the pattern and why it's benign.
3. **Checklist items, especially state-specific ones.** The tax review checklists are
   deliberately federal-leaning. State conformity, decoupling, and filing quirks are where
   returns actually go wrong.
4. **New anomaly tests, or better weights.** If you have a pattern that reliably surfaces real
   findings, we want it. Say roughly how often it fires and how often it's a true positive.
5. **New skills.** See below.

## Reporting a bug

Open an [issue](https://github.com/adoptai/cpa-skills/issues) with:

- Which skill and what you asked it to do
- What you expected, what happened
- A **synthetic** input that reproduces it
- Your OS, Python version, and which agent you're using

## Proposing a new skill

Open an issue first so we can talk about scope before you write it. A good proposal includes:

- The task, phrased the way a practitioner would ask for it
- Why it needs a skill rather than a plain prompt — usually a multi-step procedure, a
  non-obvious professional standard, or a proof that must be enforced
- **The output gate** (see below)
- What it deliberately will not do

## Requirements for any new or changed skill

### It must gate its own output

This is the design property that makes this repository different from a prompt collection.
Every skill proves its work and refuses to produce a clean deliverable when the proof fails.

Existing gates: the statement extract must tie to the statement's own balances; the
reconciliation's unexplained difference must be zero; every reviewed item needs a citation;
variance explanation components must sum to the variance; every journal entry must balance
before any anomaly test runs.

A skill with no gate will not be merged. If you cannot name a check that would make the skill
refuse to deliver, the task probably doesn't need a skill.

### It must not state statutory figures from memory

No thresholds, rates, phase-outs, limits, mileage rates, dormancy periods, or deadlines. They
change annually and a stale figure in a skill will be relied on. Write the instruction as
*"verify [limit] for TY[year] against current instructions for Form X line Y"* and cite the
form and line.

This is the rule contributors most often miss. A checklist item that says "confirm the
contribution is within the annual limit, per current-year Form 5498 instructions" is correct. One
that names a dollar amount is not.

### It must stop where judgment starts

Automate matching, footing, extraction, pattern detection, formatting. Do not automate the
professional conclusion. Report attributes, not intent — an entry posted at 2 a.m. by its own
approver is a characteristic requiring investigation, not an accusation.

### It must run locally

No network calls, no cloud OCR, no uploads, no telemetry, no API keys. Standard library plus
`openpyxl` and `pdfplumber`. If it needs a remote service, it doesn't belong here.

### It must meet the spec

Per the [Agent Skills specification](https://agentskills.io/specification.md):

- `name` matches the directory exactly; lowercase, digits, hyphens only
- `description` is 1–1024 chars and includes trigger phrases
- `SKILL.md` is under 500 lines — move detail into `references/`
- Optional directories are only `references/`, `scripts/`, `assets/`

## Pull request process

```bash
# 1. edit skills/
# 2. add the skill to .claude-plugin/marketplace.json if it's new
./validate-skills.sh          # must pass
python3 build.py              # regenerate dist/ and markdown/
# 3. update VERSIONS.md
```

Commit `dist/` and `markdown/` — they're the download paths for people who don't use a
terminal, which is a meaningful share of this audience.

Use [Conventional Commits](https://www.conventionalcommits.org/): `feat: add
fixed-asset-rollforward skill`, `fix: suppress round-number flag on contractual amounts`.

### Checklist

- [ ] `./validate-skills.sh` passes
- [ ] `python3 build.py` run, `dist/` and `markdown/` committed
- [ ] New skill listed in `.claude-plugin/marketplace.json`
- [ ] Output gate defined and enforced
- [ ] No statutory figures asserted from memory
- [ ] `VERSIONS.md` updated
- [ ] Tested against a synthetic fixture with a planted error, and the gate caught it
- [ ] No client data anywhere in the diff

## Testing

Generate a synthetic fixture, plant a known error, confirm the gate catches it. A gate that has
never failed has not been tested.

```bash
# example: verify the statement proof catches a transposed digit
# change 19875.00 to 19857.00 in a transactions CSV, then:
python3 skills/bank-statement-to-excel/scripts/build_workbook.py \
  --transactions broken.csv --opening 42150.22 --closing 38904.71 --out out.xlsx
# expect: non-zero exit, no workbook, and the offending row identified by name
```

## A note on professional standards

Contributors are working on tools other professionals will use on live engagements. If a
procedure you add is wrong, someone signs a workpaper based on it. Where a step depends on a
professional standard, cite the standard rather than paraphrasing from memory, and prefer
"confirm X" over asserting X.

If you're unsure whether something is correct, open an issue and say so. A flagged uncertainty
is far more useful than a confident guess.

## Questions

Open an [issue](https://github.com/adoptai/cpa-skills/issues) — happy to help.
