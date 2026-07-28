# AGENTS.md

Guidelines for AI agents working in this repository.

## Repository overview

This repository contains **Agent Skills** for accounting and tax work, following the
[Agent Skills specification](https://agentskills.io/specification.md). Skills install to
`.agents/skills/` (the cross-agent standard). The repo also serves as a **Claude Code plugin
marketplace** via `.claude-plugin/marketplace.json`.

- **Name**: CPA Skills
- **GitHub**: [adoptai/cpa-skills](https://github.com/adoptai/cpa-skills)
- **Maintainer**: Adopt AI
- **License**: MIT

## Structure

```
cpa-skills/
├── .claude-plugin/
│   ├── marketplace.json   # Claude Code plugin marketplace manifest
│   └── plugin.json        # plugin metadata
├── skills/                # the skills
│   └── skill-name/
│       ├── SKILL.md       # required, under 500 lines
│       ├── scripts/       # optional, executable code
│       └── references/    # optional, detail loaded on demand
├── dist/                  # generated .skill packages (build.py)
├── markdown/              # generated single-file copies (build.py)
├── build.py               # regenerates dist/ and markdown/
├── validate-skills.sh     # spec + repo compliance checks
├── AGENTS.md
├── CLAUDE.md
├── CONTRIBUTING.md
├── VERSIONS.md
├── LICENSE
└── README.md
```

`dist/` and `markdown/` are **generated**. Never hand-edit them. Edit `skills/`, then run
`python3 build.py`.

## Commands

```bash
./validate-skills.sh    # spec compliance, manifest sync, client-data check
python3 build.py        # regenerate dist/ and markdown/ from skills/
```

Both must pass before a commit. `validate-skills.sh` exits non-zero on any error.

## The Agent Skills specification

### Required frontmatter

```yaml
---
name: skill-name
description: What this skill does and when to use it. Include trigger phrases.
---
```

| Field | Required | Constraints |
|---|---|---|
| `name` | Yes | 1–64 chars, lowercase `a-z`, digits, hyphens. Must match the directory name. No leading, trailing, or consecutive hyphens. |
| `description` | Yes | 1–1024 chars. What it does **and** when to trigger. |
| `license` | No | Defaults to MIT |
| `metadata` | No | Key–value pairs |

### Optional directories

Only `references/`, `scripts/`, and `assets/`. Anything else is not bundled by `build.py` and
is flagged as an error by the validator. Note `references/` is plural.

## Domain rules — read before editing any skill

This repository is used by accountants, tax preparers, and auditors on live client
engagements. These rules are not stylistic; violating them makes a skill unsafe to ship.

### 1. Never state a statutory figure from memory

No dollar thresholds, rates, phase-outs, standard deductions, contribution limits, mileage
rates, dormancy periods, or filing deadlines. They change annually and some change mid-year. A
skill that asserts a stale figure will be relied on and will be wrong.

Instead: recompute using the figure **as shown on the return or in the client's own
worksheet**, test internal consistency, and where the figure itself must be confirmed, write
the instruction as *"verify [limit] for TY[year] against current instructions for Form X line
Y."* Cite the form and line, never a remembered number.

### 2. Every skill must gate its own output

Each skill proves its work before delivering it, and refuses to produce a clean deliverable
when the proof fails. This is the core design property of the repository:

| Skill | Gate |
|---|---|
| `bank-statement-to-excel` | Extract must tie to the statement's opening and closing balance |
| `bank-rec-to-gl` | Unexplained difference must be exactly zero |
| `tax-return-review` | Every tested item needs a form/line reference and a source document |
| `return-yoy-variance` | Explanation components must sum to the variance |
| `journal-entry-anomaly-scan` | Every entry must balance before any anomaly test runs |

When adding a skill, define its gate. A skill with no gate does not belong here.

### 3. Never plug, never guess a number

If a figure is illegible, missing, or ambiguous, it becomes a visible exception with a
reference. It is never back-solved, interpolated, or absorbed into "other." A silently
completed workpaper is worse than an incomplete one because it gets signed.

### 4. Stop where judgment starts

Skills automate the mechanical layer — matching, footing, extraction, pattern detection,
formatting. They do not render professional conclusions. The reconciliation matches
transactions but does not classify a reconciling item as a deposit in transit versus an
unidentified withdrawal. The journal entry scan reports attributes and never characterizes
intent. Preserve this boundary.

### 5. Local only

No network calls, no cloud OCR, no uploads, no telemetry, no API keys. Client financial
records must not leave the machine. If a capability requires a remote service, it does not go
in this repository.

### 6. Never commit client data

No spreadsheets, PDFs, accounting exports, or TIN-shaped strings — not in fixtures, not in
issues, not in test files. Synthetic examples only. The validator checks for this.

## Writing style

- `SKILL.md` under 500 lines; move detail to `references/`
- Second person, direct and instructional
- Short paragraphs, two to four sentences
- Tables for reference data, code blocks for commands
- Bold for key terms; no decorative emoji
- State the reason a step exists when it is not obvious — a reviewer following an unexplained
  procedure will skip it under time pressure

### Description field

Discovery depends entirely on this field. Include what the skill does, the phrases that should
trigger it, and the adjacent skill for scope boundaries. Lean pushy — under-triggering is the
common failure.

## Git workflow

**Branches**: `feature/skill-name`, `fix/skill-name-description`, `docs/description`

**Commits**: [Conventional Commits](https://www.conventionalcommits.org/) —
`feat: add fixed-asset-rollforward skill`, `fix: tighten round-number test in je scan`

**Pull request checklist**

- [ ] `./validate-skills.sh` passes
- [ ] `python3 build.py` run and `dist/` + `markdown/` committed
- [ ] `name` matches the directory
- [ ] `description` under 1024 chars with trigger phrases
- [ ] `SKILL.md` under 500 lines
- [ ] New skill added to `.claude-plugin/marketplace.json`
- [ ] No statutory figures asserted from memory
- [ ] The skill's output gate is defined and enforced
- [ ] `VERSIONS.md` updated
- [ ] No client data anywhere in the diff

## Testing a skill change

Scripts are validated with synthetic data. Generate a fixture, plant a known error, and
confirm the gate catches it — a gate that never fails has not been tested. Example: transpose
a digit in a statement extract and verify the line-continuity test names the offending row.

Do not commit fixtures containing anything resembling real client data.
