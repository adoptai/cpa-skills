# CLAUDE.md

Instructions for Claude Code working in this repository.

**Read [AGENTS.md](AGENTS.md) first.** It contains the repository structure, the Agent Skills
specification constraints, the domain rules, and the git workflow. Everything there applies
here. This file covers only what is specific to Claude Code.

## Quick reference

```bash
./validate-skills.sh    # must pass before committing
python3 build.py        # regenerate dist/ and markdown/ after editing skills/
```

Edit `skills/`. Never hand-edit `dist/` or `markdown/` — both are generated.

## The five non-negotiables

Restated because they matter more than anything else in this repo:

1. **No statutory figures from memory.** No thresholds, rates, limits, or deadlines. Cite the
   form and line to verify instead.
2. **Every skill gates its own output** and refuses to deliver when the proof fails.
3. **Never plug or guess a number.** Unknowns become visible exceptions.
4. **Stop where judgment starts.** Automate the mechanical layer, not the professional
   conclusion.
5. **Local only, and never commit client data.**

See AGENTS.md for the full reasoning on each.

## Claude Code-specific notes

### Keep SKILL.md cross-agent compatible

These skills are meant to run on Claude Code, Codex, Cursor, Windsurf, and anything else
implementing the Agent Skills spec. Claude Code supports dynamic shell injection in skill
bodies using `` !`command` `` syntax, which other agents render as literal garbled text.

**Do not put `` !`command` `` in any `SKILL.md` in this repository.** If you want that
behavior locally, add it to your own `.claude/skills/` override rather than to the shared
skill.

### Plugin marketplace

This repo is installable as a Claude Code plugin:

```bash
/plugin marketplace add adoptai/cpa-skills
/plugin install cpa-skills
```

`.claude-plugin/marketplace.json` must list every skill directory. The validator fails if a
skill is missing from it or if the manifest lists a directory that no longer exists — a skill
absent from the manifest silently fails to install, which is hard to notice.

See the [Claude Code plugins documentation](https://code.claude.com/docs/en/plugins.md).

### When a user asks you to run one of these skills

Read the skill's `SKILL.md` and follow it as written rather than improvising a faster path. The
procedures encode audit and review discipline that looks like overhead until it catches
something:

- Run the completeness or integrity check **before** the analysis, not after. Anomaly results
  computed on a truncated population create false comfort.
- Do not suppress a failing gate to produce output. The failure is the finding.
- Do not classify items the skill deliberately leaves to the user.
- Report what you did not receive. A review missing the diagnostics list or prior-year
  carryforwards is limited in scope and must say so.

### Testing changes

Generate synthetic fixtures, plant a known error, and confirm the gate catches it. A gate that
has never failed has not been tested. Never use real client data.
