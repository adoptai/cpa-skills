<!--
Do NOT include real client data in this PR — no statements, GL extracts, returns,
spreadsheets, PDFs, client names, or TINs. Synthetic fixtures only.
-->

## What this changes

<!-- One or two sentences. -->

## Why

<!-- What was wrong, or what this now makes possible. -->

## Type

- [ ] New skill
- [ ] Fix to an existing skill
- [ ] New parser format support
- [ ] False-positive suppression / test tuning
- [ ] Checklist addition
- [ ] Documentation

## How it was tested

<!--
Generate a synthetic fixture, plant a known error, confirm the gate catches it.
A gate that has never failed has not been tested. Paste the relevant output.
-->

```
```

## Behavior change

<!--
Does this alter output for a run that previously succeeded? Materiality defaults,
test weights, scoring thresholds, and matching tolerances all change what lands in a
workpaper. If so, say what changes and note it in VERSIONS.md — a reviewer comparing a
re-run against a workpaper already in the file will need to know why they disagree.
-->

- [ ] No change to existing output
- [ ] Output changes — described above and recorded in VERSIONS.md

## Checklist

- [ ] `./validate-skills.sh` passes
- [ ] `python3 build.py` run, `dist/` and `markdown/` committed
- [ ] New skill added to `.claude-plugin/marketplace.json`
- [ ] `SKILL.md` under 500 lines, `description` under 1024 chars, `name` matches directory
- [ ] Output gate defined and enforced
- [ ] No statutory thresholds, rates, limits, or deadlines asserted from memory
- [ ] Judgment boundary preserved — no professional conclusions automated
- [ ] Runs entirely locally; no network calls added
- [ ] `VERSIONS.md` updated
- [ ] No client data anywhere in the diff
