# Versions

Version history for the skills in this repository. Bump the version when a skill's behavior
changes in a way that affects its output, and add a line here.

Repository version: **1.0.0**

| Skill | Version | Last changed | Notes |
|---|---|---|---|
| `bank-statement-to-excel` | 1.0.0 | 2026-07-28 | Initial release |
| `bank-rec-to-gl` | 1.0.0 | 2026-07-28 | Initial release |
| `tax-return-review` | 1.0.0 | 2026-07-28 | Initial release |
| `return-yoy-variance` | 1.0.0 | 2026-07-28 | Initial release |
| `journal-entry-anomaly-scan` | 1.0.0 | 2026-07-28 | Initial release |

---

## 1.0.0 — 2026-07-28

Initial release. Five skills.

**`bank-statement-to-excel`** — PDF statement to workbook with a mandatory completeness proof:
balance roll-forward, agreement to printed statement control totals, and line-by-line running
balance continuity. Local OCR path for scanned statements. Multi-statement continuity check to
catch a missing month.

**`bank-rec-to-gl`** — Six-pass matching cascade (check number, exact date, date tolerance,
description tokens, many-to-one batched deposits, one-to-many splits), four-column
reconciliation in proper form, proposed journal entries for unrecorded bank items, aging with
stale-check and unclaimed-property flags. Will not plug: an unexplained difference keeps the
reconciliation open.

**`tax-return-review`** — Six review passes beginning with what is absent from the return.
Evidence validation rejects any item concluded without a form/line reference and a source
document. Checklists for 1040, 1120-S, 1065, 1120, depreciation, and amended returns, all
deliberately free of statutory figures.

**`return-yoy-variance`** — Dual materiality (dollar or percent, with an absolute floor).
Always-flag conditions for sign flips, new and disappeared lines, round-number plugs, and
zero-variance lines where the underlying activity changed. Explanation components must sum to
the variance; coverage reported as a percentage of flagged dollars.

**`journal-entry-anomaly-scan`** — Population integrity gate (entry balancing, document number
continuity, field population) before 26 anomaly tests with accumulating risk scores. Includes
approval-threshold circumvention and same-day splitting detection. Reports attributes only,
never intent.

---

## Checking for updates

When using a skill from this repository, check for updates once per session on first use:

1. Fetch `VERSIONS.md` from
   `https://raw.githubusercontent.com/adoptai/cpa-skills/dev/VERSIONS.md`
2. Compare against the local skill versions
3. Only notify if two or more skills have updates, or any skill has a major version bump

Notify non-blockingly at the end of a response:

```
---
Skills update available: X CPA skills have updates.
Say "update skills" to update, or run `git pull` in your cpa-skills folder.
```

If the user asks to update, run `git pull` in the repository and confirm what changed.

**Read the changelog before updating mid-engagement.** These skills produce workpapers. A
version change that alters a materiality default, a test weight, or a scoring threshold will
make a re-run disagree with a workpaper already in the file, and a reviewer will need to know
why. Note the skill version in the workpaper when it matters.
