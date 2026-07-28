---
name: New skill proposal
about: Propose a new accounting or tax skill
title: "feat: add  skill"
labels: new-skill
---

<!--
Please open a proposal before writing the skill so we can agree on scope first.
Do NOT include real client data anywhere in this issue.
-->

**The task, phrased the way a practitioner would ask for it**
<!-- e.g. "Reconcile the payroll register to the payroll tax filings and the GL." -->

**Who does this today, and what it costs them**
<!-- Role, rough frequency, rough time per occurrence. -->

**Why this needs a skill rather than a plain prompt**
<!-- Usually: a multi-step procedure, a non-obvious professional standard, or a proof
     that has to be enforced rather than suggested. -->

**The output gate** ← required
<!--
Every skill here proves its work and refuses to deliver a clean result when the proof
fails. Name the check that would make this skill refuse.

Examples from existing skills:
  - the extract must tie to the statement's own opening and closing balance
  - the unexplained difference must be exactly zero
  - every tested item must carry a form/line reference and a source document
  - explanation components must sum to the variance
  - every journal entry must balance before any anomaly test runs

If you can't name one, this task may not need a skill.
-->

**What it deliberately will NOT do**
<!-- Where does judgment start? What stays with the professional? -->

**Inputs it needs**
<!-- Documents, exports, prior-period files. Note which are commonly unavailable. -->

**Deliverable**
<!-- Workbook tabs, memo, journal entries, exception listing. -->

**Anything it must not assert**
<!-- Reminder: no statutory thresholds, rates, limits, or deadlines from memory —
     these change annually. Cite the form and line to verify instead. -->

**Confirmation**

- [ ] I've read [CONTRIBUTING.md](../../CONTRIBUTING.md)
- [ ] This proposal contains no real client data
- [ ] I've named an output gate
