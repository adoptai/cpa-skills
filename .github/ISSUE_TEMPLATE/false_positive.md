---
name: False positive
about: A test flags something benign, repeatedly
title: "fix: false positive in "
labels: false-positive
---

<!--
Do NOT include real client data. Describe the pattern, not the file.
-->

False positives matter more than they look. An exception report that cries wolf every month
trains people to skim it, and then a real finding gets skimmed too. So this is a genuine
defect, not a nitpick.

**Skill and test**
<!-- e.g. journal-entry-anomaly-scan, ROUNDCOMPUTED -->

**What gets flagged**
<!-- Describe the shape of the transaction or entry, with synthetic figures. -->

**Why it's benign in your environment**
<!-- e.g. "our depreciation is contractual and lands on round thousands by design",
     "our shared-service team in Manila posts at 22:00 local, so everything looks after-hours" -->

**How often**

- Roughly how many times per period:
- Roughly what share of that test's flags:

**Suggested handling**

- [ ] Suppress by default
- [ ] Keep, but lower the weight
- [ ] Make it configurable
- [ ] Keep as-is, just document the expected cause
- [ ] Not sure

**Confirmation**

- [ ] This report contains no real client data
