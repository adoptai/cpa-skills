---
name: pbc-assembler
description: Work through an auditor's PBC (provided-by-client) request list against a local document folder listing, matching each request to candidate files by client-supplied keywords, and tracking completeness without ever claiming a request is satisfied just because a file was found. Use this whenever the user mentions a PBC list, PBC request, provided-by-client documents, audit documentation requests, assembling the audit package, gathering support for the auditors, or tracking what's outstanding for fieldwork. Also trigger on "work through the PBC list," "find the documents for this audit request," "what's still outstanding for the auditors," "assemble the audit package," "match these requests to our files," or when a PBC request list and a folder listing are provided together. Pairs with `audit-workpaper`, which documents a procedure once a request is fulfilled; this skill only locates candidates and reads a folder listing you supply — never a remote API. Runs fully local.
---

# Audit PBC Assembler

Audit season turns into hours of opening folders and guessing which file answers which line on
the request list. This skill does the matching mechanically — it takes the keywords the client
or engagement team says would identify a responsive document, and checks them against a folder
listing, so a human reviews candidates instead of searching blind. It does not decide what counts
as a match beyond the keywords it's given, and it does not decide whether a matched document
actually answers the request — both are left to the person who signs off.

## The gate

The script refuses to produce a clean assembly when:

1. **A request_id is duplicated or blank.** Every request must resolve to exactly one row in the
   output — a duplicate ID means the request list itself is unreliable and nothing downstream can
   be trusted until it's fixed.
2. **A status doesn't resolve cleanly.** Every request ends in exactly one of three states:
   `matched` (at least one candidate file found), `gap` (keywords supplied, nothing found), or
   `no_keywords_supplied` (the request has no keywords at all). A request with no keywords is
   never silently treated as a gap — that would claim a search happened when it didn't.
3. **A match has no evidence trail.** Every `matched` row names the specific file path(s) and the
   keyword that triggered each match, so a reviewer can check it isn't a false positive — a
   filename that happens to contain a common word.

## Inputs

1. **The PBC request list** — one row per request, with a `keywords` column the client or
   engagement team supplies. The script never derives its own keywords from the free-text
   description; what counts as a responsive filename is a judgment call made by a person, not
   inferred from "November 2025 bank statement" into search terms.
2. **A local document folder listing.** Generate this yourself — e.g. `find . -type f
   -printf "%p,%f\n"` into a CSV with a header row — and supply it as input. The script does not
   walk the filesystem itself; that keeps the match fully auditable from two flat files.
3. **A prior-year file index** (optional), for cross-referencing recurring requests to where they
   were satisfied last year.

## Step 1 — Build the inputs

```bash
python3 scripts/pbc_assembler.py --template-requests pbc_requests.csv
python3 scripts/pbc_assembler.py --template-folder document_folder.csv
```

Fill in `keywords` per request as a pipe-separated list of terms that would appear in a
responsive filename or path (`bank statement|bank stmt|Nov 2025 bank`). Generate the folder
listing from the actual engagement folder structure.

## Step 2 — Run the assembler

```bash
python3 scripts/pbc_assembler.py --requests pbc_requests.csv \
  --folder document_folder.csv --prior-year prior_year_index.csv \
  --client "Northwind Systems Inc" --period-end 2025-12-31 \
  --out "Northwind - FY2025 PBC Status.xlsx"
```

Matching is a case-insensitive substring check of each keyword against the file name and path.
A request can match multiple files; a file can satisfy multiple requests.

## Step 3 — Deliver

**Workbook tabs:**

1. **Workpaper Summary** — result counts, exceptions, and a conclusion block that states plainly
   that a match is candidate evidence, not a confirmed response.
2. **PBC Status** — every request with its status, matched file(s), the triggering keyword(s),
   due date, and who it's requested from.
3. **Gap List** — every `gap` and `no_keywords_supplied` request, sorted by due date.
4. **Prior-Year Carryforward** — where a prior-year index was supplied, what satisfied the same
   request last time.
5. **Exceptions** — structural problems in the request list itself.

**Then, in chat:** the gap count and the nearest due dates first, then anything with no
keywords supplied (since those haven't actually been searched yet), then note explicitly that
matches are candidates for review, not confirmed responses.

## What to escalate

- **A request with no keywords supplied and an imminent due date** — nothing has actually been
  searched for it yet, and that's easy to mistake for "in progress."
- **A request matched to many unrelated files** — a keyword that's too generic (a common word, a
  short acronym) produces noise that looks like progress but isn't.
- **A recurring request that matched last year but shows a gap this year** — worth confirming
  whether the document simply wasn't filed yet or the process that produces it changed.
- **Any match the preparer opens and finds does not actually respond to the request** — that's
  exactly the determination this skill does not make, and it should be corrected in the workpaper
  that documents the actual procedure, not assumed from the match alone.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry, and no
filesystem access beyond reading the CSV files you point it at — it never opens or inspects the
documents themselves, only the folder listing's file names and paths.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
