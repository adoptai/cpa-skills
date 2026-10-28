---
name: dunning-drafts
description: Prioritizes overdue AR from the aging, matches each invoice to the client's own escalation policy, and drafts a collection email per customer — never sending anything, always "drafted — pending review," with disputed invoices excluded entirely. Use this whenever the user mentions collections emails, dunning letters, AR follow-up drafts, overdue invoice reminders, escalation emails for past-due customers, or collections outreach. Also trigger on "draft collection emails from the aging," "build the dunning queue," "write follow-up emails for overdue accounts," "escalate these past-due customers," or when an AR aging and an escalation policy are supplied together. Pairs with ar-aging-tie-out, which should tie out the aging before it's fed in here — this skill takes the downstream action on an aging that's already been verified, and a human approves every send. Runs fully local — no customer or AR data leaves the machine.
---

# Collections Dunning Drafts

This produces drafts, never sends. Every email lands with status "drafted — pending review."
Nothing in this skill has the ability to send mail, and the workbook format is deliberately
built to be read and approved by a person before anything goes out.

## The gate

**Hard gate** — blocks the workbook unless `--force`:

1. **No usable escalation policy supplied.** Bucket day ranges and wording are entirely
   client-supplied; without them there is nothing to drive escalation from, and the script never
   assumes a standard 30/60/90 convention.
2. **The tie-out fails.** The dollar total appearing across all drafted emails must equal,
   exactly, the sum of the non-disputed, bucket-matched invoices used to build them. This should
   be true by construction — a break here means something is wrong internally, and it must never
   reach a reviewer unexamined.

**Soft exceptions** — routed to the Review Queue / Exceptions tab, workbook still produced:

- An invoice that matches zero or more than one policy bucket — never forced into the nearest
  one.
- A customer with no contact email on file — no draft is auto-generated for them; they become an
  exception needing contact info.
- Any invoice flagged disputed — excluded from that customer's draft entirely. If every one of a
  customer's open invoices is disputed, that customer gets no draft at all.
- A template referencing a field this script doesn't supply — that customer gets no draft rather
  than a half-rendered email.

## Inputs

1. **AR aging** — `customer, invoice_no, amount, due_date, days_past_due`.
2. **Customer contacts** — `customer, contact_name, email`.
3. **Escalation policy** — `bucket_min_days, bucket_max_days, tier_name, tone, template_text`.
   Thresholds and wording are the client's own; nothing here is assumed.
4. **Dispute flags** (optional) — `customer, invoice_no, disputed (y/n)`.

## Step 1 — Emit a blank worksheet

```bash
python3 scripts/dunning_drafts.py --template --out aging_template.csv
```

## Step 2 — Build the queue

```bash
python3 scripts/dunning_drafts.py \
  --aging aging.csv --contacts contacts.csv --policy policy.csv \
  --disputes disputes.csv \
  --client "Northwind Systems Inc" --as-of 2025-12-31 \
  --out "Northwind - Collections Queue.xlsx"
```

## Step 3 — How a customer's tier is chosen

A customer may have invoices spanning more than one escalation bucket. The script uses the tier
of the **most-escalated bucket** that applies to any of their open, non-disputed invoices, and
that single template drafts one email listing all of those invoices — not one email per bucket.
`template_text` can reference `{customer}`, `{contact_name}`, `{total}`, `{invoice_count}`,
`{invoice_list}`, `{oldest_due_date}`, `{tier_name}`, and `{tone}`.

## Deliver

**Workbook tabs:** Collection Queue (customer, tier, total overdue, oldest due date, invoice
count — sorted largest first), Drafted Emails (full rendered text per customer, status always
"drafted — pending review"), Review Queue - Exceptions (disputes, missing contacts, unmatched
buckets, template errors), Tie-Out (drafted total vs. consumed invoice total).

**Then, in chat:** the gate result, how many customers were drafted and the total dollar amount,
then the exception count with a one-line reason breakdown. Be explicit that nothing has been
sent and nothing in this skill is capable of sending.

## What to escalate

- A customer landing on the Review Queue repeatedly for a missing contact — usually a stale CRM
  record, worth fixing once rather than exceiting every run.
- A disputed invoice that's been disputed for several consecutive aging cycles — may need
  resolution outside the collections process entirely.
- A policy bucket that never matches anything, or one that matches an unexpectedly large share
  of the population — may signal the policy thresholds no longer reflect how the business ages
  receivables.
- Any customer whose escalation tier jumped more than one level between runs — worth a human
  glance before the firmer tone goes out.

## Security posture

Fully local: standard library plus `openpyxl`. No network calls, no uploads, no telemetry, and
no email-sending capability of any kind. Customer and AR data never leave the machine.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
