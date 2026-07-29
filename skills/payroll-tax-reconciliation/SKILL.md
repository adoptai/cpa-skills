---
name: payroll-tax-reconciliation
description: Reconcile the payroll register to the quarterly Forms 941, the W-2/W-3 totals, and the general ledger — a four-way tie-out of gross wages, taxable wage bases, withholding, employer taxes, and deposits, with every difference either explained or reported as an exception. Use this whenever the user mentions payroll reconciliation, reconciling payroll to the GL, tying W-2s to 941s, Form 941 versus W-3, quarterly payroll tax returns, a payroll tax notice, W-2 totals that do not agree, payroll accrual review, or year-end payroll true-up. Also trigger on "do our 941s tie to the W-3," "reconcile payroll for the year," "why is our payroll tax liability off," "IRS sent a notice about our 941," "check the payroll register against the GL," or when a payroll register and payroll tax filings are provided together. Runs fully local — no payroll or employee data leaves the machine.
---

# Payroll Tax Reconciliation

This is the reconciliation that generates notices when it is skipped, and it gets skipped
constantly because it is four-sided rather than two-sided. The payroll register, the four Forms
941, the W-2/W-3 set, and the general ledger all describe the same wages, and they legitimately
differ — the register shows gross pay, the W-2 shows federal taxable wages after pre-tax
deferrals, the GL shows expense including accruals, and the 941s show quarterly slices.

**Because they legitimately differ, "they don't match" tells you nothing.** The work is proving
they differ *only* by amounts you can name. A discrepancy between the 941s and the W-3 is
matched automatically against agency records, so an unexplained difference here is not an
internal control matter — it is an incoming notice with penalties and interest attached.

## The standing rule on rates and wage bases

**Do not state a Social Security wage base, tax rate, deposit threshold, or filing deadline from
memory.** These change annually.

This skill takes a different and more reliable approach: it **derives the applied figures from
the data itself** and asks you to confirm them.

- The Social Security wage base is *inferred* from the register — the highest Social Security
  wages figure among employees whose Medicare wages exceed their Social Security wages. That is
  the cap the payroll system actually applied. The script reports it and instructs you to
  confirm it against current-year authority.
- Effective tax rates are *computed* from the forms (`tax ÷ taxable wages`) and compared across
  quarters and against the W-2 totals. A rate that shifts between quarters is a finding
  regardless of what the correct rate is.

This catches a wrong wage base or rate without asserting one — and it catches a *correctly
applied but mis-transcribed* figure, which a rate table would not.

## Inputs

1. **Payroll register** — full-year, by employee: gross wages, pre-tax deferrals, Section 125
   amounts, federal taxable wages, Social Security wages and tax, Medicare wages and tax,
   federal income tax withheld.
2. **All four Forms 941** as filed, including any 941-X amendments. A three-quarter set cannot
   be reconciled to an annual W-3.
3. **W-3 totals**, and per-employee W-2s if available. Per-employee data enables the
   strongest test in the set.
4. **General ledger** — wage expense, employer tax expense, and the payroll tax liability
   account, with the accrual balances at both ends of the year.
5. **Reconciling items you already know about** — third-party sick pay, group-term life,
   fringe benefits, non-cash compensation, tip income, deceased-employee payments, prior-year
   accrual reversals, and any restatement.

Ask for what is missing. A reconciliation performed without the 941-X amendments will disagree
with agency records for reasons that are not errors.

## Step 1 — Run the four-way tie

```bash
python3 scripts/payroll_recon.py \
  --register register.csv --f941 f941.csv --w3 w3.csv --gl gl.csv \
  --w2 w2_by_employee.csv \
  --reconciling-items recon_items.csv \
  --client "Northgate Manufacturing Inc" --year 2025 \
  --out "Northgate - 2025 Payroll Tax Reconciliation.xlsx"
```

Nine tests run. All must pass, or the workbook is marked failed.

**Wage bridges**

- **Test 1 — Gross to federal taxable.** Register gross wages, less pre-tax deferrals and
  Section 125, plus taxable fringes, equals federal taxable wages per the register, equals the
  sum of the 941 wage line across four quarters, equals W-3 box 1. This bridge is where
  cafeteria-plan and 401(k) treatment errors surface.
- **Test 2 — Social Security wages.** Sum of the four 941 Social Security wage lines equals W-3
  box 3, and equals the register total.
- **Test 3 — Medicare wages.** Same, against W-3 box 5.

**Tax and withholding**

- **Test 4 — Federal income tax withheld.** Sum of 941s equals W-3 box 2 equals register.
- **Test 5 — Social Security tax.** The 941 reports the combined employee and employer share;
  W-3 box 4 reports the employee share only. The script computes the implied employer share and
  flags any difference. Where the employer matches at the employee rate and there are no
  special items, the 941 figure is twice box 4 — but tips, group-term life, and adjustments
  legitimately break that, so differences are reported for classification rather than
  auto-failed.
- **Test 6 — Medicare tax.** Same treatment, including Additional Medicare Tax, which is
  employee-only and therefore breaks the doubling relationship by design. The script isolates
  it rather than burying it.

**Ledger**

- **Test 7 — GL wage expense.** Register gross wages, plus ending accrual, less beginning
  accrual, equals GL wage expense. A miss here is usually an accrual nobody reversed.
- **Test 8 — Deposits and liability rollforward.** Beginning payroll tax liability, plus taxes
  incurred per the 941s, less deposits made, equals ending liability. **The ending liability
  should be the taxes owed on the final period only.** A liability balance materially larger
  than one period's taxes means a deposit was missed or misapplied, and that is a penalty
  exposure worth escalating immediately.

**Per employee**

- **Test 9 — Employee-level integrity.** For each employee: W-2 box amounts agree to register
  YTD; Social Security wages do not exceed the inferred wage base; Social Security wages do not
  exceed Medicare wages; and no employee has taxable wages with zero withholding and no
  exemption on file. Test 9 is the one that finds real errors, because aggregate totals can tie
  while two employees' figures are swapped.

## Step 2 — Classify differences

Unexplained differences are worthless; classified ones are the deliverable. Supply
`--reconciling-items` with a `cause`, `amount`, and `evidence` for each, and the components must
sum to the difference. Common legitimate causes:

- Pre-tax deferrals and Section 125 reducing box 1 but not Social Security and Medicare wages
- Employees above the Social Security wage base — expected, and the count should equal the
  number of employees whose Medicare wages exceed their Social Security wages
- Taxable fringe benefits, group-term life over the excludable amount, personal use of a
  company vehicle
- Third-party sick pay, which appears on the W-2 but may sit outside the register
- Tip income and the associated credit
- 941-X amendments filed after the original returns
- Accrual timing between the register and the GL
- Non-cash and deferred compensation
- Employees paid in one quarter and terminated in another, where the register was re-run

Anything that cannot be named goes to exceptions. Do not plug the payroll liability account —
that account is where unremitted trust-fund taxes hide, and those carry personal liability for
responsible persons. Treat an unexplained credit balance there as urgent.

## Step 3 — Deliver

**Workbook tabs:**

1. **Reconciliation** — the four-way bridge in proper form: register → 941s → W-2/W-3 → GL, with
   each reconciling item named and the residual difference reading `0.00`. The signable page.
2. **Test Results** — all nine tests with amounts, differences, and pass/fail.
3. **Quarterly Analysis** — each 941 side by side: wages, taxes, deposits, computed effective
   rates. Rate drift between quarters is visible here and nowhere else.
4. **Employee Detail** — per employee: register YTD, W-2 boxes, differences, wage-base status,
   and flags.
5. **Liability Rollforward** — beginning balance, taxes incurred, deposits, ending balance,
   with the ending balance compared to the final period's taxes.
6. **Exceptions & Memo** — unexplained differences, escalations, control observations, and the
   figures requiring confirmation against current-year authority.

**Then, in chat:** the four control totals, whether they tie, the reconciling items, and
anything unremitted. Lead with penalty exposure if any exists — a missed deposit matters more
than a reconciled wage bridge.

## What to escalate immediately

- **Payroll tax liability materially exceeding one period's taxes** — likely a missed or
  misapplied deposit. Deposit penalties escalate with lateness, so days matter.
- **Any unremitted trust-fund tax** — withheld employee tax not deposited. Responsible-person
  liability attaches personally and does not discharge in bankruptcy. This is the most serious
  finding this reconciliation can produce.
- **941s and W-3 disagreeing** — an agency mismatch notice is essentially automatic.
- **An employee with taxable wages and zero income tax withholding** and no exemption
  certificate on file.
- **Social Security wages exceeding the inferred base for any employee** — either the base was
  applied wrongly, or two employees were merged in the register.
- **A worker paid on 1099 whose pattern resembles the W-2 population** — classification
  exposure. Report the pattern; do not conclude on status.
- **Officer or shareholder compensation absent** where the entity has profitable operations and
  an active owner — a reasonable-compensation question, and one of the most examined areas for
  S corporations.

State these as facts with the supporting figures. Recommending a voluntary correction, an
amended filing, or a penalty abatement request is a decision for the person signing.

## Security posture

Fully local. Employee-level payroll data is among the most sensitive information a firm holds.
No network calls, no uploads, no telemetry. **Full SSNs are never written to output** — employee
identification uses name or employee ID and last four digits only, and the script rejects any
input containing a full SSN pattern. Output filenames carry the client name and year only.

## Dependencies

```bash
pip install openpyxl
```

---
*Built by [Adopt AI](https://adopt.ai) — free to use and modify.*
