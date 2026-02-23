---
name: accountant
description: "Full-cycle accounting skill for SME businesses using .xlsx files. Use this skill whenever the user asks to: summarize journals or books of prime entry, post to or summarize ledgers, reconcile bank accounts, create adjusting journal entries, generate trial balances (unadjusted/adjusted/post-closing), or produce financial statements (Income Statement, Balance Sheet, Cash Flow). Also trigger when the user mentions accounting cycle, closing entries, depreciation schedules, AR/AP aging, or any bookkeeping task involving .xlsx data. This skill works exclusively with .xlsx Excel files — no Google Sheets."
---

# Accountant Skill

An end-to-end accounting cycle skill for SME businesses (specifically Shwe Mandalay Cafe / K&K Finance Team). All inputs and outputs are `.xlsx` Excel files processed with Python (openpyxl + pandas).

## Quick Reference — Modules

| Module | Command Pattern | What It Does |
|--------|----------------|--------------|
| 1 | "Summarize journals for [PERIOD]" | Reads 6 journals, validates entries, produces summary |
| 2 | "Summarize ledgers for [PERIOD]" | Reads all ledgers, calculates balances, cross-checks |
| 3 | "Reconcile bank for [PERIOD]" | Matches cash book vs bank statement |
| 4 | "Process adjusting entries for [PERIOD]" | Creates depreciation, accruals, corrections |
| 5 | "Generate trial balance for [PERIOD]" | Produces unadjusted/adjusted/post-closing TB |
| 6 | "Generate financial statements for [PERIOD]" | P&L, Balance Sheet, Cash Flow |
| 7 | "Validate accounting for [PERIOD]" | Full-cycle integrity checks |
| ALL | "Run full accounting cycle for [PERIOD]" | Modules 1→7 in sequence |

## Before You Start

1. Read `references/data-schemas.md` to understand the expected .xlsx column structures
2. Read `references/chart-of-accounts.md` for the COA mapping
3. Ensure all input .xlsx files are in the user's specified working directory
4. Install dependencies: `pip install openpyxl pandas numpy`

## Accounting Cycle Flow

Execute modules in this order for a complete period close:

```
Journals (Module 1) → Ledgers (Module 2) → Unadjusted TB (Module 5)
→ Bank Recon (Module 3) → Adjustments (Module 4) → Adjusted TB (Module 5)
→ Financial Statements (Module 6) → Validation (Module 7)
```

## Module Details

Read the appropriate reference file before executing each module:

### Module 1: Books of Prime Entry Summary
- **Reference:** `references/journal-rules.md`
- **Script:** `scripts/summarize_journals.py`
- **Inputs:** sales_journal.xlsx, purchases_journal.xlsx, cash_receipts_journal.xlsx, cash_payments_journal.xlsx, payroll_journal.xlsx, general_journal.xlsx
- **Output:** books_of_prime_entry_summary_[PERIOD].xlsx

### Module 2: Ledger Summarization
- **Reference:** `references/ledger-posting-rules.md`
- **Script:** `scripts/summarize_ledgers.py`
- **Inputs:** general_ledger.xlsx, accounts_receivable_ledger.xlsx, accounts_payable_ledger.xlsx, cash_ledger.xlsx, fixed_assets_ledger.xlsx, equity_ledger.xlsx
- **Output:** ledger_summary_[PERIOD].xlsx

### Module 3: Bank Reconciliation
- **Reference:** `references/reconciliation-rules.md`
- **Script:** `scripts/reconcile_bank.py`
- **Inputs:** cash_ledger.xlsx, bank_statement.xlsx
- **Output:** bank_reconciliation_[PERIOD].xlsx

### Module 4: Journal Adjustments
- **Reference:** `references/adjustment-rules.md`
- **Script:** `scripts/journal_adjustments.py`
- **Inputs:** User instructions + fixed_asset_register.xlsx + accounts_receivable_ledger.xlsx
- **Output:** adjusting_journal_entries_[PERIOD].xlsx

### Module 5: Trial Balance
- **Reference:** `references/chart-of-accounts.md`
- **Script:** `scripts/generate_trial_balance.py`
- **Inputs:** general_ledger.xlsx, chart_of_accounts.xlsx, adjusting_journal_entries_[PERIOD].xlsx (optional)
- **Output:** trial_balance_[PERIOD].xlsx (multi-sheet: unadjusted, adjusted, post-closing)

### Module 6: Financial Statements
- **Reference:** `references/financial-report-formats.md`
- **Script:** `scripts/generate_financials.py`
- **Inputs:** trial_balance_[PERIOD].xlsx (adjusted), chart_of_accounts.xlsx
- **Output:** financial_statements_[PERIOD].xlsx (Income Statement, Balance Sheet, Cash Flow, Schedules)

### Module 7: Validation
- **Script:** `scripts/validate_accounting.py`
- **Inputs:** All outputs from Modules 1-6
- **Output:** audit_validation_[PERIOD].xlsx

## Critical Rules

1. **Double-entry integrity** — Every journal entry MUST have Debit = Credit. Stop and report if not.
2. **COA validation** — Every account code must exist in chart_of_accounts.xlsx. Flag unknown codes.
3. **Period filtering** — Always filter data by the user-specified date range before processing.
4. **Excel formulas** — Use Excel formulas (=SUM, =SUMIF, etc.) in output files, not hardcoded Python calculations, so spreadsheets remain dynamic.
5. **Professional formatting** — All outputs must have headers, borders, number formatting, and color coding per the xlsx skill standards.
6. **Myanmar text** — Handle Myanmar Unicode text in descriptions and narrations without errors.
7. **Error-first** — Validate inputs before processing. If a required file is missing or malformed, stop and tell the user what's wrong.
8. **Idempotent** — Re-running a module for the same period produces the same output. Never append duplicates.

## Output Formatting Standards

Follow these formatting rules for all generated .xlsx files:
- **Header row**: Bold, dark blue background (#1F4E79), white text, centered
- **Number format**: #,##0 for amounts; 0.0% for percentages
- **Negative numbers**: Red font, parentheses format: (#,##0)
- **Borders**: Thin borders on all data cells; thick bottom border on totals
- **Column widths**: Auto-fit to content, minimum 12 characters
- **Sheet tab colors**: Green for summary sheets, Blue for detail sheets, Red for exceptions
- **Freeze panes**: Freeze header row and first column on all sheets
