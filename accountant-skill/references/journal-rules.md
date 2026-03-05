# Journal Rules — Books of Prime Entry

## Overview

The books of prime entry (journals) are the first point of recording for all business transactions. Each journal captures a specific type of transaction. All entries follow the double-entry bookkeeping principle: every debit must have an equal credit.

## The Six Journals

| Journal | Purpose | Typical Debit | Typical Credit |
|---------|---------|---------------|----------------|
| Sales Journal | Credit sales to customers | 11000 Accounts Receivable | 40000 Sales Revenue |
| Purchases Journal | Credit purchases from suppliers | 50000-50299 COGS / 12000 Inventory | 20000 Accounts Payable |
| Cash Receipts Journal | All money received | 10100 Cash at Bank | Various (11000 AR, 40000 Sales, etc.) |
| Cash Payments Journal | All money paid out | Various (20000 AP, 60000-69000 Expenses) | 10100 Cash at Bank |
| Payroll Journal | Salary & wage entries | 61000 Office Salaries | 10100 Bank / 22200 Wages Payable |
| General Journal | All other entries (adjustments, corrections, non-cash) | Various | Various |

## Validation Rules

### Per-Entry Validation
1. **Date must be present** and within the reporting period
2. **Reference number must be present** and unique within the journal
3. **At least one debit and one credit account** must be specified
4. **Debit Amount = Credit Amount** for every entry (or group of entries sharing the same reference)
5. **Account codes must exist** in chart_of_accounts.xlsx (5-digit codes)
6. **Amounts must be positive** (never negative in a journal)

### Compound Entries
Some entries have multiple debits or credits (e.g., payroll with deductions). These share the same reference/JV number and span multiple rows. Validation rule: sum of all debit amounts = sum of all credit amounts for the group.

### Cross-Journal Validation
After summarizing all journals:
- Total debits across ALL journals should equal total credits across ALL journals
- No transaction should appear in two different journals (check by reference number)

## Summarization Logic

For each journal:
1. Filter rows where Date falls within the reporting period
2. Group by account code
3. For each account: sum Debit amounts and sum Credit amounts
4. Calculate transaction count per account
5. Compute journal totals: grand total debits, grand total credits
6. Flag exceptions: unbalanced entries, missing dates, missing references, duplicate references

For the consolidated summary:
1. Merge all journal summaries
2. For each account code, sum debits and credits across all journals
3. Verify grand total debits = grand total credits
4. Sort by account code

## Output Structure

The summary workbook should contain:
- **Dashboard sheet**: High-level metrics (total transactions, total amounts, balance check)
- **One sheet per journal**: Account breakdown with debits/credits/count
- **Consolidated sheet**: All accounts across all journals
- **Exceptions sheet**: Any validation failures with journal name, row number, and description of issue
