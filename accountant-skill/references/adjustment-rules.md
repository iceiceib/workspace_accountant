# Adjustment Rules — Period-End Adjusting Entries

## Overview

Adjusting entries are made at the end of an accounting period to ensure revenues and expenses are recorded in the correct period (accrual basis of accounting). They are recorded in the General Journal and posted to the General Ledger before preparing the adjusted trial balance.

## Types of Adjustments

### 1. Accrued Expenses (Expenses incurred but not yet paid)

Expenses that have been incurred during the period but payment hasn't been made yet.

| Example | Debit | Credit |
|---------|-------|--------|
| Wages earned by employees, not yet paid | 5100 Salaries & Wages | 2030 Accrued Wages |
| Utilities consumed, bill not received | 5210 Utilities Expense | 2020 Accrued Expenses |
| Interest on loan, not yet due | 5910 Interest Expense | 2020 Accrued Expenses |
| Rent for current period, not yet paid | 5200 Rent Expense | 2020 Accrued Expenses |

**Calculation**: Determine the amount of expense attributable to the current period based on time, usage, or contractual terms.

### 2. Accrued Revenue (Revenue earned but not yet received)

Revenue that has been earned during the period but cash hasn't been received yet.

| Example | Debit | Credit |
|---------|-------|--------|
| Services delivered, not yet invoiced | 1100 Accounts Receivable | 4040 Sales Revenue |
| Interest earned on bank deposit | 1020 Cash at Bank | 4110 Interest Income |

### 3. Prepaid Expenses (Expenses paid in advance, now used up)

Payments made in advance for future benefits. At period end, the used portion becomes an expense.

| Example | Debit | Credit |
|---------|-------|--------|
| Insurance used this month | 5700 Insurance Expense | 1320 Prepaid Insurance |
| Rent paid in advance, now used | 5200 Rent Expense | 1310 Prepaid Rent |
| Supplies consumed | 5410 Office Supplies | 1300 Prepaid Expenses |

**Calculation**: (Total Prepaid / Total Months) × Months Used in Period

### 4. Unearned Revenue (Revenue received in advance, now earned)

Cash received before the service/goods have been delivered. At period end, the earned portion becomes revenue.

| Example | Debit | Credit |
|---------|-------|--------|
| Deposit received, now goods delivered | 2040 Unearned Revenue | 4040 Sales Revenue |

### 5. Depreciation

Systematic allocation of the cost of a non-current asset over its useful life.

**Straight-Line Method:**
```
Annual Depreciation = (Cost - Salvage Value) / Useful Life
Monthly Depreciation = Annual Depreciation / 12
```

**Reducing Balance Method:**
```
Depreciation = Net Book Value × Depreciation Rate
Where: Depreciation Rate = 1 - (Salvage Value / Cost)^(1/Useful Life)
```

| Entry | Debit | Credit |
|-------|-------|--------|
| Monthly depreciation — Buildings | 5300 Depreciation Expense | 1611 Accum. Depr. — Buildings |
| Monthly depreciation — Plant & Machinery | 5300 Depreciation Expense | 1621 Accum. Depr. — P&M |
| Monthly depreciation — Furniture | 5300 Depreciation Expense | 1631 Accum. Depr. — F&F |
| Monthly depreciation — Vehicles | 5300 Depreciation Expense | 1641 Accum. Depr. — Vehicles |
| Monthly depreciation — Equipment | 5300 Depreciation Expense | 1651 Accum. Depr. — Equipment |

**Source data**: Read from `fixed_asset_register.xlsx` to calculate depreciation per asset, then aggregate by category.

### 6. Bad Debt Provision

An estimate of accounts receivable that may not be collected.

**Methods:**
- **Percentage of receivables**: Allowance = Total AR × Estimated Bad Debt %
- **Aging method**: Different % applied to each aging bucket (30d, 60d, 90d, 120d+)

| Entry | Debit | Credit |
|-------|-------|--------|
| Increase provision for bad debts | 5800 Bad Debt Expense | 1110 Allowance for Doubtful Debts |
| Decrease provision (if over-provided) | 1110 Allowance for Doubtful Debts | 5800 Bad Debt Expense |
| Write off specific bad debt | 1110 Allowance for Doubtful Debts | 1100 Accounts Receivable |

### 7. Inventory Adjustments

Adjustments when physical stock count differs from book records.

| Scenario | Debit | Credit |
|----------|-------|--------|
| Stock shortage (book > actual) | 5010 Raw Materials Used / 5900 Misc Expense | 1200 Inventory — Raw Materials |
| Stock surplus (actual > book) | 1200 Inventory — Raw Materials | 5010 Raw Materials Used |
| COGS adjustment (periodic system) | 5010 Raw Materials Used | 1200 Inventory — Raw Materials |

### 8. Error Corrections

Corrections for mistakes discovered during the period.

| Error Type | Treatment |
|------------|-----------|
| Wrong account debited/credited | Reverse the incorrect entry, record the correct one |
| Wrong amount | Record the difference (additional debit/credit) |
| Omitted transaction | Record the entry as if it were made on the original date |
| Duplicate entry | Reverse one of the duplicate entries |

### 9. Bank Reconciliation Adjustments

Entries identified during bank reconciliation (Module 3) that need to be recorded in the cash book.

These are generated automatically by Module 3's "Required Adjusting Entries" output.

## Adjusting Entry Format

Each adjusting entry should include:

| Field | Description |
|-------|-------------|
| Entry No | Sequential: ADJ-001, ADJ-002, etc. |
| Date | Last day of the reporting period |
| Type | Accrual, Prepaid, Depreciation, Bad Debt, Inventory, Error, Bank Recon |
| Description | Clear narration of what and why |
| Debit Account | Account code |
| Debit Amount | Amount |
| Credit Account | Account code |
| Credit Amount | Amount (must equal Debit Amount) |
| Supporting Reference | Source document, calculation, or reason |

## Closing Entries (End of Financial Year)

At year-end, revenue and expense accounts are closed to Retained Earnings:

1. **Close revenue accounts**: Debit all revenue accounts, Credit 3040 Current Year Earnings
2. **Close expense accounts**: Debit 3040 Current Year Earnings, Credit all expense accounts
3. **Transfer net income**: Debit 3040 Current Year Earnings, Credit 3030 Retained Earnings

After closing entries, only permanent accounts (Assets, Liabilities, Equity) have balances.

## Output Structure

The adjusting entries workbook should contain:
- **Summary**: Count and total of adjustments by type
- **Detailed Entries**: Full journal entry format (Entry No, Date, Accounts, Amounts, Description)
- **Depreciation Schedule**: Per-asset depreciation calculation (if depreciation was processed)
- **Impact Analysis**: For each account affected, show balance before adjustment, adjustment amount, and balance after adjustment
