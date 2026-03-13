# Adjustment Rules — Period-End Adjusting Entries

## Overview

Adjusting entries are made at the end of an accounting period to ensure revenues and expenses are recorded in the correct period (accrual basis of accounting). They are recorded in the General Journal and posted to the General Ledger before preparing the adjusted trial balance.

## Types of Adjustments

### 1. Accrued Expenses (Expenses incurred but not yet paid)

Expenses that have been incurred during the period but payment hasn't been made yet.

| Example | Debit | Credit |
|---------|-------|--------|
| Wages earned by employees, not yet paid | 61000 Office Salaries | 22200 Wages Payable |
| Utilities consumed, bill not received | 63000 Utilities Expense | 22000 Utility Bills |
| Interest on loan, not yet due | 68000 Other Expenses | 20000 Accounts Payable |
| Rent for current period, not yet paid | 68000 Other Expenses | 20000 Accounts Payable |

**Calculation**: Determine the amount of expense attributable to the current period based on time, usage, or contractual terms.

### 2. Accrued Revenue (Revenue earned but not yet received)

Revenue that has been earned during the period but cash hasn't been received yet.

| Example | Debit | Credit |
|---------|-------|--------|
| Services delivered, not yet invoiced | 11000 Accounts Receivable | 40000 Sales Revenue |
| Interest earned on bank deposit | 10100 Cash at Bank | 70000 Interest Income |

### 3. Prepaid Expenses (Expenses paid in advance, now used up)

Payments made in advance for future benefits. At period end, the used portion becomes an expense.

| Example | Debit | Credit |
|---------|-------|--------|
| Insurance used this month | 68000 Other Expenses | 13000 Advanced Payments |
| Rent paid in advance, now used | 68000 Other Expenses | 13000 Advanced Payments |
| Supplies consumed | 65000 Factory & Office Supplies | 13000 Advanced Payments |

**Calculation**: (Total Prepaid / Total Months) × Months Used in Period

### 4. Unearned Revenue (Revenue received in advance, now earned)

Cash received before the service/goods have been delivered. At period end, the earned portion becomes revenue.

| Example | Debit | Credit |
|---------|-------|--------|
| Deposit received, now goods delivered | 13000 Advanced Payments | 40000 Sales Revenue |

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
| Monthly depreciation — Buildings | 66000 Depreciation Expense | 15110 Accum. Depr. — Buildings |
| Monthly depreciation — Machinery | 66000 Depreciation Expense | 15210 Accum. Depr. — Machinery |
| Monthly depreciation — Office Equipment | 66000 Depreciation Expense | 15310 Accum. Depr. — Office Equipment |
| Monthly depreciation — Electrical Systems | 66000 Depreciation Expense | 15410 Accum. Depr. — Electrical |
| Monthly depreciation — Motor Vehicles | 66000 Depreciation Expense | 15510 Accum. Depr. — Motor Vehicles |

**Source data**: Read from `fixed_asset_register.xlsx` to calculate depreciation per asset, then aggregate by category.

### 6. Bad Debt Provision

An estimate of accounts receivable that may not be collected.

**Methods:**
- **Percentage of receivables**: Allowance = Total AR × Estimated Bad Debt %
- **Aging method**: Different % applied to each aging bucket (30d, 60d, 90d, 120d+)

| Entry | Debit | Credit |
|-------|-------|--------|
| Increase provision for bad debts | 67000 Inventory Write-off | 12300 Inventory Adjustments |
| Decrease provision (if over-provided) | 12300 Inventory Adjustments | 67000 Inventory Write-off |
| Write off specific bad debt | 12300 Inventory Adjustments | 11000 Accounts Receivable |

### 7. Inventory Adjustments

Adjustments when physical stock count differs from book records.

| Scenario | Debit | Credit |
|----------|-------|--------|
| Stock shortage (book > actual) | 50010 Purchases Raw Materials / 68000 Other Expenses | 12000 Inventory — Raw Materials |
| Stock surplus (actual > book) | 12000 Inventory — Raw Materials | 50010 Purchases Raw Materials |
| COGS adjustment (periodic system) | 50010 Purchases Raw Materials | 12000 Inventory — Raw Materials |

### 8. Error Corrections

Corrections for mistakes discovered during the period.

| Error Type | Treatment |
|------------|-----------|
| Wrong account debited/credited | Reverse the incorrect entry, record the correct one |
| Wrong amount | Record the difference (additional debit/credit) |
| Omitted transaction | Record the entry as if it were made on the original date |
| Duplicate entry | Reverse one of the duplicate entries |

### 9. Work-in-Progress (WIP) Adjustments

Manufacturing businesses need to track costs as raw materials are converted into finished goods. WIP adjustments record the flow of production costs.

| Entry | Debit | Credit | Description |
|-------|-------|--------|-------------|
| Opening WIP | 50300 Opening WIP | 12400 WIP Inventory | Bring forward WIP from prior period |
| Direct Materials | 50320 Direct Materials Used | 50010 Purchases Raw Materials | Materials issued to production |
| Direct Materials | 50320 Direct Materials Used | 50110 Purchases Packaging | Packaging issued to production |
| Direct Labor | 50330 Direct Labor to WIP | 53000 Direct Labor Wages | Labor costs allocated to WIP |
| Overhead Applied | 50340 Overhead Applied | 53100 Machine Maintenance | Overhead allocated to WIP |
| Overhead Applied | 50340 Overhead Applied | 53200 Production Utilities | Overhead allocated to WIP |
| Overhead Applied | 50340 Overhead Applied | 53300 Depreciation - COGS | Overhead allocated to WIP |
| Clear WIP | 50350 WIP to FG | 50300 Opening WIP | Clear opening WIP to transfer |
| Clear WIP | 50350 WIP to FG | 50320 Direct Materials Used | Clear materials to transfer |
| Clear WIP | 50350 WIP to FG | 50330 Direct Labor to WIP | Clear labor to transfer |
| Clear WIP | 50350 WIP to FG | 50340 Overhead Applied | Clear overhead to transfer |
| Closing WIP | 12400 WIP Inventory | 50310 Closing WIP | Record closing WIP asset |
| Transfer to FG | 12200 Finished Goods | 50350 WIP to FG | COGM transferred to FG |

**COGM Calculation:**
```
COGM = Opening WIP + Direct Materials + Direct Labor + Overhead Applied - Closing WIP
```

**Notes:**
- WIP accumulation accounts (50320, 50330, 50340) should clear to zero at period-end
- Account 50350 (WIP Transferred to FG) should equal the COGM amount
- Closing WIP (50310) has a credit balance — it reduces total COGS
- WIP Inventory (12400) appears on Balance Sheet as a Current Asset

For detailed WIP accounting guidance, see `references/wip-flow-guide.md`.

### 10. Bank Reconciliation Adjustments

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

1. **Close revenue accounts**: Debit all revenue accounts, Credit 32000 Retained Earnings
2. **Close expense accounts**: Debit 32000 Retained Earnings, Credit all expense accounts
3. Net Profit/Loss is automatically reflected in Retained Earnings

After closing entries, only permanent accounts (Assets, Liabilities, Equity) have balances.

## Output Structure

The adjusting entries workbook should contain:
- **Summary**: Count and total of adjustments by type
- **Detailed Entries**: Full journal entry format (Entry No, Date, Accounts, Amounts, Description)
- **Depreciation Schedule**: Per-asset depreciation calculation (if depreciation was processed)
- **Impact Analysis**: For each account affected, show balance before adjustment, adjustment amount, and balance after adjustment
