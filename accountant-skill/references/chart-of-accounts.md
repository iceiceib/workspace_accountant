# Chart of Accounts Reference

This document defines the Chart of Accounts structure for Shwe Mandalay Cafe / K&K Finance Team. All account codes, classifications, and normal balances are defined here. Every module must validate account codes against this reference.

## Account Code Structure

Format: `XXXXX` (5-digit codes) where:
- `X` = Account Type (1-7)
- `XXXX` = Sequential number within type

| Range | Type | Normal Balance |
|-------|------|---------------|
| 10000-14999 | Current Assets | Debit |
| 15000-19999 | Non-Current Assets | Debit |
| 20000-24999 | Current Liabilities | Credit |
| 25000-29999 | Non-Current Liabilities | Credit |
| 30000-39999 | Equity | Credit |
| 40000-49999 | Revenue | Credit |
| 50000-69999 | Expenses | Debit |
| 70000-79999 | Other Income / Non-Operating | Credit (Income) / Debit (Expense) |

## Master Chart of Accounts

### 10000 — Current Assets

| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 10000 | Cash in Hand | Current Asset | Debit |
| 10100 | Cash at Bank | Current Asset | Debit |
| 11000 | Accounts Receivable | Current Asset | Debit |
| 12000 | Inventory - Raw Material | Current Asset | Debit |
| 12100 | Inventory - Packaging | Current Asset | Debit |
| 12200 | Inventory - Finished Goods | Current Asset | Debit |
| 12300 | Inventory Adjustments | Current Asset (Contra) | Credit |
| 12400 | Work-in-Progress | Current Asset | Debit |
| 13000 | Advanced Payments | Current Asset | Debit |
| 14000 | Deferred Preliminary Expenses | Current Asset | Debit |

### 15000 — Non-Current Assets (Property, Plant & Equipment)

| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 15000 | Land | Non-Current Asset | Debit |
| 15100 | Buildings & Structures | Non-Current Asset | Debit |
| 15110 | Accumulated Depreciation - Buildings & Structures | Non-Current Asset (Contra) | Credit |
| 15200 | Machinery & Equipment | Non-Current Asset | Debit |
| 15210 | Accumulated Depreciation - Machinery & Equipment | Non-Current Asset (Contra) | Credit |
| 15300 | Office & Facility Equipment | Non-Current Asset | Debit |
| 15310 | Accumulated Depreciation - Office & Facility Equipment | Non-Current Asset (Contra) | Credit |
| 15400 | Electrical & Utility Systems | Non-Current Asset | Debit |
| 15410 | Accumulated Depreciation - Electrical & Utility Systems | Non-Current Asset (Contra) | Credit |
| 15500 | Construction in Progress | Non-Current Asset | Debit |
| 15510 | Accumulated Depreciation - Motor Vehicles | Non-Current Asset (Contra) | Credit |
| 15600 | Motor Vehicles | Non-Current Asset | Debit |

### 20000 — Current Liabilities

| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 20000 | Accounts Payable | Current Liability | Credit |
| 21000 | Short-term Loans | Current Liability | Credit |
| 22000 | Utility Bills | Current Liability | Credit |
| 22100 | Provision for Management Compensation | Current Liability | Credit |
| 22200 | Wages Payable | Current Liability | Credit |

### 25000 — Non-Current Liabilities

| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 25000 | Bank Loan | Non-Current Liability | Credit |

### 30000 — Equity

| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 31000 | Paid-up Capital | Equity | Credit |
| 32000 | Retained Earnings | Equity | Credit |

### 40000 — Revenue

| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 40000 | Sales Revenue | Operating Revenue | Credit |

### 50000 — Cost of Goods Sold

#### Raw Materials
| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 50000 | Opening Inventory - Raw Materials | COGS | Debit |
| 50010 | Purchases Raw Materials | COGS | Debit |
| 50020 | Closing Inventory - Raw Materials | COGS | Debit |

#### Packaging Materials
| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 50100 | Opening Inventory - Packaging | COGS | Debit |
| 50110 | Purchases Packaging | COGS | Debit |
| 50120 | Closing Inventory - Packaging | COGS | Debit |

#### Finished Goods
| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 50200 | Opening Inventory - Finished Goods | COGS | Debit |
| 50220 | Closing Inventory - Finished Goods | COGS | Debit |

#### Work-in-Progress (WIP)
| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 50300 | Opening Work-in-Progress | COGS | Debit |
| 50310 | Closing Work-in-Progress | COGS | Debit |
| 50320 | Direct Materials Used | COGS | Debit |
| 50330 | Direct Labor Transferred to WIP | COGS | Debit |
| 50340 | Manufacturing Overhead Applied | COGS | Debit |
| 50350 | WIP Transferred to Finished Goods | COGS | Debit |

#### Production Overhead Costs
| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 53000 | Direct Labor Wages | COGS | Debit |
| 53100 | Machine Maintenance & Repair Expense | COGS | Debit |
| 53200 | Production Utilities (Electricity & Water) | COGS | Debit |
| 53300 | Depreciation Expenses - COGS | COGS | Debit |

### 60000 — Operating Expenses (SG&A)

| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 60000 | Marketing & Advertising Expense | Operating Expense | Debit |
| 61000 | Office Salaries | Operating Expense | Debit |
| 62000 | Meal Allowance Expense | Operating Expense | Debit |
| 63000 | Utilities (Electricity & Water) | Operating Expense | Debit |
| 64000 | Transportation & Distribution Expense | Operating Expense | Debit |
| 65000 | Factory Buildings & Office Supplies and Maintenance | Operating Expense | Debit |
| 66000 | Depreciation Expenses - SG&A | Operating Expense | Debit |
| 67000 | Inventory Write-off | Operating Expense | Debit |
| 68000 | Other Expenses | Operating Expense | Debit |
| 69000 | Key Management Personnel and Director Compensation | Operating Expense | Debit |

### 70000 — Other Income / Non-Operating

| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 70000 | Interest Income | Other Income | Credit |

## Contra Accounts

Contra accounts have opposite normal balance to their type:

| Code | Account Name | Contra Of |
|------|-------------|-----------|
| 12300 | Inventory Adjustments | Inventory |
| 15110 | Accumulated Depreciation - Buildings | Buildings |
| 15210 | Accumulated Depreciation - Machinery | Machinery |
| 15310 | Accumulated Depreciation - Office Equipment | Office Equipment |
| 15410 | Accumulated Depreciation - Electrical | Electrical Systems |
| 15510 | Accumulated Depreciation - Motor Vehicles | Motor Vehicles |

## Account Classification for Financial Statements

When generating financial statements, map accounts as follows:

### Income Statement Mapping
- **Revenue**: 40000 (Sales Revenue)
- **Other Income**: 70000 (Interest Income)
- **COGS**: 50000-50399, 53000-53999
- **Operating Expenses (SG&A)**: 60000-69999

### Balance Sheet Mapping
- **Current Assets**: 10000-14999
- **Non-Current Assets**: 15000-19999 (net of accumulated depreciation 15110, 15210, 15310, 15410, 15510)
- **Current Liabilities**: 20000-24999
- **Non-Current Liabilities**: 25000-29999
- **Equity**: 30000-39999

## Customization

This COA is the reference for K&K Finance. When processing .xlsx files:
1. First read the user's actual `chart_of_accounts.xlsx` if available
2. Fall back to this reference for any missing classifications
3. Flag any account codes found in journals/ledgers that don't exist in either source
