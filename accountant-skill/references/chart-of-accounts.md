# Chart of Accounts Reference

This document defines the Chart of Accounts structure for Shwe Mandalay Cafe / K&K Finance Team. All account codes, classifications, and normal balances are defined here. Every module must validate account codes against this reference.

## Account Code Structure

Format: `XYYY` where:
- `X` = Account Type (1-5)
- `YYY` = Sequential number within type

| Range | Type | Normal Balance |
|-------|------|---------------|
| 1000-1999 | Assets | Debit |
| 2000-2999 | Liabilities | Credit |
| 3000-3999 | Equity | Credit |
| 4000-4999 | Revenue | Credit |
| 5000-5999 | Expenses | Debit |

## Master Chart of Accounts

### 1000 — Assets

#### Current Assets
| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 1010 | Cash on Hand | Current Asset | Debit |
| 1020 | Cash at Bank — Main Account | Current Asset | Debit |
| 1021 | Cash at Bank — Account 2 | Current Asset | Debit |
| 1022 | Cash at Bank — Account 3 | Current Asset | Debit |
| 1030 | Petty Cash | Current Asset | Debit |
| 1100 | Accounts Receivable | Current Asset | Debit |
| 1110 | Allowance for Doubtful Debts | Current Asset (Contra) | Credit |
| 1200 | Inventory — Raw Materials | Current Asset | Debit |
| 1210 | Inventory — Packaging | Current Asset | Debit |
| 1220 | Inventory — Finished Goods | Current Asset | Debit |
| 1300 | Prepaid Expenses | Current Asset | Debit |
| 1310 | Prepaid Rent | Current Asset | Debit |
| 1320 | Prepaid Insurance | Current Asset | Debit |
| 1400 | Advances to Employees | Current Asset | Debit |
| 1500 | Other Current Assets | Current Asset | Debit |

#### Non-Current Assets
| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 1600 | Land | Non-Current Asset | Debit |
| 1610 | Buildings | Non-Current Asset | Debit |
| 1611 | Accumulated Depreciation — Buildings | Non-Current Asset (Contra) | Credit |
| 1620 | Plant & Machinery | Non-Current Asset | Debit |
| 1621 | Accumulated Depreciation — Plant & Machinery | Non-Current Asset (Contra) | Credit |
| 1630 | Furniture & Fixtures | Non-Current Asset | Debit |
| 1631 | Accumulated Depreciation — Furniture & Fixtures | Non-Current Asset (Contra) | Credit |
| 1640 | Vehicles | Non-Current Asset | Debit |
| 1641 | Accumulated Depreciation — Vehicles | Non-Current Asset (Contra) | Credit |
| 1650 | Office Equipment | Non-Current Asset | Debit |
| 1651 | Accumulated Depreciation — Office Equipment | Non-Current Asset (Contra) | Credit |
| 1660 | Construction in Progress | Non-Current Asset | Debit |
| 1700 | Intangible Assets | Non-Current Asset | Debit |

### 2000 — Liabilities

#### Current Liabilities
| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 2010 | Accounts Payable | Current Liability | Credit |
| 2020 | Accrued Expenses | Current Liability | Credit |
| 2030 | Accrued Wages & Salaries | Current Liability | Credit |
| 2040 | Unearned Revenue | Current Liability | Credit |
| 2050 | Tax Payable | Current Liability | Credit |
| 2060 | Short-term Loans | Current Liability | Credit |
| 2070 | Current Portion of Long-term Debt | Current Liability | Credit |
| 2080 | Other Current Liabilities | Current Liability | Credit |

#### Non-Current Liabilities
| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 2100 | Long-term Loans | Non-Current Liability | Credit |
| 2110 | Mortgage Payable | Non-Current Liability | Credit |
| 2200 | Other Non-Current Liabilities | Non-Current Liability | Credit |

### 3000 — Equity
| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 3010 | Owner's Capital | Equity | Credit |
| 3020 | Owner's Drawings | Equity (Contra) | Debit |
| 3030 | Retained Earnings | Equity | Credit |
| 3040 | Current Year Earnings | Equity | Credit |

### 4000 — Revenue
| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 4010 | Sales Revenue — Shop 1 | Operating Revenue | Credit |
| 4020 | Sales Revenue — Shop 2 | Operating Revenue | Credit |
| 4030 | Sales Revenue — Shop 3 | Operating Revenue | Credit |
| 4040 | Sales Revenue — General | Operating Revenue | Credit |
| 4100 | Other Income | Non-Operating Revenue | Credit |
| 4110 | Interest Income | Non-Operating Revenue | Credit |
| 4120 | Gain on Disposal of Assets | Non-Operating Revenue | Credit |
| 4200 | Sales Returns & Allowances | Revenue (Contra) | Debit |
| 4210 | Sales Discounts | Revenue (Contra) | Debit |

### 5000 — Expenses

#### Cost of Goods Sold
| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 5010 | Raw Materials Used | COGS | Debit |
| 5020 | Packaging Costs | COGS | Debit |
| 5030 | Direct Labour | COGS | Debit |
| 5040 | Manufacturing Overhead | COGS | Debit |
| 5050 | Freight In | COGS | Debit |

#### Operating Expenses
| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 5100 | Salaries & Wages | Operating Expense | Debit |
| 5110 | Employee Benefits | Operating Expense | Debit |
| 5120 | Social Security Contributions | Operating Expense | Debit |
| 5200 | Rent Expense | Operating Expense | Debit |
| 5210 | Utilities Expense | Operating Expense | Debit |
| 5220 | Telephone & Internet | Operating Expense | Debit |
| 5300 | Depreciation Expense | Operating Expense | Debit |
| 5310 | Amortization Expense | Operating Expense | Debit |
| 5400 | Repairs & Maintenance | Operating Expense | Debit |
| 5410 | Office Supplies | Operating Expense | Debit |
| 5420 | Cleaning & Sanitation | Operating Expense | Debit |
| 5500 | Transportation & Delivery | Operating Expense | Debit |
| 5510 | Fuel Expense | Operating Expense | Debit |
| 5600 | Marketing & Advertising | Operating Expense | Debit |
| 5700 | Insurance Expense | Operating Expense | Debit |
| 5800 | Bad Debt Expense | Operating Expense | Debit |
| 5900 | Miscellaneous Expense | Operating Expense | Debit |

#### Non-Operating Expenses
| Code | Account Name | Sub-Type | Normal Balance |
|------|-------------|----------|---------------|
| 5910 | Interest Expense | Non-Operating Expense | Debit |
| 5920 | Bank Charges & Fees | Non-Operating Expense | Debit |
| 5930 | Loss on Disposal of Assets | Non-Operating Expense | Debit |
| 5940 | Foreign Exchange Loss | Non-Operating Expense | Debit |
| 5950 | Tax Expense | Non-Operating Expense | Debit |

## Account Classification for Financial Statements

When generating financial statements, map accounts as follows:

### Income Statement Mapping
- **Revenue**: 4010-4040 (net of 4200, 4210)
- **Other Income**: 4100-4120
- **COGS**: 5010-5050
- **Operating Expenses**: 5100-5900
- **Non-Operating Expenses**: 5910-5950

### Balance Sheet Mapping
- **Current Assets**: 1010-1500
- **Non-Current Assets**: 1600-1700 (net of accumulated depreciation)
- **Current Liabilities**: 2010-2080
- **Non-Current Liabilities**: 2100-2200
- **Equity**: 3010-3040

## Customization

This COA is a starting template. The user may have additional or different accounts. When processing .xlsx files:
1. First read the user's actual `chart_of_accounts.xlsx` if available
2. Fall back to this reference for any missing classifications
3. Flag any account codes found in journals/ledgers that don't exist in either source
