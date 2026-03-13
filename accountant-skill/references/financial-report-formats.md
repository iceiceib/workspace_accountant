# Financial Report Formats

## Overview

This reference defines the structure, layout, and formulas for the three core financial statements plus supporting schedules. All reports are generated from the Adjusted Trial Balance.

## 1. Income Statement (Profit & Loss)

### Structure

```
SHWE MANDALAY CAFE
INCOME STATEMENT
For the period ended [DATE]

                                        Current Period    (Prior Period)
REVENUE
  Sales Revenue (40000)                      XXX              XXX
                                          ─────            ─────
  NET REVENUE                               XXX              XXX

COST OF GOODS SOLD
  Opening Inventory - Raw Materials (50000)  XXX              XXX
  Purchases Raw Materials (50010)            XXX              XXX
  Closing Inventory - Raw Materials (50020) (XXX)            (XXX)
  Opening Inventory - Packaging (50100)      XXX              XXX
  Purchases Packaging (50110)                XXX              XXX
  Closing Inventory - Packaging (50120)     (XXX)            (XXX)
  Opening Inventory - Finished Goods (50200) XXX              XXX
  Closing Inventory - Finished Goods (50220)(XXX)            (XXX)
  Opening WIP (50300)                        XXX              XXX
  Direct Materials Used (50320)              XXX              XXX
  Direct Labor to WIP (50330)                XXX              XXX
  Overhead Applied (50340)                   XXX              XXX
  WIP Transferred to FG (50350)             (XXX)            (XXX)
  Closing WIP (50310)                       (XXX)            (XXX)
  Direct Labor Wages (53000)                 XXX              XXX
  Machine Maintenance (53100)                XXX              XXX
  Production Utilities (53200)               XXX              XXX
  Depreciation - COGS (53300)                XXX              XXX
                                          ─────            ─────
  TOTAL COGS                              (XXX)            (XXX)
                                          ─────            ─────
GROSS PROFIT                                XXX              XXX
  Gross Profit Margin                      XX.X%            XX.X%

OPERATING EXPENSES
  Marketing & Advertising (60000)            XXX              XXX
  Office Salaries (61000)                    XXX              XXX
  Meal Allowance (62000)                     XXX              XXX
  Utilities - SG&A (63000)                   XXX              XXX
  Transportation & Distribution (64000)      XXX              XXX
  Factory & Office Supplies (65000)          XXX              XXX
  Depreciation - SG&A (66000)                XXX              XXX
  Inventory Write-off (67000)                XXX              XXX
  Other Expenses (68000)                     XXX              XXX
  Key Management Compensation (69000)        XXX              XXX
                                          ─────            ─────
  TOTAL OPERATING EXPENSES                (XXX)            (XXX)
                                          ─────            ─────
OPERATING PROFIT                            XXX              XXX
  Operating Profit Margin                  XX.X%            XX.X%

OTHER INCOME / (EXPENSES)
  Interest Income (70000)                    XXX              XXX
                                          ─────            ─────
  NET OTHER INCOME/(EXPENSES)              ±XXX             ±XXX
                                          ─────            ─────
PROFIT BEFORE TAX                           XXX              XXX

                                          ═════            ═════
NET PROFIT / (LOSS)                         XXX              XXX
  Net Profit Margin                        XX.X%            XX.X%
```

### Account Mapping
- Revenue accounts: 40000-49999 (credit balances = positive revenue)
- COGS accounts: 50000-53999 (debit balances = shown as negative)
- Operating expense accounts: 60000-69999
- Non-operating: 70000-79999 (income), 60000-69999 (overlapping expenses)

### Formulas
- Net Revenue = Sales Revenue (40000)
- Gross Profit = Net Revenue - Total COGS
- Operating Profit = Gross Profit - Total Operating Expenses
- Net Profit = Operating Profit + Net Other Income/Expenses

---

## 2. Balance Sheet (Statement of Financial Position)

### Structure

```
SHWE MANDALAY CAFE
BALANCE SHEET
As at [DATE]

                                        Current Period    (Prior Period)
NON-CURRENT ASSETS
  Land (15000)                               XXX              XXX
  Buildings & Structures (15100)               XXX              XXX
    Less: Accum. Depreciation (15110)       (XXX)            (XXX)
  Machinery & Equipment (15200)              XXX              XXX
    Less: Accum. Depreciation (15210)       (XXX)            (XXX)
  Office & Facility Equipment (15300)        XXX              XXX
    Less: Accum. Depreciation (15310)       (XXX)            (XXX)
  Electrical & Utility Systems (15400)       XXX              XXX
    Less: Accum. Depreciation (15410)       (XXX)            (XXX)
  Motor Vehicles (15600)                     XXX              XXX
    Less: Accum. Depreciation (15510)       (XXX)            (XXX)
  Construction in Progress (15500)           XXX              XXX
                                          ─────            ─────
  TOTAL NON-CURRENT ASSETS                  XXX              XXX

CURRENT ASSETS
  Inventory — Raw Materials (12000)          XXX              XXX
  Inventory — Packaging (12100)              XXX              XXX
  Inventory — Finished Goods (12200)         XXX              XXX
  Work-in-Progress (12400)                   XXX              XXX
  Accounts Receivable (11000)                XXX              XXX
    Less: Inventory Adjustments (12300)     (XXX)            (XXX)
  Advanced Payments (13000)                  XXX              XXX
  Deferred Preliminary Expenses (14000)      XXX              XXX
  Cash in Hand (10000)                       XXX              XXX
  Cash at Bank (10100)                       XXX              XXX
                                          ─────            ─────
  TOTAL CURRENT ASSETS                      XXX              XXX
                                          ─────            ─────
TOTAL ASSETS                                XXX              XXX
                                          ═════            ═════

EQUITY
  Paid-up Capital (31000)                    XXX              XXX
  Retained Earnings (32000)                  XXX              XXX
  Current Period Net Profit/(Loss)           XXX              XXX
                                          ─────            ─────
  TOTAL EQUITY                              XXX              XXX

NON-CURRENT LIABILITIES
  Bank Loan (25000)                         XXX              XXX
                                          ─────            ─────
  TOTAL NON-CURRENT LIABILITIES             XXX              XXX

CURRENT LIABILITIES
  Accounts Payable (20000)                   XXX              XXX
  Short-term Loans (21000)                   XXX              XXX
  Utility Bills (22000)                      XXX              XXX
  Provision for Management Compensation (22100) XXX           XXX
  Wages Payable (22200)                      XXX              XXX
                                          ─────            ─────
  TOTAL CURRENT LIABILITIES                 XXX              XXX
                                          ─────            ─────
TOTAL LIABILITIES                           XXX              XXX
                                          ─────            ─────
TOTAL EQUITY & LIABILITIES                  XXX              XXX
                                          ═════            ═════

CHECK: Assets - (Equity + Liabilities) =      0
```

### Key Validation
**Total Assets MUST equal Total Equity + Total Liabilities.** If they don't, there's an error. Report the imbalance.

---

## 3. Cash Flow Statement (Indirect Method)

### Structure

```
SHWE MANDALAY CAFE
CASH FLOW STATEMENT
For the period ended [DATE]

OPERATING ACTIVITIES
  Net Profit/(Loss)                         XXX

  Adjustments for non-cash items:
    Depreciation (66000)                   +XXX

  Changes in working capital:
    (Increase)/Decrease in AR              ±XXX
    (Increase)/Decrease in Inventory       ±XXX
    (Increase)/Decrease in Prepaid Exp.    ±XXX
    Increase/(Decrease) in AP              ±XXX
    Increase/(Decrease) in Accrued Exp.    ±XXX
    Increase/(Decrease) in Unearned Rev.   ±XXX
                                          ─────
  NET CASH FROM OPERATING ACTIVITIES        XXX

INVESTING ACTIVITIES
  Purchase of Fixed Assets                 (XXX)
  Proceeds from Disposal of Assets         +XXX
  Construction in Progress expenditure     (XXX)
                                          ─────
  NET CASH FROM INVESTING ACTIVITIES       (XXX)

FINANCING ACTIVITIES
  Owner's Capital Introduced               +XXX
  Owner's Drawings                         (XXX)
  Loan Proceeds Received                   +XXX
  Loan Repayments                          (XXX)
                                          ─────
  NET CASH FROM FINANCING ACTIVITIES       ±XXX
                                          ─────
NET INCREASE/(DECREASE) IN CASH             XXX

Cash & Cash Equivalents at Start            XXX
                                          ═════
Cash & Cash Equivalents at End              XXX
```

### Working Capital Changes Calculation
For each current asset/liability account:
- Change = Closing Balance - Opening Balance
- For assets: Increase is negative (cash outflow), Decrease is positive (cash inflow)
- For liabilities: Increase is positive (cash inflow), Decrease is negative (cash outflow)

### Validation
Net Increase/Decrease in Cash + Opening Cash Balance = Closing Cash Balance (from Balance Sheet)

---

## 4. Supporting Schedules

### AR Aging Report
| Customer | Current | 1-30 Days | 31-60 Days | 61-90 Days | 90+ Days | Total |
|----------|---------|-----------|------------|------------|----------|-------|

### AP Aging Report
| Supplier | Current | 1-30 Days | 31-60 Days | 61-90 Days | 90+ Days | Total |
|----------|---------|-----------|------------|------------|----------|-------|

### Fixed Asset Schedule
| Category | Cost | Accum Depr (Open) | Depr This Period | Accum Depr (Close) | NBV |
|----------|------|-------------------|------------------|--------------------|-----|

### Cost Breakdown
| Category | Amount | % of Revenue |
|----------|--------|-------------|

---

## Excel Formatting for Financial Statements

- **Company name**: Bold, 14pt, centered, merged across columns
- **Report title**: Bold, 12pt, centered
- **Period**: Italic, 11pt, centered
- **Section headers** (REVENUE, COGS, etc.): Bold, dark background, white text
- **Line items**: Regular weight, indented with 2 spaces
- **Subtotals**: Bold, single underline above
- **Grand totals**: Bold, double underline above
- **Negative amounts**: Red font, parentheses (#,##0);(#,##0)
- **Percentages**: 0.0% format
- **Comparative column**: Right-aligned, same format, gray text if prior period
