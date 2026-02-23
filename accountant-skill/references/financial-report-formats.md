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
  Sales Revenue — Shop 1  (4010)            XXX              XXX
  Sales Revenue — Shop 2  (4020)            XXX              XXX
  Sales Revenue — Shop 3  (4030)            XXX              XXX
  Sales Revenue — General (4040)            XXX              XXX
  Less: Sales Returns (4200)               (XXX)            (XXX)
  Less: Sales Discounts (4210)             (XXX)            (XXX)
                                          ─────            ─────
  NET REVENUE                               XXX              XXX

COST OF GOODS SOLD
  Raw Materials Used (5010)                 XXX              XXX
  Packaging Costs (5020)                    XXX              XXX
  Direct Labour (5030)                      XXX              XXX
  Manufacturing Overhead (5040)             XXX              XXX
  Freight In (5050)                         XXX              XXX
                                          ─────            ─────
  TOTAL COGS                              (XXX)            (XXX)
                                          ─────            ─────
GROSS PROFIT                                XXX              XXX
  Gross Profit Margin                      XX.X%            XX.X%

OPERATING EXPENSES
  Salaries & Wages (5100)                   XXX              XXX
  Employee Benefits (5110)                  XXX              XXX
  Rent Expense (5200)                       XXX              XXX
  Utilities Expense (5210)                  XXX              XXX
  Depreciation Expense (5300)               XXX              XXX
  Repairs & Maintenance (5400)              XXX              XXX
  Transportation (5500)                     XXX              XXX
  Marketing & Advertising (5600)            XXX              XXX
  Insurance Expense (5700)                  XXX              XXX
  Bad Debt Expense (5800)                   XXX              XXX
  Miscellaneous Expense (5900)              XXX              XXX
  [Other operating expenses]                XXX              XXX
                                          ─────            ─────
  TOTAL OPERATING EXPENSES                (XXX)            (XXX)
                                          ─────            ─────
OPERATING PROFIT                            XXX              XXX
  Operating Profit Margin                  XX.X%            XX.X%

OTHER INCOME / (EXPENSES)
  Other Income (4100)                       XXX              XXX
  Interest Income (4110)                    XXX              XXX
  Interest Expense (5910)                  (XXX)            (XXX)
  Bank Charges (5920)                      (XXX)            (XXX)
  Other Non-Operating (5930-5940)          (XXX)            (XXX)
                                          ─────            ─────
  NET OTHER INCOME/(EXPENSES)              ±XXX             ±XXX
                                          ─────            ─────
PROFIT BEFORE TAX                           XXX              XXX

  Tax Expense (5950)                      (XXX)            (XXX)
                                          ═════            ═════
NET PROFIT / (LOSS)                         XXX              XXX
  Net Profit Margin                        XX.X%            XX.X%
```

### Account Mapping
- Revenue accounts: 4000-4999 (credit balances = positive revenue)
- COGS accounts: 5010-5050 (debit balances = shown as negative)
- Operating expense accounts: 5100-5900
- Non-operating: 4100-4120 (income), 5910-5950 (expenses)

### Formulas
- Net Revenue = Sum(4010:4040) - Sum(4200:4210)
- Gross Profit = Net Revenue - Total COGS
- Operating Profit = Gross Profit - Total Operating Expenses
- Net Profit = Operating Profit + Net Other Income/Expenses - Tax

---

## 2. Balance Sheet (Statement of Financial Position)

### Structure

```
SHWE MANDALAY CAFE
BALANCE SHEET
As at [DATE]

                                        Current Period    (Prior Period)
NON-CURRENT ASSETS
  Land (1600)                               XXX              XXX
  Buildings (1610)                          XXX              XXX
    Less: Accum. Depreciation (1611)       (XXX)            (XXX)
  Plant & Machinery (1620)                  XXX              XXX
    Less: Accum. Depreciation (1621)       (XXX)            (XXX)
  Furniture & Fixtures (1630)               XXX              XXX
    Less: Accum. Depreciation (1631)       (XXX)            (XXX)
  Vehicles (1640)                           XXX              XXX
    Less: Accum. Depreciation (1641)       (XXX)            (XXX)
  Office Equipment (1650)                   XXX              XXX
    Less: Accum. Depreciation (1651)       (XXX)            (XXX)
  Construction in Progress (1660)           XXX              XXX
                                          ─────            ─────
  TOTAL NON-CURRENT ASSETS                  XXX              XXX

CURRENT ASSETS
  Inventory — Raw Materials (1200)          XXX              XXX
  Inventory — Packaging (1210)              XXX              XXX
  Inventory — Finished Goods (1220)         XXX              XXX
  Accounts Receivable (1100)                XXX              XXX
    Less: Allowance for Doubtful Debts(1110)(XXX)           (XXX)
  Prepaid Expenses (1300-1320)              XXX              XXX
  Advances to Employees (1400)              XXX              XXX
  Cash on Hand (1010)                       XXX              XXX
  Cash at Bank (1020-1022)                  XXX              XXX
  Petty Cash (1030)                         XXX              XXX
                                          ─────            ─────
  TOTAL CURRENT ASSETS                      XXX              XXX
                                          ─────            ─────
TOTAL ASSETS                                XXX              XXX
                                          ═════            ═════

EQUITY
  Owner's Capital (3010)                    XXX              XXX
  Less: Owner's Drawings (3020)            (XXX)            (XXX)
  Retained Earnings (3030)                  XXX              XXX
  Current Period Net Profit/(Loss) (3040)   XXX              XXX
                                          ─────            ─────
  TOTAL EQUITY                              XXX              XXX

NON-CURRENT LIABILITIES
  Long-term Loans (2100)                    XXX              XXX
  Mortgage Payable (2110)                   XXX              XXX
                                          ─────            ─────
  TOTAL NON-CURRENT LIABILITIES             XXX              XXX

CURRENT LIABILITIES
  Accounts Payable (2010)                   XXX              XXX
  Accrued Expenses (2020)                   XXX              XXX
  Accrued Wages (2030)                      XXX              XXX
  Unearned Revenue (2040)                   XXX              XXX
  Tax Payable (2050)                        XXX              XXX
  Short-term Loans (2060)                   XXX              XXX
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
    Depreciation (5300)                    +XXX
    Bad Debt Expense (5800)                +XXX
    Loss on Disposal of Assets (5930)      +XXX
    Gain on Disposal of Assets (4120)      -XXX

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
