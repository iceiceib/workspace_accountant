# Data Schemas — Expected .xlsx Column Structures

This document defines the expected column layout for every input .xlsx file. When reading files, the skill should match columns flexibly (case-insensitive, trimmed whitespace) and map to these canonical names.

## Table of Contents
1. [Chart of Accounts](#chart-of-accounts)
2. [Journals (Books of Prime Entry)](#journals)
3. [Ledgers](#ledgers)
4. [Bank Statement](#bank-statement)
5. [Fixed Asset Register](#fixed-asset-register)

---

## Chart of Accounts
**File:** `chart_of_accounts.xlsx`

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Account Code | Text/Number | Yes | Unique account identifier (e.g., 1010, 5300) |
| Account Name | Text | Yes | Account description |
| Type | Text | Yes | Asset, Liability, Equity, Revenue, Expense |
| Sub-Type | Text | No | Current Asset, Non-Current Asset, COGS, Operating Expense, etc. |
| Normal Balance | Text | Yes | Debit or Credit |
| Status | Text | No | Active / Inactive |

**Alternate column names to accept:** "Code", "Acct Code", "No." → Account Code; "Name", "Description", "Account" → Account Name

---

## Journals

All journals share a common base structure with minor variations.

### Sales Journal
**File:** `sales_journal.xlsx`

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Date | Date | Yes | Transaction date |
| Invoice No | Text | Yes | Sales invoice reference |
| Customer | Text | Yes | Customer name |
| Description | Text | No | Transaction narration |
| Debit Account | Text/Number | Yes | Account code to debit (usually 1100 AR) |
| Credit Account | Text/Number | Yes | Account code to credit (usually 4010-4040 Revenue) |
| Amount | Number | Yes | Transaction amount |

### Purchases Journal
**File:** `purchases_journal.xlsx`

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Date | Date | Yes | Transaction date |
| Reference | Text | Yes | Purchase order / supplier invoice reference |
| Supplier | Text | Yes | Supplier name |
| Description | Text | No | Transaction narration |
| Debit Account | Text/Number | Yes | Account code to debit (e.g., 5010 Raw Materials, 1200 Inventory) |
| Credit Account | Text/Number | Yes | Account code to credit (usually 2010 AP) |
| Amount | Number | Yes | Transaction amount |

### Cash Receipts Journal
**File:** `cash_receipts_journal.xlsx`

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Date | Date | Yes | Receipt date |
| Receipt No | Text | Yes | Receipt reference number |
| Received From | Text | Yes | Payer name |
| Description | Text | No | Transaction narration |
| Debit Account | Text/Number | Yes | Account code to debit (usually 1020 Cash at Bank) |
| Credit Account | Text/Number | Yes | Account code to credit (e.g., 1100 AR, 4010 Sales) |
| Amount | Number | Yes | Amount received |
| Bank Account | Text | No | Which bank account (if multiple) |

### Cash Payments Journal
**File:** `cash_payments_journal.xlsx`

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Date | Date | Yes | Payment date |
| Payment No | Text | Yes | Cheque number / transfer reference |
| Paid To | Text | Yes | Payee name |
| Description | Text | No | Transaction narration |
| Debit Account | Text/Number | Yes | Account code to debit (e.g., 2010 AP, 5200 Rent) |
| Credit Account | Text/Number | Yes | Account code to credit (usually 1020 Cash at Bank) |
| Amount | Number | Yes | Amount paid |
| Bank Account | Text | No | Which bank account |

### Payroll Journal
**File:** `payroll_journal.xlsx`

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Date | Date | Yes | Pay period date |
| Employee / Department | Text | Yes | Employee name or department |
| Description | Text | No | Narration |
| Debit Account | Text/Number | Yes | Account code to debit (e.g., 5100 Salaries) |
| Credit Account | Text/Number | Yes | Account code to credit (e.g., 1020 Bank, 2030 Accrued Wages) |
| Debit Amount | Number | Yes | Amount debited |
| Credit Amount | Number | Yes | Amount credited |

### General Journal
**File:** `general_journal.xlsx`

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Date | Date | Yes | Entry date |
| JV No | Text | Yes | Journal voucher reference |
| Description | Text | Yes | Detailed narration |
| Debit Account | Text/Number | Yes | Account code to debit |
| Credit Account | Text/Number | Yes | Account code to credit |
| Debit Amount | Number | Yes | Amount debited |
| Credit Amount | Number | Yes | Amount credited |

**Note on compound entries:** A single JV No may span multiple rows. Group by JV No to validate that total debits = total credits for the entry.

---

## Ledgers

All ledgers follow a T-account running balance format.

### General Ledger
**File:** `general_ledger.xlsx`

May have multiple sheets (one per account) or a single sheet with all accounts.

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Account Code | Text/Number | Yes | Account identifier |
| Account Name | Text | No | Account description |
| Date | Date | Yes | Transaction date |
| Reference | Text | Yes | Source journal reference (SJ-001, CPJ-015, GJ-003, etc.) |
| Description | Text | No | Transaction narration |
| Debit | Number | No | Debit amount (blank if credit) |
| Credit | Number | No | Credit amount (blank if debit) |
| Balance | Number | No | Running balance |

### Accounts Receivable Ledger
**File:** `accounts_receivable_ledger.xlsx`

May have multiple sheets (one per customer) or a single sheet.

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Customer | Text | Yes | Customer name or code |
| Date | Date | Yes | Transaction date |
| Invoice No | Text | Yes | Invoice / receipt reference |
| Description | Text | No | Narration |
| Debit | Number | No | Sales / charges (increases AR) |
| Credit | Number | No | Payments / credits (decreases AR) |
| Balance | Number | No | Running balance |

### Accounts Payable Ledger
**File:** `accounts_payable_ledger.xlsx`

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Supplier | Text | Yes | Supplier name or code |
| Date | Date | Yes | Transaction date |
| Reference | Text | Yes | Invoice / payment reference |
| Description | Text | No | Narration |
| Debit | Number | No | Payments (decreases AP) |
| Credit | Number | No | Purchases / charges (increases AP) |
| Balance | Number | No | Running balance |

### Cash Ledger
**File:** `cash_ledger.xlsx`

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Bank Account | Text | No | Bank account name/number (if multiple) |
| Date | Date | Yes | Transaction date |
| Reference | Text | Yes | Cheque no / transfer ref |
| Description | Text | No | Narration |
| Debit | Number | No | Money in (receipts) |
| Credit | Number | No | Money out (payments) |
| Balance | Number | No | Running balance |

### Fixed Assets Ledger
**File:** `fixed_assets_ledger.xlsx`

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Asset ID | Text | Yes | Unique asset identifier |
| Description | Text | Yes | Asset description |
| Account Code | Text/Number | Yes | GL account code (1610-1660) |
| Date Acquired | Date | Yes | Purchase/acquisition date |
| Cost | Number | Yes | Original cost |
| Useful Life (Years) | Number | Yes | Estimated useful life |
| Salvage Value | Number | No | Residual value at end of life (default 0) |
| Depreciation Method | Text | Yes | "Straight-Line" or "Reducing Balance" |
| Accumulated Depreciation | Number | Yes | Total depreciation to date |
| Net Book Value | Number | Yes | Cost - Accumulated Depreciation |

### Equity Ledger
**File:** `equity_ledger.xlsx`

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Account Code | Text/Number | Yes | Equity account code (3010-3040) |
| Date | Date | Yes | Transaction date |
| Description | Text | Yes | Narration |
| Debit | Number | No | Drawings, losses |
| Credit | Number | No | Capital introduced, earnings |
| Balance | Number | No | Running balance |

---

## Bank Statement
**File:** `bank_statement.xlsx`

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Date | Date | Yes | Transaction date |
| Description | Text | Yes | Bank narration |
| Reference | Text | No | Cheque no / transfer ref |
| Debit | Number | No | Withdrawals (money out of bank) |
| Credit | Number | No | Deposits (money into bank) |
| Balance | Number | No | Running balance |

**Note:** Bank statement uses opposite conventions to cash book. A "credit" on bank statement = money deposited = a "debit" in the cash book.

---

## Fixed Asset Register
**File:** `fixed_asset_register.xlsx`

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Asset ID | Text | Yes | Unique identifier |
| Description | Text | Yes | Asset description |
| Category | Text | Yes | Buildings, Plant & Machinery, Furniture, Vehicles, Equipment |
| Location | Text | No | Physical location (shop name, factory) |
| Date Acquired | Date | Yes | Purchase date |
| Cost | Number | Yes | Original purchase cost |
| Salvage Value | Number | No | Estimated residual value |
| Useful Life (Years) | Number | Yes | Estimated useful life |
| Depreciation Method | Text | Yes | Straight-Line or Reducing Balance |
| Annual Depreciation | Number | Yes | Yearly depreciation amount |
| Monthly Depreciation | Number | Yes | Monthly depreciation amount |
| Accumulated Depreciation | Number | Yes | Total depreciation charged to date |
| Net Book Value | Number | Yes | Cost minus Accumulated Depreciation |
| Status | Text | No | Active, Disposed, Fully Depreciated |

---

## Column Matching Strategy

When reading any .xlsx file, use this approach:
1. Read the header row (usually row 1)
2. Strip whitespace, convert to lowercase
3. Match against the canonical column names above (also lowercase)
4. Accept common aliases (defined per schema above)
5. If a required column is missing, report an error with the expected column name
6. If extra columns exist, ignore them
7. If the user's file has a different structure, ask for a column mapping
