# Data Schemas — Expected .xlsx Column Structures

This document defines the expected column layout for every input .xlsx file. When reading files, the skill should match columns flexibly (case-insensitive, trimmed whitespace) and map to these canonical names.

## Table of Contents
1. [Chart of Accounts](#chart-of-accounts)
2. [Journals (Books of Prime Entry)](#journals)
3. [Ledgers](#ledgers)
4. [Inventory Sub-Ledgers](#inventory-sub-ledgers)
5. [Bank Statement](#bank-statement)
6. [Fixed Asset Register](#fixed-asset-register)

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

### Account Code Ranges

| Range | Type | Sub-Type | Normal Balance |
|-------|------|----------|----------------|
| 10000-10999 | Asset | Current Asset - Cash | Debit |
| 11000-11999 | Asset | Current Asset - Receivables | Debit |
| 12000-12999 | Asset | Current Asset - Inventory | Debit |
| 12400 | Asset | Current Asset - WIP Inventory | Debit |
| 13000-14999 | Asset | Current Asset - Prepayments | Debit |
| 15000-19999 | Asset | Non-Current Asset - PPE | Debit |
| 20000-24999 | Liability | Current Liability | Credit |
| 25000-29999 | Liability | Non-Current Liability | Credit |
| 30000-39999 | Equity | Equity | Credit |
| 40000-49999 | Revenue | Operating Revenue | Credit |
| 50000-50299 | Expense | COGS - Inventory | Debit |
| 50300-50399 | Expense | COGS - Work-in-Progress | Debit |
| 53000-53999 | Expense | COGS - Production Costs | Debit |
| 60000-69999 | Expense | Operating Expense (SG&A) | Debit |
| 70000-79999 | Revenue/Expense | Other Income/Non-Operating | Mixed |

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
| Debit Account | Text/Number | Yes | Account code to debit (usually 11000 AR) |
| Credit Account | Text/Number | Yes | Account code to credit (usually 40000-4040 Revenue) |
| Amount | Number | Yes | Transaction amount |

### Purchases Journal
**File:** `purchases_journal.xlsx`

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Date | Date | Yes | Transaction date |
| Reference | Text | Yes | Purchase order / supplier invoice reference |
| Supplier | Text | Yes | Supplier name |
| Description | Text | No | Transaction narration |
| Debit Account | Text/Number | Yes | Account code to debit (e.g., 50000 Raw Materials, 12000 Inventory) |
| Credit Account | Text/Number | Yes | Account code to credit (usually 20000 AP) |
| Amount | Number | Yes | Transaction amount |

### Cash Receipts Journal
**File:** `cash_receipts_journal.xlsx`

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Date | Date | Yes | Receipt date |
| Receipt No | Text | Yes | Receipt reference number |
| Received From | Text | Yes | Payer name |
| Description | Text | No | Transaction narration |
| Debit Account | Text/Number | Yes | Account code to debit (usually 10100 Cash at Bank) |
| Credit Account | Text/Number | Yes | Account code to credit (e.g., 11000 AR, 40000 Sales) |
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
| Debit Account | Text/Number | Yes | Account code to debit (e.g., 20000 AP, 5200 Rent) |
| Credit Account | Text/Number | Yes | Account code to credit (usually 10100 Cash at Bank) |
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
| Credit Account | Text/Number | Yes | Account code to credit (e.g., 10100 Bank, 2030 Accrued Wages) |
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

## Inventory Sub-Ledgers

Inventory sub-ledgers track both **quantity** and **value** for each inventory item. They reconcile to GL control accounts 12000 (Raw Materials) and 12100 (Packaging Materials).

### Inventory Items Master
**File:** `inventory_items.xlsx`

Master list of all inventory items with their account codes and unit measures.

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Item Code | Text/Number | Yes | Unique item identifier (e.g., 12001, 12100) |
| Item Name | Text | Yes | Item description |
| Account Code | Text/Number | Yes | GL account code (12000-12099 for raw materials, 12100-12199 for packaging) |
| Account Name | Text | No | GL account name |
| Unit Measure | Text | Yes | Unit of measurement (Bag, Pack, Gram, Bottle, etc.) |
| Category | Text | No | Category for grouping (Raw Materials, Packaging, etc.) |
| Status | Text | No | Active / Inactive |

### Raw Materials Ledger
**File:** `raw_materials_ledger.xlsx`

Sub-ledger for raw materials (account codes 12000-12099). May have multiple sheets (one per item) or a single consolidated sheet.

**Per-Item Sheet Structure:**

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Date | Date | Yes | Transaction date |
| Reference | Text | Yes | PO/Invoice/Production reference |
| Description | Text | No | Narration (supplier, production batch, etc.) |
| Received Qty | Number | No | Units received (increases inventory) |
| Issued Qty | Number | No | Units issued to production (decreases inventory) |
| Balance Qty | Number | Yes | Running quantity balance |
| Unit Cost | Number | Yes | Weighted Average Cost per unit |
| Received Value | Number | No | Value of goods received (Received Qty × Unit Cost) |
| Issued Value | Number | No | Value of goods issued (Issued Qty × WAC) |
| Balance Value | Number | Yes | Running value balance |

**Accounting Treatment:**
- **Received (Dr)**: Increases inventory quantity and value
- **Issued (Cr)**: Decreases inventory (transferred to WIP/COGS)
- **Normal Balance**: Debit (asset)

### Packaging Materials Ledger
**File:** `packaging_ledger.xlsx`

Sub-ledger for packaging materials (account codes 12100-12199). Same structure as Raw Materials Ledger.

**Per-Item Sheet Structure:**

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Date | Date | Yes | Transaction date |
| Reference | Text | Yes | PO/Invoice/Production reference |
| Description | Text | No | Narration |
| Received Qty | Number | No | Units received |
| Issued Qty | Number | No | Units issued to production |
| Balance Qty | Number | Yes | Running quantity balance |
| Unit Cost | Number | Yes | Weighted Average Cost per unit |
| Received Value | Number | No | Value of goods received |
| Issued Value | Number | No | Value of goods issued |
| Balance Value | Number | Yes | Running value balance |

### Work-in-Progress Ledger
**File:** `wip_ledger.xlsx`

Sub-ledger for WIP inventory (account code 12400). Tracks production costs accumulated during manufacturing.

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Date | Date | Yes | Transaction date |
| Reference | Text | Yes | Production batch/job reference |
| Description | Text | No | Narration |
| Direct Materials | Number | No | Raw materials used (from RM Ledger) |
| Direct Labor | Number | No | Labor costs allocated |
| Overhead Applied | Number | No | Manufacturing overhead allocated |
| Completed Goods | Number | No | Cost of goods completed (Cr - transferred to FG) |
| Balance | Number | Yes | Running WIP balance |

### Finished Goods Ledger
**File:** `finished_goods_ledger.xlsx`

Sub-ledger for finished goods inventory (account code 12200).

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Date | Date | Yes | Transaction date |
| Reference | Text | Yes | Production batch/Sales reference |
| Description | Text | No | Narration |
| Produced Qty | Number | No | Units produced (from WIP) |
| Sold Qty | Number | No | Units sold (to COGS) |
| Balance Qty | Number | Yes | Running quantity |
| Unit Cost | Number | Yes | Cost per unit |
| Produced Value | Number | No | Value of goods produced |
| Sold Value | Number | No | Cost of goods sold |
| Balance Value | Number | Yes | Running value balance |

### Inventory Reconciliation

Sub-ledger totals must reconcile to GL control accounts:

| Sub-Ledger | GL Control Account | Reconciliation |
|------------|-------------------|----------------|
| Raw Materials Ledger (sum of all items) | 12000 Inventory - Raw Materials | Total Balance Value = GL Balance |
| Packaging Ledger (sum of all items) | 12100 Inventory - Packaging | Total Balance Value = GL Balance |
| WIP Ledger | 12400 Work-in-Progress | Balance = GL Balance |
| Finished Goods Ledger | 12200 Inventory - Finished Goods | Total Balance Value = GL Balance |

### Weighted Average Cost (WAC) Calculation

```
WAC = (Opening Stock Value + Purchases Value) / (Opening Stock Qty + Purchases Qty)

Cost of Goods Issued = Units Issued × WAC
Closing Stock Value = Closing Stock Qty × WAC
```

**Note:** WAC is recalculated after each purchase. Issues to production use the current WAC.

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
