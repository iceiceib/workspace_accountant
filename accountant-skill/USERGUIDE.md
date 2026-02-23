# Accountant Skill — User Guide
### K&K Finance Team
*Last updated: 2026-02-23 | Covers Modules 1–4*

---

## Table of Contents

1. [What This System Does](#1-what-this-system-does)
2. [How the Accounting Cycle Works](#2-how-the-accounting-cycle-works)
3. [Setup & Installation](#3-setup--installation)
4. [Folder Structure](#4-folder-structure)
5. [Module 1 — Summarize Journals](#5-module-1--summarize-journals)
6. [Module 2 — Summarize Ledgers](#6-module-2--summarize-ledgers)
7. [Module 3 — Bank Reconciliation](#7-module-3--bank-reconciliation)
8. [Module 4 — Journal Adjustments](#8-module-4--journal-adjustments)
9. [Input File Formats](#9-input-file-formats)
10. [Output Formatting Standards](#10-output-formatting-standards)
11. [Troubleshooting & Common Errors](#11-troubleshooting--common-errors)

---

## 1. What This System Does

This is a **full-cycle accounting automation system** for K&K Business. It reads your bookkeeping data from Excel files (`.xlsx`), processes it through a structured accounting workflow, and produces professional formatted reports — all without manual Excel formulas or copy-pasting.

**Everything is Excel in, Excel out.** No databases, no cloud services, no Google Sheets. Just Python processing your `.xlsx` files.

### What it replaces

| Previously manual task | Now automated by |
|---|---|
| Totalling and cross-checking 6 journals | Module 1 |
| Summarising all ledger balances | Module 2 |
| Comparing cash book against bank statement | Module 3 |
| Calculating depreciation, posting bank charges | Module 4 |
| Preparing trial balance (coming) | Module 5 |
| Producing P&L and Balance Sheet (coming) | Module 6 |
| Full-cycle integrity audit (coming) | Module 7 |

---

## 2. How the Accounting Cycle Works

The modules must be run **in order** because each one feeds the next.

```
┌─────────────────────────────────────────────────────────────────┐
│                    FULL ACCOUNTING CYCLE                        │
│                                                                 │
│  Source Documents (invoices, receipts, payroll slips)           │
│         │                                                       │
│         ▼                                                       │
│  MODULE 1 ── Summarize Journals                                 │
│  Reads 6 journals, validates double-entry, produces             │
│  summary + Profit Center / Cost Center breakdown                │
│         │                                                       │
│         ▼                                                       │
│  MODULE 2 ── Summarize Ledgers                                  │
│  Reads all ledgers, calculates balances, cross-checks           │
│  AR/AP/Cash control accounts against subsidiary ledgers         │
│         │                                                       │
│         ▼                                                       │
│  MODULE 3 ── Bank Reconciliation                                │
│  Compares cash book vs bank statement, identifies timing        │
│  differences, generates adjusting entries needed                │
│         │                                                       │
│         ▼                                                       │
│  MODULE 4 ── Journal Adjustments                                │
│  Posts depreciation, imports bank recon entries, references     │
│  accruals and prepayments already in the journals               │
│         │                                                       │
│         ▼                                                       │
│  MODULE 5 ── Trial Balance  (coming)                            │
│  Unadjusted → apply adjustments → Adjusted Trial Balance        │
│         │                                                       │
│         ▼                                                       │
│  MODULE 6 ── Financial Statements  (coming)                     │
│  Income Statement, Balance Sheet, Cash Flow Statement           │
│         │                                                       │
│         ▼                                                       │
│  MODULE 7 ── Validation  (coming)                               │
│  Full integrity check across all modules                        │
└─────────────────────────────────────────────────────────────────┘
```

> **Rule:** Always complete the modules in sequence for a given period. Do not skip ahead.

---

## 3. Setup & Installation

### Requirements

- Python 3.12 or later
- The following Python packages:

```bash
pip install openpyxl pandas numpy
```

### First-time setup (test data for January 2026)

```bash
cd accountant-skill

# Generate all journal test data
python scripts/create_test_data.py data/Jan2026

# Generate ledger test data
python scripts/create_ledger_test_data.py data/Jan2026

# Generate bank statement test data
python scripts/create_bank_statement.py data/Jan2026
```

You only need to run these once. After that, the actual data for each month comes from your real `.xlsx` files placed in the appropriate folder.

---

## 4. Folder Structure

```
accountant-skill/
│
├── SKILL.md                    ← Skill configuration
├── USERGUIDE.md                ← This file
│
├── data/
│   └── Jan2026/                ← One folder per accounting period
│       │
│       │  INPUT FILES (you provide or generate these):
│       ├── sales_journal.xlsx
│       ├── purchases_journal.xlsx
│       ├── cash_receipts_journal.xlsx
│       ├── cash_payments_journal.xlsx
│       ├── payroll_journal.xlsx
│       ├── general_journal.xlsx
│       ├── chart_of_accounts.xlsx
│       ├── profit_cost_centers.xlsx
│       ├── general_ledger.xlsx
│       ├── accounts_receivable_ledger.xlsx
│       ├── accounts_payable_ledger.xlsx
│       ├── cash_ledger.xlsx
│       ├── fixed_assets_ledger.xlsx
│       ├── bank_statement.xlsx
│       │
│       │  OUTPUT FILES (generated by the modules):
│       ├── books_of_prime_entry_Jan2026.xlsx    ← Module 1
│       ├── ledger_summary_Jan2026.xlsx          ← Module 2
│       ├── bank_reconciliation_Jan2026.xlsx     ← Module 3
│       └── adjusting_entries_Jan2026.xlsx       ← Module 4
│
├── references/                 ← Accounting rules and schemas
│   ├── journal-rules.md
│   ├── ledger-posting-rules.md
│   ├── reconciliation-rules.md
│   ├── adjustment-rules.md
│   ├── chart-of-accounts.md
│   ├── financial-report-formats.md
│   └── data-schemas.md
│
└── scripts/                    ← Python scripts
    ├── summarize_journals.py
    ├── summarize_ledgers.py
    ├── reconcile_bank.py
    ├── journal_adjustments.py
    └── utils/
        ├── coa_mapper.py
        ├── pc_cc_mapper.py
        ├── excel_reader.py
        ├── excel_writer.py
        └── double_entry.py
```

### Naming convention for data folders

Always use `MonYYYY` format for period folders:

```
data/Jan2026/
data/Feb2026/
data/Mar2026/
```

---

## 5. Module 1 — Summarize Journals

### What it does

Reads all six Books of Prime Entry for the period, validates that every entry is double-entry balanced (Debits = Credits), then produces a consolidated summary workbook with per-journal detail, a profit centre breakdown, and a cost centre breakdown.

### When to run it

At the start of each period close, after all journal entries for the month have been recorded in the six source journals.

### Input files required

All must be in your `data/[PERIOD]/` folder:

| File | What it contains |
|---|---|
| `sales_journal.xlsx` | Credit sales to customers |
| `purchases_journal.xlsx` | Credit purchases from suppliers |
| `cash_receipts_journal.xlsx` | All cash/bank receipts |
| `cash_payments_journal.xlsx` | All cash/bank payments |
| `payroll_journal.xlsx` | Salary and wage entries |
| `general_journal.xlsx` | Adjustments, corrections, non-cash entries |
| `chart_of_accounts.xlsx` | Master account list (required for validation) |
| `profit_cost_centers.xlsx` | PC and CC definitions (required for PC/CC sheets) |

### How to run

```bash
cd accountant-skill

python scripts/summarize_journals.py \
  data/Jan2026 \
  2026-01-01 \
  2026-01-31 \
  data/Jan2026/books_of_prime_entry_Jan2026.xlsx \
  data/Jan2026/chart_of_accounts.xlsx \
  data/Jan2026/profit_cost_centers.xlsx
```

**Arguments in order:**
1. Data directory (`data/Jan2026`)
2. Period start date (`2026-01-01`)
3. Period end date (`2026-01-31`)
4. Output file path
5. Chart of accounts file
6. Profit/Cost centres file

### Output file — sheets explained

| Sheet | Tab Colour | What to look at |
|---|---|---|
| **Dashboard** | Green | Grand totals for all 6 journals. Confirm Total Debits = Total Credits. |
| **Sales** | Blue | Every sales journal entry with account breakdown. |
| **Purchases** | Blue | Every purchases journal entry. |
| **Cash Receipts** | Blue | All cash and bank receipts. |
| **Cash Payments** | Blue | All cash and bank payments. |
| **Payroll** | Blue | All salary and wage entries. |
| **General** | Blue | Adjustments, depreciation (manual), corrections. |
| **PC Summary** | Orange | Profit Centre P&L: Revenue – COGS – Operating Expenses = Net Profit per segment. |
| **CC Summary** | Orange | Cost Centre spending breakdown by department. |
| **Exceptions** | Red | Only appears if errors were found. Review all items here first. |

### What to check in the output

1. **Dashboard — Balanced:** Must show `True`. If `False`, stop. Do not proceed to Module 2 until all journals balance.
2. **Dashboard — Journals processed:** Must show `6/6`. If any journal is missing, it will show an error.
3. **Exceptions sheet:** If this tab is present (red), open it first and fix all issues before continuing.
4. **PC Summary:** Revenue and costs should be logically distributed across PC01 (Soft Drink), PC02 (Drinking Water), and PC99 (Shared/Corporate).

### Key validation rules enforced

- Every journal entry must have Debit Amount = Credit Amount
- Every account code must exist in `chart_of_accounts.xlsx`
- Revenue accounts (4000–4999) must have a Profit Centre code
- Expense accounts (5000–5999) must have both a Profit Centre and a Cost Centre code
- Dates must fall within the specified period

---

## 6. Module 2 — Summarize Ledgers

### What it does

Reads all ledger files, calculates opening and closing balances for every account, and cross-checks the control accounts (AR, AP, Cash) against their subsidiary ledgers. Flags any mismatches.

### When to run it

After Module 1 is complete and balanced.

### Input files required

| File | What it contains |
|---|---|
| `general_ledger.xlsx` | All account movements (debits and credits) |
| `accounts_receivable_ledger.xlsx` | Individual customer balances |
| `accounts_payable_ledger.xlsx` | Individual supplier balances |
| `cash_ledger.xlsx` | Bank account movements with running balance |
| `fixed_assets_ledger.xlsx` | Fixed asset register |
| `chart_of_accounts.xlsx` | For account classification |

### How to run

```bash
python scripts/summarize_ledgers.py \
  data/Jan2026 \
  2026-01-01 \
  2026-01-31 \
  data/Jan2026/ledger_summary_Jan2026.xlsx \
  data/Jan2026/chart_of_accounts.xlsx
```

**Arguments in order:**
1. Data directory
2. Period start date
3. Period end date
4. Output file path
5. Chart of accounts file (optional but recommended)

### Output file — sheets explained

| Sheet | Tab Colour | What to look at |
|---|---|---|
| **Dashboard** | Green | Ledger status (OK / ERROR for each ledger). Control account reconciliation summary. |
| **GL Balances** | Blue | Opening balance, total debits, total credits, closing balance for every account. Accounts with >50% balance movement are flagged for review. |
| **AR by Customer** | Blue | Accounts receivable balance per customer. |
| **AP by Supplier** | Blue | Accounts payable balance per supplier. |
| **Cash by Bank** | Blue | Cash ledger balance per bank account. |
| **Fixed Assets** | Blue | Asset register: cost, accumulated depreciation, net book value. |
| **Control Acct Recon** | Red/Blue | The critical check: GL control account balance vs subsidiary ledger total. Must all show MATCH. |
| **Exceptions** | Red | Any issues found during processing. |

### What to check in the output

1. **Control Account Reconciliation — all MATCH:**
   - `AR (1100): GL balance = Sum of customer balances`
   - `AP (2010): GL balance = Sum of supplier balances`
   - `Cash (1020): GL balance = Cash ledger closing balance`
   - If any show MISMATCH, investigate before proceeding.

2. **GL Balances — REVIEW flags:**
   - Accounts flagged as REVIEW have moved by more than 50% from their opening balance.
   - These are not errors — they are prompts to verify the movements are expected.

3. **Fixed Assets — NBV positive:**
   - Net Book Value (Cost − Accumulated Depreciation) should be positive for all active assets.

### Understanding balance directions

- **Debit-normal accounts** (Assets, Expenses): a positive balance means the account has a debit balance — normal.
- **Credit-normal accounts** (Liabilities, Equity, Revenue): a positive balance means the account has a credit balance — normal.
- If a debit-normal account shows a negative closing balance, investigate — it may indicate a data entry error.

---

## 7. Module 3 — Bank Reconciliation

### What it does

Compares the company's internal cash records (`cash_ledger.xlsx`) against the bank's statement (`bank_statement.xlsx`) for the same period. Identifies:
- Transactions that match on both sides
- **Deposits in Transit** — recorded in the cash book but not yet showing on the bank statement
- **Outstanding Cheques** — payments recorded in the cash book but not yet cleared by the bank
- **Bank-only items** — bank charges or credits that appear on the statement but are not yet in the cash book

At the end, both the adjusted bank balance and the adjusted cash book balance must agree. The module also auto-generates the journal entries needed to bring the cash book up to date.

### When to run it

After Module 2 is complete.

### Input files required

| File | What it contains |
|---|---|
| `cash_ledger.xlsx` | Internal cash book (company records) |
| `bank_statement.xlsx` | Bank's statement (external records) |

### How to run

```bash
python scripts/reconcile_bank.py \
  data/Jan2026 \
  2026-01-01 \
  2026-01-31 \
  data/Jan2026/bank_reconciliation_Jan2026.xlsx
```

**Arguments in order:**
1. Data directory
2. Period start date
3. Period end date
4. Output file path

### Output file — sheets explained

| Sheet | Tab Colour | What to look at |
|---|---|---|
| **Dashboard** | Green | Status (RECONCILED / NOT RECONCILED), adjusted balances, item counts. |
| **Reconciliation** | Blue | The formal reconciliation statement in standard format. Both sides must arrive at the same adjusted balance. |
| **Matched Items** | Blue | All transactions found on both sides. Match type shown: Exact (same reference + amount) or Probable (same amount, dates within 3 days). |
| **Outstanding Cheques** | Blue | Payments recorded in the cash book but not yet cleared by the bank. These are timing differences — no action needed. |
| **Deposits in Transit** | Blue | Receipts recorded in the cash book but not yet credited by the bank. Timing differences — no action needed. |
| **Bank-Only Items** | Orange | Items on the bank statement that are not in the cash book. These **require action** — journal entries must be posted. |
| **Adjusting Entries** | Blue | The journal entries generated from Bank-Only Items. These are imported automatically by Module 4. |
| **Exceptions** | Red | Any issues (e.g. opening balance mismatch, unreconciled difference). |

### What to check in the output

1. **Dashboard — Status: RECONCILED** and **Difference = 0.00**
   - If NOT RECONCILED: check the Exceptions sheet, then look at the Reconciliation sheet to see which side is off.

2. **Reconciliation sheet — the formal statement:**
   ```
   Bank Statement Balance          1,495,500
   + Deposits in Transit             500,000
   - Outstanding Cheques            (560,000)
   ─────────────────────────────────────────
   Adjusted Bank Balance           1,435,500
   ═════════════════════════════════════════

   Cash Book Balance               1,438,500
   + Bank Credits not in Book         15,000
   - Bank Debits not in Book         (18,000)
   ─────────────────────────────────────────
   Adjusted Cash Book Balance      1,435,500
   ═════════════════════════════════════════
   Difference:                             0
   ```

3. **Adjusting Entries sheet** — these will be imported automatically into Module 4. You do not need to manually post them.

4. **Outstanding Cheques and Deposits in Transit** — no journal entries needed now. These will appear again next month's bank statement when they clear.

### How the matching works

The system matches transactions using two passes:

| Pass | Logic | Result label |
|---|---|---|
| 1 (Exact) | Same reference number AND same amount on both sides | `Exact` |
| 2 (Probable) | Same amount, dates within ±3 days | `Probable` |

Transactions remaining unmatched after both passes are classified by their direction (receipt vs payment) and which side they appear on.

### Cash book vs bank statement — perspective reminder

The two records use opposite sign conventions:

| Direction | In cash book | On bank statement |
|---|---|---|
| Money received | **Debit** (increases cash) | **Credit** (bank deposits) |
| Money paid out | **Credit** (decreases cash) | **Debit** (bank withdrawals) |

The script normalizes both sides to the same convention before matching, so this is handled automatically.

---

## 8. Module 4 — Journal Adjustments

### What it does

Generates the period-end adjusting entries that must be posted before preparing the trial balance. There are two categories:

**Auto-generated (new entries created by this module):**
1. **Depreciation** — reads the fixed asset register and calculates the monthly straight-line depreciation for every active asset, grouped by category
2. **Bank Reconciliation entries** — imports the adjusting entries generated by Module 3 (bank charges, interest, direct debits not yet in the cash book)

**Reference only (already posted via journals, shown for completeness):**
3. **Accruals** — reads the general ledger for movements in accrued expense accounts (e.g. accrued wages, accrued expenses)
4. **Prepayments** — reads the general ledger for movements in prepaid asset accounts (e.g. prepaid insurance)

### When to run it

After Module 3 is complete and the bank reconciliation status is RECONCILED.

### Input files required

| File | What it contains |
|---|---|
| `fixed_assets_ledger.xlsx` | Asset register with cost, salvage value, useful life, method |
| `chart_of_accounts.xlsx` | For account name lookup |
| `bank_reconciliation_Jan2026.xlsx` | Module 3 output — Adjusting Entries sheet is imported automatically |
| `general_ledger.xlsx` | For reading accrual/prepaid movements (reference sections) |

### How to run

```bash
python scripts/journal_adjustments.py \
  data/Jan2026 \
  2026-01-01 \
  2026-01-31 \
  data/Jan2026/adjusting_entries_Jan2026.xlsx
```

**Arguments in order:**
1. Data directory
2. Period start date
3. Period end date
4. Output file path

### Output file — sheets explained

| Sheet | Tab Colour | What to look at |
|---|---|---|
| **Dashboard** | Green | Entry counts by type, grand total Dr = total Cr, PASS/FAIL double-entry check. |
| **Depreciation Schedule** | Blue | Per-asset calculation table showing cost, salvage, useful life, annual and monthly depreciation. Plus the grouped journal entries (one entry per asset category). |
| **Bank Recon Entries** | Blue | The 2 entries imported from Module 3: bank interest and software subscription. |
| **Accruals** | Blue | Reference: movements in accrued expense accounts already in the GL this period. |
| **Prepayments** | Blue | Reference: movements in prepaid accounts already in the GL this period. |
| **All Entries** | Blue | The master journal — every new adjusting entry (ADJ-001 onwards) with double-entry check at the bottom. |
| **Account Impact** | Orange | Before and after balances for every account affected by the new entries. |
| **Exceptions** | Red | Warnings (e.g. bank reconciliation file not found). |

### What to check in the output

1. **Dashboard — Double-entry check: PASS**
   - Total Debits must equal Total Credits across all adjusting entries.
   - For January 2026: Grand Total = 87,750 (Dr) = 87,750 (Cr).

2. **Depreciation Schedule** — verify the per-asset calculations look right:
   - Formula: `Monthly Depreciation = (Cost − Salvage Value) / Useful Life (years) / 12`
   - Example: Buildings FA-001: (3,000,000 − 300,000) / 20 / 12 = **11,250 per month**

3. **All Entries** — the complete journal before posting to the ledger:

| Entry | Description | Debit | Credit | Amount |
|---|---|---|---|---|
| ADJ-001 | Depreciation — Buildings | 5300 Depr. Expense | 1611 Accum. Depr. | 18,750 |
| ADJ-002 | Depreciation — Plant & Machinery | 5300 Depr. Expense | 1621 Accum. Depr. | 22,500 |
| ADJ-003 | Depreciation — Furniture & Fixtures | 5300 Depr. Expense | 1631 Accum. Depr. | 6,000 |
| ADJ-004 | Depreciation — Office Equipment | 5300 Depr. Expense | 1651 Accum. Depr. | 7,500 |
| ADJ-005 | Bank interest earned | 1020 Cash at Bank | 4110 Interest Income | 15,000 |
| ADJ-006 | Software subscription auto-debit | 5220 Tel & Internet | 1020 Cash at Bank | 18,000 |

4. **Account Impact** — confirms the net effect on cash:
   - Account 1020 (Cash at Bank): `+15,000 (interest) − 18,000 (software) = −3,000 net`
   - Pre-adjustment: 1,438,500 → Post-adjustment: **1,435,500**

### Depreciation — account codes used

| Asset category | Expense account | Accumulated depreciation account |
|---|---|---|
| Buildings | 5300 Depreciation Expense | 1611 Accum. Depr. — Buildings |
| Plant & Machinery | 5300 Depreciation Expense | 1621 Accum. Depr. — P&M |
| Furniture & Fixtures | 5300 Depreciation Expense | 1631 Accum. Depr. — F&F |
| Vehicles | 5300 Depreciation Expense | 1641 Accum. Depr. — Vehicles |
| Office Equipment | 5300 Depreciation Expense | 1651 Accum. Depr. — Equipment |

---

## 9. Input File Formats

All input files are `.xlsx` with a **single header row** at row 1. The scripts use flexible column matching — minor variations in column names (e.g. `Debit Amount` vs `Debit`) are handled automatically.

### Required columns by file

**All journals share:**

| Column | Required | Notes |
|---|---|---|
| Date | Yes | Any recognisable date format |
| Debit Account | Yes | Account code (numeric) |
| Credit Account | Yes | Account code (numeric) |
| Debit Amount | Yes | Numeric, no commas or currency symbols |
| Credit Amount | Yes | Numeric |
| Description | No | Can contain Myanmar text |
| Profit Centre | No | Required for accounts 4000–5999 |
| Cost Centre | No | Required for accounts 5000–5999 |

**Cash Ledger** (`cash_ledger.xlsx`):

| Column | Notes |
|---|---|
| Bank Account | Bank account name or number |
| Date | Transaction date |
| Reference | Cheque number, receipt number, etc. |
| Description | Narration |
| Debit | Money IN (increases cash balance) |
| Credit | Money OUT (decreases cash balance) |
| Balance | Running balance |

**Bank Statement** (`bank_statement.xlsx`):

| Column | Notes |
|---|---|
| Date | Transaction date |
| Reference | Bank reference |
| Description | Bank narration |
| Debit | Money OUT / withdrawal |
| Credit | Money IN / deposit |
| Balance | Running balance |

> **Important:** The cash ledger and bank statement use **opposite conventions** for Debit/Credit. The scripts handle this automatically — do not try to adjust your bank statement to match the cash book format.

**Fixed Assets Ledger** (`fixed_assets_ledger.xlsx`):

| Column | Required | Notes |
|---|---|---|
| Asset ID | Yes | Unique identifier (e.g. FA-001) |
| Description | Yes | Asset name |
| Account Code | Yes | Asset account code (e.g. 1610 for Buildings) |
| Cost | Yes | Original purchase cost |
| Useful Life (Years) | Yes | For depreciation calculation |
| Salvage Value | No | Residual value (default 0 if omitted) |
| Depreciation Method | No | `Straight-Line` (default) or `Reducing Balance` |
| Accumulated Depreciation | No | Total depreciation to date |
| Status | No | `Active` (default) or `Disposed` |

---

## 10. Output Formatting Standards

All output files follow a consistent professional format:

| Element | Standard |
|---|---|
| Header row | Bold, dark blue background (#1F4E79), white text, centred |
| Data rows | Thin borders on all cells |
| Number format | `#,##0` (commas, no decimals) for amounts |
| Negative numbers | Red font, shown in parentheses: `(123,456)` |
| Percentage format | `0.0%` |
| Date format | `YYYY-MM-DD` |
| Total rows | Bold text, thick bottom border |
| Grand total rows | Bold text, double bottom border |
| Freeze panes | Header row and first column on all sheets |
| Column widths | Auto-fit to content, minimum 12 characters |

**Sheet tab colour coding:**

| Colour | Meaning |
|---|---|
| Green | Summary / Dashboard |
| Blue | Detailed data |
| Orange | PC/CC or impact analysis |
| Red | Exceptions / errors |

---

## 11. Troubleshooting & Common Errors

### "File not found" errors

**Symptom:** `ERROR: File not found: data/Jan2026/sales_journal.xlsx`

**Fix:** Check that:
- The file exists in the correct folder
- The filename matches exactly (case-sensitive on Linux/Mac)
- You are running the script from the `accountant-skill/` directory

---

### "Missing required columns" errors

**Symptom:** `ERROR: Missing required columns: ['Debit Account']. Available columns: [...]`

**Fix:**
- Open the input file and check the column headers
- The script shows you what columns it found — use those names or rename your columns to match the expected names
- Column matching is flexible: `Debit Amount`, `Debit`, `Dr` are all accepted for the debit column

---

### "Journals not balanced" (Module 1)

**Symptom:** Dashboard shows `Balanced: False` or Exceptions sheet has "Entry not balanced" rows.

**Fix:**
- Open the Exceptions sheet — it will list exactly which entries are unbalanced and by how much
- Go back to the source journal file and correct the entry
- Re-run Module 1

---

### "Control account mismatch" (Module 2)

**Symptom:** Control Acct Recon sheet shows MISMATCH for AR, AP, or Cash.

**Common causes:**
- A transaction was posted to the general ledger but not to the subsidiary ledger (or vice versa)
- An entry was posted to the wrong account code

**Fix:**
- Compare the GL closing balance against the subsidiary totals
- Identify the missing or extra transaction
- Correct the source data and re-run Module 2

---

### "NOT RECONCILED" (Module 3)

**Symptom:** Bank Reconciliation Dashboard shows `Status: NOT RECONCILED` with a non-zero difference.

**Steps to investigate:**

1. Check the **Reconciliation sheet** — which side has the larger balance?
2. Check **Outstanding Cheques** — are the amounts and dates plausible?
3. Check **Deposits in Transit** — was the deposit actually made before period end?
4. Check **Bank-Only Items** — is there a bank charge or credit you missed in the cash book?
5. Check if the **opening balances** match between cash book and bank statement (Exceptions sheet will flag this)

---

### "Bank recon entries not loading" (Module 4)

**Symptom:** Console shows warning: *"No bank_reconciliation*.xlsx found"*

**Fix:**
- Ensure Module 3 has been run and the output file is in the same `data/[PERIOD]/` folder
- The script searches for any file matching `bank_reconciliation*.xlsx` in the data directory
- Check the filename matches this pattern

---

### "Account XXXX" showing instead of account name (Module 4)

**Symptom:** Account Impact sheet shows `Account 1631` instead of `Accum. Depr. — F&F`

**Cause:** The account is not listed by name in `chart_of_accounts.xlsx`. The system falls back to a generic name but still processes the account correctly.

**Fix:** Add the account to your `chart_of_accounts.xlsx` with its proper name. The accounting results are not affected — only the display name is missing.

---

### Unicode / encoding errors on Windows

**Symptom:** `UnicodeEncodeError: 'charmap' codec can't encode character`

**Cause:** Windows console (cp1252) cannot display certain Unicode characters (box-drawing characters, Myanmar script) in the terminal output.

**Fix:** The Excel output files are not affected — they handle Unicode correctly. The error is only in the console display. If Myanmar text is in your data, it will be written to the Excel file correctly regardless.

---

### Re-running a module for the same period

All modules are **idempotent** — running them again for the same period will overwrite the previous output file with a fresh result. They do not append duplicates. This is safe to do if you have corrected source data.

---

## Quick Reference — All Commands

```bash
# From the accountant-skill/ directory:

# MODULE 1 — Summarize Journals
python scripts/summarize_journals.py \
  data/Jan2026 2026-01-01 2026-01-31 \
  data/Jan2026/books_of_prime_entry_Jan2026.xlsx \
  data/Jan2026/chart_of_accounts.xlsx \
  data/Jan2026/profit_cost_centers.xlsx

# MODULE 2 — Summarize Ledgers
python scripts/summarize_ledgers.py \
  data/Jan2026 2026-01-01 2026-01-31 \
  data/Jan2026/ledger_summary_Jan2026.xlsx \
  data/Jan2026/chart_of_accounts.xlsx

# MODULE 3 — Bank Reconciliation
python scripts/reconcile_bank.py \
  data/Jan2026 2026-01-01 2026-01-31 \
  data/Jan2026/bank_reconciliation_Jan2026.xlsx

# MODULE 4 — Journal Adjustments
python scripts/journal_adjustments.py \
  data/Jan2026 2026-01-01 2026-01-31 \
  data/Jan2026/adjusting_entries_Jan2026.xlsx
```

---

*This guide will be updated as Modules 5 (Trial Balance), 6 (Financial Statements), and 7 (Validation) are completed.*
