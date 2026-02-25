# Accountant Skill — User Guide
### K&K Finance Team
*Last updated: 2026-02-25 | Covers Modules 1–7*

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
9. [Module 5 — Trial Balance](#9-module-5--trial-balance)
10. [Module 6 — Financial Statements](#10-module-6--financial-statements)
11. [Module 7 — Full-Cycle Validation](#11-module-7--full-cycle-validation)
12. [Input File Formats](#12-input-file-formats)
13. [Output Formatting Standards](#13-output-formatting-standards)
14. [Troubleshooting & Common Errors](#14-troubleshooting--common-errors)
    - [Module 7 — Validation Failures](#module-7--validation-failures)
15. [Quick Reference — All Commands](#quick-reference--all-commands)

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
| Preparing trial balance | Module 5 |
| Producing P&L and Balance Sheet | Module 6 |
| Full-cycle integrity audit | Module 7 |

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
│  MODULE 5 ── Trial Balance                                      │
│  Unadjusted → apply adjustments → Adjusted Trial Balance        │
│         │                                                       │
│         ▼                                                       │
│  MODULE 6 ── Financial Statements                               │
│  Income Statement, Balance Sheet, Cash Flow Statement           │
│         │                                                       │
│         ▼                                                       │
│  MODULE 7 ── Full-Cycle Validation                              │
│  Comprehensive audit: double-entry, control accounts,           │
│  cross-module flow, financial statement integrity               │
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
│       ├── adjusting_entries_Jan2026.xlsx       ← Module 4
│       ├── trial_balance_Jan2026.xlsx           ← Module 5
│       ├── financial_statements_Jan2026.xlsx    ← Module 6
│       └── audit_validation_Jan2026.xlsx        ← Module 7
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
    ├── generate_trial_balance.py
    ├── generate_financials.py
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

## 9. Module 5 — Trial Balance

### What it does

Reads the General Ledger closing balances (unadjusted) and the adjusting entries from Module 4, then produces:

1. **Unadjusted Trial Balance** — GL closing balances before any adjustments
2. **Adjustments** — the adjusting entries from Module 4 (individual list + per-account summary)
3. **Adjusted Trial Balance** — balances after applying all adjusting entries
4. **TB Worksheet** — a 6-column combined view: Unadjusted | Adj Entries | Adjusted
5. **Dashboard** — Dr = Cr validation for both the unadjusted and adjusted TB

### When to run it

After Module 4 is complete and all adjusting entries are posted.

### Input files required

| File | What it contains |
|---|---|
| `general_ledger.xlsx` | All account movements (debits, credits, running balance) |
| `adjusting_entries_Jan2026.xlsx` | Module 4 output — "All Entries" sheet is read automatically |
| `chart_of_accounts.xlsx` | For account name and classification lookup |

### How to run

```bash
python scripts/generate_trial_balance.py \
  data/Jan2026 \
  2026-01-01 \
  2026-01-31 \
  data/Jan2026/trial_balance_Jan2026.xlsx
```

**Arguments in order:**
1. Data directory
2. Period start date
3. Period end date
4. Output file path

### Output file — sheets explained

| Sheet | Tab Colour | What to look at |
|---|---|---|
| **Dashboard** | Green | Dr = Cr check for both TBs, total adjusting entries, warnings. |
| **Unadjusted TB** | Blue | GL closing balance per account before adjustments. Dr = Cr check at bottom. |
| **Adjustments** | Blue | Section 1: each individual ADJ- entry. Section 2: per-account net adjustment. |
| **Adjusted TB** | Blue | Final balances after all adjusting entries. This sheet feeds Module 6. |
| **TB Worksheet** | Orange | 6-column combined view: Unadj Debit/Credit | Adj Dr/Cr | Adjusted Debit/Credit. |
| **Exceptions** | Red | Only appears if the TB is not balanced or errors are found. |

### What to check in the output

1. **Dashboard — Adjusted TB: Dr = Cr: PASS**
   - Both the unadjusted and adjusted totals must balance (Total Debit = Total Credit).
   - For January 2026: Adjusted TB totals Dr 21,762,750 = Cr 21,762,750.

2. **Adjustments sheet — per-account summary** — verify the net effect on key accounts:
   - Account 5300 (Depreciation Expense): +54,750 Dr (4 entries)
   - Account 1020 (Cash at Bank): net −3,000 (bank interest +15,000, software −18,000)

3. **Adjusted TB** — review that all IS accounts (4xxx, 5xxx) and BS accounts (1xxx, 2xxx, 3xxx) appear with sensible balances before proceeding to Module 6.

---

## 10. Module 6 — Financial Statements

### What it does

Reads the Adjusted Trial Balance from Module 5 and produces three core financial statements plus a dashboard summary:

1. **Income Statement** — Revenue, COGS, Operating Expenses, Other Income/Expenses, Net Profit with gross/operating/net margins
2. **Balance Sheet** — Non-Current Assets, Current Assets, Equity, Non-Current Liabilities, Current Liabilities, with Assets = Equity + Liabilities check
3. **Cash Flow Statement** — Indirect method: Net Profit → non-cash adjustments → working capital changes → investing activities → financing activities → reconciles to closing cash balance
4. **Dashboard** — key metrics at a glance, plus validation checks for the BS and CF

### When to run it

After Module 5 is complete and the Adjusted Trial Balance balances (Dr = Cr: PASS).

### Input files required

| File | What it contains |
|---|---|
| `trial_balance_Jan2026.xlsx` | Module 5 output — "Adjusted TB" sheet is read automatically |
| `general_ledger.xlsx` | For opening cash balances (used in Cash Flow Statement) |
| `adjusting_entries_Jan2026.xlsx` | For identifying non-cash items (depreciation, bad debt) in Cash Flow |

### How to run

```bash
python scripts/generate_financials.py \
  data/Jan2026 \
  2026-01-01 \
  2026-01-31 \
  data/Jan2026/financial_statements_Jan2026.xlsx
```

**Arguments in order:**
1. Data directory
2. Period start date
3. Period end date
4. Output file path

### Output file — sheets explained

| Sheet | Tab Colour | What to look at |
|---|---|---|
| **Dashboard** | Green | Key metrics (Revenue, Gross Profit, Net Profit, margins, Total Assets), BS check (PASS/FAIL), CF check (PASS/FAIL), Cash Flow summary. |
| **Income Statement** | Blue | Full P&L: Revenue → Net Revenue → Gross Profit → Operating Profit → Net Profit, with percentage margins. |
| **Balance Sheet** | Blue | Full balance sheet with Assets = Equity + Liabilities check at the bottom. |
| **Cash Flow** | Blue | Indirect method CF: Net Profit + non-cash items + WC changes + investing + financing = Net change in cash. Reconciles to closing cash. |
| **Exceptions** | Red | Only appears if the BS does not balance or the CF does not reconcile. |

### What to check in the output

1. **Dashboard — Balance Sheet check: PASS**
   - Total Assets must equal Total Equity + Total Liabilities. Difference must be 0.00.
   - For January 2026: Total Assets = 13,215,750 = Equity 9,142,750 + Liabilities 4,073,000.

2. **Dashboard — Cash Flow check: PASS**
   - Opening Cash + Net Change in Cash must equal Closing Cash (from Balance Sheet).
   - For January 2026: 1,550,000 + 35,500 = 1,585,500.

3. **Income Statement — Net Profit** must be the same figure shown in the Balance Sheet under Equity as "Current Period Net Profit/(Loss)".

4. **Balance Sheet — Contra accounts** (Allowance for Doubtful Debts, Accumulated Depreciation, Owner's Drawings) are shown indented and in parentheses (negative). This is correct presentation.

5. **Exceptions sheet** — if present, address all items before using the statements for reporting.

### How the Income Statement is structured

| Section | Account range | Treatment |
|---|---|---|
| Revenue | 4000–4099 | Credit balance = positive revenue |
| Less: Contra Revenue | 4200–4299 | Debit balance = shown as deduction (negative) |
| **Net Revenue** | | Gross Revenue minus contra items |
| Cost of Goods Sold | 5000–5099 | Debit balance = shown as cost (negative subtotal) |
| **Gross Profit** | | Net Revenue minus COGS |
| Operating Expenses | 5100–5899 | Debit balance = shown as cost |
| **Operating Profit** | | Gross Profit minus Operating Expenses |
| Other Income | 4100–4199 | Credit balance = positive |
| Other Expenses | 5900–5949 | Debit balance = shown as negative |
| Tax Expense | 5950 | Debit balance = shown as negative |
| **Net Profit** | | Operating Profit + Net Other − Tax |

### How the Cash Flow (indirect method) works

Starting from Net Profit, the module makes three sets of adjustments:

**1. Non-cash items (add back):**
- Depreciation expense (account 5300) — identified from adjusting entries
- Bad debt expense (account 5800) — identified from adjusting entries

**2. Working capital changes:**
- Current asset increase → negative CF (cash tied up in receivables/inventory)
- Current asset decrease → positive CF (cash released)
- Current liability increase → positive CF (cash not yet paid)
- Current liability decrease → negative CF (cash paid)
- Includes the Allowance for Doubtful Debts (1110) as a credit-normal working capital item

**3. Investing and Financing:**
- Investing: changes in fixed asset accounts (1600–1660)
- Financing: changes in capital (3010), drawings (3020), loans (2060, 2100, 2110)

---

## 11. Module 7 — Full-Cycle Validation

### What it does

Performs a comprehensive audit of the entire accounting cycle by reading all module outputs (Modules 1–6) and validating:

1. **Double-Entry Integrity** — Every journal, adjusting entry, and trial balance has Debits = Credits
2. **Control Account Reconciliation** — AR, AP, and Cash GL balances match their subsidiary ledgers
3. **Cross-Module Flow** — Data flows correctly between modules (e.g., Module 3 adjusting entries appear in Module 4, Module 4 entries flow to Module 5, Module 5 net profit ties to Module 6)
4. **Financial Statement Validation** — Balance Sheet balances (Assets = Equity + Liabilities), Cash Flow reconciles (Opening + Net Change = Closing)

The module produces an audit validation report with a dashboard summary showing PASS/FAIL status for every check.

### When to run it

After Module 6 is complete and all financial statements have been generated. Module 7 is the final quality gate before issuing financial reports.

### Input files required

Module 7 automatically reads all prior module outputs from the data directory:

| File | Source | What it validates |
|---|---|---|
| `books_of_prime_entry_*.xlsx` | Module 1 | Journal double-entry balance |
| `ledger_summary_*.xlsx` | Module 2 | Control account reconciliation |
| `bank_reconciliation_*.xlsx` | Module 3 | Bank recon adjusting entries |
| `adjusting_entries_*.xlsx` | Module 4 | Adjusting entry balance |
| `trial_balance_*.xlsx` | Module 5 | TB balance, data flow from Module 4 |
| `financial_statements_*.xlsx` | Module 6 | BS balance, CF reconciliation |
| `chart_of_accounts.xlsx` | Master data | Account classification |

### How to run

```bash
python scripts/validate_accounting.py \
  data/Jan2026 \
  2026-01-01 \
  2026-01-31 \
  data/Jan2026/audit_validation_Jan2026.xlsx
```

**Arguments in order:**
1. Data directory
2. Period start date
3. Period end date
4. Output file path

### Output file — sheets explained

| Sheet | Tab Colour | What to look at |
|---|---|---|
| **Dashboard** | Green | Summary of all validation checks with PASS/FAIL status. Overall result at top. |
| **Double-Entry Checks** | Blue | Detailed results for each journal and TB balance check. Shows Dr total, Cr total, difference. |
| **Control Account Recon** | Blue | AR, AP, Cash reconciliation: GL balance vs subsidiary ledger total, with difference. |
| **Cross-Module Flow** | Orange | Data flow validation between modules. Shows item counts and ties (e.g., net profit tie-out). |
| **Financial Validation** | Blue | Balance Sheet and Cash Flow validation results with detailed calculations. |
| **Exceptions** | Red | Only appears if there are FAIL or WARN results. Lists all items requiring attention. |

### What to check in the output

1. **Dashboard — Overall Result: PASS**
   - All critical checks must pass. If any show FAIL, do not proceed with financial reporting.
   - Summary shows: X/Y checks passed, with breakdown by status (PASS/FAIL/WARN/SKIP).

2. **Double-Entry Checks — all PASS:**
   - Module 1 journals: Each of the 6 journals must balance
   - Module 4 adjusting entries: All ADJ- entries combined must balance
   - Module 5 trial balances: Both Unadjusted and Adjusted TB must balance

3. **Control Account Recon — all MATCH:**
   - AR (1100): GL balance = Sum of customer balances
   - AP (2010): GL balance = Sum of supplier balances
   - Cash (1020): GL balance = Cash ledger closing balance

4. **Cross-Module Flow — data continuity:**
   - Module 3 → Module 4: Bank reconciliation adjusting entries appear in Module 4
   - Module 4 → Module 5: All adjusting entries flow to the Trial Balance
   - Module 5 → Module 6: Net profit from TB matches IS net profit

5. **Financial Validation — key equations:**
   - Balance Sheet: Assets = Equity + Liabilities (Difference must be 0.00)
   - Cash Flow: Opening Cash + Net Change = Closing Cash (Difference must be 0.00)

6. **Exceptions sheet — if present:**
   - Review all FAIL items first — these are critical errors that must be fixed
   - WARN items indicate potential issues that should be investigated
   - SKIP means data was not available for that check (e.g., subsidiary ledger not found)

### Validation check categories

| Category | Checks | What it validates |
|---|---|---|
| Double-Entry | 4+ checks | Every journal, adjusting entry, and TB has Dr = Cr |
| Control Account Recon | 3 checks | AR, AP, Cash GL balances match subsidiary ledgers |
| Cross-Module Flow | 3 checks | Data flows correctly between modules |
| Financial Validation | 4+ checks | BS balances, CF reconciles, dashboard checks |

### Example output (January 2026)

```
============================================================
  MODULE 7 -- FULL-CYCLE ACCOUNTING VALIDATION
  Period : 2026-01-01 to 2026-01-31
  Data   : data/Jan2026
  Output : data/Jan2026/audit_validation_Jan2026.xlsx
============================================================

Loading module outputs...
  All module outputs loaded successfully.

Running validation checks...
  - Checking double-entry integrity...
  - Reconciling control accounts...
  - Validating cross-module data flow...
  - Validating financial statements...

Validation complete: 5/5 passed, 0 failed, 0 warnings

============================================================
  OUTPUT  : data/Jan2026/audit_validation_Jan2026.xlsx
  Sheets  : Dashboard | Double-Entry Checks | Control Account Recon
            Cross-Module Flow | Financial Validation | Exceptions
  RESULT  : PASS (5/5 checks passed)
============================================================
```

### Understanding check statuses

| Status | Meaning | Action required |
|---|---|---|
| PASS | Check succeeded | None |
| FAIL | Check failed | Investigate and fix the root cause before proceeding |
| WARN | Check passed with warnings | Review the warning details; may not block reporting |
| SKIP | Check could not be performed | Data not available (e.g., subsidiary ledger not provided) |

### Common failure scenarios

| Failure | Likely cause | Fix |
|---|---|---|
| Double-entry FAIL | Journal entry with Dr ≠ Cr | Go back to source journal, correct the entry, re-run Module 1 |
| Control account MISMATCH | GL and subsidiary ledger totals differ | Find the missing/duplicate transaction, correct source data, re-run Module 2 |
| Cross-module flow WARN | Adjusting entries not flowing to next module | Check that prior module output files exist and are correctly formatted |
| Balance Sheet FAIL | BS does not balance | Check for unknown account codes in Adjusted TB; verify Module 5 output is balanced |
| Cash Flow FAIL | CF does not reconcile to closing cash | Check that all BS account changes are captured in CF sections |

---

## 12. Input File Formats

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

## 13. Output Formatting Standards

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

## 14. Troubleshooting & Common Errors

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

### Trial Balance not balanced (Module 5)

**Symptom:** Dashboard shows `Adjusted TB: Dr = Cr: FAIL` or Exceptions sheet appears.

**Steps to investigate:**

1. Check **Dashboard — Unadjusted TB** first. If it is also unbalanced, the error is in the GL data, not the adjustments. Fix Module 2 / Module 4 inputs and re-run.
2. If only the **Adjusted TB** is unbalanced, check the adjusting entries: open `adjusting_entries_Jan2026.xlsx` → All Entries sheet → confirm the Dr = Cr check at the bottom shows PASS.
3. Re-run Module 4 to regenerate the adjusting entries, then re-run Module 5.

---

### Balance Sheet does not balance (Module 6)

**Symptom:** Dashboard shows `Balance Sheet check: FAIL` with a non-zero difference. Exceptions sheet appears.

**Common causes:**
- An account code appears in the Adjusted TB but is not classified (falls in an unknown range)
- Net profit was not added to equity — this would show as Assets > Equity + Liabilities by exactly the net profit amount
- A balance sheet account has an abnormal balance (e.g. a debit balance on a liability)

**Steps to investigate:**

1. Open the Balance Sheet sheet — look for any section where totals seem wrong
2. Open the Exceptions sheet — it shows the exact difference
3. Verify the Adjusted TB is balanced (Dr = Cr) before running Module 6 — an unbalanced TB will cause the BS to fail
4. Check for any unknown account codes in the console output — unknown accounts are excluded and will cause an imbalance

---

### Cash Flow does not reconcile (Module 6)

**Symptom:** Dashboard shows `Cash Flow check: FAIL`. The difference shown = Opening Cash + Net Change − Closing Cash.

**Common causes:**
- A balance sheet account changed during the period but is not captured in any CF section (operating WC, investing, or financing)
- The GL opening balance file (`general_ledger.xlsx`) is not available — opening cash defaults to 0

**Steps to investigate:**

1. Check the console output for warnings about `general_ledger.xlsx` not found
2. Review the Cash Flow sheet — check each section total looks reasonable
3. The difference value tells you how much is unaccounted. Cross-check which BS accounts changed by that amount

> **Note:** A small CF discrepancy (e.g. rounding of less than 1.00) can be ignored. Differences larger than 1.00 require investigation.

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

### Module 7 — Validation failures

**Symptom:** Module 7 Dashboard shows FAIL for one or more checks.

**By check type:**

| Check Type | Common Cause | Fix |
|---|---|---|
| Double-entry FAIL | Journal or TB has Dr ≠ Cr | Go to source module, fix the unbalanced entry, re-run that module |
| Control Account FAIL | GL ≠ Subsidiary Ledger | Trace missing/duplicate transaction, correct source data, re-run Module 2 |
| Cross-module WARN | Data gap between modules | Check prior module outputs exist; may be expected if no adjusting entries |
| Balance Sheet FAIL | BS doesn't balance | Verify Adjusted TB is balanced; check for unknown account codes |
| Cash Flow FAIL | CF doesn't reconcile | Check all BS account changes are captured in CF sections |

**General approach:**
1. Open the Exceptions sheet — it lists all FAIL and WARN items
2. Fix FAIL items in priority order: Double-entry → Control Accounts → Financial Validation
3. Re-run the affected module(s) upstream, then re-run Module 7
4. WARN items can often be reviewed and accepted if the explanation is satisfactory

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

# MODULE 5 — Trial Balance
python scripts/generate_trial_balance.py \
  data/Jan2026 2026-01-01 2026-01-31 \
  data/Jan2026/trial_balance_Jan2026.xlsx

# MODULE 6 — Financial Statements
python scripts/generate_financials.py \
  data/Jan2026 2026-01-01 2026-01-31 \
  data/Jan2026/financial_statements_Jan2026.xlsx

# MODULE 7 — Full-Cycle Validation
python scripts/validate_accounting.py \
  data/Jan2026 2026-01-01 2026-01-31 \
  data/Jan2026/audit_validation_Jan2026.xlsx
```

---

*All 7 modules are now complete. The accounting cycle is fully automated from source journals through financial statements with end-to-end validation.*
