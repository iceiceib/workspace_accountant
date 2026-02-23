# Ledger Posting & Summarization Rules

## Overview

The ledger is the principal book of account. Journal entries are posted (transferred) to individual ledger accounts. Each account accumulates debits and credits to maintain a running balance.

## Posting Process

Journal → Ledger posting follows this flow:
1. For each journal entry, identify the debit account and credit account
2. Post the debit amount to the debit side of the debit account's ledger
3. Post the credit amount to the credit side of the credit account's ledger
4. Include the journal reference, date, and description

## Ledger Account Structure (T-Account)

Each ledger account follows this format:

```
Account: [Code] [Name]
Opening Balance: [Amount] (Dr/Cr side based on normal balance)

| Date | Reference | Description | Debit | Credit | Balance |
|------|-----------|-------------|-------|--------|---------|
| ...  | ...       | ...         | ...   | ...    | ...     |

Closing Balance: [Amount]
```

## Balance Calculation

For each account:
- **Opening Balance** = Closing balance from prior period (or zero if new)
- **Closing Balance** = Opening Balance + Sum(Debits) - Sum(Credits) for debit-normal accounts
- **Closing Balance** = Opening Balance + Sum(Credits) - Sum(Debits) for credit-normal accounts

Quick rule:
- Assets & Expenses (debit-normal): Balance increases with debits, decreases with credits
- Liabilities, Equity & Revenue (credit-normal): Balance increases with credits, decreases with debits

## Control Account Reconciliation

Subsidiary ledgers must reconcile to their control accounts in the General Ledger:

| Subsidiary Ledger | GL Control Account |
|---|---|
| Accounts Receivable Ledger (sum of all customer balances) | 1100 Accounts Receivable |
| Accounts Payable Ledger (sum of all supplier balances) | 2010 Accounts Payable |
| Cash Ledger (sum of all bank accounts) | 1020+1021+1022 Cash at Bank |

If these don't match, there's a posting error that must be investigated.

## Summarization Logic

For the Ledger Summary report:

### Per-Account Summary
1. Determine opening balance for the period
2. Sum all debits posted during the period
3. Sum all credits posted during the period
4. Calculate closing balance
5. Determine if the balance is normal (debit balance for assets/expenses, credit for liabilities/equity/revenue)

### Cross-Check Against Journals
After summarizing all ledger accounts:
- Total debits posted to all ledger accounts should equal total debits from all journals
- Total credits posted to all ledger accounts should equal total credits from all journals
- If they don't match, report the difference and investigate

### Movement Analysis
For each account, calculate:
- Period movement = Closing Balance - Opening Balance
- Percentage change = Movement / |Opening Balance| × 100
- Flag accounts with >50% change for review (unusual movements)

## Output Structure

The ledger summary workbook should contain:
- **Summary Dashboard**: Total accounts, total debit balances, total credit balances, balance check
- **General Ledger Balances**: All accounts with opening/debits/credits/closing
- **AR Ledger by Customer**: Customer-level balances
- **AP Ledger by Supplier**: Supplier-level balances
- **Cash Ledger by Bank**: Bank account-level balances
- **Fixed Assets Summary**: Asset-level NBV schedule
- **Equity Summary**: Equity account movements
- **Control Account Reconciliation**: Subsidiary vs GL comparison
- **Journal-to-Ledger Cross-Check**: Total comparison with variance
