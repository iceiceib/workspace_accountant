# Bank Reconciliation Rules

## Overview

Bank reconciliation is the process of comparing the business's cash book (internal records) with the bank statement (external records) to identify and explain any differences. At the end of reconciliation, the adjusted balances must agree.

## Why Balances Differ

The cash book and bank statement may show different balances because:

1. **Timing differences** — Transactions recorded in one but not yet in the other
2. **Bank charges/interest** — Appear on bank statement first, not yet in cash book
3. **Direct debits/credits** — Standing orders, direct deposits appear on bank statement
4. **Errors** — In either the cash book or the bank statement

## Reconciling Items

### Items in Cash Book but NOT on Bank Statement
| Item | Description | Treatment |
|------|-------------|-----------|
| Outstanding cheques | Cheques written and recorded but not yet cleared | Deduct from bank balance |
| Deposits in transit | Deposits recorded but not yet credited by bank | Add to bank balance |

### Items on Bank Statement but NOT in Cash Book
| Item | Description | Treatment |
|------|-------------|-----------|
| Bank charges | Service fees, transaction fees | Deduct from cash book balance → needs adjusting entry |
| Bank interest earned | Interest credited by bank | Add to cash book balance → needs adjusting entry |
| Direct debits | Automated payments (insurance, loans) | Deduct from cash book balance → needs adjusting entry |
| Direct credits | Customer payments made directly to bank | Add to cash book balance → needs adjusting entry |
| Dishonoured cheques | Customer cheques that bounced | Deduct from cash book balance → needs adjusting entry |

### Errors
| Error | Treatment |
|-------|-----------|
| Transposition error in cash book | Correct cash book balance |
| Wrong amount recorded | Correct the side with the error |
| Recorded in wrong bank account | Transfer between accounts |

## Reconciliation Statement Format

```
BANK RECONCILIATION STATEMENT
As at [DATE]

Balance per Bank Statement                           XXX

Add: Deposits in transit
     [Date] [Reference] [Amount]                    +XXX
                                                    +XXX
Less: Outstanding cheques
     [Date] [Reference] [Amount]                    (XXX)
                                                    (XXX)
                                                   ─────
Adjusted Bank Balance                                XXX
                                                   ═════

Balance per Cash Book                                XXX

Add: Items to be recorded in cash book
     Bank interest [Date]                           +XXX
     Direct credits [Date] [From]                   +XXX
                                                    +XXX
Less: Items to be recorded in cash book
     Bank charges [Date]                            (XXX)
     Direct debits [Date] [To]                      (XXX)
     Dishonoured cheque [Date] [Customer]           (XXX)
                                                    (XXX)

Add/Less: Error corrections                         ±XXX
                                                   ─────
Adjusted Cash Book Balance                           XXX
                                                   ═════

Difference (must be ZERO)                              0
```

## Matching Algorithm

When auto-matching transactions between cash book and bank statement:

1. **Exact match**: Same date + same amount + same reference → matched
2. **Amount match**: Same amount within ±3 days, no reference match → probable match (flag for review)
3. **Reference match**: Same reference but different date/amount → flag as error
4. **Unmatched cash book items**: Likely outstanding cheques or deposits in transit
5. **Unmatched bank items**: Likely bank charges, interest, direct debits/credits

## Adjusting Entries from Reconciliation

Items on the bank statement that are NOT in the cash book require adjusting journal entries. Generate these automatically:

| Reconciling Item | Debit | Credit |
|---|---|---|
| Bank charges | 5920 Bank Charges | 1020 Cash at Bank |
| Bank interest earned | 1020 Cash at Bank | 4110 Interest Income |
| Direct debit (e.g., insurance) | 5700 Insurance Expense | 1020 Cash at Bank |
| Direct credit (customer payment) | 1020 Cash at Bank | 1100 Accounts Receivable |
| Dishonoured cheque | 1100 Accounts Receivable | 1020 Cash at Bank |

These entries should be output as "Required Adjusting Entries" and can feed directly into Module 4.

## Output Structure

The reconciliation workbook should contain:
- **Reconciliation Statement**: The formal reconciliation (format above)
- **Matched Transactions**: All successfully matched items
- **Outstanding Cheques**: Unmatched cash book payments
- **Deposits in Transit**: Unmatched cash book receipts
- **Bank-Only Items**: Charges, interest, direct debits/credits
- **Unmatched / Errors**: Items needing investigation
- **Required Adjusting Entries**: Journal entries to update cash book
