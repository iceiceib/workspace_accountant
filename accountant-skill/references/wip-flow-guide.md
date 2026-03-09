# Work-in-Progress (WIP) Flow Guide

## Overview

WIP accounts track manufacturing costs as raw materials are converted into finished goods. This guide shows how WIP accounts flow through the accounting cycle using the **periodic inventory system**.

---

## WIP Account Codes

| Code | Account Name | Type | Normal Balance |
|------|-------------|------|----------------|
| 50300 | Opening Work-in-Progress | COGS | Debit |
| 50310 | Closing Work-in-Progress | COGS | Debit |
| 50320 | Direct Materials Used | COGS | Debit |
| 50330 | Direct Labor Transferred to WIP | COGS | Debit |
| 50340 | Manufacturing Overhead Applied | COGS | Debit |
| 50350 | WIP Transferred to Finished Goods | COGS | Debit (contra effect) |

**Balance Sheet Account:**
- 12400: Work-in-Progress (Asset - Inventory)

---

## The Manufacturing Cost Flow

```
Raw Materials → WIP → Finished Goods → Cost of Goods Sold
     ↓           ↓          ↓            ↓
  50010     50320      50350        (Sale)
  50110     50330
            50340
```

---

## Step-by-Step Journal Entries

### Step 1: Record Opening WIP Balance (Period Start)

At the beginning of the period, bring forward the WIP balance:

| Dr/Cr | Account Code | Account Name | Amount |
|-------|-------------|--------------|--------|
| Dr | 50300 | Opening Work-in-Progress | 50,000 |
| Cr | 12400 | Work-in-Progress (Asset) | 50,000 |

**Purpose:** Recognizes the value of partially completed goods from the prior period as a cost of the current period.

---

### Step 2: Direct Materials Used in Production

Transfer raw materials and packaging into production:

| Dr/Cr | Account Code | Account Name | Amount |
|-------|-------------|--------------|--------|
| Dr | 50320 | Direct Materials Used | 150,000 |
| Cr | 50010 | Purchases Raw Materials | 120,000 |
| Cr | 50110 | Purchases Packaging | 30,000 |

**Purpose:** Moves material costs from inventory purchases into the WIP cost pool.

---

### Step 3: Direct Labor Transferred to WIP

Allocate direct production labor to WIP:

| Dr/Cr | Account Code | Account Name | Amount |
|-------|-------------|--------------|--------|
| Dr | 50330 | Direct Labor Transferred to WIP | 80,000 |
| Cr | 53000 | Direct Labor Wages | 80,000 |

**Purpose:** Reclassifies direct labor from a general production cost to specific WIP allocation.

---

### Step 4: Manufacturing Overhead Applied to WIP

Apply overhead costs (utilities, depreciation, indirect labor) to WIP:

| Dr/Cr | Account Code | Account Name | Amount |
|-------|-------------|--------------|--------|
| Dr | 50340 | Manufacturing Overhead Applied | 45,000 |
| Cr | 53100 | Machine Maintenance & Repair | 15,000 |
| Cr | 53200 | Production Utilities | 20,000 |
| Cr | 53300 | Depreciation Expenses - COGS | 10,000 |

**Purpose:** Allocates indirect manufacturing costs to WIP based on a predetermined overhead rate.

---

### Step 5: WIP Transferred to Finished Goods

When goods are completed, transfer total cost to finished goods:

| Dr/Cr | Account Code | Account Name | Amount |
|-------|-------------|--------------|--------|
| Dr | 50200 | Opening Inventory - Finished Goods | 200,000 |
| Cr | 50300 | Opening Work-in-Progress | 50,000 |
| Cr | 50320 | Direct Materials Used | 150,000 |
| Cr | 50330 | Direct Labor Transferred to WIP | 80,000 |
| Cr | 50340 | Manufacturing Overhead Applied | 45,000 |
| Dr | 50350 | WIP Transferred to Finished Goods | 325,000 |

**Note:** Account 50350 is a memo/transfers account. The net effect closes out all WIP accumulation accounts.

**Purpose:** Moves completed goods cost out of WIP and into finished goods inventory.

---

### Step 6: Record Closing WIP Balance (Period End)

Physical count shows ending WIP value. Record as an asset:

| Dr/Cr | Account Code | Account Name | Amount |
|-------|-------------|--------------|--------|
| Dr | 12400 | Work-in-Progress (Asset) | 65,000 |
| Cr | 50310 | Closing Work-in-Progress | 65,000 |

**Purpose:** Recognizes the value of partially completed goods still in production at period-end as an asset (not an expense).

---

## T-Account Visualization

### Work-in-Progress (Asset Account 12400)
```
                    Work-in-Progress (12400)
---------------------------------------------------
Beg Balance    50,000 | Step 1    50,000
Step 6         65,000 |
---------------------------------------------------
End Balance    65,000 |
```

### WIP Cost Accumulation (COGS Accounts)
```
  Direct Materials Used (50320)        Direct Labor to WIP (50330)
---------------------------------      ---------------------------------
Step 2    150,000 | Step 5  150,000    Step 3     80,000 | Step 5   80,000
---------------------------------      ---------------------------------
Balance        0 |                      Balance         0 |


  Manufacturing Overhead Applied (50340)
---------------------------------
Step 4     45,000 | Step 5   45,000
---------------------------------
Balance         0 |


  Opening WIP (50300)              Closing WIP (50310)
---------------------------------      ---------------------------------
Step 1    50,000 | Step 5   50,000    Step 6     65,000 | (carried forward)
---------------------------------      ---------------------------------
Balance        0 |                      Balance    65,000 | (asset)


  WIP to Finished Goods (50350)
---------------------------------
Step 5   325,000 | (transfer out)
---------------------------------
(Net effect: closes WIP accumulation)
```

---

## Cost of Goods Manufactured (COGM) Formula

```
COGM = Opening WIP + Total Manufacturing Costs - Closing WIP

Where:
  Total Manufacturing Costs = Direct Materials + Direct Labor + Manufacturing Overhead

Example:
  Opening WIP                          50,000
  + Direct Materials Used             150,000
  + Direct Labor Transferred           80,000
  + Manufacturing Overhead Applied     45,000
  -----------------------------------------
  Total Manufacturing Costs           325,000

  - Closing WIP                        65,000
  -----------------------------------------
  Cost of Goods Manufactured          260,000
```

---

## Income Statement Presentation

```
K&K Business - Income Statement (Partial)
For the Period Ended XXXX

Revenue                                    500,000
Cost of Goods Sold:
  Opening Inventory - Raw Materials         30,000
  Purchases Raw Materials                  200,000
  Closing Inventory - Raw Materials        (40,000)
  Opening Inventory - Packaging             10,000
  Purchases Packaging                       50,000
  Closing Inventory - Packaging            (15,000)
  Opening Inventory - Finished Goods        80,000
  Opening Work-in-Progress                  50,000
  Direct Materials Used                    150,000
  Direct Labor Transferred to WIP           80,000
  Manufacturing Overhead Applied            45,000
  WIP Transferred to Finished Goods       (325,000)
  Closing Inventory - Finished Goods      (120,000)
  Closing Work-in-Progress                 (65,000)
  -----------------------------------------
  Total Cost of Goods Sold                 100,000
  -----------------------------------------
Gross Profit                               400,000
```

---

## Module Integration

### Module 1 (Summarize Journals)
- Captures all WIP-related journal entries
- Validates Dr = Cr for each WIP transfer entry

### Module 2 (Summarize Ledgers)
- Posts WIP entries to individual ledger accounts
- Shows WIP movement by period

### Module 5 (Trial Balance)
- Includes WIP accounts in unadjusted and adjusted TB
- Account 12400 appears in Asset section
- WIP COGS accounts (50300-50350) appear in expense section

### Module 6 (Financial Statements)
- COGS section includes WIP cost flow accounts
- Balance Sheet shows WIP (12400) as part of Inventory
- COGM can be shown as supporting schedule

### Module 7 (Validation)
- Verifies WIP transfers balance (Dr = Cr)
- Confirms COGM calculation ties to Finished Goods

---

## Common WIP Scenarios

### Scenario A: No WIP (Trading Company)
If there's no manufacturing, all WIP accounts remain at zero.

### Scenario B: Continuous Production
WIP balances carry forward each period as production is ongoing.

### Scenario C: Job Order Costing
Each job has its own WIP sub-ledger. Total WIP = sum of all incomplete jobs.

### Scenario D: Process Costing
WIP is tracked by department/process. Costs accumulate by stage.

---

## Audit Trail

For proper WIP tracking, maintain:

1. **Material Requisition Forms** - Support Direct Materials Used (50320)
2. **Labor Time Cards** - Support Direct Labor Transferred (50330)
3. **Overhead Allocation Worksheet** - Support Overhead Applied (50340)
4. **Production Completion Reports** - Support WIP to FG Transfer (50350)
5. **Period-End WIP Count Sheets** - Support Closing WIP (50310 / 12400)
