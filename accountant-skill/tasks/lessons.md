# Lessons Learned

## Patterns to Remember

### Windows Console Encoding
- **Issue:** Unicode box/check characters (`─`, `✓`, `✗`) cause `UnicodeEncodeError` on Windows cp1252 console
- **Solution:** Use plain ASCII only in `print()` statements (`-`, `/`, `[PASS]`, `[FAIL]`)
- **Applies to:** All module scripts

### NaN Handling with pandas
- **Issue:** `float('nan')` is truthy in Python - `val or ''` does NOT catch NaN
- **Solution:** Always use `pd.isna()` or `math.isnan()` to check for NaN values
- **Applies to:** All data processing code

### Excel Account Code Normalization
- **Issue:** Excel reads numeric account codes as floats (e.g., `10100.0` instead of `10100`)
- **Solution:** Normalize with `str(int(float(str(val).strip())))`
- **Applies to:** All modules reading account codes from Excel

### Header Row Detection
- **Issue:** Output sheets have `write_title()` block before column headers
- **Solution:** When reading back, scan for header row by searching for key column name (e.g., 'Date')
- **Applies to:** Module 4 reading Module 3 output, Module 5 reading Module 4 output, Module 7 reading all outputs

### Grand Total Row Filtering
- **Issue:** Grand-total rows in output sheets have non-numeric Dr/Cr codes
- **Solution:** Filter with `if not _norm_code(row['Dr Code']).isdigit(): continue`
- **Applies to:** Any module reading another module's output

## Module-Specific Lessons

### Module 1 (Summarize Journals)
- Profit Center/Cost Center data is optional - sheets only generated if data exists
- Single Amount column journals (sales, purchases) are implicitly balanced

### Module 3 (Bank Reconciliation)
- Bank statement and cash ledger use opposite debit/credit conventions
- The script handles this automatically - don't manually adjust data
- Matching uses two passes: Exact (ref + amount) then Probable (amount only, dates within ±3 days)

### Module 4 (Adjustments)
- Bank reconciliation adjusting entries are auto-imported from Module 3 output
- Depreciation is calculated per asset, then grouped by category for journal entries

### Module 5 (Trial Balance)
- Adjusting entries from Module 4 are read from "All Entries" sheet
- Filter out TOTALS row by checking Dr/Cr codes are numeric
- Balance display: accounts show on normal balance side (abnormal balances show on opposite)

### Module 6 (Financial Statements)
- Reads Adjusted TB from Module 5
- Net Profit flows from Income Statement to Balance Sheet Equity
- Cash Flow uses indirect method starting from Net Profit
- Balance Sheet totals may be in the last column (Total column), not first value column

### Module 7 (Full-Cycle Validation) - NEW
- Reads ALL prior module outputs automatically from data directory
- Validates 4 categories: Double-Entry, Control Account Recon, Cross-Module Flow, Financial Validation
- Dashboard sheet shows PASS/FAIL summary with color coding
- Exceptions sheet only created if there are FAIL or WARN results
- Balance Sheet parsing: look for "TOTAL ASSETS" but exclude "TOTAL NON-CURRENT ASSETS" and "TOTAL CURRENT ASSETS"
- Equity parsing: exclude "TOTAL EQUITY & LIABILITIES" row when finding "TOTAL EQUITY"

## Cross-Module Data Flow

```
M1 (Journals) → M2 (Ledgers) → M5 (Unadjusted TB)
                    ↓
M3 (Bank Recon) → M4 (Adjustments) → M5 (Adjusted TB)
                                          ↓
                              M6 (Financial Statements)
                                          ↓
                              M7 (Validation - reads M1-M6)
```

## Testing Strategy

1. Generate test data with `create_test_data.py` and `create_bank_statement.py`
2. Run modules 1-7 in sequence
3. Verify Module 7 Dashboard shows all PASS
4. Key validation points:
   - M1: Grand Total Dr = Cr
   - M2: AR/AP/Cash control accounts MATCH
   - M3: Status = RECONCILED
   - M4: Double-entry check PASS
   - M5: Both TBs balance
   - M6: BS balances, CF reconciles
   - M7: All checks PASS

## WIP (Work-in-Progress) Accounting - Added 2026-03-06

### Key Lessons
1. WIP accounts (50300-50350, 12400) must be added to both `general_journal.xlsx` AND `general_ledger.xlsx`
2. WIP accumulation accounts (50320, 50330, 50340) correctly clear to zero - they are filtered from TB
3. COGM formula: Opening WIP + Direct Materials + Direct Labor + Overhead - Closing WIP
4. Account 50350 (WIP Transferred to FG) should equal COGM amount
5. Closing WIP (50310) has credit normal balance - it's a contra-expense that reduces COGS

### Pre-existing Test Data Issues
- The 1,120,000 TB imbalance existed before WIP entries were added
- This is due to periodic inventory system without closing inventory entries recorded
- WIP entries themselves are perfectly balanced

### Files Updated for WIP
- `references/chart-of-accounts.md` - Added WIP account structure
- `references/wip-flow-guide.md` - New comprehensive WIP guide
- `scripts/utils/coa_mapper.py` - Added 50300-50399 range to COGS classification
- `data/Jan2026/chart_of_accounts.xlsx` - Added 6 WIP accounts
- `data/Jan2026/general_journal.xlsx` - Added 12 WIP journal entries
- `data/Jan2026/general_ledger.xlsx` - Added 18 WIP transaction rows

## Code Fixes (2026-03-10)

### Fixed: Account Code Ranges in pc_cc_mapper.py
- **Issue:** PC_REQUIRED_RANGES and CC_REQUIRED_RANGES used 4-digit codes (4000-4999, 5000-5999)
  but the COA uses 5-digit codes (40000-49999, 50000-69999)
- **Fix:** Updated all ranges to 5-digit format
- **Files affected:** `scripts/utils/pc_cc_mapper.py`

### Fixed: Control Account Codes in summarize_ledgers.py
- **Issue:** AR_GL_ACCOUNT=1100, AP_GL_ACCOUNT=2010, CASH_GL_ACCOUNTS=[1020,1021,1022]
  but the COA uses 5-digit codes (11000, 20000, 10100)
- **Fix:** Updated to correct 5-digit codes
- **Impact:** Control account reconciliation now correctly finds GL balances
- **Files affected:** `scripts/summarize_ledgers.py`

### Lesson: Always Use Centralized Account Code Constants
- Account codes should be defined in one place to prevent inconsistencies
- Recommend creating `scripts/utils/constants.py` for all account codes

## Input File Population from Reference Files (2026-03-13)

### Script: scripts/create_input_files.py
- Reads reference General Ledger and General Journal from `Exisitng Accounting Workflow _ reference files/`
- Populates input files: cash_receipts_journal.xlsx, cash_payments_journal.xlsx, general_journal.xlsx
- Handles 3 matching patterns for counterparty identification:
  1. **Single counterparty**: Immediate neighbor with exact matching amount
  2. **Multiple counterparties**: Grouped between cash entries, total amounts match
  3. **Monthly summaries**: Description-based matching (e.g., "140ml" in both cash and revenue entries)

### GL Transaction Matching Logic
- Cash at Bank (10100) DEBITED = Cash Receipt (money in)
- Cash at Bank (10100) CREDITED = Cash Payment (money out)
- Counterparties are typically immediate neighbors (before/after cash entry)
- Some monthly summaries have counterparties BEFORE the cash entry with different descriptions
- Use ALL cash entries as boundaries (not just debits or credits) when searching for counterparties

### Important: Input vs Output Separation
- Input files (`data/input/`) should remain as source data
- Processing scripts should ONLY write to output files (`data/output/`)
- The `batch_process_all_months.py` had a bug where it wrote to input files
- This overwrites source data and breaks audit trail

### Results Verified
- Cash Receipts: 207 entries, 1,732,445,359 MMK
- Cash Payments: 129 entries, 1,469,287,817 MMK
- General Journal: 176 adjustment entries
- Difference from reference GL: 0.00 MMK (all transactions captured)
