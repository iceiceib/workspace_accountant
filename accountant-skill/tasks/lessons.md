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
- **Issue:** Excel reads numeric account codes as floats (e.g., `1020.0` instead of `1020`)
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
