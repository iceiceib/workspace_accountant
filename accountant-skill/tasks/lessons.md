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
- **Applies to:** Module 4 reading Module 3 output, Module 5 reading Module 4 output

### Grand Total Row Filtering
- **Issue:** Grand-total rows in output sheets have non-numeric Dr/Cr codes
- **Solution:** Filter with `if not _norm_code(row['Dr Code']).isdigit(): continue`
- **Applies to:** Any module reading another module's output

## Module-Specific Lessons

### Module 3 (Bank Reconciliation)
- Bank statement and cash ledger use opposite debit/credit conventions
- The script handles this automatically - don't manually adjust data

### Module 4 (Adjustments)
- Bank reconciliation adjusting entries are auto-imported from Module 3 output
- Depreciation is calculated per asset, then grouped by category for journal entries

### Module 5 (Trial Balance)
- Adjusting entries from Module 4 are read from "All Entries" sheet
- Filter out TOTALS row by checking Dr/Cr codes are numeric

### Module 6 (Financial Statements)
- Reads Adjusted TB from Module 5
- Net Profit flows from Income Statement to Balance Sheet Equity
- Cash Flow uses indirect method starting from Net Profit
