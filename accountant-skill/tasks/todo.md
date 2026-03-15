## TODO
- [x] Add WIP (Work-in-Progress) accounts to COA structure
- [x] Update COA mapper (coa_mapper.py) for WIP account range 50300-50399
- [x] Create WIP flow documentation (references/wip-flow-guide.md)
- [x] Add WIP journal entries to general_journal.xlsx
- [x] Add WIP transactions to general_ledger.xlsx
- [x] Update USERGUIDE.md with WIP accounting section
- [x] Update data-schemas.md with WIP account code ranges
- [x] Update lessons.md with WIP implementation notes
- [x] Run full accounting cycle to verify WIP entries
- [x] Update adjustment-rules.md with WIP adjustment examples
- [x] Fix account code ranges in pc_cc_mapper.py (4-digit to 5-digit)
- [x] Fix control account codes in summarize_ledgers.py (4-digit to 5-digit)
- [x] Create script to populate input files from reference files (2026-03-13)
- [x] Update USERGUIDE.md to cover all modules (2026-03-14)
- [x] Fix batch_process_all_months.py to not overwrite input files (2026-03-15)
- [x] Populate ledger files from reference General_Ledger_edited.xlsx (2026-03-15)
- [x] Clean sample data from AR/AP ledgers (2026-03-15)

## Completed This Session (2026-03-15)
1. Fixed `batch_process_all_months.py` - removed unused `OUTPUT_INPUT` variable
2. Populated `general_ledger.xlsx` with 758 transactions from reference file (Feb-Oct 2025)
3. Derived `cash_ledger.xlsx` from GL with 242 transactions
4. Cleaned AR/AP ledgers to empty templates (headers only)

## Notes
- Reference GL uses cash-based system (no AR/AP accounts - all sales/purchases through Cash 10100)
- Inventory sub-ledgers (raw_materials, packaging) await separate data from user

## Links to check later
-
