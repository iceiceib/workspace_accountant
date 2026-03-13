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
- [ ] Update USERGUIDE.md to cover all modules (1-7) - pending since 2026-02-23
- [ ] Fix batch_process_all_months.py to not overwrite input files

## Random thoughts
- Input files should remain as source data - processing scripts should only write to output files
- The batch_process_all_months.py still has code that writes to data/input/ - needs cleanup

## Links to check later
-
