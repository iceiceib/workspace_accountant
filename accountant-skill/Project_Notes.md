# Project Notes - K&K Business Accounting Automation

## Project Overview
- **Client:** K&K Business / Shwe Mandalay Cafe
- **Period:** Monthly accounting cycle automation
- **Tech Stack:** Python 3.12, openpyxl, pandas, numpy
- **I/O Format:** Excel .xlsx files only (no databases, no Google Sheets)

---

## Session Log

### 2026-02-25 - Current Session
**Goal:** Review project structure and prepare for Module 6 implementation

**Actions Taken:**
1. Read SKILL.md - skill configuration for accountant skill
2. Read USERGUIDE.md - comprehensive user guide (covers Modules 1-6)
3. Created tasks/ directory with:
   - tasks/todo.md - task tracking with checkable items
   - tasks/lessons.md - patterns and lessons learned
4. Created this Project_Notes.md file

**Module Status:**
| Module | Status | Script |
|--------|--------|--------|
| 1 - Summarize Journals | COMPLETE | scripts/summarize_journals.py |
| 2 - Summarize Ledgers | COMPLETE | scripts/summarize_ledgers.py |
| 3 - Bank Reconciliation | COMPLETE | scripts/reconcile_bank.py |
| 4 - Journal Adjustments | COMPLETE | scripts/journal_adjustments.py |
| 5 - Trial Balance | COMPLETE | scripts/generate_trial_balance.py |
| 6 - Financial Statements | PENDING | scripts/generate_financials.py |
| 7 - Full-Cycle Validation | PENDING | scripts/validate_accounting.py |

**Next Steps:**
1. Implement Module 6 (Financial Statements)
2. Implement Module 7 (Validation)
3. Update USERGUIDE.md to cover all modules

---

## Technical Architecture

### Accounting Cycle Flow
```
Journals (M1) → Ledgers (M2) → Unadjusted TB (M5)
→ Bank Recon (M3) → Adjustments (M4) → Adjusted TB (M5)
→ Financial Statements (M6) → Validation (M7)
```

### Shared Utilities (scripts/utils/)
- **coa_mapper.py** - COAMapper for chart of accounts
- **pc_cc_mapper.py** - PCCCMapper for profit/cost centers
- **excel_reader.py** - read_xlsx(), filter_by_period()
- **excel_writer.py** - write_title(), write_header_row(), formatting
- **double_entry.py** - validate_journal_balance()

### Critical Rules
1. Double-entry: Dr = Cr for every entry
2. COA validation: every account code must exist
3. Period filtering: filter by date range before processing
4. Excel formulas: use =SUM() not hardcoded values
5. Idempotent: re-running produces same output
6. Myanmar text: handle Unicode without errors
7. Account code normalization: Excel reads as floats

---

## Reference Files
- `references/journal-rules.md` → Module 1
- `references/ledger-posting-rules.md` → Module 2
- `references/reconciliation-rules.md` → Module 3
- `references/adjustment-rules.md` → Module 4
- `references/chart-of-accounts.md` → Module 5
- `references/financial-report-formats.md` → Module 6
- `references/data-schemas.md` → Input column schemas

---

## Pending User Request
- Update USERGUIDE.md to cover all modules (1-7) after completion
- Requested: 2026-02-23
