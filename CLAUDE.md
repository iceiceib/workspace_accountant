# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Full-cycle accounting automation for K&K Business (Shwe Mandalay Cafe). All inputs and outputs are `.xlsx` Excel files processed with Python (`openpyxl` + `pandas`). No Google Sheets, no databases.

**Tech stack:** Python 3.12, openpyxl, pandas, numpy

**Install:** `pip install openpyxl pandas numpy`

---

## Module Run Commands

All scripts run from the `accountant-skill/` directory. Replace `data/Jan2026` and dates with the actual period.

**Module 1 — Summarize Journals:**
```
python scripts/summarize_journals.py data/Jan2026 2026-01-01 2026-01-31 \
  data/Jan2026/books_of_prime_entry_Jan2026.xlsx \
  data/Jan2026/chart_of_accounts.xlsx \
  data/Jan2026/profit_cost_centers.xlsx
```

**Module 2 — Summarize Ledgers:**
```
python scripts/summarize_ledgers.py data/Jan2026 2026-01-01 2026-01-31 \
  data/Jan2026/ledger_summary_Jan2026.xlsx \
  data/Jan2026/chart_of_accounts.xlsx
```

**Module 3 — Bank Reconciliation:**
```
python scripts/reconcile_bank.py data/Jan2026 2026-01-01 2026-01-31 \
  data/Jan2026/bank_reconciliation_Jan2026.xlsx
```

**Module 4 — Journal Adjustments:**
```
python scripts/journal_adjustments.py data/Jan2026 2026-01-01 2026-01-31 \
  data/Jan2026/adjusting_entries_Jan2026.xlsx
```

**Module 5 — Trial Balance:**
```
python scripts/generate_trial_balance.py data/Jan2026 2026-01-01 2026-01-31 \
  data/Jan2026/trial_balance_Jan2026.xlsx
```

**Module 6 — Financial Statements (pending):**
```
python scripts/generate_financials.py data/Jan2026 2026-01-01 2026-01-31 \
  data/Jan2026/financial_statements_Jan2026.xlsx
```

**Generate test data (once only):**
```
python scripts/create_test_data.py data/Jan2026
python scripts/create_bank_statement.py data/Jan2026
```

---

## Accounting Cycle Flow

Modules must run in order — each feeds the next:

```
Module 1 (Journals) → Module 2 (Ledgers) → Module 5 (Unadjusted TB)
→ Module 3 (Bank Recon) → Module 4 (Adjustments) → Module 5 (Adjusted TB)
→ Module 6 (Financial Statements) → Module 7 (Validation)
```

---

## Architecture

### `scripts/utils/` — Shared Utilities

All module scripts import from `scripts/utils/`:

- **`coa_mapper.py` — `COAMapper`**: Loads `chart_of_accounts.xlsx`. Falls back to built-in defaults. Key methods: `get_type(code)`, `is_debit_normal(code)`, `classify_for_financial_statements(code)`. Account code ranges: 1xxx=Assets, 2xxx=Liabilities, 3xxx=Equity, 4xxx=Revenue, 5xxx=Expenses.

- **`pc_cc_mapper.py` — `PCCCMapper`**: Loads `profit_cost_centers.xlsx` (Sheet 1: Profit Centers, Sheet 2: Cost Centers). PC required on accounts 4000–5999; CC required on 5000–5999. Critical NaN fix: always use `_clean(val)` with `math.isnan()` — never `val or ''` because `float('nan')` is truthy.

- **`excel_reader.py`**: `read_xlsx(filepath, required_columns, optional_columns, date_columns)` → `dict(data=df, error, warnings)`. `filter_by_period(df, date_col, start, end)`.

- **`excel_writer.py`**: Enforces formatting standards. Key functions: `write_title`, `write_header_row`, `write_data_row`, `write_total_row`, `auto_fit_columns`, `save_workbook`.

- **`double_entry.py`**: `validate_journal_balance(df, debit_col, credit_col, group_col)` → checks Dr = Cr.

### Data Flow Between Modules

- Module 3 → Module 4: `bank_reconciliation_Jan2026.xlsx` "Adjusting Entries" sheet (header detection needed — title block precedes column headers; scan for row containing 'Date').
- Module 4 → Module 5: `adjusting_entries_Jan2026.xlsx` "All Entries" sheet (filter out TOTALS row by checking Dr/Cr codes are numeric).
- Module 5 → Module 6: `trial_balance_Jan2026.xlsx` "Adjusted TB" sheet.

---

## Critical Rules (enforce in all modules)

1. **Double-entry**: Every entry must have Debit = Credit. Stop and report if not.
2. **COA validation**: Every account code must exist in `chart_of_accounts.xlsx`. Flag unknowns.
3. **Period filtering**: Filter data by date range before any processing.
4. **Excel formulas**: Use `=SUM`, `=SUMIF` etc. in output files — not hardcoded Python values.
5. **Idempotent**: Re-running for same period produces same output. Never append duplicates.
6. **Myanmar text**: Handle Myanmar Unicode in descriptions without errors.
7. **Account code normalization**: Excel reads numeric codes as floats. Always normalize: `str(int(float(str(val).strip())))`.

---

## Output Formatting Standards

| Element | Standard |
|---------|----------|
| Header row | Bold, `#1F4E79` background, white text, centered |
| Numbers | `#,##0` format; `0.0%` for percentages |
| Negatives | Red font, parentheses `(#,##0)` |
| Borders | Thin on data; thick bottom on totals |
| Column widths | Auto-fit, min 12 chars |
| Tab colors | Green `#00B050`=Dashboard, Blue `#4472C4`=Detail, Orange `#70AD47`=PC/CC, Red `#FF0000`=Exceptions |
| Freeze panes | Header row + first column on all sheets |

---

## Current Status

- **Complete:** Modules 1–5
- **Next:** Module 6 (`scripts/generate_financials.py`) — reads `trial_balance_Jan2026.xlsx` (Adjusted TB sheet), outputs Dashboard, Income Statement, Balance Sheet, Exceptions. See `references/financial-report-formats.md`.
- **After Module 6:** Module 7 (`validate_accounting.py`) full-cycle integrity checks, then update USERGUIDE.md.

---

## Reference Files

Before implementing a module, read the corresponding reference:
- `references/journal-rules.md` → Module 1
- `references/ledger-posting-rules.md` → Module 2
- `references/reconciliation-rules.md` → Module 3
- `references/adjustment-rules.md` → Module 4
- `references/chart-of-accounts.md` → Module 5
- `references/financial-report-formats.md` → Module 6
- `references/data-schemas.md` → input column schemas for all modules

## Windows Console Note

Do NOT use Unicode box/check characters (`─`, `✓`, `✗`) in `print()` statements — Windows cp1252 console will fail. Use plain ASCII instead.

---

## Workflow Orchestration

### 1. Plan Node Default
- Enter plan mode for ANY non-trivial task (3+ steps or architectural decisions)
- If something goes sideways, STOP and re-plan immediately — don't keep pushing
- Use plan mode for verification steps, not just building
- Write detailed specs upfront to reduce ambiguity

### 2. Subagent Strategy
- Use subagents liberally to keep main context window clean
- Offload research, exploration, and parallel analysis to subagents
- For complex problems, throw more compute at it via subagents
- One task per subagent for focused execution

### 3. Self-Improvement Loop
- After ANY correction from the user: update `tasks/lessons.md` with the pattern
- Write rules for yourself that prevent the same mistake
- Ruthlessly iterate on these lessons until mistake rate drops
- Review lessons at session start for relevant project

### 4. Verification Before Done
- Never mark a task complete without proving it works
- Diff behavior between main and your changes when relevant
- Ask yourself: "Would a staff engineer approve this?"
- Run tests, check logs, demonstrate correctness

### 5. Demand Elegance (Balanced)
- For non-trivial changes: pause and ask "is there a more elegant way?"
- If a fix feels hacky: "Knowing everything I know now, implement the elegant solution"
- Skip this for simple, obvious fixes — don't over-engineer
- Challenge your own work before presenting it

### 6. Autonomous Bug Fixing
- When given a bug report: just fix it. Don't ask for hand-holding
- Point at logs, errors, failing tests — then resolve them
- Zero context switching required from the user
- Go fix failing CI tests without being told how

---

## Task Management

- **Plan First:** Write plan to `tasks/todo.md` with checkable items
- **Verify Plan:** Check in before starting implementation
- **Track Progress:** Mark items complete as you go
- **Explain Changes:** High-level summary at each step
- **Document Results:** Add review section to `tasks/todo.md`
- **Capture Lessons:** Update `tasks/lessons.md` after corrections

---

## Core Principles

- **Simplicity First:** Make every change as simple as possible. Impact minimal code.
- **No Laziness:** Find root causes. No temporary fixes. Senior developer standards.
- **Minimal Impact:** Changes should only touch what's necessary. Avoid introducing bugs.
