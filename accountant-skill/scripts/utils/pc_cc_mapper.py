"""
Profit Center & Cost Center Mapper
Validates PC/CC codes and classifies accounts for segment reporting.

Loads definitions from profit_cost_centers.xlsx (two sheets: Profit Centers, Cost Centers).
Falls back to built-in defaults when no file is provided.
"""
import math
from pathlib import Path


def _clean(val):
    """Return val as uppercase stripped string; treat None/NaN/empty as ''."""
    if val is None:
        return ''
    try:
        if math.isnan(float(val)):
            return ''
    except (ValueError, TypeError):
        pass
    s = str(val).strip().upper()
    return '' if s in ('NAN', 'NONE', 'N/A', '-') else s


# ── Built-in defaults (used when no file is loaded) ───────────────────────────

DEFAULT_PROFIT_CENTERS = {
    'PC01': 'Soft Drink',
    'PC02': 'Drinking Water',
    'PC99': 'Shared / Corporate',
}

DEFAULT_COST_CENTERS = {
    'CC101': {'name': 'Soft Drink Production',       'default_pc': 'PC01'},
    'CC102': {'name': 'Drinking Water Production',   'default_pc': 'PC02'},
    'CC103': {'name': 'Water Treatment',             'default_pc': 'PC02'},
    'CC104': {'name': 'Preform Production',          'default_pc': None},
    'CC105': {'name': 'Filling & Packaging',         'default_pc': None},
    'CC201': {'name': 'Factory Utilities',           'default_pc': 'PC99'},
    'CC202': {'name': 'Maintenance',                 'default_pc': 'PC99'},
    'CC301': {'name': 'Sales & Marketing',           'default_pc': 'PC99'},
    'CC302': {'name': 'Administration',              'default_pc': 'PC99'},
}

# Account ranges that REQUIRE a Profit Center tag
PC_REQUIRED_RANGES = [
    (4000, 4999),   # Revenue
    (5000, 5999),   # All expenses (COGS + Opex + Non-Op)
]

# Account ranges that REQUIRE a Cost Center tag
CC_REQUIRED_RANGES = [
    (5000, 5999),   # All expenses must specify where the cost was incurred
]

# Account type classification for P&L grouping
ACCOUNT_SEGMENTS = {
    'revenue':         (4000, 4199),
    'revenue_contra':  (4200, 4299),
    'cogs':            (5000, 5099),
    'opex':            (5100, 5899),
    'nonop':           (5900, 5999),
}


# ── Mapper class ──────────────────────────────────────────────────────────────

class PCCCMapper:
    """Profit Center / Cost Center validator and lookup.

    Usage:
        pcc = PCCCMapper()                                  # use built-in defaults
        pcc = PCCCMapper('data/Jan2026/profit_cost_centers.xlsx')  # load from file
    """

    def __init__(self, filepath=None):
        # Start from built-in defaults
        self.profit_centers = dict(DEFAULT_PROFIT_CENTERS)
        self.cost_centers   = {k: dict(v) for k, v in DEFAULT_COST_CENTERS.items()}

        if filepath:
            self.load_pcc(filepath)

    # ── File loading ──────────────────────────────────────────────────────────

    def load_pcc(self, filepath):
        """Load Profit Centers and Cost Centers from an .xlsx file.

        Expected sheets:
          'Profit Centers'  — columns: PC Code | Name
          'Cost Centers'    — columns: CC Code | Name | Default PC
        """
        import pandas as pd

        filepath = Path(filepath)
        if not filepath.exists():
            print(f"Warning: PC/CC file not found: {filepath}. Using defaults.")
            return

        try:
            xl = pd.ExcelFile(filepath)

            # ── Profit Centers sheet ──────────────────────────────────────
            pc_sheet = self._find_sheet(xl.sheet_names, ['profit centers', 'profit center', 'pc'])
            if pc_sheet:
                df = xl.parse(pc_sheet)
                df.columns = [str(c).strip().lower() for c in df.columns]

                code_col = self._find_col(df.columns, ['pc code', 'code', 'pc'])
                name_col = self._find_col(df.columns, ['name', 'description'])

                if code_col:
                    self.profit_centers = {}
                    for _, row in df.iterrows():
                        code = _clean(row[code_col])
                        name = str(row[name_col]).strip() if name_col and pd.notna(row.get(name_col)) else code
                        if code:
                            self.profit_centers[code] = name
            else:
                print("Warning: 'Profit Centers' sheet not found. Using defaults.")

            # ── Cost Centers sheet ────────────────────────────────────────
            cc_sheet = self._find_sheet(xl.sheet_names, ['cost centers', 'cost center', 'cc'])
            if cc_sheet:
                df = xl.parse(cc_sheet)
                df.columns = [str(c).strip().lower() for c in df.columns]

                code_col    = self._find_col(df.columns, ['cc code', 'code', 'cc'])
                name_col    = self._find_col(df.columns, ['name', 'description'])
                def_pc_col  = self._find_col(df.columns, ['default pc', 'default profit center', 'pc'])

                if code_col:
                    self.cost_centers = {}
                    for _, row in df.iterrows():
                        code = _clean(row[code_col])
                        if not code:
                            continue
                        name = str(row[name_col]).strip() if name_col and pd.notna(row.get(name_col)) else code
                        def_pc = _clean(row[def_pc_col]) if def_pc_col and pd.notna(row.get(def_pc_col)) else None
                        self.cost_centers[code] = {'name': name, 'default_pc': def_pc or None}
            else:
                print("Warning: 'Cost Centers' sheet not found. Using defaults.")

        except Exception as e:
            print(f"Warning: Error loading PC/CC file: {e}. Using defaults.")

    @staticmethod
    def _find_sheet(sheet_names, candidates):
        """Return the first sheet name that matches any candidate (case-insensitive)."""
        lower_names = [s.lower() for s in sheet_names]
        for c in candidates:
            if c in lower_names:
                return sheet_names[lower_names.index(c)]
        return None

    @staticmethod
    def _find_col(columns, candidates):
        """Return the first column name that matches any candidate."""
        for c in candidates:
            if c in columns:
                return c
        return None

    # ── Validation ────────────────────────────────────────────────────────────

    def validate_pc(self, code):
        return _clean(code) in self.profit_centers

    def validate_cc(self, code):
        return _clean(code) in self.cost_centers

    def get_pc_name(self, code):
        return self.profit_centers.get(_clean(code), f'Unknown PC ({code})')

    def get_cc_name(self, code):
        info = self.cost_centers.get(_clean(code))
        return info['name'] if info else f'Unknown CC ({code})'

    def get_cc_default_pc(self, cc_code):
        info = self.cost_centers.get(_clean(cc_code))
        return info['default_pc'] if info else None

    # ── Account classification ────────────────────────────────────────────────

    def classify_account(self, account_code):
        """
        Classify an account code for P&L segmentation.
        Returns: 'revenue' | 'revenue_contra' | 'cogs' | 'opex' | 'nonop' | 'balance_sheet'
        """
        try:
            code = int(account_code)
        except (ValueError, TypeError):
            return 'unknown'
        for segment, (lo, hi) in ACCOUNT_SEGMENTS.items():
            if lo <= code <= hi:
                return segment
        return 'balance_sheet'

    def is_pc_required(self, account_code):
        try:
            code = int(account_code)
        except (ValueError, TypeError):
            return False
        return any(lo <= code <= hi for lo, hi in PC_REQUIRED_RANGES)

    def is_cc_required(self, account_code):
        try:
            code = int(account_code)
        except (ValueError, TypeError):
            return False
        return any(lo <= code <= hi for lo, hi in CC_REQUIRED_RANGES)

    # ── Journal validation ────────────────────────────────────────────────────

    def validate_journal_rows(self, df, journal_name):
        """
        Validate PC/CC tags across a journal DataFrame.
        Returns list of exception dicts.
        """
        exceptions = []
        has_pc = 'Profit Center' in df.columns
        has_cc = 'Cost Center' in df.columns

        if not has_pc:
            exceptions.append({
                'journal': journal_name,
                'row': 'ALL',
                'issue': "Missing 'Profit Center' column — add this column to the file"
            })
            return exceptions

        valid_pcs = ', '.join(self.profit_centers)
        valid_ccs = ', '.join(self.cost_centers)

        for idx, row in df.iterrows():
            excel_row = idx + 2
            pc = _clean(row.get('Profit Center', ''))
            cc = _clean(row.get('Cost Center', '')) if has_cc else ''

            dr_acct = row.get('Debit Account', '')
            cr_acct = row.get('Credit Account', '')

            if pc and not self.validate_pc(pc):
                exceptions.append({
                    'journal': journal_name, 'row': excel_row,
                    'issue': f"Unknown Profit Center '{pc}'. Valid: {valid_pcs}"
                })

            if cc and not self.validate_cc(cc):
                exceptions.append({
                    'journal': journal_name, 'row': excel_row,
                    'issue': f"Unknown Cost Center '{cc}'. Valid: {valid_ccs}"
                })

            try:
                dr_code = int(float(str(dr_acct)))
                if self.is_pc_required(dr_code) and not pc:
                    exceptions.append({
                        'journal': journal_name, 'row': excel_row,
                        'issue': f"Debit account {dr_code} (expense/revenue) requires a Profit Center"
                    })
                if self.is_cc_required(dr_code) and not cc:
                    exceptions.append({
                        'journal': journal_name, 'row': excel_row,
                        'issue': f"Debit account {dr_code} (expense) requires a Cost Center"
                    })
            except (ValueError, TypeError):
                pass

            try:
                cr_code = int(float(str(cr_acct)))
                if self.is_pc_required(cr_code) and not pc:
                    exceptions.append({
                        'journal': journal_name, 'row': excel_row,
                        'issue': f"Credit account {cr_code} (revenue) requires a Profit Center"
                    })
            except (ValueError, TypeError):
                pass

        return exceptions

    # ── Summarization helpers ─────────────────────────────────────────────────

    def build_pc_summary(self, journal_dfs):
        """
        Build a profit center P&L summary from all journal DataFrames.

        Returns:
            pc_summary  — dict { pc_code: { revenue, cogs, opex, nonop } }
            cc_summary  — dict { cc_code: { debits, credits } }
        """
        import pandas as pd

        pc_summary = {pc: {'revenue': 0.0, 'cogs': 0.0, 'opex': 0.0, 'nonop': 0.0}
                      for pc in self.profit_centers}
        cc_summary = {cc: {'debits': 0.0, 'credits': 0.0}
                      for cc in self.cost_centers}

        for journal_name, df in journal_dfs.items():
            if df is None or len(df) == 0:
                continue
            if 'Profit Center' not in df.columns:
                continue

            if '_debit' not in df.columns:
                if 'Debit Amount' in df.columns and 'Credit Amount' in df.columns:
                    df = df.copy()
                    df['_debit']  = pd.to_numeric(df['Debit Amount'],  errors='coerce').fillna(0)
                    df['_credit'] = pd.to_numeric(df['Credit Amount'], errors='coerce').fillna(0)
                elif 'Amount' in df.columns:
                    df = df.copy()
                    df['_debit']  = pd.to_numeric(df['Amount'], errors='coerce').fillna(0)
                    df['_credit'] = pd.to_numeric(df['Amount'], errors='coerce').fillna(0)
                else:
                    continue

            for _, row in df.iterrows():
                pc = _clean(row.get('Profit Center', ''))
                cc = _clean(row.get('Cost Center',   ''))
                dr_amt = float(row.get('_debit',  0) or 0)
                cr_amt = float(row.get('_credit', 0) or 0)

                try:
                    dr_code = int(float(str(row.get('Debit Account', 0))))
                    dr_seg  = self.classify_account(dr_code)
                except (ValueError, TypeError):
                    dr_seg = 'unknown'

                try:
                    cr_code = int(float(str(row.get('Credit Account', 0))))
                    cr_seg  = self.classify_account(cr_code)
                except (ValueError, TypeError):
                    cr_seg = 'unknown'

                if pc in pc_summary:
                    if cr_seg == 'revenue':
                        pc_summary[pc]['revenue'] += cr_amt
                    if dr_seg == 'revenue_contra':
                        pc_summary[pc]['revenue'] -= dr_amt
                    for seg in ('cogs', 'opex', 'nonop'):
                        if dr_seg == seg:
                            pc_summary[pc][seg] += dr_amt

                if cc in cc_summary:
                    cc_summary[cc]['debits']  += dr_amt
                    cc_summary[cc]['credits'] += cr_amt

        return pc_summary, cc_summary
