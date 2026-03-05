"""
Chart of Accounts Mapper — Lookup and classify accounts.
"""
import pandas as pd
from pathlib import Path


# Default account classification based on 5-digit code ranges
# Matches K&K Finance Chart of Accounts structure
DEFAULT_CLASSIFICATIONS = {
    # Assets (10000-19999)
    (10000, 10999): {'type': 'Asset', 'sub_type': 'Current Asset', 'normal': 'Debit', 'group': 'Cash & Equivalents'},
    (11000, 11999): {'type': 'Asset', 'sub_type': 'Current Asset', 'normal': 'Debit', 'group': 'Accounts Receivable'},
    (12000, 12999): {'type': 'Asset', 'sub_type': 'Current Asset', 'normal': 'Debit', 'group': 'Inventory'},
    (13000, 13999): {'type': 'Asset', 'sub_type': 'Current Asset', 'normal': 'Debit', 'group': 'Advanced Payments & Prepayments'},
    (14000, 14999): {'type': 'Asset', 'sub_type': 'Current Asset', 'normal': 'Debit', 'group': 'Deferred Expenses'},
    (15000, 15999): {'type': 'Asset', 'sub_type': 'Non-Current Asset', 'normal': 'Debit', 'group': 'Property, Plant & Equipment'},
    (15100, 15199): {'type': 'Asset', 'sub_type': 'Non-Current Asset', 'normal': 'Debit', 'group': 'Buildings & Structures'},
    (15200, 15299): {'type': 'Asset', 'sub_type': 'Non-Current Asset', 'normal': 'Debit', 'group': 'Machinery & Equipment'},
    (15300, 15399): {'type': 'Asset', 'sub_type': 'Non-Current Asset', 'normal': 'Debit', 'group': 'Office & Facility Equipment'},
    (15400, 15499): {'type': 'Asset', 'sub_type': 'Non-Current Asset', 'normal': 'Debit', 'group': 'Electrical & Utility Systems'},
    (15500, 15599): {'type': 'Asset', 'sub_type': 'Non-Current Asset', 'normal': 'Debit', 'group': 'Construction in Progress'},
    (15600, 15699): {'type': 'Asset', 'sub_type': 'Non-Current Asset', 'normal': 'Debit', 'group': 'Motor Vehicles'},
    # Accumulated Depreciation (Contra-Asset)
    (15110, 15119): {'type': 'Asset', 'sub_type': 'Non-Current Asset', 'normal': 'Credit', 'group': 'Accumulated Depreciation - Buildings'},
    (15210, 15219): {'type': 'Asset', 'sub_type': 'Non-Current Asset', 'normal': 'Credit', 'group': 'Accumulated Depreciation - Machinery'},
    (15310, 15319): {'type': 'Asset', 'sub_type': 'Non-Current Asset', 'normal': 'Credit', 'group': 'Accumulated Depreciation - Office Equipment'},
    (15410, 15419): {'type': 'Asset', 'sub_type': 'Non-Current Asset', 'normal': 'Credit', 'group': 'Accumulated Depreciation - Electrical'},
    (15510, 15519): {'type': 'Asset', 'sub_type': 'Non-Current Asset', 'normal': 'Credit', 'group': 'Accumulated Depreciation - Vehicles'},
    (17000, 19999): {'type': 'Asset', 'sub_type': 'Non-Current Asset', 'normal': 'Debit', 'group': 'Intangible & Other Non-Current Assets'},

    # Liabilities (20000-29999)
    (20000, 20999): {'type': 'Liability', 'sub_type': 'Current Liability', 'normal': 'Credit', 'group': 'Accounts Payable'},
    (21000, 21999): {'type': 'Liability', 'sub_type': 'Current Liability', 'normal': 'Credit', 'group': 'Short-term Loans'},
    (22000, 22999): {'type': 'Liability', 'sub_type': 'Current Liability', 'normal': 'Credit', 'group': 'Accrued Expenses'},
    (25000, 25999): {'type': 'Liability', 'sub_type': 'Non-Current Liability', 'normal': 'Credit', 'group': 'Long-term Bank Loans'},
    (26000, 29999): {'type': 'Liability', 'sub_type': 'Non-Current Liability', 'normal': 'Credit', 'group': 'Other Non-Current Liabilities'},

    # Equity (30000-39999)
    (30000, 30999): {'type': 'Equity', 'sub_type': 'Equity', 'normal': 'Credit', 'group': 'Paid-up Capital'},
    (31000, 31999): {'type': 'Equity', 'sub_type': 'Equity', 'normal': 'Credit', 'group': 'Share Capital'},
    (32000, 32999): {'type': 'Equity', 'sub_type': 'Equity', 'normal': 'Credit', 'group': 'Retained Earnings'},
    (33000, 39999): {'type': 'Equity', 'sub_type': 'Equity', 'normal': 'Credit', 'group': 'Other Equity'},

    # Revenue (40000-49999)
    (40000, 40999): {'type': 'Revenue', 'sub_type': 'Operating Revenue', 'normal': 'Credit', 'group': 'Sales Revenue'},
    (41000, 41999): {'type': 'Revenue', 'sub_type': 'Non-Operating Revenue', 'normal': 'Credit', 'group': 'Other Income'},

    # Cost of Goods Sold (50000-52999)
    (50000, 50099): {'type': 'Expense', 'sub_type': 'COGS', 'normal': 'Debit', 'group': 'Inventory - Raw Materials'},
    (50100, 50199): {'type': 'Expense', 'sub_type': 'COGS', 'normal': 'Debit', 'group': 'Inventory - Packaging'},
    (50200, 50299): {'type': 'Expense', 'sub_type': 'COGS', 'normal': 'Debit', 'group': 'Inventory - Finished Goods'},
    (53000, 53999): {'type': 'Expense', 'sub_type': 'COGS', 'normal': 'Debit', 'group': 'Production Costs'},

    # Operating Expenses - SG&A (60000-69999)
    (60000, 60999): {'type': 'Expense', 'sub_type': 'Operating Expense', 'normal': 'Debit', 'group': 'Marketing & Advertising'},
    (61000, 61999): {'type': 'Expense', 'sub_type': 'Operating Expense', 'normal': 'Debit', 'group': 'Office Salaries'},
    (62000, 62999): {'type': 'Expense', 'sub_type': 'Operating Expense', 'normal': 'Debit', 'group': 'Employee Benefits'},
    (63000, 63999): {'type': 'Expense', 'sub_type': 'Operating Expense', 'normal': 'Debit', 'group': 'Utilities'},
    (64000, 64999): {'type': 'Expense', 'sub_type': 'Operating Expense', 'normal': 'Debit', 'group': 'Transportation & Distribution'},
    (65000, 65999): {'type': 'Expense', 'sub_type': 'Operating Expense', 'normal': 'Debit', 'group': 'Facility & Office Supplies'},
    (66000, 66999): {'type': 'Expense', 'sub_type': 'Operating Expense', 'normal': 'Debit', 'group': 'Depreciation - SG&A'},
    (67000, 67999): {'type': 'Expense', 'sub_type': 'Operating Expense', 'normal': 'Debit', 'group': 'Inventory Write-offs'},
    (68000, 68999): {'type': 'Expense', 'sub_type': 'Operating Expense', 'normal': 'Debit', 'group': 'Other Operating Expenses'},
    (69000, 69999): {'type': 'Expense', 'sub_type': 'Operating Expense', 'normal': 'Debit', 'group': 'Management Compensation'},

    # Other Income/Expense (70000-79999)
    (70000, 70999): {'type': 'Revenue', 'sub_type': 'Non-Operating Revenue', 'normal': 'Credit', 'group': 'Interest Income'},
    (71000, 79999): {'type': 'Expense', 'sub_type': 'Non-Operating Expense', 'normal': 'Debit', 'group': 'Other Non-Operating Items'},
}

# Contra accounts have opposite normal balance (5-digit codes)
CONTRA_ACCOUNTS = {
    12300,  # Inventory Adjustments
    15110, 15111,  # Accumulated Depreciation - Buildings
    15210, 15211,  # Accumulated Depreciation - Machinery
    15310, 15311,  # Accumulated Depreciation - Office Equipment
    15410, 15411,  # Accumulated Depreciation - Electrical
    15510, 15511,  # Accumulated Depreciation - Vehicles
    30200,  # Owner's Drawings
}


class COAMapper:
    """Chart of Accounts lookup and classification."""
    
    def __init__(self, coa_filepath=None):
        """
        Initialize with optional COA file.
        Falls back to default classifications if no file provided.
        """
        self.coa_df = None
        self.coa_dict = {}
        
        if coa_filepath:
            self.load_coa(coa_filepath)
    
    def load_coa(self, filepath):
        """Load COA from .xlsx file."""
        filepath = Path(filepath)
        if not filepath.exists():
            print(f"Warning: COA file not found: {filepath}. Using defaults.")
            return
        
        try:
            df = pd.read_excel(filepath)
            # Normalize column names
            df.columns = [str(c).strip().lower() for c in df.columns]
            
            # Try to find account code column
            code_col = None
            for candidate in ['account code', 'code', 'acct code', 'no.', 'account_code']:
                if candidate in df.columns:
                    code_col = candidate
                    break
            
            if code_col is None:
                print(f"Warning: Cannot find account code column in COA file. Using defaults.")
                return
            
            name_col = None
            for candidate in ['account name', 'name', 'description', 'account']:
                if candidate in df.columns:
                    name_col = candidate
                    break
            
            self.coa_df = df
            for _, row in df.iterrows():
                code = int(row[code_col]) if pd.notna(row[code_col]) else None
                if code is None:
                    continue
                entry = {'code': code}
                if name_col:
                    entry['name'] = str(row[name_col]) if pd.notna(row[name_col]) else ''
                for col in ['type', 'sub-type', 'sub_type', 'subtype', 'normal balance', 'normal_balance', 'status']:
                    if col in df.columns and pd.notna(row[col]):
                        key = col.replace('-', '_').replace(' ', '_')
                        entry[key] = str(row[col])
                self.coa_dict[code] = entry
        except Exception as e:
            print(f"Warning: Error loading COA: {e}. Using defaults.")
    
    def get_account(self, code):
        """
        Get account info by code.
        Returns dict with: code, name, type, sub_type, normal_balance, group
        """
        try:
            code = int(code)
        except (ValueError, TypeError):
            return None
        
        # Check loaded COA first
        if code in self.coa_dict:
            info = self.coa_dict[code]
            result = {
                'code': code,
                'name': info.get('name', f'Account {code}'),
                'type': info.get('type', ''),
                'sub_type': info.get('sub_type', info.get('subtype', '')),
                'normal_balance': info.get('normal_balance', ''),
                'group': ''
            }
            # Fill from defaults if missing
            if not result['type'] or not result['normal_balance']:
                default = self._get_default(code)
                if default:
                    if not result['type']:
                        result['type'] = default['type']
                    if not result['sub_type']:
                        result['sub_type'] = default['sub_type']
                    if not result['normal_balance']:
                        result['normal_balance'] = default['normal']
                    result['group'] = default['group']
            return result
        
        # Fall back to defaults
        default = self._get_default(code)
        if default:
            return {
                'code': code,
                'name': f'Account {code}',
                'type': default['type'],
                'sub_type': default['sub_type'],
                'normal_balance': default['normal'],
                'group': default['group']
            }
        return None
    
    def _get_default(self, code):
        """Get default classification from code range."""
        # Handle contra accounts
        if code in CONTRA_ACCOUNTS:
            base = self._get_range_default(code)
            if base:
                base = base.copy()
                base['normal'] = 'Credit' if base['normal'] == 'Debit' else 'Debit'
            return base
        return self._get_range_default(code)
    
    def _get_range_default(self, code):
        for (lo, hi), info in DEFAULT_CLASSIFICATIONS.items():
            if lo <= code <= hi:
                return info
        return None
    
    def get_normal_balance(self, code):
        """Get the normal balance side for an account code."""
        info = self.get_account(code)
        return info['normal_balance'].lower() if info else None
    
    def is_debit_normal(self, code):
        return self.get_normal_balance(code) == 'debit'
    
    def is_credit_normal(self, code):
        return self.get_normal_balance(code) == 'credit'
    
    def get_type(self, code):
        info = self.get_account(code)
        return info['type'] if info else None
    
    def is_income_statement_account(self, code):
        """Returns True for Revenue and Expense accounts (temporary accounts)."""
        t = self.get_type(code)
        return t in ('Revenue', 'Expense')
    
    def is_balance_sheet_account(self, code):
        """Returns True for Asset, Liability, Equity accounts (permanent accounts)."""
        t = self.get_type(code)
        return t in ('Asset', 'Liability', 'Equity')
    
    def validate_code(self, code):
        """Check if an account code is valid (exists in COA or falls in known range)."""
        return self.get_account(code) is not None
    
    def classify_for_financial_statements(self, code):
        """
        Classify an account for financial statement placement.

        Returns: dict with 'statement' (IS/BS), 'section', 'subsection', 'sign'
        """
        info = self.get_account(code)
        if not info:
            return None

        code = int(code)
        t = info['type']

        if t == 'Revenue':
            # 5-digit codes: 41000-41999 = Other Income, 70000-70999 = Interest Income
            section = 'Other Income' if (41000 <= code <= 41999 or 70000 <= code <= 70999) else 'Revenue'
            sign = -1 if code in CONTRA_ACCOUNTS else 1
            return {'statement': 'IS', 'section': section, 'sign': sign}

        if t == 'Expense':
            # 5-digit codes for COGS: 50000-50299 (Inventory), 53000-53999 (Production)
            if 50000 <= code <= 50299 or 53000 <= code <= 53999:
                section = 'COGS'
            # 5-digit codes for Operating Expenses: 60000-69999 (SG&A)
            elif 60000 <= code <= 69999:
                section = 'Operating Expenses'
            # Other expenses: 71000-79999 (Non-Operating)
            elif 71000 <= code <= 79999:
                section = 'Non-Operating Expenses'
            else:
                section = 'COGS'  # Default for 50xxx codes
            return {'statement': 'IS', 'section': section, 'sign': -1}

        if t == 'Asset':
            # 5-digit codes: 15000-19999 = Non-Current, 10000-14999 = Current
            section = 'Non-Current Assets' if 15000 <= code <= 19999 else 'Current Assets'
            sign = -1 if code in CONTRA_ACCOUNTS else 1
            return {'statement': 'BS', 'section': section, 'sign': sign}

        if t == 'Liability':
            # 5-digit codes: 25000-29999 = Non-Current, 20000-24999 = Current
            section = 'Non-Current Liabilities' if 25000 <= code <= 29999 else 'Current Liabilities'
            return {'statement': 'BS', 'section': section, 'sign': 1}

        if t == 'Equity':
            sign = -1 if code in CONTRA_ACCOUNTS else 1
            return {'statement': 'BS', 'section': 'Equity', 'sign': sign}

        return None
