"""
Chart of Accounts Mapper — Lookup and classify accounts.
"""
import pandas as pd
from pathlib import Path


# Default account classification based on code ranges
DEFAULT_CLASSIFICATIONS = {
    (1000, 1099): {'type': 'Asset', 'sub_type': 'Current Asset', 'normal': 'Debit', 'group': 'Cash & Equivalents'},
    (1100, 1199): {'type': 'Asset', 'sub_type': 'Current Asset', 'normal': 'Debit', 'group': 'Receivables'},
    (1200, 1299): {'type': 'Asset', 'sub_type': 'Current Asset', 'normal': 'Debit', 'group': 'Inventory'},
    (1300, 1499): {'type': 'Asset', 'sub_type': 'Current Asset', 'normal': 'Debit', 'group': 'Prepaid & Other Current'},
    (1500, 1599): {'type': 'Asset', 'sub_type': 'Current Asset', 'normal': 'Debit', 'group': 'Other Current Assets'},
    (1600, 1699): {'type': 'Asset', 'sub_type': 'Non-Current Asset', 'normal': 'Debit', 'group': 'Fixed Assets'},
    (1700, 1999): {'type': 'Asset', 'sub_type': 'Non-Current Asset', 'normal': 'Debit', 'group': 'Intangible Assets'},
    (2000, 2099): {'type': 'Liability', 'sub_type': 'Current Liability', 'normal': 'Credit', 'group': 'Current Liabilities'},
    (2100, 2999): {'type': 'Liability', 'sub_type': 'Non-Current Liability', 'normal': 'Credit', 'group': 'Non-Current Liabilities'},
    (3000, 3999): {'type': 'Equity', 'sub_type': 'Equity', 'normal': 'Credit', 'group': 'Equity'},
    (4000, 4099): {'type': 'Revenue', 'sub_type': 'Operating Revenue', 'normal': 'Credit', 'group': 'Sales Revenue'},
    (4100, 4199): {'type': 'Revenue', 'sub_type': 'Non-Operating Revenue', 'normal': 'Credit', 'group': 'Other Income'},
    (4200, 4299): {'type': 'Revenue', 'sub_type': 'Revenue Contra', 'normal': 'Debit', 'group': 'Sales Returns & Discounts'},
    (5000, 5099): {'type': 'Expense', 'sub_type': 'COGS', 'normal': 'Debit', 'group': 'Cost of Goods Sold'},
    (5100, 5899): {'type': 'Expense', 'sub_type': 'Operating Expense', 'normal': 'Debit', 'group': 'Operating Expenses'},
    (5900, 5999): {'type': 'Expense', 'sub_type': 'Non-Operating Expense', 'normal': 'Debit', 'group': 'Non-Operating Expenses'},
}

# Contra accounts have opposite normal balance
CONTRA_ACCOUNTS = {1110, 1611, 1621, 1631, 1641, 1651, 3020, 4200, 4210}


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
            section = 'Other Income' if 4100 <= code <= 4199 else 'Revenue'
            sign = -1 if code in CONTRA_ACCOUNTS else 1
            return {'statement': 'IS', 'section': section, 'sign': sign}
        
        if t == 'Expense':
            if 5010 <= code <= 5050:
                section = 'COGS'
            elif 5100 <= code <= 5899:
                section = 'Operating Expenses'
            else:
                section = 'Non-Operating Expenses'
            return {'statement': 'IS', 'section': section, 'sign': -1}
        
        if t == 'Asset':
            section = 'Non-Current Assets' if 1600 <= code <= 1999 else 'Current Assets'
            sign = -1 if code in CONTRA_ACCOUNTS else 1
            return {'statement': 'BS', 'section': section, 'sign': sign}
        
        if t == 'Liability':
            section = 'Non-Current Liabilities' if 2100 <= code <= 2999 else 'Current Liabilities'
            return {'statement': 'BS', 'section': section, 'sign': 1}
        
        if t == 'Equity':
            sign = -1 if code in CONTRA_ACCOUNTS else 1
            return {'statement': 'BS', 'section': 'Equity', 'sign': sign}
        
        return None
