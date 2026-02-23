"""
Double-Entry Validation Utility — Ensures accounting integrity.
"""
import pandas as pd


def validate_entry_balance(debit_amount, credit_amount, tolerance=0.01):
    """Check if a single entry balances (debit = credit)."""
    return abs(debit_amount - credit_amount) <= tolerance


def validate_journal_balance(df, debit_col='Debit Amount', credit_col='Credit Amount', 
                              group_col=None, tolerance=0.01):
    """
    Validate that journal entries balance.
    
    If group_col is provided, validates per-group (compound entries).
    Otherwise validates per-row.
    
    Returns: dict with 'balanced' (bool), 'total_debit', 'total_credit', 
             'difference', 'unbalanced_entries' (list)
    """
    # Handle single Amount column journals (sales, purchases, cash receipts/payments)
    if debit_col not in df.columns and credit_col not in df.columns and 'Amount' in df.columns:
        total = df['Amount'].sum()
        return {
            'balanced': True,
            'total_debit': total,
            'total_credit': total,
            'difference': 0,
            'unbalanced_entries': []
        }
    
    df = df.copy()
    df[debit_col] = pd.to_numeric(df[debit_col], errors='coerce').fillna(0)
    df[credit_col] = pd.to_numeric(df[credit_col], errors='coerce').fillna(0)
    
    unbalanced = []
    
    if group_col and group_col in df.columns:
        for name, group in df.groupby(group_col):
            dr_total = group[debit_col].sum()
            cr_total = group[credit_col].sum()
            if abs(dr_total - cr_total) > tolerance:
                unbalanced.append({
                    'reference': name,
                    'debit_total': dr_total,
                    'credit_total': cr_total,
                    'difference': dr_total - cr_total
                })
    else:
        for idx, row in df.iterrows():
            dr = row[debit_col]
            cr = row[credit_col]
            if abs(dr - cr) > tolerance and dr > 0 and cr > 0:
                unbalanced.append({
                    'row': idx + 2,  # +2 for Excel row (1-indexed + header)
                    'debit': dr,
                    'credit': cr,
                    'difference': dr - cr
                })
    
    total_debit = df[debit_col].sum()
    total_credit = df[credit_col].sum()
    
    return {
        'balanced': abs(total_debit - total_credit) <= tolerance and len(unbalanced) == 0,
        'total_debit': total_debit,
        'total_credit': total_credit,
        'difference': total_debit - total_credit,
        'unbalanced_entries': unbalanced
    }


def validate_trial_balance(accounts_df, debit_col='Debit', credit_col='Credit', tolerance=0.01):
    """
    Validate that a trial balance balances.
    
    Returns: dict with 'balanced', 'total_debit', 'total_credit', 'difference'
    """
    total_debit = pd.to_numeric(accounts_df[debit_col], errors='coerce').fillna(0).sum()
    total_credit = pd.to_numeric(accounts_df[credit_col], errors='coerce').fillna(0).sum()
    
    return {
        'balanced': abs(total_debit - total_credit) <= tolerance,
        'total_debit': total_debit,
        'total_credit': total_credit,
        'difference': total_debit - total_credit
    }


def validate_balance_sheet(total_assets, total_equity, total_liabilities, tolerance=0.01):
    """Validate the accounting equation: Assets = Equity + Liabilities."""
    equity_plus_liabilities = total_equity + total_liabilities
    difference = total_assets - equity_plus_liabilities
    return {
        'balanced': abs(difference) <= tolerance,
        'total_assets': total_assets,
        'total_equity_liabilities': equity_plus_liabilities,
        'difference': difference
    }


def calculate_balance(opening, debits, credits, normal_balance='debit'):
    """
    Calculate closing balance for an account.
    
    For debit-normal accounts: Closing = Opening + Debits - Credits
    For credit-normal accounts: Closing = Opening + Credits - Debits
    """
    if normal_balance.lower() == 'debit':
        return opening + debits - credits
    else:
        return opening + credits - debits


def check_control_account(subsidiary_total, gl_control_balance, account_name, tolerance=0.01):
    """
    Check if subsidiary ledger total matches GL control account.
    
    Returns: dict with 'matched', 'subsidiary_total', 'gl_balance', 'difference'
    """
    difference = subsidiary_total - gl_control_balance
    return {
        'matched': abs(difference) <= tolerance,
        'account': account_name,
        'subsidiary_total': subsidiary_total,
        'gl_balance': gl_control_balance,
        'difference': difference
    }
