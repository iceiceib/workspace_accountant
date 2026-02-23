"""
Excel Reader Utility — Standardized .xlsx reading with flexible column matching.
"""
import pandas as pd
import os
from pathlib import Path


# Column name aliases for flexible matching
COLUMN_ALIASES = {
    'account code': ['code', 'acct code', 'acct_code', 'account_code', 'no.', 'account no', 'account no.'],
    'account name': ['name', 'description', 'account', 'account_name', 'acct name'],
    'date': ['date', 'trans date', 'transaction date', 'entry date', 'post date'],
    'reference': ['ref', 'ref no', 'ref.', 'reference no', 'voucher no', 'jv no', 'invoice no', 'receipt no', 'payment no'],
    'description': ['desc', 'narration', 'narrative', 'memo', 'particulars', 'details'],
    'debit': ['debit', 'dr', 'debit amount', 'debit_amount'],
    'credit': ['credit', 'cr', 'credit amount', 'credit_amount'],
    'amount': ['amount', 'total', 'value', 'sum'],
    'balance': ['balance', 'running balance', 'bal'],
    'debit account': ['debit account', 'debit_account', 'dr account', 'dr_account', 'account debit'],
    'credit account': ['credit account', 'credit_account', 'cr account', 'cr_account', 'account credit'],
    'customer': ['customer', 'client', 'buyer', 'debtor', 'received from', 'from'],
    'supplier': ['supplier', 'vendor', 'seller', 'creditor', 'paid to', 'to'],
    'type': ['type', 'account type', 'category', 'classification'],
    'sub-type': ['sub-type', 'sub type', 'subtype', 'sub_type', 'subcategory'],
    'normal balance': ['normal balance', 'normal_balance', 'norm bal', 'normal'],
    'bank account': ['bank account', 'bank_account', 'bank', 'bank name'],
    'asset id': ['asset id', 'asset_id', 'asset no', 'asset code', 'id'],
    'cost': ['cost', 'original cost', 'purchase cost', 'acquisition cost'],
    'useful life (years)': ['useful life (years)', 'useful life', 'useful_life', 'life years', 'life'],
    'salvage value': ['salvage value', 'salvage_value', 'residual value', 'residual', 'scrap value'],
    'depreciation method': ['depreciation method', 'depreciation_method', 'method', 'depr method'],
    'accumulated depreciation': ['accumulated depreciation', 'accumulated_depreciation', 'accum depr', 'accum depreciation', 'total depreciation'],
    'net book value': ['net book value', 'net_book_value', 'nbv', 'book value', 'carrying value'],
    'date acquired': ['date acquired', 'date_acquired', 'acquisition date', 'purchase date'],
    'status': ['status', 'active', 'state'],
    'employee / department': ['employee / department', 'employee', 'department', 'dept', 'staff'],
    'location': ['location', 'site', 'branch', 'shop'],
    'category': ['category', 'asset category', 'class', 'group'],
    'annual depreciation': ['annual depreciation', 'annual_depreciation', 'yearly depreciation'],
    'monthly depreciation': ['monthly depreciation', 'monthly_depreciation'],
}


def normalize_column(col_name):
    """Normalize a column name for matching."""
    return str(col_name).strip().lower()


def map_columns(df, required_columns, optional_columns=None):
    """
    Map DataFrame columns to canonical names using aliases.
    
    Args:
        df: pandas DataFrame
        required_columns: list of canonical column names that must be present
        optional_columns: list of canonical column names that are nice to have
    
    Returns:
        tuple: (mapped_df, missing_required, mapping_used)
    """
    if optional_columns is None:
        optional_columns = []

    normalized_headers = {normalize_column(c): c for c in df.columns}
    mapping = {}
    missing_required = []

    all_needed = required_columns + optional_columns
    for canonical in all_needed:
        canon_lower = canonical.lower()
        # Direct match
        if canon_lower in normalized_headers:
            mapping[canonical] = normalized_headers[canon_lower]
            continue
        # Alias match
        aliases = COLUMN_ALIASES.get(canon_lower, [])
        found = False
        for alias in aliases:
            if alias in normalized_headers:
                mapping[canonical] = normalized_headers[alias]
                found = True
                break
        if not found and canonical in required_columns:
            missing_required.append(canonical)

    if missing_required:
        return None, missing_required, mapping

    # Rename columns
    reverse_map = {v: k for k, v in mapping.items()}
    mapped_df = df.rename(columns=reverse_map)
    return mapped_df, [], mapping


def read_xlsx(filepath, sheet_name=0, required_columns=None, optional_columns=None, date_columns=None):
    """
    Read an .xlsx file with column validation.
    
    Args:
        filepath: path to .xlsx file
        sheet_name: sheet name or index (default 0)
        required_columns: list of required canonical column names
        optional_columns: list of optional canonical column names
        date_columns: list of columns to parse as dates
    
    Returns:
        dict with keys: 'data' (DataFrame or None), 'error' (str or None), 'mapping' (dict)
    """
    if required_columns is None:
        required_columns = []
    if optional_columns is None:
        optional_columns = []

    filepath = Path(filepath)
    if not filepath.exists():
        return {'data': None, 'error': f"File not found: {filepath}", 'mapping': {}}
    if filepath.suffix.lower() not in ['.xlsx', '.xls']:
        return {'data': None, 'error': f"Not an Excel file: {filepath}", 'mapping': {}}

    try:
        parse_dates = date_columns if date_columns else False
        df = pd.read_excel(filepath, sheet_name=sheet_name, parse_dates=parse_dates)
    except Exception as e:
        return {'data': None, 'error': f"Error reading {filepath}: {str(e)}", 'mapping': {}}

    if not required_columns:
        return {'data': df, 'error': None, 'mapping': {}}

    mapped_df, missing, mapping = map_columns(df, required_columns, optional_columns)
    if missing:
        available = [str(c) for c in df.columns]
        return {
            'data': None,
            'error': f"Missing required columns: {missing}. Available columns: {available}",
            'mapping': mapping
        }

    return {'data': mapped_df, 'error': None, 'mapping': mapping}


def read_all_sheets(filepath):
    """Read all sheets from an xlsx file into a dict of DataFrames."""
    filepath = Path(filepath)
    if not filepath.exists():
        return {'data': None, 'error': f"File not found: {filepath}"}
    try:
        sheets = pd.read_excel(filepath, sheet_name=None)
        return {'data': sheets, 'error': None}
    except Exception as e:
        return {'data': None, 'error': f"Error reading {filepath}: {str(e)}"}


def find_xlsx_files(directory, pattern=None):
    """Find all .xlsx files in a directory, optionally matching a pattern."""
    directory = Path(directory)
    if not directory.exists():
        return []
    files = list(directory.glob('*.xlsx'))
    if pattern:
        files = [f for f in files if pattern.lower() in f.stem.lower()]
    return sorted(files)


def filter_by_period(df, date_column, start_date, end_date):
    """Filter DataFrame rows to a date range."""
    df = df.copy()
    df[date_column] = pd.to_datetime(df[date_column], errors='coerce')
    mask = (df[date_column] >= pd.Timestamp(start_date)) & (df[date_column] <= pd.Timestamp(end_date))
    return df[mask]
