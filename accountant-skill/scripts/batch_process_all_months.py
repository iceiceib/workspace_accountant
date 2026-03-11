#!/usr/bin/env python
"""
Batch process all months from Feb 2025 to Sep 2025.
Chains opening balances from month to month.
Generates Trial Balance in format: Opening | Debits | Credits | Ending
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
import subprocess
import sys

# Paths
SOURCE_DIR = Path('Exisitng Accounting Workflow _ reference files')
OUTPUT_INPUT = Path('data/input')
OUTPUT_OUTPUT = Path('data/output')
MASTER_DIR = Path('data/input/master')

# Months to process (in order for balance chaining)
MONTHS = [
    ('2025-02-01', '2025-02-28', 'Feb2025'),
    ('2025-03-01', '2025-03-31', 'Mar2025'),
    ('2025-04-01', '2025-04-30', 'Apr2025'),
    ('2025-05-01', '2025-05-31', 'May2025'),
    ('2025-06-01', '2025-06-30', 'Jun2025'),
    ('2025-07-01', '2025-07-31', 'Jul2025'),
    ('2025-08-01', '2025-08-31', 'Aug2025'),
    ('2025-09-01', '2025-09-30', 'Sep2025'),
    ('2025-10-01', '2025-10-31', 'Oct2025'),
]

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Style constants
HEADER_FILL = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
HEADER_FONT = Font(bold=True, color='FFFFFF')
TOTAL_FONT = Font(bold=True)
THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
THICK_BOTTOM = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='double')
)


def load_coa():
    """Load Chart of Accounts."""
    coa_path = MASTER_DIR / 'chart_of_accounts.xlsx'
    df = pd.read_excel(coa_path)
    df.columns = [c.strip() for c in df.columns]

    accounts = {}
    for _, row in df.iterrows():
        code = int(row['Account Code'])
        accounts[code] = {
            'code': code,
            'name': row['Account Name'],
            'type': row['Type'],
            'sub_type': row.get('Sub-Type', ''),
            'normal_balance': row['Normal Balance'].lower(),
            'status': row.get('Status', 'Active')
        }
    return accounts


def write_simple_excel(data, filepath, sheet_name='Sheet1'):
    """Write a simple Excel file from list of dicts or DataFrame."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    if isinstance(data, list):
        df = pd.DataFrame(data) if data else pd.DataFrame()
    else:
        df = data

    for col_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal='center')
        cell.border = THIN_BORDER

    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = THIN_BORDER
            if isinstance(value, (int, float)) and not pd.isna(value):
                cell.number_format = '#,##0.00'

    for col in ws.columns:
        max_length = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_length + 2, 50)

    wb.save(filepath)


def extract_period_data(gl_df, coa, start_date, end_date, opening_balances=None):
    """
    Extract period data from GL and classify into journals.
    Returns: (cash_receipts, cash_payments, general_journal, period_gl, ending_balances)
    """
    # Filter GL for this period
    period_gl = gl_df[(gl_df['Date'] >= start_date) & (gl_df['Date'] <= end_date)].copy()

    # Initialize balances with opening - include ALL accounts from COA plus any in GL
    balances = {}
    for code, info in coa.items():
        balances[code] = {
            'name': info['name'],
            'type': info['type'],
            'normal_balance': info['normal_balance'],
            'opening': opening_balances.get(code, 0.0) if opening_balances else 0.0,
            'period_dr': 0.0,
            'period_cr': 0.0,
            'closing': opening_balances.get(code, 0.0) if opening_balances else 0.0,
        }

    # Also add accounts that appear in GL but not in COA
    gl_codes = period_gl['COA Account Number'].dropna().unique()
    for code_val in gl_codes:
        code = int(float(code_val))
        if code not in balances:
            # Get name from GL
            name_rows = period_gl[period_gl['COA Account Number'] == code_val]
            name = str(name_rows['Account Name'].iloc[0]) if len(name_rows) > 0 and pd.notna(name_rows['Account Name'].iloc[0]) else f'Account {code}'

            # Determine type and normal balance from code range
            if code >= 80000:
                acct_type = 'Expense'
                normal_balance = 'debit'
            elif code >= 70000:
                acct_type = 'Revenue'
                normal_balance = 'credit'
            elif code >= 60000:
                acct_type = 'Expense'
                normal_balance = 'debit'
            elif code >= 50000:
                acct_type = 'Expense'
                normal_balance = 'debit'
            elif code >= 40000:
                acct_type = 'Revenue'
                normal_balance = 'credit'
            elif code >= 30000:
                acct_type = 'Equity'
                normal_balance = 'credit'
            elif code >= 25000:
                acct_type = 'Liability'
                normal_balance = 'credit'
            elif code >= 20000:
                acct_type = 'Liability'
                normal_balance = 'credit'
            else:
                acct_type = 'Asset'
                normal_balance = 'debit'

            balances[code] = {
                'name': name,
                'type': acct_type,
                'normal_balance': normal_balance,
                'opening': opening_balances.get(code, 0.0) if opening_balances else 0.0,
                'period_dr': 0.0,
                'period_cr': 0.0,
                'closing': opening_balances.get(code, 0.0) if opening_balances else 0.0,
            }

    # Process GL transactions
    for idx, row in period_gl.iterrows():
        code = int(row['COA Account Number'])
        dr = float(row['Debit (MMK)']) if pd.notna(row['Debit (MMK)']) else 0.0
        cr = float(row['Credit (MMK)']) if pd.notna(row['Credit (MMK)']) else 0.0

        if code in balances:
            balances[code]['period_dr'] += dr
            balances[code]['period_cr'] += cr

            # Update closing balance
            if balances[code]['normal_balance'] == 'debit':
                balances[code]['closing'] = balances[code]['opening'] + balances[code]['period_dr'] - balances[code]['period_cr']
            else:
                balances[code]['closing'] = balances[code]['opening'] - balances[code]['period_dr'] + balances[code]['period_cr']

    # Create journals
    cash_receipts = []
    cash_payments = []
    general_journal = []

    # Sales Revenue -> Cash Receipts
    for idx, row in period_gl[period_gl['COA Account Number'] == 40000].iterrows():
        desc = str(row['Descritpion']) if pd.notna(row['Descritpion']) else ''
        amount = float(row['Credit (MMK)']) if pd.notna(row['Credit (MMK)']) else 0.0
        cash_receipts.append({
            'Date': row['Date'],
            'Receipt No': f'CR-{row["Date"].strftime("%m%d")}-{len(cash_receipts)+1:03d}',
            'Received From': 'Cash Sales',
            'Description': desc or 'Sale of Drinking Water',
            'Amount': amount,
            'Bank Account': 'Main',
            'Debit Account': 10100,
            'Credit Account': 40000,
        })

    # Interest Income, Capital -> Cash Receipts
    for acct in [70000, 31000]:
        for idx, row in period_gl[period_gl['COA Account Number'] == acct].iterrows():
            desc = str(row['Descritpion']) if pd.notna(row['Descritpion']) else ''
            amount = float(row['Credit (MMK)']) if pd.notna(row['Credit (MMK)']) else 0.0
            cash_receipts.append({
                'Date': row['Date'],
                'Receipt No': f'CR-{row["Date"].strftime("%m%d")}-{len(cash_receipts)+1:03d}',
                'Received From': desc[:50] if desc else 'Other Receipt',
                'Description': desc,
                'Amount': amount,
                'Bank Account': 'Main',
                'Debit Account': 10100,
                'Credit Account': acct,
            })

    # Purchases, Expenses, CIP -> Cash Payments
    for acct in [50010, 50110, 53000, 53100, 53200, 65000, 14000, 15500, 15200, 13000]:
        for idx, row in period_gl[period_gl['COA Account Number'] == acct].iterrows():
            desc = str(row['Descritpion']) if pd.notna(row['Descritpion']) else ''
            amount = float(row['Debit (MMK)']) if pd.notna(row['Debit (MMK)']) else 0.0
            if amount > 0:
                cash_payments.append({
                    'Date': row['Date'],
                    'Payment No': f'CP-{row["Date"].strftime("%m%d")}-{len(cash_payments)+1:03d}',
                    'Paid To': desc[:50] if desc else f'Payment for {acct}',
                    'Description': desc,
                    'Amount': amount,
                    'Bank Account': 'Main',
                    'Debit Account': acct,
                    'Credit Account': 10100,
                })

    # Inventory adjustments -> General Journal
    inv_accounts = [50000, 50020, 50100, 50120, 50200, 50220, 12000, 12100, 12200]
    for idx, row in period_gl[period_gl['COA Account Number'].isin(inv_accounts)].iterrows():
        desc = str(row['Descritpion']) if pd.notna(row['Descritpion']) else ''
        dr = float(row['Debit (MMK)']) if pd.notna(row['Debit (MMK)']) else 0.0
        cr = float(row['Credit (MMK)']) if pd.notna(row['Credit (MMK)']) else 0.0
        if dr > 0 or cr > 0:
            general_journal.append({
                'Date': row['Date'],
                'JV No': f'JV-{row["Date"].strftime("%m")}-{len(general_journal)+1:03d}',
                'Description': desc,
                'Debit Account': int(row['COA Account Number']),
                'Credit Account': int(row['COA Account Number']),
                'Debit Amount': dr,
                'Credit Amount': cr,
            })

    # Depreciation -> General Journal
    dep_accounts = [15110, 15210, 15410, 66000, 53300]
    for idx, row in period_gl[period_gl['COA Account Number'].isin(dep_accounts)].iterrows():
        desc = str(row['Descritpion']) if pd.notna(row['Descritpion']) else 'Depreciation'
        dr = float(row['Debit (MMK)']) if pd.notna(row['Debit (MMK)']) else 0.0
        cr = float(row['Credit (MMK)']) if pd.notna(row['Credit (MMK)']) else 0.0
        if dr > 0 or cr > 0:
            general_journal.append({
                'Date': row['Date'],
                'JV No': f'JV-{row["Date"].strftime("%m")}-{len(general_journal)+1:03d}',
                'Description': desc,
                'Debit Account': int(row['COA Account Number']),
                'Credit Account': int(row['COA Account Number']),
                'Debit Amount': dr,
                'Credit Amount': cr,
            })

    # Prepare GL output
    gl_output = period_gl[['Date', 'COA Account Number', 'Account Name', 'Descritpion', 'Debit (MMK)', 'Credit (MMK)', 'Account Balance (MMK)']].copy()
    gl_output.columns = ['Date', 'Account Code', 'Account Name', 'Description', 'Debit', 'Credit', 'Balance']
    gl_output = gl_output.sort_values(['Account Code', 'Date'])

    # Ending balances for chaining
    ending_balances = {code: data['closing'] for code, data in balances.items()}

    return cash_receipts, cash_payments, general_journal, gl_output, balances, ending_balances


def create_trial_balance_xlsx(coa, balances, period_name, start_date, end_date, output_path):
    """Create Trial Balance Excel file with all accounts including zeros."""
    wb = Workbook()

    # Dashboard sheet
    ws = wb.active
    ws.title = 'Dashboard'

    total_opening_dr = 0
    total_opening_cr = 0
    total_period_dr = 0
    total_period_cr = 0
    total_ending_dr = 0
    total_ending_cr = 0

    for code, data in balances.items():
        opening = data['opening']
        period_dr = data['period_dr']
        period_cr = data['period_cr']
        closing = data['closing']
        normal = data['normal_balance']

        # Opening balance display
        if normal == 'debit':
            if opening >= 0:
                total_opening_dr += opening
            else:
                total_opening_cr += abs(opening)
        else:
            if opening >= 0:
                total_opening_cr += opening
            else:
                total_opening_dr += abs(opening)

        # Period movements
        total_period_dr += period_dr
        total_period_cr += period_cr

        # Ending balance display
        if normal == 'debit':
            if closing >= 0:
                total_ending_dr += closing
            else:
                total_ending_cr += abs(closing)
        else:
            if closing >= 0:
                total_ending_cr += closing
            else:
                total_ending_dr += abs(closing)

    # Write dashboard
    ws['A1'] = 'TRIAL BALANCE VALIDATION'
    ws['A1'].font = Font(bold=True, size=14)
    ws['A3'] = f'Period: {start_date} to {end_date}'
    ws['A5'] = 'Opening Balance Check:'
    ws['B5'] = 'PASS' if abs(total_opening_dr - total_opening_cr) < 0.01 else 'FAIL'
    ws['C5'] = f'Dr: {total_opening_dr:,.2f} | Cr: {total_opening_cr:,.2f}'
    ws['A6'] = 'Period Movements Check:'
    ws['B6'] = 'PASS' if abs(total_period_dr - total_period_cr) < 0.01 else 'FAIL'
    ws['C6'] = f'Dr: {total_period_dr:,.2f} | Cr: {total_period_cr:,.2f}'
    ws['A7'] = 'Ending Balance Check:'
    ws['B7'] = 'PASS' if abs(total_ending_dr - total_ending_cr) < 0.01 else 'FAIL'
    ws['C7'] = f'Dr: {total_ending_dr:,.2f} | Cr: {total_ending_cr:,.2f}'

    # Trial Balance sheet
    ws_tb = wb.create_sheet('Trial Balance')

    # Title
    ws_tb['A1'] = 'Trial Balance'
    ws_tb['A1'].font = Font(bold=True, size=14)
    ws_tb['A2'] = f'Period: {start_date} to {end_date}'
    ws_tb['A3'] = ''

    # Headers
    headers = ['Account Code', 'Account Name', 'Opening Balance', 'Debits', 'Credits', 'Ending Balance']
    for col, header in enumerate(headers, 1):
        cell = ws_tb.cell(row=4, column=col, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = THIN_BORDER

    row = 5
    for code in sorted(balances.keys()):
        data = balances[code]
        opening = data['opening']
        period_dr = data['period_dr']
        period_cr = data['period_cr']
        closing = data['closing']
        normal = data['normal_balance']

        # Display opening balance (positive on normal side)
        if normal == 'debit':
            opening_display = opening if opening >= 0 else -opening
        else:
            opening_display = opening if opening >= 0 else -opening

        # Display ending balance
        if normal == 'debit':
            ending_display = closing if closing >= 0 else -closing
        else:
            ending_display = closing if closing >= 0 else -closing

        ws_tb.cell(row=row, column=1, value=code).border = THIN_BORDER
        ws_tb.cell(row=row, column=2, value=data['name']).border = THIN_BORDER

        cell = ws_tb.cell(row=row, column=3, value=opening_display if opening_display != 0 else None)
        cell.border = THIN_BORDER
        cell.number_format = '#,##0.00'

        cell = ws_tb.cell(row=row, column=4, value=period_dr if period_dr != 0 else None)
        cell.border = THIN_BORDER
        cell.number_format = '#,##0.00'

        cell = ws_tb.cell(row=row, column=5, value=period_cr if period_cr != 0 else None)
        cell.border = THIN_BORDER
        cell.number_format = '#,##0.00'

        cell = ws_tb.cell(row=row, column=6, value=ending_display if ending_display != 0 else None)
        cell.border = THIN_BORDER
        cell.number_format = '#,##0.00'

        row += 1

    # Total row
    ws_tb.cell(row=row, column=1, value='TOTAL').font = TOTAL_FONT
    ws_tb.cell(row=row, column=1).border = THICK_BOTTOM
    ws_tb.cell(row=row, column=2).border = THICK_BOTTOM

    cell = ws_tb.cell(row=row, column=3, value=total_opening_dr if total_opening_dr >= total_opening_cr else total_opening_cr)
    cell.font = TOTAL_FONT
    cell.border = THICK_BOTTOM
    cell.number_format = '#,##0.00'

    cell = ws_tb.cell(row=row, column=4, value=total_period_dr)
    cell.font = TOTAL_FONT
    cell.border = THICK_BOTTOM
    cell.number_format = '#,##0.00'

    cell = ws_tb.cell(row=row, column=5, value=total_period_cr)
    cell.font = TOTAL_FONT
    cell.border = THICK_BOTTOM
    cell.number_format = '#,##0.00'

    cell = ws_tb.cell(row=row, column=6, value=total_ending_dr if total_ending_dr >= total_ending_cr else total_ending_cr)
    cell.font = TOTAL_FONT
    cell.border = THICK_BOTTOM
    cell.number_format = '#,##0.00'

    # Auto-fit
    for col in range(1, 7):
        ws_tb.column_dimensions[get_column_letter(col)].width = 20

    wb.save(output_path)
    return total_period_dr, total_period_cr


def create_financial_statements_xlsx(coa, balances, period_name, start_date, end_date, output_path):
    """Create Financial Statements Excel file."""
    wb = Workbook()

    # Income Statement
    ws_is = wb.active
    ws_is.title = 'Income Statement'
    ws_is['A1'] = 'Income Statement'
    ws_is['A1'].font = Font(bold=True, size=14)
    ws_is['A2'] = f'For the period {start_date} to {end_date}'

    # Calculate totals
    revenue = 0
    cogs = 0
    opex = 0
    other_income = 0

    for code, data in balances.items():
        closing = data['closing']
        period_cr = data['period_cr']
        period_dr = data['period_dr']

        if code == 40000:  # Sales Revenue
            revenue = data['closing']  # Credit balance (negative for calculation)
        elif code in [50000, 50010, 50020, 50100, 50110, 50120, 50200, 50220, 53000, 53100, 53200, 53300]:  # COGS
            cogs += data['closing']
        elif code in [60000, 61000, 62000, 63000, 64000, 65000, 66000, 67000, 68000, 69000]:  # SG&A
            opex += data['closing']
        elif code == 70000:  # Interest Income
            other_income = data['closing']

    gross_profit = -revenue - cogs  # Revenue is negative (credit), COGS is positive (debit)
    operating_profit = gross_profit - opex
    net_profit = operating_profit - other_income

    row = 4
    ws_is.cell(row=row, column=1, value='Sales Revenue').font = Font(bold=True)
    ws_is.cell(row=row, column=2, value=abs(-revenue))
    ws_is.cell(row=row, column=2).number_format = '#,##0.00'
    row += 2

    ws_is.cell(row=row, column=1, value='Cost of Goods Sold')
    ws_is.cell(row=row, column=2, value=cogs)
    ws_is.cell(row=row, column=2).number_format = '#,##0.00'
    row += 1

    ws_is.cell(row=row, column=1, value='Gross Profit').font = Font(bold=True)
    ws_is.cell(row=row, column=2, value=gross_profit)
    ws_is.cell(row=row, column=2).number_format = '#,##0.00'
    ws_is.cell(row=row, column=2).font = Font(bold=True)
    row += 2

    ws_is.cell(row=row, column=1, value='Operating Expenses')
    ws_is.cell(row=row, column=2, value=opex)
    ws_is.cell(row=row, column=2).number_format = '#,##0.00'
    row += 1

    ws_is.cell(row=row, column=1, value='Operating Profit').font = Font(bold=True)
    ws_is.cell(row=row, column=2, value=operating_profit)
    ws_is.cell(row=row, column=2).number_format = '#,##0.00'
    ws_is.cell(row=row, column=2).font = Font(bold=True)
    row += 2

    ws_is.cell(row=row, column=1, value='Other Income (Interest)')
    ws_is.cell(row=row, column=2, value=abs(other_income))
    ws_is.cell(row=row, column=2).number_format = '#,##0.00'
    row += 1

    ws_is.cell(row=row, column=1, value='Net Profit/(Loss)').font = Font(bold=True, size=12)
    ws_is.cell(row=row, column=2, value=net_profit)
    ws_is.cell(row=row, column=2).number_format = '#,##0.00'
    ws_is.cell(row=row, column=2).font = Font(bold=True, size=12)

    ws_is.column_dimensions['A'].width = 30
    ws_is.column_dimensions['B'].width = 15

    # Balance Sheet
    ws_bs = wb.create_sheet('Balance Sheet')
    ws_bs['A1'] = 'Balance Sheet'
    ws_bs['A1'].font = Font(bold=True, size=14)
    ws_bs['A2'] = f'As at {end_date}'

    total_assets = 0
    total_liabilities = 0
    total_equity = 0

    row = 4
    ws_bs.cell(row=row, column=1, value='ASSETS').font = Font(bold=True)
    row += 1

    for code in sorted(balances.keys()):
        data = balances[code]
        if data['type'] == 'Asset':
            closing = data['closing']
            # For contra accounts (accumulated depreciation), show as negative
            if data['normal_balance'] == 'credit':
                closing = -closing
            if closing != 0:
                ws_bs.cell(row=row, column=1, value=data['name'])
                ws_bs.cell(row=row, column=2, value=abs(closing))
                ws_bs.cell(row=row, column=2).number_format = '#,##0.00'
                total_assets += closing
                row += 1

    ws_bs.cell(row=row, column=1, value='TOTAL ASSETS').font = Font(bold=True)
    ws_bs.cell(row=row, column=2, value=abs(total_assets)).font = Font(bold=True)
    ws_bs.cell(row=row, column=2).number_format = '#,##0.00'
    row += 2

    ws_bs.cell(row=row, column=1, value='LIABILITIES').font = Font(bold=True)
    row += 1

    for code in sorted(balances.keys()):
        data = balances[code]
        if data['type'] == 'Liability':
            closing = data['closing']
            if closing != 0:
                ws_bs.cell(row=row, column=1, value=data['name'])
                ws_bs.cell(row=row, column=2, value=abs(closing))
                ws_bs.cell(row=row, column=2).number_format = '#,##0.00'
                total_liabilities += closing
                row += 1

    ws_bs.cell(row=row, column=1, value='TOTAL LIABILITIES').font = Font(bold=True)
    ws_bs.cell(row=row, column=2, value=abs(total_liabilities)).font = Font(bold=True)
    ws_bs.cell(row=row, column=2).number_format = '#,##0.00'
    row += 2

    ws_bs.cell(row=row, column=1, value='EQUITY').font = Font(bold=True)
    row += 1

    for code in sorted(balances.keys()):
        data = balances[code]
        if data['type'] == 'Equity':
            closing = data['closing']
            if closing != 0:
                ws_bs.cell(row=row, column=1, value=data['name'])
                ws_bs.cell(row=row, column=2, value=abs(closing))
                ws_bs.cell(row=row, column=2).number_format = '#,##0.00'
                total_equity += closing
                row += 1

    ws_bs.cell(row=row, column=1, value='TOTAL EQUITY').font = Font(bold=True)
    ws_bs.cell(row=row, column=2, value=abs(total_equity)).font = Font(bold=True)
    ws_bs.cell(row=row, column=2).number_format = '#,##0.00'
    row += 2

    ws_bs.cell(row=row, column=1, value='TOTAL LIABILITIES & EQUITY').font = Font(bold=True)
    ws_bs.cell(row=row, column=2, value=abs(total_liabilities + total_equity)).font = Font(bold=True)
    ws_bs.cell(row=row, column=2).number_format = '#,##0.00'

    ws_bs.column_dimensions['A'].width = 40
    ws_bs.column_dimensions['B'].width = 15

    wb.save(output_path)


def main():
    print("="*60)
    print("BATCH PROCESSING WITH BALANCE CHAINING")
    print("Feb 2025 - Oct 2025")
    print("="*60)

    # Load GL once
    print("\nLoading General Ledger...")
    gl_df = pd.read_excel(SOURCE_DIR / 'Ledger Accounts' / 'General_Ledger_edited.xlsx', header=3)
    gl_df = gl_df.dropna(how='all')
    gl_df['Date'] = pd.to_datetime(gl_df['Date'], errors='coerce')
    gl_df = gl_df[gl_df['Date'].notna()]
    print(f"  Total GL rows: {len(gl_df)}")

    # Load COA
    print("Loading Chart of Accounts...")
    coa = load_coa()
    print(f"  Total accounts: {len(coa)}")

    # Initialize opening balances (all zeros for Feb 2025)
    opening_balances = {code: 0.0 for code in coa.keys()}

    results = []

    for start_date, end_date, period_name in MONTHS:
        print(f"\n{'='*60}")
        print(f"Processing: {period_name}")
        print('='*60)

        output_dir = OUTPUT_OUTPUT / period_name
        output_dir.mkdir(parents=True, exist_ok=True)

        # Extract data with opening balances
        cash_receipts, cash_payments, general_journal, gl_output, balances, ending_balances = extract_period_data(
            gl_df, coa, start_date, end_date, opening_balances
        )

        print(f"  GL transactions: {len(gl_output)}")
        print(f"  Cash Receipts: {len(cash_receipts)}")
        print(f"  Cash Payments: {len(cash_payments)}")
        print(f"  General Journal: {len(general_journal)}")

        # Write input files (for modules that need them)
        journals_dir = OUTPUT_INPUT / 'journals'
        write_simple_excel(cash_receipts, journals_dir / 'cash_receipts_journal.xlsx')
        write_simple_excel(cash_payments, journals_dir / 'cash_payments_journal.xlsx')
        write_simple_excel(general_journal, journals_dir / 'general_journal.xlsx')
        write_simple_excel(gl_output, OUTPUT_INPUT / 'ledgers' / 'general_ledger.xlsx')

        # Run Module 1
        print(f"  Running Module 1...")
        cmd = f'python scripts/summarize_journals.py data/input/journals {start_date} {end_date} data/output/{period_name}/books_of_prime_entry_{period_name}.xlsx data/input/master'
        subprocess.run(cmd, shell=True, capture_output=True)

        # Run Module 2
        print(f"  Running Module 2...")
        cmd = f'python scripts/summarize_ledgers.py data/input/ledgers {start_date} {end_date} data/output/{period_name}/ledger_summary_{period_name}.xlsx data/input/master'
        subprocess.run(cmd, shell=True, capture_output=True)

        # Create Trial Balance directly (not using Module 5 to include all accounts)
        print(f"  Creating Trial Balance...")
        total_dr, total_cr = create_trial_balance_xlsx(
            coa, balances, period_name, start_date, end_date,
            output_dir / f'trial_balance_{period_name}.xlsx'
        )
        print(f"    Period Dr: {total_dr:,.2f} | Period Cr: {total_cr:,.2f}")

        # Create Financial Statements
        print(f"  Creating Financial Statements...")
        create_financial_statements_xlsx(
            coa, balances, period_name, start_date, end_date,
            output_dir / f'financial_statements_{period_name}.xlsx'
        )

        # Chain balances to next period
        opening_balances = ending_balances.copy()
        results.append((period_name, total_dr, total_cr, abs(total_dr - total_cr) < 0.01))

    # Summary
    print("\n" + "="*60)
    print("SUMMARY")
    print("="*60)
    print(f"{'Period':12} | {'Period Dr':>15} | {'Period Cr':>15} | Balanced")
    print("-"*60)
    for period_name, dr, cr, balanced in results:
        print(f"{period_name:12} | {dr:>15,.2f} | {cr:>15,.2f} | {'YES' if balanced else 'NO'}")


if __name__ == '__main__':
    main()