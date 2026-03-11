"""
Accounting Constants - Centralized account codes and configuration.

All modules should import from this file to ensure consistency.
The K&K Finance COA uses 5-digit account codes.
"""

# ============================================================================
# Account Codes (5-digit COA)
# ============================================================================

# Cash Accounts
CASH_IN_HAND = 10000
CASH_AT_BANK = 10100

# Receivables
ACCOUNTS_RECEIVABLE = 11000

# Inventory
INVENTORY_RAW_MATERIAL = 12000
INVENTORY_PACKAGING = 12100
INVENTORY_FINISHED_GOODS = 12200
INVENTORY_ADJUSTMENTS = 12300
WORK_IN_PROGRESS = 12400

# Prepayments
ADVANCED_PAYMENTS = 13000
DEFERRED_PRELIMINARY_EXPENSES = 14000

# Non-Current Assets
LAND = 15000
BUILDINGS = 15100
ACCUM_DEPR_BUILDINGS = 15110
MACHINERY_EQUIPMENT = 15200
ACCUM_DEPR_MACHINERY = 15210
OFFICE_EQUIPMENT = 15300
ACCUM_DEPR_OFFICE = 15310
ELECTRICAL_SYSTEMS = 15400
ACCUM_DEPR_ELECTRICAL = 15410
CONSTRUCTION_IN_PROGRESS = 15500
ACCUM_DEPR_VEHICLES = 15510
MOTOR_VEHICLES = 15600

# Current Liabilities
ACCOUNTS_PAYABLE = 20000
SHORT_TERM_LOANS = 21000
UTILITY_BILLS = 22000
WAGES_PAYABLE = 22200

# Non-Current Liabilities
BANK_LOAN = 25000

# Equity
PAID_UP_CAPITAL = 31000
RETAINED_EARNINGS = 32000

# Revenue
SALES_REVENUE = 40000

# COGS - Raw Materials
OPENING_INV_RAW_MATERIALS = 50000
PURCHASES_RAW_MATERIALS = 50010
CLOSING_INV_RAW_MATERIALS = 50020

# COGS - Packaging
OPENING_INV_PACKAGING = 50100
PURCHASES_PACKAGING = 50110
CLOSING_INV_PACKAGING = 50120

# COGS - Finished Goods
OPENING_INV_FINISHED_GOODS = 50200
CLOSING_INV_FINISHED_GOODS = 50220

# COGS - WIP
OPENING_WIP = 50300
CLOSING_WIP = 50310
DIRECT_MATERIALS_USED = 50320
DIRECT_LABOR_TO_WIP = 50330
MANUFACTURING_OVERHEAD = 50340
WIP_TO_FINISHED_GOODS = 50350

# COGS - Production
DIRECT_LABOR_WAGES = 53000
MACHINE_MAINTENANCE = 53100
PRODUCTION_UTILITIES = 53200
DEPRECIATION_COGS = 53300

# Operating Expenses (SG&A)
MARKETING_EXPENSE = 60000
OFFICE_SALARIES = 61000
MEAL_ALLOWANCE = 62000
UTILITIES = 63000
TRANSPORTATION = 64000
FACTORY_OFFICE_SUPPLIES = 65000
DEPRECIATION_SGA = 66000
INVENTORY_WRITE_OFF = 67000
OTHER_EXPENSES = 68000
MANAGEMENT_COMPENSATION = 69000

# Other Income
INTEREST_INCOME = 70000

# ============================================================================
# Control Account Codes
# ============================================================================

AR_GL_ACCOUNT = ACCOUNTS_RECEIVABLE
AP_GL_ACCOUNT = ACCOUNTS_PAYABLE
CASH_GL_ACCOUNTS = [CASH_AT_BANK]

# ============================================================================
# Account Code Ranges
# ============================================================================

# By Type
CURRENT_ASSET_RANGE = (10000, 14999)
NON_CURRENT_ASSET_RANGE = (15000, 19999)
CURRENT_LIABILITY_RANGE = (20000, 24999)
NON_CURRENT_LIABILITY_RANGE = (25000, 29999)
EQUITY_RANGE = (30000, 39999)
REVENUE_RANGE = (40000, 49999)
EXPENSE_RANGE = (50000, 69999)
OTHER_INCOME_RANGE = (70000, 79999)

# By Sub-Type for P&L
COGS_RANGE = (50000, 53999)  # Includes WIP accounts
SGA_RANGE = (60000, 69999)

# ============================================================================
# Profit/Cost Center Requirements
# ============================================================================

# Accounts that require Profit Center tagging
PC_REQUIRED_RANGES = [
    REVENUE_RANGE,
    EXPENSE_RANGE,
]

# Accounts that require Cost Center tagging
CC_REQUIRED_RANGES = [
    EXPENSE_RANGE,
]

# ============================================================================
# Contra Accounts (opposite normal balance)
# ============================================================================

CONTRA_ACCOUNTS = {
    INVENTORY_ADJUSTMENTS,
    ACCUM_DEPR_BUILDINGS,
    ACCUM_DEPR_MACHINERY,
    ACCUM_DEPR_OFFICE,
    ACCUM_DEPR_ELECTRICAL,
    ACCUM_DEPR_VEHICLES,
}

# ============================================================================
# Accumulated Depreciation Mapping
# Maps asset account code -> (accum_depr_code, accum_depr_name)
# ============================================================================

ACCUM_DEPR_MAP = {
    BUILDINGS: (ACCUM_DEPR_BUILDINGS, 'Accum. Depr. - Buildings & Structures'),
    MACHINERY_EQUIPMENT: (ACCUM_DEPR_MACHINERY, 'Accum. Depr. - Machinery & Equipment'),
    OFFICE_EQUIPMENT: (ACCUM_DEPR_OFFICE, 'Accum. Depr. - Office & Facility Equipment'),
    ELECTRICAL_SYSTEMS: (ACCUM_DEPR_ELECTRICAL, 'Accum. Depr. - Electrical & Utility Systems'),
    MOTOR_VEHICLES: (ACCUM_DEPR_VEHICLES, 'Accum. Depr. - Motor Vehicles'),
    CONSTRUCTION_IN_PROGRESS: ('15610', 'Accum. Depr. - Construction in Progress'),
}