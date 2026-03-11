"""
Inventory Mapper - Lookup inventory items, calculate WAC, and track inventory movements.

This module provides:
- InventoryMapper class for loading and querying inventory items
- WAC (Weighted Average Cost) calculation functions
- Inventory movement tracking utilities
"""

import pandas as pd
import math
from pathlib import Path
from typing import Optional, Dict, List, Tuple


# Default inventory item definitions based on K&K Finance structure
# Item codes 12001-12018 = Raw Materials, 12100-12106 = Packaging
DEFAULT_INVENTORY_ITEMS = {
    # Raw Materials (12001-12018)
    12001: {'name': 'Coffee Beans (Premium)', 'category': 'Raw Materials', 'unit': 'Bag', 'account_code': 12000},
    12002: {'name': 'Sugar', 'category': 'Raw Materials', 'unit': 'Bag', 'account_code': 12000},
    12003: {'name': 'Milk Powder', 'category': 'Raw Materials', 'unit': 'Bag', 'account_code': 12000},
    12004: {'name': 'Creamer', 'category': 'Raw Materials', 'unit': 'Bag', 'account_code': 12000},
    12005: {'name': 'Tea Leaves', 'category': 'Raw Materials', 'unit': 'Gram', 'account_code': 12000},
    12006: {'name': 'Condensed Milk', 'category': 'Raw Materials', 'unit': 'Gram', 'account_code': 12000},
    12007: {'name': 'Evaporated Milk', 'category': 'Raw Materials', 'unit': 'Gram', 'account_code': 12000},
    12008: {'name': 'Flavoring Syrup', 'category': 'Raw Materials', 'unit': 'Bottles', 'account_code': 12000},
    12009: {'name': 'Chocolate Powder', 'category': 'Raw Materials', 'unit': 'Gram', 'account_code': 12000},
    12010: {'name': 'Honey', 'category': 'Raw Materials', 'unit': 'Gram', 'account_code': 12000},
    12011: {'name': 'Butter', 'category': 'Raw Materials', 'unit': 'Gram', 'account_code': 12000},
    12012: {'name': 'Flour', 'category': 'Raw Materials', 'unit': 'Bag', 'account_code': 12000},
    12013: {'name': 'Baking Powder', 'category': 'Raw Materials', 'unit': 'Gram', 'account_code': 12000},
    12014: {'name': 'Salt', 'category': 'Raw Materials', 'unit': 'Gram', 'account_code': 12000},
    12015: {'name': 'Vanilla Extract', 'category': 'Raw Materials', 'unit': 'Gram', 'account_code': 12000},
    12016: {'name': 'Eggs', 'category': 'Raw Materials', 'unit': 'Pack', 'account_code': 12000},
    12017: {'name': 'Oil', 'category': 'Raw Materials', 'unit': 'Gram', 'account_code': 12000},
    12018: {'name': 'Other Ingredients', 'category': 'Raw Materials', 'unit': 'Gram', 'account_code': 12000},

    # Packaging Materials (12100-12106)
    12100: {'name': 'Packing Bags', 'category': 'Packaging', 'unit': 'Pack', 'account_code': 12100},
    12101: {'name': 'Small Cups', 'category': 'Packaging', 'unit': 'Pack', 'account_code': 12100},
    12102: {'name': 'Large Cups', 'category': 'Packaging', 'unit': 'Pack', 'account_code': 12100},
    12103: {'name': 'Lids', 'category': 'Packaging', 'unit': 'Pack', 'account_code': 12100},
    12104: {'name': 'Straws', 'category': 'Packaging', 'unit': 'Pack', 'account_code': 12100},
    12105: {'name': 'Napkins', 'category': 'Packaging', 'unit': 'Pack', 'account_code': 12100},
    12106: {'name': 'Carry Bags', 'category': 'Packaging', 'unit': 'Pack', 'account_code': 12100},
}

# Account code ranges for inventory categories
INVENTORY_ACCOUNT_RANGES = {
    'raw_materials': (12000, 12099),
    'packaging': (12100, 12199),
    'finished_goods': (12200, 12299),
    'wip': (12400, 12499),
}


def _clean(val):
    """Clean value, handling NaN properly."""
    if val is None:
        return ''
    if isinstance(val, float) and math.isnan(val):
        return ''
    return str(val).strip() if pd.notna(val) else ''


def _clean_numeric(val):
    """Clean numeric value, returning 0 for NaN/None."""
    if val is None:
        return 0.0
    if isinstance(val, float) and math.isnan(val):
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def calculate_wac(opening_qty: float, opening_value: float,
                  purchase_qty: float, purchase_value: float) -> float:
    """
    Calculate Weighted Average Cost.

    WAC = (Opening Value + Purchase Value) / (Opening Qty + Purchase Qty)

    Returns 0 if total quantity is 0.
    """
    total_qty = opening_qty + purchase_qty
    if total_qty <= 0:
        return 0.0

    total_value = opening_value + purchase_value
    return total_value / total_qty


class InventoryMapper:
    """
    Inventory item lookup and management.

    Loads inventory items from file or uses defaults.
    Provides methods to query items by code, category, or account.
    """

    def __init__(self, items_filepath=None):
        """
        Initialize with optional inventory items file.

        Args:
            items_filepath: Path to inventory_items.xlsx (optional)
        """
        self.items_df = None
        self.items_dict = {}

        if items_filepath:
            self.load_items(items_filepath)
        else:
            # Use defaults
            self._load_defaults()

    def _load_defaults(self):
        """Load default inventory items."""
        self.items_dict = {}
        for code, info in DEFAULT_INVENTORY_ITEMS.items():
            self.items_dict[code] = {
                'code': code,
                'name': info['name'],
                'category': info['category'],
                'unit': info['unit'],
                'account_code': info['account_code'],
                'status': 'Active'
            }

    def load_items(self, filepath):
        """
        Load inventory items from .xlsx file.

        Expected columns: Item Code, Item Name, Account Code, Unit Measure, Category, Status
        """
        filepath = Path(filepath)
        if not filepath.exists():
            print(f"Warning: Inventory items file not found: {filepath}. Using defaults.")
            self._load_defaults()
            return

        try:
            df = pd.read_excel(filepath)
            df.columns = [str(c).strip().lower() for c in df.columns]

            # Find columns
            code_col = None
            for candidate in ['item code', 'code', 'item_code', 'no.']:
                if candidate in df.columns:
                    code_col = candidate
                    break

            if code_col is None:
                print("Warning: Cannot find item code column. Using defaults.")
                self._load_defaults()
                return

            name_col = None
            for candidate in ['item name', 'name', 'description', 'item_name']:
                if candidate in df.columns:
                    name_col = candidate
                    break

            account_col = None
            for candidate in ['account code', 'account_code', 'acct code', 'gl code']:
                if candidate in df.columns:
                    account_col = candidate
                    break

            unit_col = None
            for candidate in ['unit measure', 'unit', 'uom', 'unit_measure']:
                if candidate in df.columns:
                    unit_col = candidate
                    break

            category_col = None
            for candidate in ['category', 'type', 'item_type']:
                if candidate in df.columns:
                    category_col = candidate
                    break

            self.items_df = df
            for _, row in df.iterrows():
                try:
                    code = int(float(_clean(row[code_col]))) if pd.notna(row[code_col]) else None
                except (ValueError, TypeError):
                    continue

                if code is None:
                    continue

                entry = {
                    'code': code,
                    'name': _clean(row[name_col]) if name_col else f'Item {code}',
                    'account_code': int(float(_clean(row[account_col]))) if account_col and pd.notna(row[account_col]) else self._get_default_account(code),
                    'unit': _clean(row[unit_col]) if unit_col else 'Unit',
                    'category': _clean(row[category_col]) if category_col else self._get_default_category(code),
                    'status': 'Active'
                }
                self.items_dict[code] = entry

            # Merge with defaults for any missing items
            for code, info in DEFAULT_INVENTORY_ITEMS.items():
                if code not in self.items_dict:
                    self.items_dict[code] = {
                        'code': code,
                        'name': info['name'],
                        'category': info['category'],
                        'unit': info['unit'],
                        'account_code': info['account_code'],
                        'status': 'Active'
                    }

        except Exception as e:
            print(f"Warning: Error loading inventory items: {e}. Using defaults.")
            self._load_defaults()

    def _get_default_account(self, code):
        """Get default GL account code for an item code."""
        if 12001 <= code <= 12099:
            return 12000  # Raw Materials
        elif 12100 <= code <= 12199:
            return 12100  # Packaging
        elif 12200 <= code <= 12299:
            return 12200  # Finished Goods
        elif 12400 <= code <= 12499:
            return 12400  # WIP
        return None

    def _get_default_category(self, code):
        """Get default category for an item code."""
        if 12001 <= code <= 12099:
            return 'Raw Materials'
        elif 12100 <= code <= 12199:
            return 'Packaging'
        elif 12200 <= code <= 12299:
            return 'Finished Goods'
        elif 12400 <= code <= 12499:
            return 'WIP'
        return 'Other'

    def get_item(self, code) -> Optional[Dict]:
        """
        Get inventory item by code.

        Returns dict with: code, name, category, unit, account_code, status
        """
        try:
            code = int(float(code))
        except (ValueError, TypeError):
            return None

        return self.items_dict.get(code)

    def get_item_name(self, code) -> str:
        """Get item name by code."""
        item = self.get_item(code)
        return item['name'] if item else f'Item {code}'

    def get_item_unit(self, code) -> str:
        """Get item unit of measure by code."""
        item = self.get_item(code)
        return item['unit'] if item else 'Unit'

    def get_item_category(self, code) -> str:
        """Get item category by code."""
        item = self.get_item(code)
        return item['category'] if item else 'Other'

    def get_items_by_category(self, category: str) -> List[Dict]:
        """Get all items in a category."""
        return [item for item in self.items_dict.values()
                if item.get('category', '').lower() == category.lower()]

    def get_items_by_account(self, account_code: int) -> List[Dict]:
        """Get all items for a GL account code."""
        try:
            account_code = int(float(account_code))
        except (ValueError, TypeError):
            return []

        return [item for item in self.items_dict.values()
                if item.get('account_code') == account_code]

    def get_items_by_range(self, code_min: int, code_max: int) -> List[Dict]:
        """Get all items within a code range."""
        return [item for item in self.items_dict.values()
                if code_min <= item['code'] <= code_max]

    def get_raw_materials(self) -> List[Dict]:
        """Get all raw material items (codes 12001-12099)."""
        return self.get_items_by_range(12001, 12099)

    def get_packaging(self) -> List[Dict]:
        """Get all packaging items (codes 12100-12199)."""
        return self.get_items_by_range(12100, 12199)

    def validate_item_code(self, code) -> bool:
        """Check if an item code is valid."""
        return self.get_item(code) is not None

    def is_raw_material(self, code) -> bool:
        """Check if item is a raw material."""
        try:
            code = int(float(code))
        except (ValueError, TypeError):
            return False
        return 12001 <= code <= 12099

    def is_packaging(self, code) -> bool:
        """Check if item is packaging material."""
        try:
            code = int(float(code))
        except (ValueError, TypeError):
            return False
        return 12100 <= code <= 12199


class InventoryLedger:
    """
    Tracks inventory quantities and values with WAC calculation.
    """

    def __init__(self, item_code: int, item_name: str, unit: str,
                 opening_qty: float = 0, opening_value: float = 0):
        """
        Initialize an inventory ledger for a single item.

        Args:
            item_code: The inventory item code
            item_name: The item name/description
            unit: Unit of measure
            opening_qty: Opening quantity balance
            opening_value: Opening value balance
        """
        self.item_code = item_code
        self.item_name = item_name
        self.unit = unit
        self.opening_qty = opening_qty
        self.opening_value = opening_value
        self.wac = opening_value / opening_qty if opening_qty > 0 else 0.0

        # Current running balance
        self.current_qty = opening_qty
        self.current_value = opening_value

        # Transaction history
        self.transactions = []

    def receive(self, date, reference, qty: float, unit_cost: float,
                description: str = '') -> Dict:
        """
        Record receipt of inventory (purchase).

        Args:
            date: Transaction date
            reference: PO/Invoice reference
            qty: Quantity received
            unit_cost: Cost per unit
            description: Optional description

        Returns:
            Dict with transaction details
        """
        value = qty * unit_cost

        # Recalculate WAC
        self.wac = calculate_wac(
            self.current_qty, self.current_value,
            qty, value
        )

        # Update balance
        self.current_qty += qty
        self.current_value += value

        txn = {
            'date': date,
            'reference': reference,
            'description': description,
            'txn_type': 'RECEIVE',
            'received_qty': qty,
            'issued_qty': 0,
            'balance_qty': self.current_qty,
            'unit_cost': unit_cost,
            'received_value': value,
            'issued_value': 0,
            'balance_value': self.current_value,
            'wac': self.wac
        }
        self.transactions.append(txn)
        return txn

    def issue(self, date, reference, qty: float,
              description: str = '') -> Dict:
        """
        Record issue of inventory (to production).

        Uses current WAC for valuation.

        Args:
            date: Transaction date
            reference: Production batch reference
            qty: Quantity issued
            description: Optional description

        Returns:
            Dict with transaction details
        """
        if qty > self.current_qty:
            raise ValueError(f"Cannot issue {qty} units. Only {self.current_qty} available.")

        value = qty * self.wac

        # Update balance
        self.current_qty -= qty
        self.current_value -= value

        txn = {
            'date': date,
            'reference': reference,
            'description': description,
            'txn_type': 'ISSUE',
            'received_qty': 0,
            'issued_qty': qty,
            'balance_qty': self.current_qty,
            'unit_cost': self.wac,
            'received_value': 0,
            'issued_value': value,
            'balance_value': self.current_value,
            'wac': self.wac
        }
        self.transactions.append(txn)
        return txn

    def get_balance(self) -> Tuple[float, float, float]:
        """
        Get current balance.

        Returns:
            Tuple of (quantity, value, wac)
        """
        return (self.current_qty, self.current_value, self.wac)

    def get_transactions_df(self) -> pd.DataFrame:
        """Get all transactions as a DataFrame."""
        if not self.transactions:
            return pd.DataFrame(columns=[
                'Date', 'Reference', 'Description', 'Received Qty', 'Issued Qty',
                'Balance Qty', 'Unit Cost', 'Received Value', 'Issued Value',
                'Balance Value', 'WAC'
            ])

        df = pd.DataFrame(self.transactions)
        df = df.rename(columns={
            'date': 'Date',
            'reference': 'Reference',
            'description': 'Description',
            'received_qty': 'Received Qty',
            'issued_qty': 'Issued Qty',
            'balance_qty': 'Balance Qty',
            'unit_cost': 'Unit Cost',
            'received_value': 'Received Value',
            'issued_value': 'Issued Value',
            'balance_value': 'Balance Value',
            'wac': 'WAC'
        })
        return df[[
            'Date', 'Reference', 'Description', 'Received Qty', 'Issued Qty',
            'Balance Qty', 'Unit Cost', 'Received Value', 'Issued Value',
            'Balance Value', 'WAC'
        ]]

    def get_period_summary(self) -> Dict:
        """
        Get period summary.

        Returns:
            Dict with opening, movements, and closing balances
        """
        total_received_qty = sum(t['received_qty'] for t in self.transactions)
        total_received_value = sum(t['received_value'] for t in self.transactions)
        total_issued_qty = sum(t['issued_qty'] for t in self.transactions)
        total_issued_value = sum(t['issued_value'] for t in self.transactions)

        return {
            'item_code': self.item_code,
            'item_name': self.item_name,
            'unit': self.unit,
            'opening_qty': self.opening_qty,
            'opening_value': self.opening_value,
            'received_qty': total_received_qty,
            'received_value': total_received_value,
            'issued_qty': total_issued_qty,
            'issued_value': total_issued_value,
            'closing_qty': self.current_qty,
            'closing_value': self.current_value,
            'wac': self.wac
        }