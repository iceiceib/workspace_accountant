"""Convert general_ledger.xlsx to 5-digit codes."""
import openpyxl

OLD_TO_NEW = {
    1010: 10000, 1020: 10100, 1030: 10000,
    1100: 11000, 1110: 11000,
    1200: 12000, 1300: 13000, 1320: 13000,
    1610: 15100, 1611: 15110, 1620: 15200, 1621: 15210,
    1630: 15300, 1650: 15300, 1651: 15310,
    2010: 20000, 2020: 22000, 2030: 22200, 2040: 20000, 2060: 21000, 2100: 25000,
    3010: 31000, 3020: 30200, 3030: 32000, 3040: 32000,
    4010: 40000, 4020: 40000, 4030: 40000, 4040: 40000, 4100: 40000, 4110: 70000,
    5010: 50000, 5020: 50100, 5030: 53000, 5040: 53999,
    5100: 61000, 5110: 62000, 5200: 65000, 5210: 63000, 5220: 65000,
    5300: 66000, 5400: 65000, 5410: 65000, 5420: 65000,
    5500: 64000, 5600: 60000, 5700: 65000, 5800: 67000,
    5900: 68000, 5910: 71000, 5920: 65000,
}

wb = openpyxl.load_workbook('data/Jan2026/general_ledger.xlsx')
sheet = wb['General Ledger']

converted = 0
for row in sheet.iter_rows():
    for cell in row:
        if cell.value is not None and isinstance(cell.value, (int, float)):
            val = int(cell.value)
            if val in OLD_TO_NEW:
                cell.value = OLD_TO_NEW[val]
                converted += 1

wb.save('data/Jan2026/general_ledger.xlsx')
print(f'general_ledger.xlsx: {converted} cells converted')
print('Saved successfully!')
