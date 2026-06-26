import pandas as pd
import sys

FILE = 'Dashboard Culinary Depurado Santiago.xlsx'
xl = pd.ExcelFile(FILE)
print('=== HOJAS ===')
print(xl.sheet_names)
print()

for sheet in xl.sheet_names:
    df = pd.read_excel(FILE, sheet_name=sheet, header=None)
    print(f'=== HOJA: {sheet} === ({df.shape[0]} filas x {df.shape[1]} cols)')
    # Print first 5 rows raw
    print(df.head(10).to_string())
    print()
