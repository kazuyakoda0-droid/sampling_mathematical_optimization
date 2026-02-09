# -*- coding: utf-8 -*-
import pandas as pd

# Mathematical Optimization_sampling_2026.xlsx
xl1 = pd.ExcelFile('Mathematical Optimization_sampling_2026.xlsx')
for sheet in xl1.sheet_names:
    df = pd.read_excel(xl1, sheet_name=sheet)
    df.to_csv(f'data_{sheet}.csv', index=False, encoding='utf-8-sig')
    print(f'Exported: data_{sheet}.csv')

# sample.xlsx  
xl2 = pd.ExcelFile('sample.xlsx')
for sheet in xl2.sheet_names:
    df = pd.read_excel(xl2, sheet_name=sheet)
    df.to_csv(f'schedule_{sheet}.csv', index=False, encoding='utf-8-sig')
    print(f'Exported: schedule_{sheet}.csv')

print("Done!")
