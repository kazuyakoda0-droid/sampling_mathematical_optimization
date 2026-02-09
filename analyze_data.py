# -*- coding: utf-8 -*-
import pandas as pd
import sys

# Excel files analysis
print("=" * 60)
print("Mathematical Optimization_sampling_2026.xlsx の分析")
print("=" * 60)

xl1 = pd.ExcelFile('Mathematical Optimization_sampling_2026.xlsx')
print(f"シート名: {xl1.sheet_names}")

for sheet in xl1.sheet_names:
    df = pd.read_excel(xl1, sheet_name=sheet)
    print(f"\n--- {sheet} シート ---")
    print(f"列名: {list(df.columns)}")
    print(f"行数: {len(df)}")
    print("\nデータ:")
    print(df.to_string())

print("\n" + "=" * 60)
print("sample.xlsx の分析")
print("=" * 60)

xl2 = pd.ExcelFile('sample.xlsx')
print(f"シート名: {xl2.sheet_names}")

for sheet in xl2.sheet_names:
    df = pd.read_excel(xl2, sheet_name=sheet)
    print(f"\n--- {sheet} シート ---")
    print(f"列名: {list(df.columns)}")
    print(f"行数: {len(df)}")
    print("\nデータ:")
    print(df.to_string())
