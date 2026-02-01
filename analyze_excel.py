# -*- coding: utf-8 -*-
import pandas as pd
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

excel_path = r'D:\Code\d2l-zh\my-new-site\assets\(数据表)表格视图.xlsx'

# 读取Excel文件
df = pd.read_excel(excel_path, sheet_name=0)

print("=== Excel Analysis ===\n")
print(f"Total rows: {len(df)}")
print(f"\nColumns: {list(df.columns)}")
print(f"\nColumn count: {len(df.columns)}")

print("\n=== First 5 rows ===")
print(df.head())

print("\n=== Data types ===")
print(df.dtypes)

print("\n=== Null values ===")
print(df.isnull().sum())

# 检查是否有图片相关的列
print("\n=== Image-related columns ===")
for col in df.columns:
    col_str = str(col)
    if '图' in col_str or 'image' in col_str.lower() or 'img' in col_str.lower() or 'picture' in col_str.lower():
        print(f"Found image column: {col}")
        print(f"Values: {df[col].dropna().tolist()}")

