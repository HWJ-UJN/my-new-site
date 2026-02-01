# -*- coding: utf-8 -*-
import openpyxl
import os
from openpyxl.utils import get_column_letter

excel_path = r'D:\Code\d2l-zh\my-new-site\assets\(数据表)表格视图.xlsx'
images_dir = r'D:\Code\d2l-zh\my-new-site\images\excel'

# 创建图片目录
if not os.path.exists(images_dir):
    os.makedirs(images_dir)

# 加载工作簿
wb = openpyxl.load_workbook(excel_path)
ws = wb.active

print(f"Images found in worksheet: {len(ws._images)}")
print("Extracting images...\n")

for idx, image in enumerate(ws._images, 1):
    # 获取图片位置
    anchor = image.anchor
    if hasattr(anchor, 'from_'):
        col = anchor.from_.col
        row = anchor.from_.row
        print(f"Image {idx} at column {col}, row {row}")

    # 保存图片
    img_name = f"note_{idx}"
    img_path = os.path.join(images_dir, f"{img_name}.png")

    with open(img_path, 'wb') as f:
        f.write(image._data())

    print(f"Saved: {img_path}\n")

print("Extraction complete!")
