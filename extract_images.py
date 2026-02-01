# -*- coding: utf-8 -*-
import openpyxl
from openpyxl_image_loader import SheetImageLoader
import os

excel_path = r'D:\Code\d2l-zh\my-new-site\assets\(数据表)表格视图.xlsx'
images_dir = r'D:\Code\d2l-zh\my-new-site\images\excel'

# 创建图片目录
if not os.path.exists(images_dir):
    os.makedirs(images_dir)

# 加载工作簿
wb = openpyxl.load_workbook(excel_path)
ws = wb.active

print("Extracting images from Excel...")

# 尝试提取图片
try:
    from openpyxl_image_loader import SheetImageLoader
    image_loader = SheetImageLoader(ws)

    # 获取所有图片
    images = image_loader.images
    print(f"Found {len(images)} images")

    for img_name, img in images.items():
        # 保存图片
        img_path = os.path.join(images_dir, f"{img_name}.png")
        img.save(img_path)
        print(f"Saved: {img_path}")

except ImportError:
    print("Installing openpyxl-image-loader...")
    os.system("pip install openpyxl-image-loader")
except Exception as e:
    print(f"Error: {e}")
    print("Trying alternative method...")

    # 备用方法：直接从工作表提取图片
    from openpyxl.utils import get_column_letter
    image_count = 0

    for image in ws._images:
        image_count += 1
        # 获取图片位置
        img_name = f"excel_image_{image_count}"
        img_path = os.path.join(images_dir, f"{img_name}.png")

        # 保存图片
        with open(img_path, 'wb') as f:
            f.write(image._data())
        print(f"Saved: {img_path}")

print("\nDone!")
