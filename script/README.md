# Excel 日记增量提取工具使用说明

## 功能特点

✅ **自动增量提取** - 只添加新数据，避免重复
✅ **图片自动映射** - 自动提取并映射到对应行
✅ **多文件支持** - 支持批量处理多个 xlsx 文件
✅ **数据去重** - 基于日期自动去重
✅ **智能排序** - 自动按日期倒序排列

## 使用方法

### 1. 添加新日记文件

将新的 xlsx 文件放到 `script/` 文件夹中，文件名需要包含 "log"，例如：
- `day_log_2026_01.xlsx`
- `daily_log_jan.xlsx`

### 2. 运行提取脚本

```bash
cd script
python excel_analyse.py
```

### 3. 查看结果

- **数据文件**: `day_log_data.json`
- **图片文件夹**: `images/excel/`

## 工作流程

```
新 xlsx 文件 → 提取数据 → 检查日期 → 去重 → 合并数据 → 保存 JSON
     ↓
提取图片 → 映射行号 → 保存到 images/excel/
```

## 输出格式

### JSON 数据结构
```json
{
  "headers": ["日期", "文本", "感悟-想法", "图片", "链接", "Time"],
  "rows": [
    {
      "日期": "2026-01-15",
      "文本": "今天学习了...",
      "图片": "day_log_row2.png"
    }
  ],
  "image_count": 9
}
```

### 图片命名规则
- 格式: `{文件名}_row{行号}.{扩展名}`
- 例如: `day_log_row2.png`, `day_log_row5.jpeg`

## 配置选项

在 `excel_analyse.py` 中可以修改：

```python
# Excel 文件存放目录
EXCEL_DIR = r'D:\Code\d2l-zh\my-new-site\script'

# 图片输出目录
IMAGES_OUTPUT_DIR = r'D:\Code\d2l-zh\my-new-site\images\excel'

# 数据输出文件
DATA_OUTPUT_FILE = r'D:\Code\d2l-zh\my-new-site\day_log_data.json'
```

## 批量处理

脚本会自动处理所有匹配的文件：

```python
# 处理所有包含 "log" 的 xlsx 文件
extractor.run('*log*.xlsx')

# 处理特定文件
extractor.run('day_log.xlsx')

# 处理所有 xlsx 文件
extractor.run('*.xlsx')
```

## 常见问题

### Q: 如何重新生成所有数据？
A: 删除 `day_log_data.json` 后重新运行脚本

### Q: 图片重复了怎么办？
A: 脚本会自动使用 `{文件名}_row{行号}` 格式避免冲突

### Q: 如何修改列名？
A: 直接在 Excel 文件中修改表头，脚本会自动识别

### Q: 支持哪些图片格式？
A: png, jpg, jpeg, gif, bmp

## 下一步

现在数据已准备好，可以在前端页面中使用：
- `about.html` - 笔记展示页面
- 使用 JavaScript 读取 `day_log_data.json`
- 动态渲染日记内容
