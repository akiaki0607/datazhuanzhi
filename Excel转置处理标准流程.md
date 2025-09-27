# Excel合并单元格转置处理标准流程

## 概述

本文档记录了处理包含合并单元格的Excel文件转置操作的标准流程。该流程经过实际验证，能够正确处理复杂的合并单元格结构，确保数据完整性。

## 处理流程

### 第一步：文件结构分析

#### 1.1 基础信息获取
```python
# 使用openpyxl读取工作簿
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb[sheet_name]

# 获取基础信息
print(f"工作表尺寸: {ws.max_row} 行 x {ws.max_column} 列")
```

#### 1.2 合并单元格识别
```python
# 分析合并单元格
merged_ranges = list(ws.merged_cells.ranges)
print(f"找到 {len(merged_ranges)} 个合并单元格区域")

# 创建品牌到列索引的映射
brand_columns = {}
for merged_range in merged_ranges:
    top_left_cell = ws[merged_range.min_row][merged_range.min_col-1]
    brand_name = top_left_cell.value
    
    if brand_name and isinstance(brand_name, str):
        brand_columns[brand_name] = {
            'start_col': merged_range.min_col,
            'end_col': merged_range.max_col,
            'row': merged_range.min_row
        }
```

#### 1.3 子标题行分析
```python
# 获取子标题行（通常是第2行）
sub_headers_row = 2
sub_headers = []
for col_idx in range(1, ws.max_column + 1):
    cell = ws.cell(row=sub_headers_row, column=col_idx)
    sub_headers.append(cell.value)
```

### 第二步：数据映射建立

#### 2.1 品牌列映射
```python
# 为每个品牌确定其子标题
brand_to_columns = {}
for brand_name, col_info in brand_columns.items():
    start_col = col_info['start_col']
    end_col = col_info['end_col']
    
    # 获取该品牌对应的子标题
    brand_sub_headers = []
    for i in range(start_col, end_col + 1):
        if i <= len(sub_headers):
            brand_sub_headers.append(sub_headers[i-1])
    
    brand_to_columns[brand_name] = {
        'start_col': start_col,
        'end_col': end_col,
        'sub_headers': brand_sub_headers
    }
```

### 第三步：数据提取

#### 3.1 逐行数据提取
```python
# 提取所有数据
data_rows = []
data_start_row = 3  # 数据从第3行开始（第1行品牌，第2行子标题）

for row_idx in range(data_start_row, ws.max_row + 1):
    # 获取信源平台名称（第一列）
    source_platform = ws.cell(row=row_idx, column=1).value
    
    if source_platform is None or source_platform == '':
        continue
    
    # 为每个品牌提取数据
    for brand_name, col_info in brand_to_columns.items():
        row_data = {
            '信源平台名称': source_platform,
            '品牌': brand_name
        }
        
        # 提取该品牌对应的数据列
        start_col = col_info['start_col']
        end_col = col_info['end_col']
        sub_headers = col_info['sub_headers']
        
        for i, sub_header in enumerate(sub_headers):
            col_idx = start_col + i
            if col_idx <= end_col and col_idx <= ws.max_column:
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                row_data[sub_header] = cell_value
            else:
                row_data[sub_header] = None
        
        data_rows.append(row_data)
```

### 第四步：数据验证

#### 4.1 完整性检查
```python
# 创建DataFrame
df = pd.DataFrame(data_rows)

# 检查是否有非空数据
non_empty_count = 0
for col in ['DeepSeek', 'Kimi', '元宝', '豆包']:
    if col in df.columns:
        non_empty = df[col].notna() & (df[col] != '0.0%') & (df[col] != 0) & (df[col] != '0')
        non_empty_count += non_empty.sum()

print(f"非空数据条目数量: {non_empty_count}")
```

#### 4.2 数据统计
```python
# 统计信息
print(f"总行数: {len(df)}")
print(f"总列数: {len(df.columns)}")
print(f"唯一信源平台数: {df['信源平台名称'].nunique()}")
print(f"唯一品牌数: {df['品牌'].nunique()}")

# 各列非空数据统计
for col in ['DeepSeek', 'Kimi', '元宝', '豆包']:
    if col in df.columns:
        non_empty = df[col].notna() & (df[col] != '0.0%') & (df[col] != 0) & (df[col] != '0')
        print(f"{col}列非空数据: {non_empty.sum()} 条")
```

### 第五步：文件保存

#### 5.1 保存为Excel
```python
# 创建新的Excel文件
output_file = "转置后数据.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='转置后数据', index=False)

print(f"数据已保存到: {output_file}")
```

## 关键要点

### 1. 数据行识别
- **第1行**：品牌名称（合并单元格）
- **第2行**：子标题（DeepSeek、Kimi、元宝、豆包等）
- **第3行开始**：实际数据

### 2. 列映射逻辑
- 每个品牌对应一个合并单元格区域
- 合并单元格的列范围确定该品牌的数据列范围
- 子标题行提供列名信息

### 3. 数据提取策略
- 为每个信源平台创建多行数据（每个品牌一行）
- 保持原始数据格式（百分比、数值等）
- 确保所有数据都被提取，不遗漏

### 4. 验证标准
- 检查非空数据数量
- 验证数据完整性
- 确认转换后的数据结构正确

## 常见问题及解决方案

### 问题1：数据为空
**原因**：列映射错误
**解决**：仔细检查合并单元格的列范围，确保正确映射到子标题

### 问题2：数据缺失
**原因**：数据起始行识别错误
**解决**：检查前几行，确定实际数据开始的行号

### 问题3：列名不匹配
**原因**：子标题行识别错误
**解决**：检查第2行的内容，确认子标题格式

## 标准脚本模板

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel合并单元格转置工具 - 标准版
"""

import pandas as pd
import openpyxl
import os

def process_excel_transpose(input_file, output_file):
    """
    标准Excel转置处理函数
    """
    try:
        # 第一步：文件结构分析
        wb = openpyxl.load_workbook(input_file, data_only=True)
        ws = wb[wb.sheetnames[0]]
        
        # 识别合并单元格
        merged_ranges = list(ws.merged_cells.ranges)
        brand_columns = {}
        for merged_range in merged_ranges:
            top_left_cell = ws[merged_range.min_row][merged_range.min_col-1]
            brand_name = top_left_cell.value
            if brand_name and isinstance(brand_name, str):
                brand_columns[brand_name] = {
                    'start_col': merged_range.min_col,
                    'end_col': merged_range.max_col
                }
        
        # 获取子标题
        sub_headers = []
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=2, column=col_idx)
            sub_headers.append(cell.value)
        
        # 第二步：建立映射
        brand_to_columns = {}
        for brand_name, col_info in brand_columns.items():
            start_col = col_info['start_col']
            end_col = col_info['end_col']
            brand_sub_headers = []
            for i in range(start_col, end_col + 1):
                if i <= len(sub_headers):
                    brand_sub_headers.append(sub_headers[i-1])
            brand_to_columns[brand_name] = {
                'start_col': start_col,
                'end_col': end_col,
                'sub_headers': brand_sub_headers
            }
        
        # 第三步：数据提取
        data_rows = []
        for row_idx in range(3, ws.max_row + 1):
            source_platform = ws.cell(row=row_idx, column=1).value
            if source_platform is None or source_platform == '':
                continue
            
            for brand_name, col_info in brand_to_columns.items():
                row_data = {
                    '信源平台名称': source_platform,
                    '品牌': brand_name
                }
                
                start_col = col_info['start_col']
                end_col = col_info['end_col']
                sub_headers = col_info['sub_headers']
                
                for i, sub_header in enumerate(sub_headers):
                    col_idx = start_col + i
                    if col_idx <= end_col and col_idx <= ws.max_column:
                        cell_value = ws.cell(row=row_idx, column=col_idx).value
                        row_data[sub_header] = cell_value
                    else:
                        row_data[sub_header] = None
                
                data_rows.append(row_data)
        
        # 第四步：数据验证和保存
        df = pd.DataFrame(data_rows)
        
        # 保存文件
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='转置后数据', index=False)
        
        print(f"处理完成！")
        print(f"输入文件: {input_file}")
        print(f"输出文件: {output_file}")
        print(f"转换后数据行数: {len(df)}")
        print(f"转换后数据列数: {len(df.columns)}")
        
        return df
        
    except Exception as e:
        print(f"处理过程中出现错误: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    # 使用示例
    input_file = "待处理文件.xlsx"
    output_file = "转置后文件.xlsx"
    
    df = process_excel_transpose(input_file, output_file)
```

## 使用说明

1. **准备环境**：安装pandas和openpyxl
2. **调用函数**：使用`process_excel_transpose(input_file, output_file)`
3. **检查结果**：验证输出文件的数据完整性
4. **处理异常**：根据错误信息调整参数

## 注意事项

- 确保输入文件格式正确（第1行品牌，第2行子标题，第3行开始数据）
- 合并单元格必须正确设置
- 子标题行必须包含所有需要的列名
- 建议在处理前备份原始文件

---

**创建时间**：2025年1月27日  
**版本**：1.0  
**适用场景**：包含合并单元格的Excel文件转置操作

