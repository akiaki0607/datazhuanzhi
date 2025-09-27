import pandas as pd
import os
from datetime import datetime

def read_excel_file(file_path):
    """读取Excel文件"""
    try:
        # 尝试读取Excel文件的所有工作表
        excel_file = pd.ExcelFile(file_path)
        print(f"工作表名称: {excel_file.sheet_names}")
        
        # 读取第一个工作表
        df = pd.read_excel(file_path, sheet_name=0, header=None)
        print(f"原始数据形状: {df.shape}")
        print("原始数据预览:")
        print(df.head(10))
        return df
    except Exception as e:
        print(f"读取文件 {file_path} 时出错: {e}")
        return None

def analyze_example_structure(example_df):
    """分析示例文件的结构"""
    print("\n=== 示例文件结构分析 ===")
    print(f"数据形状: {example_df.shape}")
    print("数据内容:")
    print(example_df.to_string())
    
    # 查找非空数据的范围
    non_null_mask = example_df.notna()
    rows_with_data = non_null_mask.any(axis=1)
    cols_with_data = non_null_mask.any(axis=0)
    
    print(f"\n有数据的行数: {rows_with_data.sum()}")
    print(f"有数据的列数: {cols_with_data.sum()}")
    
    return example_df

def transpose_data(df):
    """转置数据"""
    print("\n=== 开始转置操作 ===")
    
    # 移除完全为空的行和列
    df_cleaned = df.dropna(how='all').dropna(axis=1, how='all')
    print(f"清理后数据形状: {df_cleaned.shape}")
    
    # 转置数据
    df_transposed = df_cleaned.T
    print(f"转置后数据形状: {df_transposed.shape}")
    
    # 重置索引
    df_transposed.reset_index(drop=True, inplace=True)
    df_transposed.columns = range(len(df_transposed.columns))
    
    return df_transposed

def save_transposed_file(df, original_path, output_dir="outputs"):
    """保存转置后的文件"""
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 生成输出文件名
    base_name = os.path.splitext(os.path.basename(original_path))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"{base_name}_转置_{timestamp}.xlsx"
    output_path = os.path.join(output_dir, output_filename)
    
    # 保存文件
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='转置数据', index=False, header=False)
        print(f"转置文件已保存到: {output_path}")
        return output_path
    except Exception as e:
        print(f"保存文件时出错: {e}")
        return None

def main():
    # 文件路径
    example_file = "示例/三次测试对内报表_仅仅示例_副本5.xlsx"
    target_file = "待处理文件/2025926移山科技循环10次采集对内报表_副本.xlsx"
    
    print("=== Excel文件转置处理程序 ===\n")
    
    # 检查文件是否存在
    if not os.path.exists(example_file):
        print(f"示例文件不存在: {example_file}")
        return
    
    if not os.path.exists(target_file):
        print(f"待处理文件不存在: {target_file}")
        return
    
    # 读取示例文件
    print("1. 读取示例文件...")
    example_df = read_excel_file(example_file)
    if example_df is None:
        return
    
    # 分析示例文件结构
    analyze_example_structure(example_df)
    
    # 读取待处理文件
    print("\n2. 读取待处理文件...")
    target_df = read_excel_file(target_file)
    if target_df is None:
        return
    
    # 转置待处理文件
    print("\n3. 转置待处理文件...")
    transposed_df = transpose_data(target_df)
    
    print("\n转置后数据预览:")
    print(transposed_df.head(10))
    
    # 保存转置后的文件
    print("\n4. 保存转置文件...")
    output_path = save_transposed_file(transposed_df, target_file)
    
    if output_path:
        print(f"\n✅ 处理完成！转置文件已保存到: {output_path}")
        
        # 显示转置前后的对比信息
        print(f"\n📊 转置对比:")
        print(f"原始文件形状: {target_df.shape}")
        print(f"转置后形状: {transposed_df.shape}")
    else:
        print("\n❌ 处理失败！")

if __name__ == "__main__":
    main()