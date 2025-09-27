#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动测试所有转置文件
批量验证转置结果
"""

import os
import glob
import subprocess
import sys
from datetime import datetime

def find_transposed_files():
    """查找所有转置后的文件"""
    patterns = [
        "*_转置后.xlsx",
        "*_转置完成.xlsx",
        "*转置后*.xlsx"
    ]
    
    transposed_files = []
    for pattern in patterns:
        files = glob.glob(pattern)
        transposed_files.extend(files)
    
    return list(set(transposed_files))  # 去重

def find_original_files():
    """查找原始文件"""
    original_files = []
    
    # 查找待处理文件目录
    if os.path.exists("待处理文件"):
        for file in os.listdir("待处理文件"):
            if file.endswith(".xlsx") and not file.startswith("~"):
                original_files.append(os.path.join("待处理文件", file))
    
    # 查找当前目录的原始文件
    for file in os.listdir("."):
        if (file.endswith(".xlsx") and 
            not file.startswith("~") and 
            "转置" not in file and 
            "测试" not in file and
            "示例" not in file):
            original_files.append(file)
    
    return original_files

def match_original_and_transposed():
    """匹配原始文件和转置文件"""
    original_files = find_original_files()
    transposed_files = find_transposed_files()
    
    matches = []
    
    for original in original_files:
        original_base = os.path.splitext(os.path.basename(original))[0]
        
        for transposed in transposed_files:
            transposed_base = os.path.splitext(os.path.basename(transposed))[0]
            
            # 检查转置文件是否包含原始文件的基础名称
            if original_base in transposed_base:
                matches.append((original, transposed))
                break
    
    return matches

def run_validation_test(original_file, transposed_file):
    """运行单个验证测试"""
    print(f"\n{'='*60}")
    print(f"测试: {os.path.basename(original_file)} -> {os.path.basename(transposed_file)}")
    print(f"{'='*60}")
    
    try:
        # 运行验证测试
        result = subprocess.run([
            sys.executable, "test_transpose_validation.py", 
            original_file, transposed_file
        ], capture_output=True, text=True, encoding='utf-8')
        
        print(result.stdout)
        if result.stderr:
            print("错误信息:", result.stderr)
        
        return result.returncode == 0
        
    except Exception as e:
        print(f"测试执行失败: {str(e)}")
        return False

def generate_summary_report(test_results):
    """生成总结报告"""
    report_file = f"批量测试总结报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    
    total_tests = len(test_results)
    passed_tests = sum(1 for result in test_results if result['passed'])
    failed_tests = total_tests - passed_tests
    
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write("批量转置验证测试总结报告\n")
        f.write("=" * 50 + "\n")
        f.write(f"测试时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"总测试数: {total_tests}\n")
        f.write(f"通过测试: {passed_tests}\n")
        f.write(f"失败测试: {failed_tests}\n")
        f.write(f"通过率: {passed_tests/total_tests*100:.1f}%\n\n")
        
        f.write("详细测试结果:\n")
        f.write("-" * 30 + "\n")
        for result in test_results:
            status = "通过" if result['passed'] else "失败"
            f.write(f"{status} {result['original']} -> {result['transposed']}\n")
            if not result['passed']:
                f.write(f"  错误: {result.get('error', '未知错误')}\n")
        f.write("\n")
    
    print(f"\n总结报告已保存到: {report_file}")
    return report_file

def main():
    """主函数"""
    print("=" * 60)
    print("批量转置验证测试工具")
    print("=" * 60)
    
    # 查找匹配的文件对
    matches = match_original_and_transposed()
    
    if not matches:
        print("未找到匹配的原始文件和转置文件对")
        return
    
    print(f"找到 {len(matches)} 对文件:")
    for original, transposed in matches:
        print(f"  {os.path.basename(original)} -> {os.path.basename(transposed)}")
    
    # 运行所有测试
    test_results = []
    
    for original, transposed in matches:
        passed = run_validation_test(original, transposed)
        test_results.append({
            'original': os.path.basename(original),
            'transposed': os.path.basename(transposed),
            'passed': passed
        })
    
    # 生成总结报告
    report_file = generate_summary_report(test_results)
    
    # 显示最终结果
    total_tests = len(test_results)
    passed_tests = sum(1 for result in test_results if result['passed'])
    failed_tests = total_tests - passed_tests
    
    print(f"\n{'='*60}")
    print("批量测试完成")
    print(f"{'='*60}")
    print(f"总测试数: {total_tests}")
    print(f"通过测试: {passed_tests}")
    print(f"失败测试: {failed_tests}")
    print(f"通过率: {passed_tests/total_tests*100:.1f}%")
    
    if failed_tests > 0:
        print("\n失败的测试:")
        for result in test_results:
            if not result['passed']:
                print(f"  ❌ {result['original']} -> {result['transposed']}")
    
    if passed_tests == total_tests:
        print("\n🎉 所有测试通过！所有转置数据验证成功！")
        return 0
    else:
        print(f"\n⚠️  {failed_tests} 个测试失败，请检查转置结果！")
        return 1

if __name__ == "__main__":
    sys.exit(main())

