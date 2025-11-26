#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
综合数据分析脚本 - 提供payroll_details表的全面数据概览
包括数据质量评估、分布分析、异常检测等
"""

import sqlite3
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import os
import sys

# Import database configuration
sys.path.append(os.path.join(os.path.dirname(__file__), 'excel_processor'))
from config import DATABASE_PATH

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False

def get_comprehensive_data_analysis():
    """
    获取payroll_details表的全面数据分析
    """
    try:
        # 连接到SQLite数据库
        conn = sqlite3.connect(DATABASE_PATH)
        
        print("=" * 80)
        print("PAYROLL_DETAILS 表综合数据分析报告")
        print("=" * 80)
        
        # 读取所有数据
        df = pd.read_sql_query("SELECT * FROM payroll_details;", conn)
        
        print(f"\n1. 数据基本信息")
        print("-" * 40)
        print(f"总记录数: {len(df):,}")
        print(f"列数: {len(df.columns)}")
        print(f"数据时间跨度: {df['文件名'].min()} 到 {df['文件名'].max()}")
        
        print(f"\n2. 表结构信息")
        print("-" * 40)
        print("列名\t\t数据类型\t\t非空值数\t空值数\t空值比例")
        print("-" * 80)
        
        for col in df.columns:
            non_null_count = df[col].notna().sum()
            null_count = df[col].isna().sum()
            null_percentage = (null_count / len(df)) * 100
            dtype = df[col].dtype
            
            print(f"{col:<15}\t{str(dtype):<15}\t{non_null_count:<10}\t{null_count:<8}\t{null_percentage:.2f}%")
        
        print(f"\n3. 数据质量评估")
        print("-" * 40)
        
        # 关键字段的空值分析
        key_columns = ['职员全名', '日期', '型号', '工序全名', '金额']
        print("关键字段空值分析:")
        for col in key_columns:
            null_count = df[col].isna().sum()
            null_percentage = (null_count / len(df)) * 100
            print(f"  {col}: {null_count} 个空值 ({null_percentage:.2f}%)")
        
        # 重复记录分析
        duplicate_count = df.duplicated().sum()
        duplicate_percentage = (duplicate_count / len(df)) * 100
        print(f"\n重复记录: {duplicate_count} 条 ({duplicate_percentage:.2f}%)")
        
        # 异常值检测
        print(f"\n4. 数值字段异常值检测")
        print("-" * 40)
        
        numeric_columns = ['计件数量', '系数', '定额', '金额']
        for col in numeric_columns:
            if col in df.columns:
                # 转换为数值类型
                df[col] = pd.to_numeric(df[col], errors='coerce')
                
                # 统计信息
                mean_val = df[col].mean()
                std_val = df[col].std()
                min_val = df[col].min()
                max_val = df[col].max()
                
                # 异常值检测 (使用3σ原则)
                lower_bound = mean_val - 3 * std_val
                upper_bound = mean_val + 3 * std_val
                outliers = df[(df[col] < lower_bound) | (df[col] > upper_bound)][col].count()
                
                print(f"  {col}:")
                print(f"    范围: {min_val:.2f} - {max_val:.2f}")
                print(f"    均值: {mean_val:.2f}, 标准差: {std_val:.2f}")
                print(f"    异常值: {outliers} 个 ({outliers/len(df)*100:.2f}%)")
        
        print(f"\n5. 分类字段分布分析")
        print("-" * 40)
        
        # 职员分布
        print("职员分布 (前10名):")
        employee_counts = df['职员全名'].value_counts().head(10)
        for employee, count in employee_counts.items():
            percentage = (count / len(df)) * 100
            print(f"  {employee}: {count} 条记录 ({percentage:.2f}%)")
        
        # 型号分布
        print(f"\n型号种类: {df['型号'].nunique()} 种")
        print("最常用型号 (前10名):")
        model_counts = df['型号'].value_counts().head(10)
        for model, count in model_counts.items():
            percentage = (count / len(df)) * 100
            print(f"  {model}: {count} 条记录 ({percentage:.2f}%)")
        
        # 工序分布
        print(f"\n工序种类: {df['工序全名'].nunique()} 种")
        print("最常用工序 (前10名):")
        process_counts = df['工序全名'].value_counts().head(10)
        for process, count in process_counts.items():
            percentage = (count / len(df)) * 100
            print(f"  {process}: {count} 条记录 ({percentage:.2f}%)")
        
        # 文件分布
        print(f"\n文件分布:")
        file_counts = df['文件名'].value_counts()
        print(f"  文件总数: {len(file_counts)}")
        print(f"  平均每个文件记录数: {len(df)/len(file_counts):.1f}")
        print(f"  最多记录的文件: {file_counts.idxmax()} ({file_counts.max()} 条记录)")
        print(f"  最少记录的文件: {file_counts.idxmin()} ({file_counts.min()} 条记录)")
        
        print(f"\n6. 金额分析")
        print("-" * 40)
        
        amount_stats = df['金额'].describe()
        print(f"金额统计:")
        print(f"  最小值: {amount_stats['min']:.2f}")
        print(f"  最大值: {amount_stats['max']:.2f}")
        print(f"  平均值: {amount_stats['mean']:.2f}")
        print(f"  中位数: {amount_stats['50%']:.2f}")
        print(f"  标准差: {amount_stats['std']:.2f}")
        
        # 金额分布区间
        amount_bins = [0, 10, 50, 100, 200, 500, 1000, float('inf')]
        amount_labels = ['0-10', '10-50', '50-100', '100-200', '200-500', '500-1000', '1000+']
        amount_distribution = pd.cut(df['金额'], bins=amount_bins, labels=amount_labels).value_counts().sort_index()
        
        print(f"\n金额分布区间:")
        for bin_label, count in amount_distribution.items():
            percentage = (count / len(df)) * 100
            print(f"  {bin_label}: {count} 条记录 ({percentage:.2f}%)")
        
        print(f"\n7. 时间趋势分析")
        print("-" * 40)
        
        # 从文件名提取年份信息
        df['年份'] = df['文件名'].str.extract(r'(\d{4})')[0]
        year_counts = df['年份'].value_counts().sort_index()
        
        print("按年份的记录分布:")
        for year, count in year_counts.items():
            percentage = (count / len(df)) * 100
            print(f"  {year}: {count} 条记录 ({percentage:.2f}%)")
        
        print(f"\n8. 数据完整性评估")
        print("-" * 40)
        
        # 计算完整记录的比例
        complete_records = df.dropna(subset=key_columns)
        complete_percentage = (len(complete_records) / len(df)) * 100
        print(f"完整记录 (所有关键字段非空): {len(complete_records)} 条 ({complete_percentage:.2f}%)")
        
        # 数据质量评分
        quality_score = (
            (1 - duplicate_percentage/100) * 0.2 +
            (complete_percentage/100) * 0.3 +
            (1 - df['金额'].isna().sum()/len(df)) * 0.3 +
            (1 - df['职员全名'].isna().sum()/len(df)) * 0.2
        ) * 100
        
        print(f"\n数据质量综合评分: {quality_score:.1f}/100")
        
        if quality_score >= 90:
            print("数据质量: 优秀")
        elif quality_score >= 80:
            print("数据质量: 良好")
        elif quality_score >= 70:
            print("数据质量: 一般")
        else:
            print("数据质量: 需要改进")
        
        print(f"\n9. 建议和改进点")
        print("-" * 40)
        
        issues = []
        if duplicate_count > 0:
            issues.append(f"- 存在 {duplicate_count} 条重复记录，建议进行数据去重")
        
        if df['职员全名'].isna().sum() > 0:
            issues.append(f"- 职员全名字段有 {df['职员全名'].isna().sum()} 个空值，影响数据分析")
        
        if df['金额'].isna().sum() > 0:
            issues.append(f"- 金额字段有 {df['金额'].isna().sum()} 个空值，影响财务分析")
        
        if len(issues) == 0:
            print("- 数据质量良好，无明显问题")
        else:
            for issue in issues:
                print(issue)
        
        print(f"\n" + "=" * 80)
        print("数据分析完成")
        print("=" * 80)
        
        # 生成可视化图表
        generate_visualizations(df)
        
    except sqlite3.Error as e:
        print(f"数据库错误: {e}")
    except Exception as e:
        print(f"发生错误: {e}")
    finally:
        if conn:
            conn.close()

def generate_visualizations(df):
    """
    生成数据可视化图表
    """
    try:
        print(f"\n生成数据可视化图表...")
        
        # 创建图表目录
        os.makedirs('data_analysis_charts', exist_ok=True)
        
        # 1. 金额分布直方图
        plt.figure(figsize=(12, 8))
        
        plt.subplot(2, 2, 1)
        # 过滤掉极端值以便更好地显示分布
        amount_filtered = df[(df['金额'] > 0) & (df['金额'] <= 500)]['金额']
        plt.hist(amount_filtered, bins=50, alpha=0.7, color='skyblue', edgecolor='black')
        plt.title('金额分布 (0-500元)')
        plt.xlabel('金额 (元)')
        plt.ylabel('频次')
        
        # 2. 职员工作量分布
        plt.subplot(2, 2, 2)
        top_employees = df['职员全名'].value_counts().head(15)
        top_employees.plot(kind='bar', color='lightcoral')
        plt.title('工作量最多的前15名职员')
        plt.xlabel('职员姓名')
        plt.ylabel('记录数量')
        plt.xticks(rotation=45)
        
        # 3. 年份趋势
        plt.subplot(2, 2, 3)
        df['年份'] = df['文件名'].str.extract(r'(\d{4})')[0]
        year_counts = df['年份'].value_counts().sort_index()
        year_counts.plot(kind='line', marker='o', color='green')
        plt.title('按年份的记录数量趋势')
        plt.xlabel('年份')
        plt.ylabel('记录数量')
        
        # 4. 金额箱线图
        plt.subplot(2, 2, 4)
        amount_data = df[df['金额'] <= 500]['金额']  # 过滤极端值
        plt.boxplot(amount_data)
        plt.title('金额分布箱线图')
        plt.ylabel('金额 (元)')
        
        plt.tight_layout()
        plt.savefig('data_analysis_charts/data_overview.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        print("图表已保存至: data_analysis_charts/data_overview.png")
        
    except Exception as e:
        print(f"生成图表时出错: {e}")

if __name__ == "__main__":
    get_comprehensive_data_analysis()
