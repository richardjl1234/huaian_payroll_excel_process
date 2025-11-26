#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
金额关系分析脚本 - 分析计件数量、系数、定额与金额之间的关系
"""

import sqlite3
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
import os
import sys

# Import database configuration
sys.path.append(os.path.join(os.path.dirname(__file__), 'excel_processor'))
from config import DATABASE_PATH

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False

def analyze_amount_relationships():
    """
    分析计件数量、系数、定额与金额之间的关系
    """
    try:
        # 连接到SQLite数据库
        conn = sqlite3.connect(DATABASE_PATH)
        
        print("=" * 80)
        print("计件数量、系数、定额与金额关系分析")
        print("=" * 80)
        
        # 读取所有数据
        df = pd.read_sql_query("SELECT * FROM payroll_details;", conn)
        
        # 确保数值字段是数值类型
        numeric_columns = ['计件数量', '系数', '定额', '金额']
        for col in numeric_columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # 过滤掉无效数据
        df_clean = df.dropna(subset=numeric_columns)
        df_clean = df_clean[(df_clean['金额'] > 0) & (df_clean['金额'] <= 1000)]  # 过滤极端值
        
        print(f"分析数据量: {len(df_clean):,} 条记录")
        
        print(f"\n1. 基本统计信息")
        print("-" * 40)
        for col in numeric_columns:
            stats = df_clean[col].describe()
            print(f"{col}:")
            print(f"  最小值: {stats['min']:.2f}")
            print(f"  最大值: {stats['max']:.2f}")
            print(f"  平均值: {stats['mean']:.2f}")
            print(f"  中位数: {stats['50%']:.2f}")
            print(f"  标准差: {stats['std']:.2f}")
        
        print(f"\n2. 相关性分析")
        print("-" * 40)
        
        # 计算相关系数矩阵
        corr_matrix = df_clean[numeric_columns].corr()
        print("相关系数矩阵:")
        print(corr_matrix.round(3))
        
        # 金额与其他变量的相关性
        print(f"\n金额与其他变量的相关性:")
        for col in ['计件数量', '系数', '定额']:
            correlation = df_clean['金额'].corr(df_clean[col])
            print(f"  金额 vs {col}: {correlation:.3f}")
        
        print(f"\n3. 金额计算公式分析")
        print("-" * 40)
        
        # 检查可能的计算公式
        print("可能的计算公式验证:")
        
        # 公式1: 金额 = 计件数量 × 系数
        calculated_amount_1 = df_clean['计件数量'] * df_clean['系数']
        correlation_1 = df_clean['金额'].corr(calculated_amount_1)
        print(f"  金额 = 计件数量 × 系数: 相关性 = {correlation_1:.3f}")
        
        # 公式2: 金额 = 计件数量 × 定额
        calculated_amount_2 = df_clean['计件数量'] * df_clean['定额']
        correlation_2 = df_clean['金额'].corr(calculated_amount_2)
        print(f"  金额 = 计件数量 × 定额: 相关性 = {correlation_2:.3f}")
        
        # 公式3: 金额 = 计件数量 × 系数 × 定额
        calculated_amount_3 = df_clean['计件数量'] * df_clean['系数'] * df_clean['定额']
        correlation_3 = df_clean['金额'].corr(calculated_amount_3)
        print(f"  金额 = 计件数量 × 系数 × 定额: 相关性 = {correlation_3:.3f}")
        
        # 公式4: 金额 = 定额 × 系数
        calculated_amount_4 = df_clean['定额'] * df_clean['系数']
        correlation_4 = df_clean['金额'].corr(calculated_amount_4)
        print(f"  金额 = 定额 × 系数: 相关性 = {correlation_4:.3f}")
        
        print(f"\n4. 不同工序类型的金额计算模式")
        print("-" * 40)
        
        # 分析主要工序类型的金额计算模式
        top_processes = df_clean['工序全名'].value_counts().head(5).index
        print("主要工序类型的金额计算模式:")
        
        for process in top_processes:
            process_data = df_clean[df_clean['工序全名'] == process]
            if len(process_data) > 100:  # 只分析有足够数据的工序
                # 计算该工序下最可能的计算公式
                correlations = {}
                correlations['计件数量×系数'] = process_data['金额'].corr(process_data['计件数量'] * process_data['系数'])
                correlations['计件数量×定额'] = process_data['金额'].corr(process_data['计件数量'] * process_data['定额'])
                correlations['定额×系数'] = process_data['金额'].corr(process_data['定额'] * process_data['系数'])
                
                best_formula = max(correlations, key=correlations.get)
                print(f"  {process}: {len(process_data)} 条记录, 最佳公式: {best_formula} (相关性: {correlations[best_formula]:.3f})")
        
        print(f"\n5. 异常模式检测")
        print("-" * 40)
        
        # 检测不符合常见计算模式的记录
        df_clean['calculated_amount'] = df_clean['计件数量'] * df_clean['系数'] * df_clean['定额']
        df_clean['amount_diff'] = abs(df_clean['金额'] - df_clean['calculated_amount'])
        
        # 定义异常阈值
        threshold = df_clean['amount_diff'].quantile(0.95)  # 前5%的差异视为异常
        anomalies = df_clean[df_clean['amount_diff'] > threshold]
        
        print(f"异常记录数量 (金额与计算值差异较大): {len(anomalies)} 条 ({len(anomalies)/len(df_clean)*100:.2f}%)")
        
        if len(anomalies) > 0:
            print("异常记录示例:")
            sample_anomalies = anomalies.head(5)
            for idx, row in sample_anomalies.iterrows():
                print(f"  计件数量: {row['计件数量']}, 系数: {row['系数']}, 定额: {row['定额']}, 金额: {row['金额']:.2f}, 计算值: {row['calculated_amount']:.2f}, 差异: {row['amount_diff']:.2f}")
        
        print(f"\n6. 金额计算规则总结")
        print("-" * 40)
        
        # 分析最常见的金额计算模式
        print("金额计算规则分析:")
        
        # 检查金额为0的记录
        zero_amount_records = df_clean[df_clean['金额'] == 0]
        print(f"  金额为0的记录: {len(zero_amount_records)} 条")
        if len(zero_amount_records) > 0:
            print(f"    其中计件数量为0: {len(zero_amount_records[zero_amount_records['计件数量'] == 0])} 条")
            print(f"    其中系数为0: {len(zero_amount_records[zero_amount_records['系数'] == 0])} 条")
            print(f"    其中定额为0: {len(zero_amount_records[zero_amount_records['定额'] == 0])} 条")
        
        # 分析金额与各变量的散点关系
        generate_relationship_visualizations(df_clean)
        
        print(f"\n" + "=" * 80)
        print("金额关系分析完成")
        print("=" * 80)
        
    except sqlite3.Error as e:
        print(f"数据库错误: {e}")
    except Exception as e:
        print(f"发生错误: {e}")
    finally:
        if conn:
            conn.close()

def generate_relationship_visualizations(df):
    """
    生成金额关系可视化图表
    """
    try:
        print(f"\n生成关系可视化图表...")
        
        # 创建图表目录
        os.makedirs('relationship_analysis_charts', exist_ok=True)
        
        # 1. 相关系数热力图
        plt.figure(figsize=(10, 8))
        
        numeric_columns = ['计件数量', '系数', '定额', '金额']
        corr_matrix = df[numeric_columns].corr()
        
        plt.subplot(2, 2, 1)
        sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', center=0, 
                   square=True, fmt='.3f', cbar_kws={'shrink': 0.8})
        plt.title('变量相关性热力图')
        
        # 2. 金额 vs 计件数量散点图
        plt.subplot(2, 2, 2)
        sample_df = df.sample(min(1000, len(df)))  # 抽样显示，避免过度密集
        plt.scatter(sample_df['计件数量'], sample_df['金额'], alpha=0.5, s=10)
        plt.xlabel('计件数量')
        plt.ylabel('金额')
        plt.title('金额 vs 计件数量')
        
        # 3. 金额 vs 系数散点图
        plt.subplot(2, 2, 3)
        plt.scatter(sample_df['系数'], sample_df['金额'], alpha=0.5, s=10)
        plt.xlabel('系数')
        plt.ylabel('金额')
        plt.title('金额 vs 系数')
        
        # 4. 金额 vs 定额散点图
        plt.subplot(2, 2, 4)
        plt.scatter(sample_df['定额'], sample_df['金额'], alpha=0.5, s=10)
        plt.xlabel('定额')
        plt.ylabel('金额')
        plt.title('金额 vs 定额')
        
        plt.tight_layout()
        plt.savefig('relationship_analysis_charts/amount_relationships.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        # 5. 金额计算公式验证图
        plt.figure(figsize=(12, 4))
        
        # 验证金额 = 计件数量 × 系数 × 定额
        calculated_amount = df['计件数量'] * df['系数'] * df['定额']
        
        plt.subplot(1, 3, 1)
        plt.scatter(calculated_amount, df['金额'], alpha=0.3, s=10)
        plt.plot([0, df['金额'].max()], [0, df['金额'].max()], 'r--', alpha=0.8)
        plt.xlabel('计算金额 (计件数量×系数×定额)')
        plt.ylabel('实际金额')
        plt.title('金额计算公式验证')
        
        # 金额分布对比
        plt.subplot(1, 3, 2)
        plt.hist(df['金额'], bins=50, alpha=0.7, label='实际金额', color='blue')
        plt.hist(calculated_amount, bins=50, alpha=0.7, label='计算金额', color='red')
        plt.xlabel('金额')
        plt.ylabel('频次')
        plt.legend()
        plt.title('金额分布对比')
        
        # 差异分布
        plt.subplot(1, 3, 3)
        amount_diff = df['金额'] - calculated_amount
        plt.hist(amount_diff, bins=50, alpha=0.7, color='green')
        plt.xlabel('金额差异 (实际-计算)')
        plt.ylabel('频次')
        plt.title('金额计算差异分布')
        
        plt.tight_layout()
        plt.savefig('relationship_analysis_charts/amount_calculation_validation.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        print("关系分析图表已保存至: relationship_analysis_charts/")
        
    except Exception as e:
        print(f"生成图表时出错: {e}")

if __name__ == "__main__":
    analyze_amount_relationships()
