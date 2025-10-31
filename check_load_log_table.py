#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
检查SQLite数据库中的load_log表内容，并提供表结构和完整记录信息
"""

import sqlite3
import pandas as pd
import sys
import os

# Import database configuration
sys.path.append(os.path.join(os.path.dirname(__file__), 'excel_processor'))
from config import DATABASE_PATH


def print_table_with_format(df, max_rows=20):
    """
    以更好的格式打印DataFrame，确保所有列都能完整显示
    """
    # 计算终端宽度以优化显示
    try:
        # 获取终端宽度
        terminal_width = pd.util.terminal.get_terminal_size()[0]
    except:
        terminal_width = 150  # 默认宽度
    
    # 设置pandas显示选项以获得更好的格式化输出
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', 50)  # 增加列宽以显示更多内容
    pd.set_option('display.width', terminal_width)
    pd.set_option('display.float_format', '{:.2f}'.format)
    
    # 只显示前max_rows行
    display_df = df.head(max_rows)
    
    # 使用pandas的to_string方法打印表格
    print(display_df.to_string(index=False))


def get_table_structure(cursor, table_name):
    """
    获取表的结构信息
    """
    cursor.execute(f"PRAGMA table_info({table_name});")
    columns = cursor.fetchall()
    
    print("列名\t\t\t数据类型\t\t\t是否非空\t\t主键")
    print("=" * 80)
    
    for col in columns:
        column_name = col[1]
        data_type = col[2]
        not_null = "是" if col[3] else "否"
        primary_key = "是" if col[5] else "否"
        
        # 格式化输出，确保对齐
        print(f"{column_name:<20}\t{data_type:<20}\t{not_null:<10}\t{primary_key}")


def get_basic_statistics(cursor, table_name):
    """
    获取表的基本统计信息
    """
    stats = {}
    
    # 总记录数
    cursor.execute(f"SELECT COUNT(*) FROM {table_name};")
    stats['total_records'] = cursor.fetchone()[0]
    
    # 获取所有列名
    cursor.execute(f"SELECT * FROM {table_name} LIMIT 1;")
    column_names = [desc[0] for desc in cursor.description]
    
    # 打印统计信息
    print("\n基本统计信息:")
    print("统计项\t\t\t\t值")
    print("=" * 40)
    print(f"总记录数\t\t\t{stats['total_records']}")
    print(f"列数\t\t\t\t{len(column_names)}")
    print(f"列名\t\t\t\t{', '.join(column_names)}")
    
    return stats


def check_load_log_table():
    """
    连接到SQLite数据库，检查load_log表的结构和内容
    """
    table_name = "load_log"
    
    try:
        # 连接到SQLite数据库
        conn = sqlite3.connect(DATABASE_PATH)
        cursor = conn.cursor()
        
        print("成功连接到数据库")
        
        # 检查load_log表是否存在
        cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}';")
        table_exists = cursor.fetchone()
        
        if not table_exists:
            print(f"错误: {table_name}表不存在于数据库中")
            return
        
        print(f"\n=== {table_name}表信息 ===")
        
        # 获取并显示表结构
        print("\n表结构:")
        get_table_structure(cursor, table_name)
        
        # 获取并显示基本统计信息
        get_basic_statistics(cursor, table_name)
        
        # 显示所有记录（最多显示50条）
        print(f"\n{table_name}表中的记录）:")
        query = f"SELECT * FROM {table_name} ;"
        df = pd.read_sql_query(query, conn)
        
        if not df.empty:
            print_table_with_format(df, max_rows=50)
            
            # 额外显示列的非空统计信息
            print("\n各列的非空值统计:")
            print("列名\t\t\t非空值数量")
            print("=" * 30)
            for col in df.columns:
                non_null_count = df[col].count()
                print(f"{col:<20}\t{non_null_count}")
        else:
            print(f"{table_name}表中没有记录")
        
    except sqlite3.Error as e:
        print(f"数据库错误: {e}")
    except Exception as e:
        print(f"发生错误: {e}")
    finally:
        # 关闭数据库连接
        if conn:
            conn.close()
            print("\n数据库连接已关闭")


if __name__ == "__main__":
    check_load_log_table()
