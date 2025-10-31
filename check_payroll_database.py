#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
检查SQLite数据库中的payroll_details表内容，并提供多种查询选项
"""

import sqlite3
import pandas as pd
import os
import sys

# Import database configuration
sys.path.append(os.path.join(os.path.dirname(__file__), 'excel_processor'))
from config import DATABASE_PATH


def print_table_with_format(df, max_rows=10):
    """
    以更好的格式打印DataFrame
    """
    # 设置pandas显示选项以获得更好的格式化输出
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', 20)
    pd.set_option('display.width', 1000)
    pd.set_option('display.float_format', '{:.2f}'.format)
    
    # 只显示前max_rows行
    display_df = df.head(max_rows)
    
    # 使用pandas的to_string方法打印表格
    print(display_df.to_string(index=False))


def get_table_structure(cursor):
    """
    获取表的结构信息
    """
    cursor.execute("PRAGMA table_info(payroll_details);")
    columns = cursor.fetchall()
    
    print("列名\t\t数据类型\t\t是否非空\t主键")
    print("-" * 60)
    
    for col in columns:
        column_name = col[1]
        data_type = col[2]
        not_null = "是" if col[3] else "否"
        primary_key = "是" if col[5] else "否"
        
        # 格式化输出，确保对齐
        print(f"{column_name:<15}\t{data_type:<15}\t{not_null:<8}\t{primary_key}")


def get_basic_statistics(cursor):
    """
    获取表的基本统计信息
    """
    stats = {}
    
    # 总记录数
    cursor.execute("SELECT COUNT(*) FROM payroll_details;")
    stats['total_records'] = cursor.fetchone()[0]
    
    # 不同文件名数量
    cursor.execute("SELECT COUNT(DISTINCT 文件名) FROM payroll_details;")
    stats['unique_files'] = cursor.fetchone()[0]
    
    # 不同sheet名数量
    cursor.execute("SELECT COUNT(DISTINCT sheet名) FROM payroll_details;")
    stats['unique_sheets'] = cursor.fetchone()[0]
    
    # 不同职员数量
    cursor.execute("SELECT COUNT(DISTINCT 职员全名) FROM payroll_details;")
    stats['unique_employees'] = cursor.fetchone()[0]
    
    # 不同型号数量
    cursor.execute("SELECT COUNT(DISTINCT 型号) FROM payroll_details;")
    stats['unique_models'] = cursor.fetchone()[0]
    
    # 金额统计
    cursor.execute("SELECT MIN(金额), MAX(金额), AVG(金额), SUM(金额) FROM payroll_details;")
    min_amount, max_amount, avg_amount, sum_amount = cursor.fetchone()
    stats['min_amount'] = min_amount
    stats['max_amount'] = max_amount
    stats['avg_amount'] = avg_amount
    stats['sum_amount'] = sum_amount
    
    # 打印统计信息
    print("\n基本统计信息:")
    print("统计项\t\t\t值")
    print("-" * 40)
    print(f"总记录数\t\t{stats['total_records']}")
    print(f"不同文件数\t\t{stats['unique_files']}")
    print(f"不同Sheet数\t\t{stats['unique_sheets']}")
    print(f"不同职员数\t\t{stats['unique_employees']}")
    print(f"不同型号数\t\t{stats['unique_models']}")
    print(f"最小金额\t\t{stats['min_amount']:.2f}")
    print(f"最大金额\t\t{stats['max_amount']:.2f}")
    print(f"平均金额\t\t{stats['avg_amount']:.2f}")
    print(f"总金额\t\t{stats['sum_amount']:.2f}")
    
    return stats


def query_by_employee(cursor, employee_name):
    """
    按职员查询记录
    """
    query = """
    SELECT * FROM payroll_details 
    WHERE 职员全名 = ?
    LIMIT 20;
    """
    
    return pd.read_sql_query(query, cursor.connection, params=(employee_name,))


def query_by_date(cursor, date):
    """
    按日期查询记录
    """
    query = """
    SELECT * FROM payroll_details 
    WHERE 日期 = ?
    LIMIT 20;
    """
    
    return pd.read_sql_query(query, cursor.connection, params=(date,))


def query_by_model(cursor, model):
    """
    按型号查询记录
    """
    query = """
    SELECT * FROM payroll_details 
    WHERE 型号 LIKE ?
    LIMIT 20;
    """
    
    return pd.read_sql_query(query, cursor.connection, params=(f"%{model}%",))


def check_payroll_details_table():
    """
    连接到SQLite数据库，检查payroll_details表的结构和内容
    并提供交互式查询功能
    """
    try:
        # 连接到SQLite数据库
        conn = sqlite3.connect(DATABASE_PATH)
        cursor = conn.cursor()
        
        print("成功连接到数据库")
        
        # 检查payroll_details表是否存在
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='payroll_details';")
        table_exists = cursor.fetchone()
        
        if not table_exists:
            print("错误: payroll_details表不存在于数据库中")
            return
        
        print("\n=== payroll_details表信息 ===")
        
        # 获取并显示表结构
        print("\n表结构:")
        get_table_structure(cursor)
        
        # 获取并显示基本统计信息
        stats = get_basic_statistics(cursor)
        
        # 显示随机的10条记录
        print("\n随机10条记录:")
        random_query = """
        SELECT * FROM payroll_details 
        ORDER BY RANDOM() 
        LIMIT 10;
        """
        random_df = pd.read_sql_query(random_query, conn)
        print_table_with_format(random_df)
        
        # 显示最近的10条记录（按文件名排序，假设文件名包含日期信息）
        print("\n最近的10条记录（按文件名排序）:")
        recent_query = """
        SELECT * FROM payroll_details 
        ORDER BY 文件名 DESC 
        LIMIT 10;
        """
        recent_df = pd.read_sql_query(recent_query, conn)
        print_table_with_format(recent_df)
        
        # 交互式查询功能
        while True:
            print("\n=== 交互式查询菜单 ===")
            print("1. 显示所有记录（限前50条）")
            print("2. 按职员姓名查询")
            print("3. 按日期查询")
            print("4. 按型号查询")
            print("0. 退出")
            
            choice = input("请选择操作 (0-4): ")
            
            if choice == '0':
                print("退出查询")
                break
            elif choice == '1':
                df = pd.read_sql_query("SELECT * FROM payroll_details LIMIT 50;", conn)
                print_table_with_format(df, max_rows=50)
            elif choice == '2':
                employee = input("请输入职员姓名: ")
                df = query_by_employee(cursor, employee)
                if not df.empty:
                    print(f"\n找到 {len(df)} 条关于 '{employee}' 的记录:")
                    print_table_with_format(df)
                else:
                    print(f"未找到关于 '{employee}' 的记录")
            elif choice == '3':
                date = input("请输入日期: ")
                df = query_by_date(cursor, date)
                if not df.empty:
                    print(f"\n找到 {len(df)} 条关于日期 '{date}' 的记录:")
                    print_table_with_format(df)
                else:
                    print(f"未找到关于日期 '{date}' 的记录")
            elif choice == '4':
                model = input("请输入型号关键词: ")
                df = query_by_model(cursor, model)
                if not df.empty:
                    print(f"\n找到 {len(df)} 条包含型号 '{model}' 的记录:")
                    print_table_with_format(df)
                else:
                    print(f"未找到包含型号 '{model}' 的记录")
            else:
                print("无效的选择，请重试")
        
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
    check_payroll_details_table()
