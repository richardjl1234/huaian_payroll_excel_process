#!/usr/bin/env python3
"""
清洁工资数据库中的异常记录 - Step 3。

删除包含"合计"的汇总行，这些行是各职员的小计/合计记录，不属于个人工资明细。

用法:
    python cleansing_outliers_step3.py [--dry-run]

参数:
    --dry-run  仅导出到Excel，不执行删除
    无参数    执行删除操作
"""

import sqlite3
import sys
import argparse
from pathlib import Path
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

DB_PATH = Path(__file__).parent.parent / "payroll_database.db"
OUTPUT_PATH = Path(__file__).parent / "outliers_to_be_deleted_step3.html"
PAYROLL_TABLE = "payroll_details"


def get_rows_with_合计(conn: sqlite3.Connection) -> tuple:
    """获取所有包含'合计'的记录。"""
    cursor = conn.execute(f"SELECT rowid, * FROM {PAYROLL_TABLE}")
    columns = [description[0] for description in cursor.description]
    all_rows = cursor.fetchall()

    text_cols = [c for c in columns if c not in ['rowid', '计件数量', '系数', '定额', '金额']]
    conditions = ' OR '.join([f"{col} LIKE '%合计%'" for col in text_cols])

    cursor = conn.execute(f"SELECT rowid, * FROM {PAYROLL_TABLE} WHERE {conditions}")
    columns = [description[0] for description in cursor.description]
    rows = cursor.fetchall()
    return columns, rows


def export_to_html(columns: list, rows: list, output_path: Path):
    """导出包含'合计'的记录到HTML文件。"""
    html = f'''<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Outliers to be deleted - 包含合计的记录</title>
<style>
table {{ border-collapse: collapse; font-size: 12px; }}
th, td {{ border: 1px solid #ddd; padding: 8px; }}
th {{ background-color: #4472C4; color: white; position: sticky; top: 0; }}
tr:nth-child(even) {{ background-color: #f2f2f2; }}
tr:hover {{ background-color: #ddd; }}
td {{ max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
</style>
</head>
<body>
<h1>Found {len(rows)} rows containing 合计</h1>
<table>
<tr>{''.join(f'<th>{c}</th>' for c in columns)}</tr>
'''
    for row in rows:
        html += '<tr>' + ''.join(f'<td>{str(v) if v is not None else ""}</td>' for v in row) + '</tr>'

    html += '''
</table>
</body>
</html>'''

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)


def delete_rows(conn: sqlite3.Connection, rows: list) -> int:
    """删除指定记录，返回删除的行数。"""
    if not rows:
        return 0
    rowids = [row[0] for row in rows]
    placeholders = ",".join("?" * len(rowids))
    cursor = conn.execute(
        f"DELETE FROM {PAYROLL_TABLE} WHERE rowid IN ({placeholders})",
        rowids
    )
    conn.commit()
    return cursor.rowcount


def summarize_by_file(columns: list, rows: list) -> dict:
    """按文件名汇总统计。"""
    filename_idx = columns.index("文件名")
    summary = {}
    for row in rows:
        fname = row[filename_idx]
        summary[fname] = summary.get(fname, 0) + 1
    return summary


def main():
    parser = argparse.ArgumentParser(
        description="清洁工资数据库中的异常记录 - Step 3 (删除含'合计'的汇总行)"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="仅导出到HTML，不执行删除"
    )
    args = parser.parse_args()

    if not DB_PATH.exists():
        print(f"错误: 数据库文件不存在: {DB_PATH}")
        sys.exit(1)

    conn = sqlite3.connect(str(DB_PATH))

    columns, rows = get_rows_with_合计(conn)

    if not rows:
        print("未找到包含'合计'的记录。")
        conn.close()
        sys.exit(0)

    print(f"找到 {len(rows)} 条包含'合计'的记录。")

    # 按文件名统计
    by_file = summarize_by_file(columns, rows)
    print("\n按文件名分布:")
    for fname, count in sorted(by_file.items(), key=lambda x: x[1], reverse=True):
        print(f"  - {fname}: {count}")

    # 导出到HTML
    export_to_html(columns, rows, OUTPUT_PATH)
    print(f"\n已导出到: {OUTPUT_PATH}")

    def count_table_records(conn):
        cursor = conn.execute(f"SELECT COUNT(*) FROM {PAYROLL_TABLE}")
        return cursor.fetchone()[0]

    if args.dry_run:
        print("\n[DRY-RUN 模式] 未执行删除操作。")
    else:
        before_count = count_table_records(conn)
        print(f"\n[确认删除] 数据库当前共有 {before_count} 条记录")
        print(f"[确认删除] 即将删除 {len(rows)} 条包含'合计'的记录...")
        confirm = input("请输入 'yes' 确认删除: ")
        if confirm.strip().lower() == "yes":
            deleted_count = delete_rows(conn, rows)
            after_count = count_table_records(conn)
            print(f"已成功删除 {deleted_count} 条记录。")
            print(f"删除后数据库共有 {after_count} 条记录。")
        else:
            print("已取消删除操作。")

    conn.close()


if __name__ == "__main__":
    main()