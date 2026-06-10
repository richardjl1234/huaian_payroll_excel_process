#!/usr/bin/env python3
"""
清洁工资数据库中的异常记录 - Step 4。

将郁俊海的月份日期记录（如"4月"、"三月"、"7月份"）转换为该月份的第一个工作日。
工作日判断：周一至周五，且不是中国法定节假日。

用法:
    python cleansing_outliers_step4.py [--dry-run]

参数:
    --dry-run  仅预览，不执行更新
    无参数    执行更新操作
"""

import sqlite3
import sys
import argparse
import re
from pathlib import Path
from datetime import date, timedelta

try:
    import holidays
    HAS_HOLIDAYS = True
except ImportError:
    HAS_HOLIDAYS = False

DB_PATH = Path(__file__).parent.parent / "payroll_database.db"
OUTPUT_PATH = Path(__file__).parent / "outliers_to_be_updated_step4.html"
PAYROLL_TABLE = "payroll_details"


def export_to_html(columns: list, rows_with_new_date: list, output_path: Path):
    """导出需要更新的记录到HTML文件。"""
    # Get all columns from payroll_details
    html = f'''<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Outliers to be updated - Step 4 月份日期转工作日</title>
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
<h1>Step 4 - 郁俊海月份日期记录更新</h1>
<p>共 {len(rows_with_new_date)} 条记录将被更新</p>
<table>
<tr><th>原日期</th><th>新日期</th><th>说明</th>''' + ''.join(f'<th>{c}</th>' for c in columns) + '''</tr>
'''
    for old_date, new_date_str, note, row in rows_with_new_date:
        html += '<tr>'
        html += f'<td>{old_date}</td>'
        html += f'<td>{new_date_str}</td>'
        html += f'<td>{note}</td>'
        html += ''.join(f'<td>{str(v) if v is not None else ""}</td>' for v in row)
        html += '</tr>'

    html += '''
</table>
</body>
</html>'''

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)


def get_china_holidays(year: int):
    """获取指定年份的中国法定节假日。"""
    if not HAS_HOLIDAYS:
        return set()
    return holidays.China(years=year)


def get_first_working_day(year: int, month: int) -> date:
    """
    获取指定年月的第一个工作日。
    工作日 = 周一至周五 且 不是法定节假日
    """
    cn_holidays = get_china_holidays(year)

    # Start from the first day of the month
    current = date(year, month, 1)

    # If it's a holiday, move to next day
    # Keep going until we find a weekday (Mon-Fri) that's not a holiday
    max_days = 31  # Safety limit

    for _ in range(max_days):
        # Check if it's a weekday (0=Monday, 5=Saturday, 6=Sunday)
        if current.weekday() < 5 and current not in cn_holidays:
            return current
        current += timedelta(days=1)

    return current  # Fallback


def parse_month_pattern(date_str: str) -> tuple:
    """
    解析月份日期模式，返回 (year, month)。
    支持的模式: 三月, 4月, 5月, 6月, 7月份 等
    """
    date_str = str(date_str).strip()

    # Chinese month names
    chinese_months = {
        '一月': 1, '二月': 2, '三月': 3, '四月': 4, '五月': 5,
        '六月': 6, '七月': 7, '八月': 8, '九月': 9, '十月': 10,
        '十一月': 11, '十二月': 12
    }

    for cn_name, month_num in chinese_months.items():
        if date_str == cn_name or date_str == cn_name + '份':
            return month_num

    # Numeric month like "4月", "5月", "6月份"
    match = re.match(r'(\d+)月?', date_str)
    if match:
        return int(match.group(1))

    return None


def get_target_date(file_name: str, date_str: str) -> str:
    """
    根据文件名和日期字符串计算目标日期。
    返回格式: 该月份第一个工作日的"日"（即天数的字符串，如"3"）
    """
    # Extract year from file name like "202504.xls"
    year = int(file_name[:4])

    month = parse_month_pattern(date_str)
    if month is None:
        return None

    first_working_day = get_first_working_day(year, month)
    return str(first_working_day.day)  # Return just the day number


def get_records_to_update_full(conn: sqlite3.Connection) -> tuple:
    """获取所有需要更新的郁俊海的月份记录的完整信息。"""
    cursor = conn.execute(f"SELECT rowid, * FROM {PAYROLL_TABLE} WHERE 1=0")
    columns = [desc[0] for desc in cursor.description]

    cursor = conn.execute(f"""
        SELECT rowid, * FROM {PAYROLL_TABLE}
        WHERE 职员全名 = '郁俊海'
        AND (日期 LIKE '%月' OR 日期 LIKE '%月份')
    """)
    rows = cursor.fetchall()
    return columns, rows


def main():
    parser = argparse.ArgumentParser(
        description="清洁工资数据库中的异常记录 - Step 4 (月份日期转工作日)"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="仅预览更新结果，不执行更新"
    )
    args = parser.parse_args()

    if not DB_PATH.exists():
        print(f"错误: 数据库文件不存在: {DB_PATH}")
        sys.exit(1)

    if not HAS_HOLIDAYS:
        print("警告: holidays 模块未安装，将使用简单的周末判断（不考虑法定节假日）")
        print("      请运行: pip install holidays")

    conn = sqlite3.connect(str(DB_PATH))

    columns, rows = get_records_to_update_full(conn)

    if not rows:
        print("未找到需要更新的月份记录。")
        conn.close()
        sys.exit(0)

    print(f"找到 {len(rows)} 条郁俊海的月份记录需要更新。")
    print(f"\n当前 holidays 模块: {'已安装' if HAS_HOLIDAYS else '未安装'}")

    # Calculate updates and prepare HTML data
    updates = []
    rows_with_new_date = []

    print("\n" + "=" * 80)
    print(f"{'文件名':<12} {'原日期':<8} {'新日期':<12} {'说明'}")
    print("=" * 80)

    for row in rows:
        rowid = row[0]
        fname = row[columns.index('文件名')]
        old_date = row[columns.index('日期')]

        new_date = get_target_date(fname, old_date)

        if new_date:
            year = int(fname[:4])
            month = parse_month_pattern(old_date)
            first_wd = get_first_working_day(year, month)
            note = first_wd.strftime('%Y-%m-%d') + " (first working day)"

            print(f"{fname:<12} {old_date:<8} {new_date:<12} {note}")
            updates.append((new_date, rowid))
            rows_with_new_date.append((old_date, new_date, note, row))
        else:
            print(f"{fname:<12} {old_date:<8} {'N/A':<12} 无法解析月份")
            rows_with_new_date.append((old_date, "N/A", "无法解析月份", row))

    print("=" * 80)

    # Export to HTML
    export_to_html(columns, rows_with_new_date, OUTPUT_PATH)
    print(f"\n已导出到: {OUTPUT_PATH}")

    if args.dry_run:
        print(f"\n[DRY-RUN 模式] 预览: 本次可更新 {len(updates)} 条记录。")
    else:
        def count_table_records(conn):
            cursor = conn.execute(f"SELECT COUNT(*) FROM {PAYROLL_TABLE}")
            return cursor.fetchone()[0]

        before_count = count_table_records(conn)
        print(f"\n[确认更新] 数据库当前共有 {before_count} 条记录")
        print(f"[确认更新] 即将更新 {len(updates)} 条月份记录...")

        confirm = input("请输入 'yes' 确认更新: ")
        if confirm.strip().lower() == "yes":
            # Execute updates
            for new_date, rowid in updates:
                conn.execute(
                    f"UPDATE {PAYROLL_TABLE} SET 日期 = ? WHERE rowid = ?",
                    (new_date, rowid)
                )
            conn.commit()

            after_count = count_table_records(conn)
            print(f"已成功更新 {len(updates)} 条记录。")
            print(f"更新后数据库共有 {after_count} 条记录。")
        else:
            print("已取消更新操作。")

    conn.close()


if __name__ == "__main__":
    main()