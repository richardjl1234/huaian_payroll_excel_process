#!/usr/bin/env python3
"""
清洁工资数据库中的日期列 - Step 7。

处理剩余的波浪号(~)模式，将波浪号范围展开为逗号分隔的列表。
这些是 Step 0 将全角波浪号(~)转换为半角后遗留的模式。

处理模式:
- `16~31` -> `16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31`
- `1~4,15~29` -> `1,2,3,4,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29`
- `10-11-12` -> `10,11,12` (短横线分隔的非范围模式)

用法:
    python cleansing_data_handling_step7.py [--dry-run]

参数:
    --dry-run  仅预览，不执行更新
    无参数    执行更新操作
"""

import sqlite3
import sys
import argparse
import re
from pathlib import Path

DB_PATH = Path(__file__).parent.parent / "payroll_database.db"
OUTPUT_PATH = Path(__file__).parent / "date_handling_step7_output.html"
PAYROLL_TABLE = "payroll_details"


def expand_tilde_range(date_str: str) -> str:
    """
    处理波浪号范围: '16~31' -> '16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31'
    同时处理混合模式: '1~4,15~29' -> '1,2,3,4,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29'
    """
    s = str(date_str).strip()

    if '~' not in s:
        return s

    # 按逗号分割，每部分单独处理
    if ',' in s:
        parts = s.split(',')
        result_parts = []
        for part in parts:
            part = part.strip()
            if '~' in part:
                range_match = re.match(r'^(\d+)~(\d+)$', part)
                if range_match:
                    start = int(range_match.group(1))
                    end = int(range_match.group(2))
                    if 1 <= start <= 31 and 1 <= end <= 31 and start <= end:
                        days = list(range(start, end + 1))
                        result_parts.append(','.join(str(d) for d in days))
                    else:
                        result_parts.append(part)
                else:
                    result_parts.append(part)
            else:
                # 非波浪号部分，可能是数字或范围
                result_parts.append(part)
        return ','.join(result_parts)
    else:
        # 没有逗号，纯波浪号范围
        range_match = re.match(r'^(\d+)~(\d+)$', s)
        if range_match:
            start = int(range_match.group(1))
            end = int(range_match.group(2))
            if 1 <= start <= 31 and 1 <= end <= 31 and start <= end:
                days = list(range(start, end + 1))
                return ','.join(str(d) for d in days)
        return s


def expand_dash_list(date_str: str) -> str:
    """
    处理短横线分隔的非范围模式: '10-11-12' -> '10,11,12'
    这种模式的特点是有多个短横线但不是范围（如 10-12 是范围，10-11-12 不是）
    """
    s = str(date_str).strip()

    if '-' not in s:
        return s

    # 如果有逗号，先按逗号分割
    if ',' in s:
        parts = s.split(',')
        result_parts = []
        for part in parts:
            part = part.strip()
            if '-' in part:
                # 检查是否是范围模式（只有一个短横线）
                dash_count = part.count('-')
                if dash_count == 1:
                    # 范围模式，如 10-12
                    range_match = re.match(r'^(\d+)-(\d+)$', part)
                    if range_match:
                        start = int(range_match.group(1))
                        end = int(range_match.group(2))
                        if 1 <= start <= 31 and 1 <= end <= 31 and start <= end:
                            days = list(range(start, end + 1))
                            result_parts.append(','.join(str(d) for d in days))
                        else:
                            result_parts.append(part)
                    else:
                        result_parts.append(part)
                else:
                    # 多个短横线，如 10-11-12
                    dash_parts = part.split('-')
                    valid_days = []
                    for dp in dash_parts:
                        dp = dp.strip()
                        if dp.isdigit():
                            day = int(dp)
                            if 1 <= day <= 31:
                                valid_days.append(str(day))
                    if valid_days:
                        result_parts.append(','.join(valid_days))
                    else:
                        result_parts.append(part)
            else:
                result_parts.append(part)
        return ','.join(result_parts)
    else:
        # 没有逗号
        dash_count = s.count('-')
        if dash_count == 1:
            # 范围模式，已在其他步骤处理
            return s
        elif dash_count > 1:
            # 多个短横线，如 10-11-12
            dash_parts = s.split('-')
            valid_days = []
            for dp in dash_parts:
                dp = dp.strip()
                if dp.isdigit():
                    day = int(dp)
                    if 1 <= day <= 31:
                        valid_days.append(str(day))
            if valid_days:
                return ','.join(valid_days)
        return s


def needs_processing(date_str: str) -> bool:
    """检查是否需要处理"""
    s = str(date_str).strip()

    # 有波浪号
    if '~' in s:
        return True

    # 有多个短横线（不是范围模式）
    if '-' in s:
        dash_count = s.count('-')
        # 如果有逗号，检查逗号分隔的部分
        if ',' in s:
            # 如果每个逗号分隔的部分只有1个短横线，则是范围模式
            parts = s.split(',')
            for part in parts:
                if part.count('-') > 1:
                    return True
        elif dash_count > 1:
            return True

    return False


def expand_date(date_str: str) -> str:
    """展开日期模式"""
    s = str(date_str).strip()

    # 先处理波浪号
    if '~' in s:
        s = expand_tilde_range(s)

    # 再处理短横线分隔的非范围模式
    if '-' in s:
        s = expand_dash_list(s)

    return s


def export_to_html(columns: list, rows_with_changes: list, output_path: Path):
    """导出更新记录到HTML文件。"""
    html = f'''<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Step 7 - 日期处理结果</title>
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
<h1>Step 7 - 波浪号和短横线分隔符处理结果</h1>
<p>共 {len(rows_with_changes)} 条记录被更新</p>
<table>
<tr><th>原日期</th><th>新日期</th>''' + ''.join(f'<th>{c}</th>' for c in columns) + '''</tr>
'''
    for old_date, new_date, row in rows_with_changes:
        html += '<tr>'
        html += f'<td>{old_date}</td>'
        html += f'<td>{new_date}</td>'
        html += ''.join(f'<td>{str(v) if v is not None else ""}</td>' for v in row)
        html += '</tr>'

    html += '''
</table>
</body>
</html>'''

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)


def main():
    parser = argparse.ArgumentParser(
        description="清洁工资数据库中的日期列 - Step 7 (波浪号和短横线分隔符处理)"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="仅预览，不执行更新"
    )
    args = parser.parse_args()

    if not DB_PATH.exists():
        print(f"错误: 数据库文件不存在: {DB_PATH}")
        sys.exit(1)

    conn = sqlite3.connect(str(DB_PATH))

    # 获取所有记录
    cursor = conn.execute(f"SELECT rowid, * FROM {PAYROLL_TABLE}")
    columns = [desc[0] for desc in cursor.description]
    all_rows = cursor.fetchall()

    print(f"数据库共有 {len(all_rows)} 条记录")
    print(f"开始处理波浪号和短横线分隔符...")

    # 处理每条记录
    updates = []
    skipped = 0
    updated = 0

    rows_with_changes = []

    for row in all_rows:
        rowid = row[0]
        date_idx = columns.index('日期')
        old_date = row[date_idx]

        if old_date is None or str(old_date).strip() == '':
            skipped += 1
            continue

        old_date_str = str(old_date).strip()

        # 检查是否需要处理
        if not needs_processing(old_date_str):
            skipped += 1
            continue

        # 展开日期
        new_date_str = expand_date(old_date_str)

        if new_date_str != old_date_str:
            updates.append((new_date_str, rowid))
            rows_with_changes.append((old_date_str, new_date_str, row))
            updated += 1
        else:
            skipped += 1

    print(f"\n处理完成:")
    print(f"  - 已更新: {updated}")
    print(f"  - 跳过: {skipped}")

    # 按模式统计
    patterns = {}
    for old_date, new_date, row in rows_with_changes:
        key = old_date
        if key not in patterns:
            patterns[key] = {'count': 0, 'new': new_date}
        patterns[key]['count'] += 1

    print(f"\n模式统计 (前20):")
    sorted_patterns = sorted(patterns.items(), key=lambda x: x[1]['count'], reverse=True)
    for old_pat, info in sorted_patterns[:20]:
        print(f"  '{old_pat}' -> '{info['new']}': {info['count']} records")

    # 导出到HTML
    export_to_html(columns, rows_with_changes, OUTPUT_PATH)
    print(f"\n已导出到: {OUTPUT_PATH}")

    if args.dry_run:
        print(f"\n[DRY-RUN 模式] 预览: 本次可更新 {len(updates)} 条记录。")
    else:
        def count_table_records(conn):
            cursor = conn.execute(f"SELECT COUNT(*) FROM {PAYROLL_TABLE}")
            return cursor.fetchone()[0]

        before_count = count_table_records(conn)
        print(f"\n[确认更新] 数据库当前共有 {before_count} 条记录")
        print(f"[确认更新] 即将更新 {len(updates)} 条日期记录...")

        confirm = input("请输入 'yes' 确认更新: ")
        if confirm.strip().lower() == "yes":
            # 执行更新
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