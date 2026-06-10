#!/usr/bin/env python3
"""
清洁工资数据库中的日期列 - Step 6。

处理复杂混合模式（Case 6），这些模式在 Step 5 中被跳过。
前提：Step 0 已将所有全角字符转换为半角。

处理模式:
1. YY.M.D-DD 或 YYYY.M.D-DD 格式：提取年月，展开日范围
2. 点号分隔的多个部分：每个部分可能是单日或范围
3. 括号休息日标记：(X休息) 或 (X休) - 从结果中移除X
4. 双短横线规范：-- → -
5. 逗号分隔列表（含范围）：1,2,6-10 → 1,2,6,7,8,9,10

用法:
    python cleansing_date_handling_step6.py [--dry-run]

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
OUTPUT_PATH = Path(__file__).parent / "date_handling_step6_output.html"
PAYROLL_TABLE = "payroll_details"


def parse_file_year_month(file_name: str) -> tuple:
    """从文件名提取年月。例如: '201801.xls' -> (2018, 1)"""
    if len(file_name) >= 6:
        try:
            year = int(file_name[:4])
            month = int(file_name[4:6])
            return year, month
        except ValueError:
            return None, None
    return None, None


def extract_rest_days(date_str: str) -> tuple:
    """提取休息日信息，返回 (cleaned_str, rest_days)"""
    rest_days = []
    # 匹配 (数字休息) 或 (数字休)
    rest_match = re.search(r'\((\d+)\s*(?:休息?|休)\)', date_str)
    if rest_match:
        rest_days = [int(rest_match.group(1))]
        date_str = re.sub(r'\([^)]+\)', '', date_str)
    return date_str, rest_days


def expand_dash_range_in_list(parts: list, rest_days: list) -> list:
    """展开列表中的短横线范围，如 ['1-4', '6'] -> ['1', '2', '3', '4', '6']"""
    result = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            range_parts = part.split('-')
            if len(range_parts) == 2:
                try:
                    start = int(range_parts[0].strip())
                    end = int(range_parts[1].strip())
                    if 1 <= start <= 31 and 1 <= end <= 31 and start <= end:
                        for d in range(start, end + 1):
                            if d not in rest_days:
                                result.append(str(d))
                    else:
                        result.append(part)
                except ValueError:
                    result.append(part)
            else:
                result.append(part)
        else:
            if part.isdigit():
                day = int(part)
                if 1 <= day <= 31 and day not in rest_days:
                    result.append(part)
            else:
                result.append(part)
    return result


def expand_yyyymm_d_format(date_str: str, file_year: int, file_month: int, rest_days: list) -> str:
    """
    处理 YY.M.D-DD 格式，如 '19.3.1-10' -> '1,2,3,4,5,6,7,8,9,10'
    或 '18.11.1-3' -> '1,2,3'
    也处理 '11.12.14-20' -> '14,15,16,17,18,19,20'
    """
    # 首先规范化双短横线
    date_str = re.sub(r'--+', '-', date_str)

    # 模式: YY.M.D-DD (例如 19.3.1-10)
    pattern1 = r'^(\d{2})\.(\d+)\.(\d+)-(\d+)$'
    match = re.match(pattern1, date_str)
    if match:
        yy = int(match.group(1))
        mm = int(match.group(2))
        start_day = int(match.group(3))
        end_day = int(match.group(4))

        year = 2000 + yy if yy < 100 else yy

        if year == file_year and mm == file_month:
            if 1 <= start_day <= 31 and 1 <= end_day <= 31 and start_day <= end_day:
                result = [str(d) for d in range(start_day, end_day + 1) if d not in rest_days]
                return ','.join(result)
        return date_str

    # 模式: YY.MM.D-DD (例如 11.12.14-20)
    pattern2 = r'^(\d{2})\.(\d{2})\.(\d+)-(\d+)$'
    match = re.match(pattern2, date_str)
    if match:
        yy = int(match.group(1))
        mm = int(match.group(2))
        start_day = int(match.group(3))
        end_day = int(match.group(4))

        year = 2000 + yy if yy < 100 else yy

        if year == file_year and mm == file_month:
            if 1 <= start_day <= 31 and 1 <= end_day <= 31 and start_day <= end_day:
                result = [str(d) for d in range(start_day, end_day + 1) if d not in rest_days]
                return ','.join(result)

    return date_str


def expand_yyyymm_dd_format(date_str: str, file_year: int, file_month: int, rest_days: list) -> str:
    """
    处理 YYYY.M.D-DD 或 YYYY.M.D-DD 格式，如 '2018.5.2-5.7' -> '2,3,4,5,6,7'
    """
    # 模式: YYYY.M.D-DD 或 YYYY.MM.D-DD
    pattern = r'^(\d{4})\.(\d{1,2})\.(\d+)-(\d+)\.(\d+)$'
    match = re.match(pattern, date_str)
    if match:
        year = int(match.group(1))
        month = int(match.group(2))
        start_day = int(match.group(3))
        end_month = int(match.group(4))
        end_day = int(match.group(5))

        # 如果开始月和结束月不同，不处理（跨月份太复杂）
        if year == file_year and month == file_month and month == end_month:
            if 1 <= start_day <= 31 and 1 <= end_day <= 31 and start_day <= end_day:
                result = [str(d) for d in range(start_day, end_day + 1) if d not in rest_days]
                return ','.join(result)

    return date_str


def expand_dot_separated_list(date_str: str, rest_days: list) -> str:
    """
    处理点号分隔的列表，如 '1-7.9.10' -> '1,2,3,4,5,6,7,9,10'
    每个部分可能是:
    - 单日: '9', '10'
    - 范围: '1-7', '13-17'
    """
    if '.' not in date_str:
        return date_str

    parts = date_str.split('.')
    result = []

    for part in parts:
        part = part.strip()
        if not part:
            continue

        if '-' in part:
            # 范围
            range_parts = part.split('-')
            if len(range_parts) == 2:
                try:
                    start = int(range_parts[0].strip())
                    end = int(range_parts[1].strip())
                    if 1 <= start <= 31 and 1 <= end <= 31 and start <= end:
                        for d in range(start, end + 1):
                            if d not in rest_days:
                                result.append(str(d))
                    else:
                        result.append(part)
                except ValueError:
                    result.append(part)
            else:
                result.append(part)
        else:
            if part.isdigit():
                day = int(part)
                if 1 <= day <= 31 and day not in rest_days:
                    result.append(part)

    return ','.join(result) if result else date_str


def expand_comma_separated_with_range(date_str: str, rest_days: list) -> str:
    """
    处理逗号分隔的列表（含范围），如 '1,2,6-10' -> '1,2,6,7,8,9,10'
    """
    if ',' not in date_str:
        return date_str

    parts = date_str.split(',')
    result = expand_dash_range_in_list(parts, rest_days)
    return ','.join(result)


def normalize_double_dash(date_str: str) -> str:
    """规范化双短横线，如 '11.12.14--20' -> '11.12.14-20'"""
    return re.sub(r'--+', '-', date_str)


def expand_complex_date(date_str: str, file_name: str) -> tuple:
    """
    将复杂日期模式展开为逗号分隔的列表。
    返回: (expanded_str, error_msg)
    """
    original = date_str

    # 从文件名提取年月
    file_year, file_month = parse_file_year_month(file_name)
    if file_year is None:
        return date_str, f"无法从文件名 '{file_name}' 提取年月"

    # 提取休息日信息
    date_str, rest_days = extract_rest_days(date_str)

    # 规范化双短横线（提前处理，避免干扰后续匹配）
    date_str = normalize_double_dash(date_str)

    # 检查是否是 YY.M.D-DD 格式 (先规范化再匹配)
    # 注意：只有当月份部分是1-9时才匹配，避免错误匹配如 '11.12.14-20' 这种点号分隔列表
    if re.match(r'^\d{2}\.[1-9]\.\d+-\d+$', date_str):
        result = expand_yyyymm_d_format(date_str, file_year, file_month, rest_days)
        if result != original:
            return result, None
        return date_str, None

    # 检查是否是 YYYY.M.D-DD 格式
    if re.match(r'^\d{4}\.\d+\.\d+-\d+\.\d+$', date_str):
        result = expand_yyyymm_dd_format(date_str, file_year, file_month, rest_days)
        if result != original:
            return result, None
        return date_str, None

    # 如果有点号，按点号分隔处理
    if '.' in date_str:
        result = expand_dot_separated_list(date_str, rest_days)
        if result != original:
            return result, None

    # 如果有逗号，按逗号分隔处理
    if ',' in date_str:
        result = expand_comma_separated_with_range(date_str, rest_days)
        if result != original:
            return result, None

    # 如果有短横线但没有点号，按短横线范围处理
    if '-' in date_str and '.' not in date_str:
        result = ','.join(expand_dash_range_in_list([date_str], rest_days))
        if result != original:
            return result, None

    return date_str, None


def export_to_html(columns: list, rows_with_changes: list, output_path: Path):
    """导出更新记录到HTML文件。"""
    html = f'''<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Step 6 - 日期处理结果</title>
<style>
table {{ border-collapse: collapse; font-size: 12px; }}
th, td {{ border: 1px solid #ddd; padding: 8px; }}
th {{ background-color: #4472C4; color: white; position: sticky; top: 0; }}
tr:nth-child(even) {{ background-color: #f2f2f2; }}
tr:hover {{ background-color: #ddd; }}
td {{ max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
.error {{ color: red; font-weight: bold; }}
</style>
</head>
<body>
<h1>Step 6 - 复杂日期处理结果</h1>
<p>共 {len(rows_with_changes)} 条记录被更新</p>
<table>
<tr><th>原日期</th><th>新日期</th><th>错误</th>''' + ''.join(f'<th>{c}</th>' for c in columns) + '''</tr>
'''
    for old_date, new_date, error, row in rows_with_changes:
        html += '<tr>'
        html += f'<td>{old_date}</td>'
        html += f'<td>{new_date}</td>'
        error_display = error if error else ''
        html += f'<td class="error">{error_display}</td>'
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
        description="清洁工资数据库中的日期列 - Step 6 (复杂模式处理)"
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
    print(f"开始处理复杂日期模式...")

    # 处理每条记录
    updates = []
    errors = []
    skipped = 0
    updated = 0

    rows_with_changes = []

    for row in all_rows:
        rowid = row[0]
        date_idx = columns.index('日期')
        file_idx = columns.index('文件名')
        old_date = row[date_idx]
        file_name = row[file_idx]

        if old_date is None or str(old_date).strip() == '':
            skipped += 1
            continue

        old_date_str = str(old_date).strip()

        # 展开复杂日期
        new_date, error = expand_complex_date(old_date_str, file_name)

        if new_date != old_date_str:
            updates.append((new_date, rowid))
            rows_with_changes.append((old_date_str, new_date, error, row))
            updated += 1

        if error:
            errors.append((rowid, old_date_str, new_date, error))

    print(f"\n处理完成:")
    print(f"  - 已更新: {updated}")
    print(f"  - 有错误/警告: {len(errors)}")

    # 按模式统计
    patterns = {}
    for old_date, new_date, error, row in rows_with_changes:
        key = old_date
        if key not in patterns:
            patterns[key] = {'count': 0, 'new': new_date, 'error': error}
        patterns[key]['count'] += 1

    print(f"\n模式统计 (前20):")
    sorted_patterns = sorted(patterns.items(), key=lambda x: x[1]['count'], reverse=True)
    for old_pat, info in sorted_patterns[:20]:
        print(f"  '{old_pat}' -> '{info['new']}': {info['count']} records")

    # 打印错误信息
    if errors:
        print(f"\n=== 错误/警告信息 ({len(errors)} 条) ===")
        for rowid, old_date, new_date, error in errors[:50]:
            print(f"  rowid={rowid}, 日期='{old_date}' -> '{new_date}', 错误: {error}")
        if len(errors) > 50:
            print(f"  ... 还有 {len(errors) - 50} 条错误未显示")

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