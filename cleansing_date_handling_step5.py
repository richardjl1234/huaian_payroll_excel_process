#!/usr/bin/env python3
"""
清洁工资数据库中的日期列 - Step 5。

将各种日期模式统一转换为逗号分隔的列表格式。
例如: '1、2' -> '1,2', '1～5' -> '1,2,3,4,5'

处理模式:
- Case 2: 逗号分隔 (、或,) -> '1、2' -> '1,2'
- Case 3: 点号分隔 (.) -> '1.2.3' -> '1,2,3'
- Case 4: 波浪号范围 (～) -> '1～5' -> '1,2,3,4,5'
- Case 5: 短横线范围 (-) -> '1-5' -> '1,2,3,4,5'

注意: Case 6 (复杂混合模式) 不在本步骤处理，将在 Step 6 处理。

用法:
    python cleansing_date_handling_step5.py [--dry-run]

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
OUTPUT_PATH = Path(__file__).parent / "date_handling_step5_output.html"
PAYROLL_TABLE = "payroll_details"


def parse_comma_range(date_str: str) -> str:
    """
    处理逗号分隔的日期: '1、2' 或 '1,2'
    返回逗号分隔的格式: '1,2'
    """
    s = str(date_str).strip()

    # 中文逗号 或 英文逗号
    if '、' in s:
        parts = s.split('、')
    elif ',' in s:
        parts = s.split(',')
    else:
        return s  # 没有逗号，保持原样

    # 清理每个部分并过滤空值
    cleaned_parts = []
    for part in parts:
        part = part.strip()
        if part:
            cleaned_parts.append(part)

    return ','.join(cleaned_parts)


def parse_dot_list(date_str: str) -> str:
    """
    处理点号分隔的日期列表: '1.2.3' 或 '26.27'
    返回逗号分隔的格式: '1,2,3' 或 '26,27'

    复杂模式如 '14.6.3' (3个部分, 表示完整日期) 会被跳过, 在Step 6处理
    """
    s = str(date_str).strip()

    # 点号分隔
    if '.' not in s:
        return s

    parts = s.split('.')
    cleaned_parts = []
    for part in parts:
        part = part.strip()
        if part:
            cleaned_parts.append(part)

    # 如果有3个部分: [X, Y, Z]
    # 检查是否符合完整日期模式: 月(1-12), 日(1-31)
    # 例如: '14.6.3' -> 14是无效年份部分, 6是月, 3是日
    # 这种模式需要文件名上下文来确定年份, 跳过在Step 6处理
    if len(cleaned_parts) == 3:
        try:
            part1 = cleaned_parts[0]   # 可能是无效年份部分 (如 14, 18)
            part2 = int(cleaned_parts[1])  # 可能是月
            part3 = cleaned_parts[2]   # 可能是日

            # 检查是否符合日期模式: 月(1-12), 日(1-31)
            if (1 <= part2 <= 12 and
                part3.isdigit() and 1 <= int(part3) <= 31):
                # 这是完整日期格式，需要文件名上下文，返回None表示跳过
                return None  # Step 6 will handle this with filename context
        except (ValueError, IndexError):
            pass

    # 两部分点号分隔: '1.2.3.4' 或 '26.27' -> 逗号分隔列表
    return ','.join(cleaned_parts)


def parse_tilde_range(date_str: str) -> str:
    """
    处理波浪号范围: '1～5' -> '1,2,3,4,5'
    也处理 '15.12.2～4' -> '15.12.2,15.12.3,15.12.4' 格式
    """
    s = str(date_str).strip()

    if '～' not in s:
        return s

    parts = s.split('～')
    if len(parts) != 2:
        return s  # 格式不对，保持原样

    left_part = parts[0].strip()
    right_part = parts[1].strip()

    try:
        # 检查左侧是否包含点号（YYYY.M.D 格式）
        if '.' in left_part:
            left_subparts = left_part.split('.')
            # 提取日期部分（最后一个点后的部分）
            if len(left_subparts) >= 3:
                prefix = '.'.join(left_subparts[:-1])  # '15.12'
                day_start = int(left_subparts[-1])      # 2
            else:
                prefix = left_part
                day_start = 0

            day_end = int(right_part)
            if day_start < 1 or day_end > 31 or day_start > day_end:
                return s  # 无效范围，保持原样

            days = list(range(day_start, day_end + 1))
            return ','.join(f'{prefix}.{d}' for d in days)
        else:
            # 简单数字范围
            start = int(left_part)
            end = int(right_part)

            if start < 1 or end > 31 or start > end:
                return s  # 无效范围，保持原样

            days = list(range(start, end + 1))
            return ','.join(str(d) for d in days)
    except ValueError:
        return s  # 解析失败，保持原样


def parse_dash_range(date_str: str) -> str:
    """
    处理短横线范围: '1-5' 或 '11--20' 或 '1-4、6' 或 '1-4、6、8'
    '11--20' 应该被规范化为 '11-20'
    '1-4、6' -> '1,2,3,4,6'
    '1-4、6、8' -> '1,2,3,4,6,8'
    """
    s = str(date_str).strip()

    if '-' not in s:
        return s

    # 先替换中文逗号为英文逗号
    if '、' in s:
        s = s.replace('、', ',')

    # 处理双短横线情况: '11--20' -> '11-20'
    s = re.sub(r'--+', '-', s)

    # 去除首尾的短横线
    s = s.strip('-')

    if '-' not in s:
        return s

    # 处理混合模式: '1-4,6' 或 '1-4,6,8'
    # 先按逗号分割，展开每个部分，然后合并
    if ',' in s:
        parts = s.split(',')
        expanded_parts = []
        for part in parts:
            part = part.strip()
            if '-' in part:
                # 这是个范围，如 '1-4'
                range_parts = part.split('-')
                if len(range_parts) == 2:
                    try:
                        start = int(range_parts[0].strip())
                        end = int(range_parts[1].strip())
                        if 1 <= start <= 31 and 1 <= end <= 31 and start <= end:
                            days = list(range(start, end + 1))
                            expanded_parts.extend(str(d) for d in days)
                        else:
                            expanded_parts.append(part)
                    except ValueError:
                        expanded_parts.append(part)
                else:
                    expanded_parts.append(part)
            else:
                # 单独的日子
                try:
                    day = int(part)
                    if 1 <= day <= 31:
                        expanded_parts.append(str(day))
                    else:
                        expanded_parts.append(part)
                except ValueError:
                    expanded_parts.append(part)
        return ','.join(expanded_parts)

    # 简单情况: '1-5'
    parts = s.split('-')
    if len(parts) != 2:
        return s  # 格式不对，保持原样

    try:
        start = int(parts[0].strip())
        end = int(parts[1].strip())

        if start < 1 or end > 31 or start > end:
            return s  # 无效范围，保持原样

        days = list(range(start, end + 1))
        return ','.join(str(d) for d in days)
    except ValueError:
        return s  # 解析失败，保持原样


def is_complex_pattern(date_str: str) -> bool:
    """
    判断是否为复杂混合模式 (Case 6)
    复杂模式: 同时包含多种分隔符，或者包含点号和短横线的组合等
    """
    s = str(date_str).strip()

    # 点号和波浪号同时存在
    # 检查波浪号左侧部分是否包含2个或更多点号 (YYYY.M.D 格式)
    # 例如 '15.12.2～4' 左侧 '15.12.2' 有2个点号，不是复杂模式
    # 例如 '1.2～5' 左侧 '1.2' 只有1个点号，是复杂模式
    if '.' in s and '～' in s:
        tilde_idx = s.index('～')
        left_part = s[:tilde_idx]
        if left_part.count('.') >= 2:
            return False  # Not complex - this is YYYY.M.D～D format
        else:
            return True  # Complex - this could be M.D～D format

    # 中文逗号和短横线同时存在
    # 如果没有点号（如 '1、2、6-10'），可以处理：先替换逗号，再处理短横线范围
    # 如果同时有点号（如 '18.1.2、3'），才是真正的复杂模式
    if '、' in s and '-' in s:
        if '.' in s:
            return True  # Complex - has both comma and dash and dots
        else:
            return False  # NOT complex - we can handle this

    # 中文逗号和波浪号同时存在 - 复杂
    if '、' in s and '～' in s:
        return True

    # 如果同时包含多种分隔符，视为复杂模式
    delimiters = []
    if '、' in s: delimiters.append('、')
    if '，' in s: delimiters.append('，')  # 中文逗号
    if ',' in s: delimiters.append(',')
    if '.' in s: delimiters.append('.')
    if '-' in s: delimiters.append('-')
    if '～' in s: delimiters.append('～')

    # 多种分隔符组合
    if len(delimiters) >= 2:
        return True

    return False


def needs_expansion(date_str: str) -> bool:
    """
    判断日期是否需要展开
    Case 1: 单个日期 (如 '7', '22') - 不需要展开
    Case 2-5: 需要展开的模式 - 返回 True
    Case 6: 复杂模式 - 返回 False (Step 6处理)
    """
    s = str(date_str).strip()

    # 单个日期 (纯数字 1-31)
    if s.isdigit() and 1 <= int(s) <= 31:
        return False

    # 逗号分隔
    if '、' in s or ',' in s:
        # 检查是否是简单逗号分隔
        if is_complex_pattern(s):
            return False
        return True

    # 波浪号范围
    if '～' in s:
        if is_complex_pattern(s):
            return False
        return True

    # 短横线范围
    if '-' in s:
        if is_complex_pattern(s):
            return False
        return True

    # 点号分隔
    if '.' in s:
        if is_complex_pattern(s):
            return False
        return True

    return False


def expand_date(date_str: str) -> str:
    """
    将日期模式展开为逗号分隔的列表
    """
    s = str(date_str).strip()

    # 短横线范围 (Case 5) - 优先处理，因为可能和逗号混合
    if '-' in s and not is_complex_pattern(s):
        result = parse_dash_range(s)
        if result != s:
            return result

    # 波浪号范围 (Case 4)
    if '～' in s and not is_complex_pattern(s):
        return parse_tilde_range(s)

    # 逗号分隔 (Case 2) - 包括中文逗号
    if '、' in s or (',' in s and not is_complex_pattern(s)):
        return parse_comma_range(s)

    # 点号列表 (Case 3)
    if '.' in s and not is_complex_pattern(s):
        result = parse_dot_list(s)
        if result is None:
            return None  # 表示需要Step 6处理
        return result

    # 其他情况保持原样
    return s


def export_to_html(columns: list, rows_with_changes: list, output_path: Path):
    """导出更新记录到HTML文件。"""
    html = f'''<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Step 5 - 日期处理结果</title>
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
<h1>Step 5 - 日期处理结果</h1>
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
        description="清洁工资数据库中的日期列 - Step 5"
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
    print(f"开始处理日期列...")

    # 处理每条记录
    updates = []
    skipped_complex = 0
    skipped_single = 0
    updated = 0

    rows_with_changes = []

    for row in all_rows:
        rowid = row[0]
        date_idx = columns.index('日期')
        old_date = row[date_idx]

        if old_date is None or str(old_date).strip() == '':
            continue

        old_date_str = str(old_date).strip()

        # 检查是否需要展开
        if not needs_expansion(old_date_str):
            if old_date_str.isdigit() and 1 <= int(old_date_str) <= 31:
                skipped_single += 1
            else:
                skipped_complex += 1
            continue

        # 展开日期
        new_date = expand_date(old_date_str)

        # 如果返回None，表示是复杂日期模式(如14.6.3)，需要Step 6处理
        if new_date is None:
            skipped_complex += 1
            continue

        if new_date != old_date_str:
            updates.append((new_date, rowid))
            rows_with_changes.append((old_date_str, new_date, row))
            updated += 1

    print(f"\n处理完成:")
    print(f"  - 单个日期跳过: {skipped_single}")
    print(f"  - 复杂模式跳过 (Step 6处理): {skipped_complex}")
    print(f"  - 已更新: {updated}")

    # 按模式统计
    patterns = {}
    for old_date, new_date, row in rows_with_changes:
        if old_date not in patterns:
            patterns[old_date] = {'count': 0, 'new': new_date}
        patterns[old_date]['count'] += 1

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