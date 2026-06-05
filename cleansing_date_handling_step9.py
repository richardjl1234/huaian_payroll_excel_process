#!/usr/bin/env python3
"""
清洁工资数据库 Step 9 - 清理残留的 yy,m 与 m 前缀日期。

用途: 处理 Step 6 之后仍然残留的"年份月份"或"月份"前缀数据。
前提: Step 0-8 已将日期列规范化为逗号分隔的单日列表。

处理内容:
1. yy,m 前缀模式: 形如 '14,6,3' (file 201406.xls)
   - 当日期值前两位 (yy, m) 与文件名 YYYYMM 匹配时，去除前两位
2. m 前缀模式: 形如 '6,1' (file 201606.xls)
   - 当日期值第一位 (m) 与文件名月份匹配且剩余部分升序时，去除第一位
3. 两种都不匹配时，记录错误并保留原值
4. 仅处理当前为非升序（即会触发验证错误）的记录，不动已通过的记录
5. 包含短横线范围 (如 '1-3') 的记录会先展开再判断

用法:
    python cleansing_date_handling_step9.py [--dry-run]

参数:
    --dry-run  仅预览并导出HTML，不执行更新
    无参数    执行更新操作
"""

import sqlite3
import sys
import argparse
import re
from pathlib import Path

DB_PATH = Path(__file__).parent.parent / "payroll_database.db"
OUTPUT_PATH = Path(__file__).parent / "date_handling_step9_output.html"
PAYROLL_TABLE = "payroll_details"


def parse_file_year_month(file_name: str) -> tuple:
    """
    从文件名提取年月。例如: '201406.xls' -> (2014, 6)
    解析失败返回 (None, None)。
    """
    if not file_name or len(file_name) < 6:
        return None, None
    try:
        year = int(file_name[:4])
        month = int(file_name[4:6])
        if 1900 <= year <= 2100 and 1 <= month <= 12:
            return year, month
    except ValueError:
        pass
    return None, None


def expand_dash_range(part: str) -> list:
    """
    将 'D-D' 展开为 [start..end]。不是范围格式则尝试转为单 int。
    失败返回空列表。
    """
    part = part.strip()
    if '-' in part:
        range_parts = part.split('-')
        if len(range_parts) == 2:
            try:
                start = int(range_parts[0])
                end = int(range_parts[1])
                if 1 <= start <= 31 and 1 <= end <= 31 and start <= end:
                    return list(range(start, end + 1))
            except ValueError:
                pass
    try:
        return [int(part)]
    except ValueError:
        return []


def parse_date_list(date_str) -> list:
    """
    解析日期字符串为整数列表。展开短横线范围。
    失败或空值返回 None。
    """
    if date_str is None:
        return None
    s = str(date_str).strip()
    if not s:
        return None
    parts = [p.strip() for p in s.split(',')]
    nums = []
    for p in parts:
        if not p:
            continue
        expanded = expand_dash_range(p)
        if not expanded:
            return None
        nums.extend(expanded)
    return nums if nums else None


def is_strictly_ascending(nums: list) -> bool:
    """检查列表是否严格升序（不允许相等）"""
    for i in range(len(nums) - 1):
        if nums[i] >= nums[i + 1]:
            return False
    return True


def try_yy_m_pattern(nums: list, year: int, month: int) -> tuple:
    """
    尝试 yy,m 前缀模式。
    返回 (matched: bool, new_nums: list|None, error_msg: str|None)
    - len < 3 时直接返回 (False, None, None)，不视为错误
    """
    if len(nums) < 3:
        return False, None, None
    yy = year % 100
    if nums[0] != yy:
        return False, None, f"yy={nums[0]} ≠ 文件名年份后两位 {yy}"
    if nums[1] != month:
        return False, None, f"m={nums[1]} ≠ 文件名月份 {month}"
    return True, nums[2:], None


def try_m_pattern(nums: list, year: int, month: int) -> tuple:
    """
    尝试 m 前缀模式。
    返回 (matched: bool, new_nums: list|None, error_msg: str|None)
    - len < 2 时直接返回 (False, None, None)，不视为错误
    """
    if len(nums) < 2:
        return False, None, None
    if nums[0] != month:
        return False, None, f"m={nums[0]} ≠ 文件名月份 {month}"
    rest = nums[1:]
    if not is_strictly_ascending(rest):
        return False, None, f"剩余 {rest} 非升序"
    return True, rest, None


def is_prefix_candidate(nums: list) -> bool:
    """判断是否可能是 yy,m 或 m 前缀候选（首位 0-99 且至少 2 个值）"""
    if not nums or len(nums) < 2:
        return False
    return 0 <= nums[0] <= 99


def export_to_html(columns, yy_m_updates, m_updates, errors, output_path: Path):
    """导出处理结果到HTML"""
    display_cols = [c for c in columns if c != 'rowid']

    def render_row(row, extra_cells):
        display_row = [row[columns.index(c)] for c in display_cols]
        cells = ''.join(f'<td>{str(v) if v is not None else ""}</td>' for v in display_row)
        return f'<tr>{"".join(extra_cells)}{cells}</tr>'

    header = (
        '<tr><th>操作</th><th>原日期</th><th>新日期</th>'
        + ''.join(f'<th>{c}</th>' for c in display_cols)
        + '</tr>'
    )
    err_header = (
        '<tr><th>原日期</th><th>原因</th>'
        + ''.join(f'<th>{c}</th>' for c in display_cols)
        + '</tr>'
    )

    html_parts = ['''<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Step 9 - yy,m / m 前缀清理结果</title>
<style>
body { font-family: sans-serif; margin: 20px; }
table { border-collapse: collapse; font-size: 12px; margin-bottom: 30px; }
th, td { border: 1px solid #ddd; padding: 6px; }
th { background-color: #4472C4; color: white; position: sticky; top: 0; }
tr:nth-child(even) { background-color: #f2f2f2; }
tr:hover { background-color: #ffeb99; }
td { max-width: 250px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.section { font-size: 18px; font-weight: bold; margin-top: 30px; margin-bottom: 10px; }
.updated { background-color: #ffffcc; }
.error { background-color: #ffcccc; }
</style>
</head>
<body>
<h1>Step 9 - yy,m / m 前缀清理结果</h1>
''']

    html_parts.append(f'<div class="section">一、yy,m 前缀已更新记录（{len(yy_m_updates)} 条）</div>')
    html_parts.append(f'<table>{header}')
    for upd in yy_m_updates:
        cells = [
            '<td>yy,m→日</td>',
            f'<td>{upd["old"]}</td>',
            f'<td>{upd["new"]}</td>',
        ]
        html_parts.append(render_row(upd['row'], cells).replace('<tr>', '<tr class="updated">'))
    html_parts.append('</table>')

    html_parts.append(f'<div class="section">二、m 前缀已更新记录（{len(m_updates)} 条）</div>')
    html_parts.append(f'<table>{header}')
    for upd in m_updates:
        cells = [
            '<td>m→日</td>',
            f'<td>{upd["old"]}</td>',
            f'<td>{upd["new"]}</td>',
        ]
        html_parts.append(render_row(upd['row'], cells).replace('<tr>', '<tr class="updated">'))
    html_parts.append('</table>')

    html_parts.append(f'<div class="section">三、错误（未匹配任何前缀，保留原值，{len(errors)} 条）</div>')
    html_parts.append(f'<table>{err_header}')
    for err in errors:
        cells = [
            f'<td>{err["old"]}</td>',
            f'<td>{err["reason"]}</td>',
        ]
        html_parts.append(render_row(err['row'], cells).replace('<tr>', '<tr class="error">'))
    html_parts.append('</table>')

    html_parts.append('</body></html>')

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(''.join(html_parts))


def main():
    parser = argparse.ArgumentParser(
        description="清洁工资数据库 Step 9 - 清理残留的 yy,m 与 m 前缀日期"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="仅预览并导出HTML，不执行更新"
    )
    args = parser.parse_args()

    if not DB_PATH.exists():
        print(f"错误: 数据库文件不存在: {DB_PATH}")
        sys.exit(1)

    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.execute(f"SELECT rowid, * FROM {PAYROLL_TABLE}")
    columns = [desc[0] for desc in cursor.description]
    all_rows = cursor.fetchall()

    date_idx = columns.index('日期')
    filename_idx = columns.index('文件名')
    rowid_idx = 0

    print(f"数据库共有 {len(all_rows)} 条记录")

    yy_m_updates = []
    m_updates = []
    error_rows = []
    skipped = 0

    for row in all_rows:
        rowid = row[rowid_idx]
        date_val = row[date_idx]
        file_name = row[filename_idx]

        if date_val is None or str(date_val).strip() == '':
            skipped += 1
            continue

        nums = parse_date_list(date_val)
        if nums is None:
            skipped += 1
            continue

        # 仅处理当前为非升序的记录（即验证失败的）
        if is_strictly_ascending(nums):
            skipped += 1
            continue

        year, month = parse_file_year_month(file_name)
        if year is None:
            skipped += 1
            continue

        old_str = str(date_val).strip()

        # 先尝试 yy,m 模式
        matched, new_nums, ym_err = try_yy_m_pattern(nums, year, month)
        if matched:
            new_str = ','.join(str(n) for n in new_nums)
            yy_m_updates.append({
                'old': old_str, 'new': new_str, 'row': row, 'rowid': rowid,
            })
            continue

        # 再尝试 m 模式
        matched, new_nums, m_err = try_m_pattern(nums, year, month)
        if matched:
            new_str = ','.join(str(n) for n in new_nums)
            m_updates.append({
                'old': old_str, 'new': new_str, 'row': row, 'rowid': rowid,
            })
            continue

        # 都不匹配：若是前缀候选则记为错误
        if is_prefix_candidate(nums):
            reason = ym_err or m_err or "未匹配任何前缀模式"
            error_rows.append({
                'old': old_str, 'reason': reason, 'row': row, 'rowid': rowid,
            })
        else:
            skipped += 1

    total_updates = len(yy_m_updates) + len(m_updates)
    print(f"\n【更新】")
    print(f"  yy,m 模式: {len(yy_m_updates)} 条")
    print(f"  m 模式:    {len(m_updates)} 条")
    print(f"  合计:      {total_updates} 条")
    print(f"\n【错误】保留原值: {len(error_rows)} 条")
    print(f"\n【跳过】无需处理: {skipped} 条")

    # 模式统计
    if total_updates > 0:
        patterns = {}
        for upd in yy_m_updates + m_updates:
            key = upd['old']
            if key not in patterns:
                patterns[key] = {'count': 0, 'new': upd['new']}
            patterns[key]['count'] += 1
        print("\n更新模式统计（前20）:")
        for old_pat, info in sorted(patterns.items(), key=lambda x: x[1]['count'], reverse=True)[:20]:
            print(f"  '{old_pat}' → '{info['new']}': {info['count']} 条")

    if error_rows:
        print("\n错误模式统计:")
        err_patterns = {}
        for err in error_rows:
            key = (err['old'], err['reason'])
            err_patterns[key] = err_patterns.get(key, 0) + 1
        for (old, reason), count in sorted(err_patterns.items(), key=lambda x: x[1], reverse=True)[:20]:
            print(f"  '{old}': {reason} — {count} 条")

    # 导出 HTML
    export_to_html(columns, yy_m_updates, m_updates, error_rows, OUTPUT_PATH)
    print(f"\n已导出到: {OUTPUT_PATH}")

    # 执行或 dry-run
    if args.dry_run:
        print(f"\n[DRY-RUN 模式] 未执行任何数据库操作。")
    else:
        confirm = input(
            f"\n[确认] 即将更新 {total_updates} 条记录（{len(error_rows)} 条错误保留原值）。\n"
            f"请输入 'yes' 确认: "
        )
        if confirm.strip().lower() != "yes":
            print("已取消操作。")
            conn.close()
            sys.exit(0)

        for upd in yy_m_updates + m_updates:
            conn.execute(
                f"UPDATE {PAYROLL_TABLE} SET 日期 = ? WHERE rowid = ?",
                (upd['new'], upd['rowid'])
            )
        conn.commit()
        print(f"已更新 {total_updates} 条记录。")

        cursor = conn.execute(f"SELECT COUNT(*) FROM {PAYROLL_TABLE}")
        final_count = cursor.fetchone()[0]
        print(f"操作完成后数据库共有 {final_count} 条记录。")

        # 验证：检查是否还有可修复的残留
        cursor = conn.execute(
            f"SELECT rowid, 日期, 文件名 FROM {PAYROLL_TABLE}"
        )
        remaining = cursor.fetchall()
        remaining_bad = []
        for rid, date, fname in remaining:
            if date is None or str(date).strip() == '':
                continue
            nums = parse_date_list(date)
            if nums is None or len(nums) < 2:
                continue
            if is_strictly_ascending(nums):
                continue
            year, month = parse_file_year_month(fname)
            if year is None:
                continue
            yy = year % 100
            # 模式 1
            if len(nums) >= 3 and nums[0] == yy and nums[1] == month:
                remaining_bad.append((rid, date, fname, "yy,m"))
            # 模式 2
            elif nums[0] == month and is_strictly_ascending(nums[1:]):
                remaining_bad.append((rid, date, fname, "m"))

        if remaining_bad:
            print(f"\n警告: 仍有 {len(remaining_bad)} 条记录残留可修复的前缀模式:")
            for rid, date, fname, kind in remaining_bad[:10]:
                print(f"  rowid={rid}, 日期='{date}', 文件={fname} ({kind} 模式)")
            if len(remaining_bad) > 10:
                print(f"  ... 还有 {len(remaining_bad) - 10} 条")
        else:
            print("\n验证通过：无可修复的 yy,m/m 前缀模式残留。")

    conn.close()


if __name__ == "__main__":
    main()
