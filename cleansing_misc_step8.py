#!/usr/bin/env python3
"""
清洁工资数据库 Step 8 - 杂项清理（最终步骤）。

处理内容:
1. 删除包含非日期文本的记录（24条）：'日期'字样、工序数据（嵌线/排线/电机/绕线/铣）
2. 空格分隔的日期展开：空格作为分隔符，同时展开短横线范围
   '18-21 23' → '18,19,20,21,23'
3. 处理 '15-21(20休' 模式：展开范围并移除休息日
   '15-21(20休' → '15,16,17,18,19,21'
4. 处理 '10。11' 模式：将中文句号替换为逗号
   '10。11' → '10,11'
5. 处理 '1(加班)' 模式：提取日期数字
   '1(加班)' → '1'
6. 处理 '&' 分隔符：替换为逗号
   '10&12' → '10,12'

用法:
    python cleansing_misc_step8.py [--dry-run]

参数:
    --dry-run  仅预览并导出HTML，不执行操作
    无参数    执行删除和更新操作
"""

import sqlite3
import sys
import argparse
import re
from pathlib import Path

DB_PATH = Path(__file__).parent.parent / "payroll_database.db"
OUTPUT_PATH = Path(__file__).parent / "misc_step8_output.html"
PAYROLL_TABLE = "payroll_details"


def has_chinese(text: str) -> bool:
    """检查字符串是否包含中文字符"""
    return bool(re.search(r'[一-鿿]', str(text)))


def is_pure_garbage_date(date_val: str) -> bool:
    """
    判断是否为纯垃圾日期值（非日期数据，应删除）。
    这些记录不是日期数据，而是其他字段数据错误地写入日期列。
    """
    s = str(date_val).strip()
    # 包含中文字符且不是可识别的日期模式
    if not has_chinese(s):
        return False

    # 可识别的日期模式（不删除）
    # 'X-Y(Z休' - 范围+休息日标记
    if re.match(r'^\d+-\d+\(\d+休?$', s):
        return False
    # '1(加班)' - 加班标记
    if re.match(r'^\d+\(加班\)$', s):
        return False

    # 其他包含中文的都是垃圾数据
    return True


def expand_space_separated(date_str: str) -> str:
    """
    将空格分隔的日期转换为逗号分隔，同时展开范围。

    '18-21 23'  → '18,19,20,21,23'
    '10 11 13'  → '10,11,13'
    '7 9-11'    → '7,9,10,11'
    '28  31'    → '28,31'
    """
    parts = str(date_str).strip().split()
    result = []
    for part in parts:
        if '-' in part and part.count('-') == 1:
            bounds = part.split('-')
            if bounds[0].isdigit() and bounds[1].isdigit():
                s, e = int(bounds[0]), int(bounds[1])
                if 1 <= s <= 31 and 1 <= e <= 31 and s <= e:
                    result.extend(str(d) for d in range(s, e + 1))
                    continue
        result.append(part)
    return ','.join(result)


def expand_rest_day_range(date_str: str) -> str:
    """
    处理 'X-Y(Z休' 模式：展开范围并移除休息日。

    '15-21(20休' → '15,16,17,18,19,21'
    含义：15-21日，其中20日休息，因此排除20
    """
    s = str(date_str).strip()
    m = re.match(r'^(\d+)-(\d+)\((\d+)休?$', s)
    if not m:
        return s

    start, end, rest_day = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if not (1 <= start <= 31 and 1 <= end <= 31 and start <= end):
        return s
    if not (1 <= rest_day <= 31):
        return s

    days = [str(d) for d in range(start, end + 1) if d != rest_day]
    return ','.join(days) if days else s


def remove_overtime_marker(date_str: str) -> str:
    """
    处理 '1(加班)' 模式：提取日期数字。
    '1(加班)' → '1'
    """
    s = str(date_str).strip()
    m = re.match(r'^(\d+)\(加班\)$', s)
    if m:
        return m.group(1)
    return s


def replace_chinese_period(date_str: str) -> str:
    """
    将中文句号 。替换为逗号。
    '10。11' → '10,11'
    """
    return str(date_str).replace('。', ',')


def replace_ampersand(date_str: str) -> str:
    """
    将 '&' 分隔符替换为逗号。
    '10&12' → '10,12'
    '7,1&3' → '7,1,3'
    """
    return str(date_str).replace('&', ',')


def needs_space_expansion(date_val: str) -> bool:
    """检查是否需要空格分隔处理"""
    s = str(date_val).strip()
    if ' ' not in s:
        return False
    # 确认确实有可处理的数字内容
    parts = s.split()
    for part in parts:
        clean = part.replace('-', '')
        if clean.isdigit():
            return True
    return False


def needs_rest_day_expansion(date_val: str) -> bool:
    """检查是否是 'X-Y(Z休' 模式"""
    return bool(re.match(r'^\d+-\d+\(\d+休?$', str(date_val).strip()))


def needs_overtime_cleanup(date_val: str) -> bool:
    """检查是否是 'D(加班)' 模式"""
    return bool(re.match(r'^\d+\(加班\)$', str(date_val).strip()))


def needs_chinese_period_fix(date_val: str) -> bool:
    """检查是否包含中文句号"""
    return '。' in str(date_val)


def needs_ampersand_fix(date_val: str) -> bool:
    """检查是否包含 & 分隔符"""
    return '&' in str(date_val)


def export_to_html(columns: list, delete_rows: list, update_details: list, output_path: Path):
    """导出处理结果到HTML文件。"""
    # 构建列名列表（不含rowid）
    display_cols = [c for c in columns if c != 'rowid']

    html = f'''<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Step 8 - 杂项清理结果</title>
<style>
body {{ font-family: sans-serif; margin: 20px; }}
table {{ border-collapse: collapse; font-size: 12px; margin-bottom: 30px; }}
th, td {{ border: 1px solid #ddd; padding: 6px; }}
th {{ background-color: #4472C4; color: white; position: sticky; top: 0; }}
tr:nth-child(even) {{ background-color: #f2f2f2; }}
tr:hover {{ background-color: #ddd; }}
td {{ max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
.section {{ font-size: 18px; font-weight: bold; margin-top: 30px; margin-bottom: 10px; }}
.deleted {{ background-color: #ffcccc; }}
.updated {{ background-color: #ffffcc; }}
</style>
</head>
<body>
<h1>Step 8 - 杂项清理结果</h1>

<div class="section">一、已删除的非日期文本记录（{len(delete_rows)} 条）</div>
<table>
<tr>{"".join(f'<th>{c}</th>' for c in display_cols)}</tr>
'''
    for row in delete_rows:
        display_row = [row[columns.index(c)] for c in display_cols]
        html += '<tr class="deleted">' + ''.join(
            f'<td>{str(v) if v is not None else ""}</td>' for v in display_row
        ) + '</tr>'

    html += '''
</table>

<div class="section">二、已更新的空格分隔日期记录（''' + str(len([u for u in update_details if u['type'] == 'space'])) + ''' 条）</div>
<table>
<tr><th>操作</th><th>原日期</th><th>新日期</th>''' + "".join(f'<th>{c}</th>' for c in display_cols) + '''</tr>
'''
    for upd in update_details:
        if upd['type'] != 'space':
            continue
        row = upd['row']
        display_row = [row[columns.index(c)] for c in display_cols]
        html += '<tr class="updated">'
        html += f'<td>空格→逗号</td>'
        html += f'<td>{upd["old"]}</td>'
        html += f'<td>{upd["new"]}</td>'
        html += ''.join(f'<td>{str(v) if v is not None else ""}</td>' for v in display_row)
        html += '</tr>'

    html += '''
</table>

<div class="section">三、已更新的休息日标记记录（''' + str(len([u for u in update_details if u['type'] == 'rest_day'])) + ''' 条）</div>
<table>
<tr><th>操作</th><th>原日期</th><th>新日期</th>''' + "".join(f'<th>{c}</th>' for c in display_cols) + '''</tr>
'''
    for upd in update_details:
        if upd['type'] != 'rest_day':
            continue
        row = upd['row']
        display_row = [row[columns.index(c)] for c in display_cols]
        html += '<tr class="updated">'
        html += f'<td>休息日展开</td>'
        html += f'<td>{upd["old"]}</td>'
        html += f'<td>{upd["new"]}</td>'
        html += ''.join(f'<td>{str(v) if v is not None else ""}</td>' for v in display_row)
        html += '</tr>'

    html += '''
</table>

<div class="section">四、&分隔符替换记录（''' + str(len([u for u in update_details if u['type'] == 'ampersand'])) + ''' 条）</div>
<table>
<tr><th>操作</th><th>原日期</th><th>新日期</th>''' + "".join(f'<th>{c}</th>' for c in display_cols) + '''</tr>
'''
    for upd in update_details:
        if upd['type'] != 'ampersand':
            continue
        row = upd['row']
        display_row = [row[columns.index(c)] for c in display_cols]
        html += '<tr class="updated">'
        html += f'<td>&rarr;逗号</td>'
        html += f'<td>{upd["old"]}</td>'
        html += f'<td>{upd["new"]}</td>'
        html += ''.join(f'<td>{str(v) if v is not None else ""}</td>' for v in display_row)
        html += '</tr>'

    html += '''
</table>

<div class="section">五、其他更新（''' + str(len([u for u in update_details if u['type'] == 'other'])) + ''' 条）</div>
<table>
<tr><th>操作</th><th>原日期</th><th>新日期</th>''' + "".join(f'<th>{c}</th>' for c in display_cols) + '''</tr>
'''
    for upd in update_details:
        if upd['type'] != 'other':
            continue
        row = upd['row']
        display_row = [row[columns.index(c)] for c in display_cols]
        html += '<tr class="updated">'
        html += f'<td>其他</td>'
        html += f'<td>{upd["old"]}</td>'
        html += f'<td>{upd["new"]}</td>'
        html += ''.join(f'<td>{str(v) if v is not None else ""}</td>' for v in display_row)
        html += '</tr>'

    html += '''
</table>
</body>
</html>'''

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)


def summarize_by_file(columns: list, rows: list) -> dict:
    """按文件名汇总统计"""
    filename_idx = columns.index("文件名")
    summary = {}
    for row in rows:
        fname = row[filename_idx]
        summary[fname] = summary.get(fname, 0) + 1
    return summary


def main():
    parser = argparse.ArgumentParser(
        description="清洁工资数据库 Step 8 - 杂项清理"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="仅预览并导出HTML，不执行操作"
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
    rowid_idx = 0  # rowid is first column

    print(f"数据库共有 {len(all_rows)} 条记录")

    # =========================================
    # Part 1: 识别并删除非日期文本记录
    # =========================================
    delete_rows = []
    for row in all_rows:
        date_val = row[date_idx]
        if date_val is not None and is_pure_garbage_date(str(date_val)):
            delete_rows.append(row)

    print(f"\n【删除】找到 {len(delete_rows)} 条非日期文本记录")
    if delete_rows:
        by_file = summarize_by_file(columns, delete_rows)
        print("按文件名分布:")
        for fname, count in sorted(by_file.items(), key=lambda x: x[1], reverse=True):
            print(f"  - {fname}: {count}")

    # =========================================
    # Part 2: 识别并转换需要更新的记录
    # =========================================
    update_details = []  # list of dicts: {type, old, new, row, rowid}

    for row in all_rows:
        date_val = row[date_idx]
        if date_val is None or str(date_val).strip() == '':
            continue

        rowid = row[rowid_idx]
        old_date = str(date_val).strip()
        new_date = old_date

        # 2a. 空格分隔日期
        if needs_space_expansion(new_date):
            new_date = expand_space_separated(new_date)

        # 2b. 休息日标记 'X-Y(Z休'
        if needs_rest_day_expansion(new_date):
            new_date = expand_rest_day_range(new_date)

        # 2c. 加班标记 'D(加班)'
        if needs_overtime_cleanup(new_date):
            new_date = remove_overtime_marker(new_date)

        # 2d. 中文句号
        if needs_chinese_period_fix(new_date):
            new_date = replace_chinese_period(new_date)

        # 2e. & 分隔符替换为逗号
        if needs_ampersand_fix(new_date):
            new_date = replace_ampersand(new_date)

        if new_date != old_date:
            # 确定类型
            if needs_space_expansion(old_date):
                upd_type = 'space'
            elif needs_rest_day_expansion(old_date):
                upd_type = 'rest_day'
            elif needs_ampersand_fix(old_date):
                upd_type = 'ampersand'
            else:
                upd_type = 'other'

            update_details.append({
                'type': upd_type,
                'old': old_date,
                'new': new_date,
                'row': row,
                'rowid': rowid
            })

    # 按类型统计
    space_count = len([u for u in update_details if u['type'] == 'space'])
    rest_day_count = len([u for u in update_details if u['type'] == 'rest_day'])
    ampersand_count = len([u for u in update_details if u['type'] == 'ampersand'])
    other_count = len([u for u in update_details if u['type'] == 'other'])

    print(f"\n【更新】共 {len(update_details)} 条记录需要更新:")
    print(f"  - 空格分隔日期: {space_count}")
    print(f"  - 休息日标记(X-Y(Z休): {rest_day_count}")
    print(f"  - &分隔符: {ampersand_count}")
    print(f"  - 其他(加班/句号): {other_count}")

    # 模式统计
    if update_details:
        patterns = {}
        for upd in update_details:
            key = upd['old']
            if key not in patterns:
                patterns[key] = {'count': 0, 'new': upd['new']}
            patterns[key]['count'] += 1

        print("\n更新模式统计:")
        sorted_pats = sorted(patterns.items(), key=lambda x: x[1]['count'], reverse=True)
        for old_pat, info in sorted_pats[:20]:
            print(f"  '{old_pat}' → '{info['new']}': {info['count']} 条")

    # =========================================
    # 导出到HTML
    # =========================================
    export_to_html(columns, delete_rows, update_details, OUTPUT_PATH)
    print(f"\n已导出到: {OUTPUT_PATH}")

    # =========================================
    # 执行操作（非 --dry-run 模式）
    # =========================================
    if args.dry_run:
        print(f"\n[DRY-RUN 模式] 未执行任何操作。")
    else:
        confirm_all = input(f"\n[确认] 即将删除 {len(delete_rows)} 条记录并更新 {len(update_details)} 条记录。\n请输入 'yes' 确认: ")
        if confirm_all.strip().lower() != "yes":
            print("已取消操作。")
            conn.close()
            sys.exit(0)

        # 执行删除
        if delete_rows:
            delete_rowids = [row[rowid_idx] for row in delete_rows]
            placeholders = ",".join("?" * len(delete_rowids))
            conn.execute(
                f"DELETE FROM {PAYROLL_TABLE} WHERE rowid IN ({placeholders})",
                delete_rowids
            )
            conn.commit()
            print(f"已删除 {len(delete_rows)} 条记录。")

        # 执行更新
        if update_details:
            for upd in update_details:
                conn.execute(
                    f"UPDATE {PAYROLL_TABLE} SET 日期 = ? WHERE rowid = ?",
                    (upd['new'], upd['rowid'])
                )
            conn.commit()
            print(f"已更新 {len(update_details)} 条记录。")

        # 最终统计
        cursor = conn.execute(f"SELECT COUNT(*) FROM {PAYROLL_TABLE}")
        final_count = cursor.fetchone()[0]
        print(f"操作完成后数据库共有 {final_count} 条记录。")

        # 验证：检查是否还有残留
        cursor = conn.execute(f"SELECT rowid, * FROM {PAYROLL_TABLE}")
        remaining = cursor.fetchall()
        remaining_garbage = 0
        for row in remaining:
            date_val = row[date_idx]
            if date_val is not None:
                s = str(date_val).strip()
                if has_chinese(s) and not re.match(r'^\d+-\d+\(\d+休?$', s) and not re.match(r'^\d+\(加班\)$', s):
                    remaining_garbage += 1
                if '。' in s:
                    remaining_garbage += 1
                    print(f"  警告: 仍有中文句号残留: rowid={row[rowid_idx]}, 日期='{s}'")
                if '&' in s:
                    remaining_garbage += 1
                    print(f"  警告: 仍有&分隔符残留: rowid={row[rowid_idx]}, 日期='{s}'")

        if remaining_garbage > 0:
            print(f"\n警告: 仍有 {remaining_garbage} 条记录可能包含未处理的非标准日期值。")
        else:
            print("\n验证通过：无残留非标准日期记录。")

    conn.close()


if __name__ == "__main__":
    main()
