#!/usr/bin/env python3
"""
清洁工资数据库中的日期列 - Step 0 (预处理)。

将日期列中的全角字符转换为半角字符，统一格式。
此步骤在其他所有清洁步骤之前执行。

全角转半角映射:
- ， -> ,  (全角逗号 -> 半角逗号)
- 、 -> ,  (中文逗号 -> 半角逗号)
- ～ -> ~  (全角波浪号 -> 半角波浪号)
- —— -> -  (中文破折号 -> 半角短横线)
- — -> -   (中文长破折号 -> 半角短横线)
- （ -> (  (全角左括号 -> 半角左括号)
- ） -> )  (全角右括号 -> 半角右括号)
- 　 -> ' ' (全角空格 -> 半角空格)
- 　(末尾空格) -> '' (去除)

用法:
    python cleansing_data_dbcs_handling_step0.py [--dry-run]

参数:
    --dry-run  仅预览，不执行更新
    无参数    执行更新操作
"""

import sqlite3
import sys
import argparse
from pathlib import Path

DB_PATH = Path(__file__).parent.parent / "payroll_database.db"
OUTPUT_PATH = Path(__file__).parent / "dbcs_handling_step0_output.html"
PAYROLL_TABLE = "payroll_details"


# 全角转半角映射表
DBCS_MAP = {
    '，': ',',  # 全角逗号 -> 半角逗号
    '、': ',',  # 中文逗号 -> 半角逗号
    '～': '~',  # 全角波浪号 -> 半角波浪号
    '——': '-',  # 中文破折号 -> 半角短横线
    '—': '-',   # 中文长破折号 -> 半角短横线
    '（': '(',  # 全角左括号 -> 半角左括号
    '）': ')',  # 全角右括号 -> 半角右括号
    '　': ' ',  # 全角空格 -> 半角空格
    '·': ',',   # 中文点号 -> 半角逗号
}


def convert_dbcs(text: str) -> str:
    """
    将全角字符转换为半角字符。
    """
    if text is None:
        return None

    result = str(text)
    for dbcs, sbc in DBCS_MAP.items():
        result = result.replace(dbcs, sbc)

    # 去除末尾空格
    result = result.rstrip()

    return result


def export_to_html(columns: list, rows_with_changes: list, output_path: Path):
    """导出更新记录到HTML文件。"""
    html = f'''<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Step 0 - 全角转半角处理结果</title>
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
<h1>Step 0 - 全角转半角处理结果</h1>
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
        description="清洁工资数据库中的日期列 - Step 0 (全角转半角)"
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
    print(f"开始处理全角字符...")

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

        old_date_str = str(old_date)

        # 转换全角字符
        new_date_str = convert_dbcs(old_date_str)

        if new_date_str != old_date_str:
            updates.append((new_date_str, rowid))
            rows_with_changes.append((old_date_str, new_date_str, row))
            updated += 1

    # 按模式统计
    patterns = {}
    for old_date, new_date, row in rows_with_changes:
        key = old_date
        if key not in patterns:
            patterns[key] = {'count': 0, 'new': new_date}
        patterns[key]['count'] += 1

    print(f"\n处理完成:")
    print(f"  - 跳过(空值): {skipped}")
    print(f"  - 已更新: {updated}")

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