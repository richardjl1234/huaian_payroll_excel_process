#!/usr/bin/env python3
"""
清洁工资数据库中的异常日期记录 - Step 2。

使用前向填充(ffill)方法，根据 文件名、sheet名、职员全名 分组，
填充日期为空的记录。

用法:
    python cleansing_outliers_step2.py [--dry-run]

参数:
    --dry-run  仅预览填充结果，不执行更新
    无参数    执行填充操作
"""

import sqlite3
import sys
import argparse
from pathlib import Path

DB_PATH = Path(__file__).parent.parent / "payroll_database.db"
PAYROLL_TABLE = "payroll_details"


def fill_blank_dates(conn: sqlite3.Connection, dry_run: bool = False) -> list:
    """
    按 文件名、sheet名、职员全名 分组，对每组内的日期空白行进行前向填充。
    返回填充记录的列表，每项为 (filled_row, source_row)。
    """
    cursor = conn.execute(f"SELECT rowid, * FROM {PAYROLL_TABLE}")
    columns = [description[0] for description in cursor.description]
    all_rows = cursor.fetchall()

    if not all_rows:
        print("数据库为空。")
        return []

    col_indices = {col: idx for idx, col in enumerate(columns)}
    filename_idx = col_indices["文件名"]
    sheetname_idx = col_indices["sheet名"]
    employeename_idx = col_indices["职员全名"]
    date_idx = col_indices["日期"]
    rowid_idx = col_indices["rowid"]

    all_rows_sorted = sorted(all_rows, key=lambda r: (
        r[filename_idx] or "",
        r[sheetname_idx] or "",
        r[employeename_idx] or "",
    ))

    grouped = {}
    for row in all_rows_sorted:
        key = (row[filename_idx], row[sheetname_idx], row[employeename_idx])
        if key not in grouped:
            grouped[key] = []
        grouped[key].append(row)

    filled_records = []
    for key, group_rows in grouped.items():
        prev_row = None
        for row in group_rows:
            current_date = row[date_idx]
            if current_date is None or str(current_date).strip() == "":
                if prev_row is not None:
                    source_date = prev_row[date_idx]
                    if source_date is not None and str(source_date).strip() != "":
                        filled_records.append((row, prev_row))
                        if not dry_run:
                            rowid = row[rowid_idx]
                            conn.execute(
                                f"UPDATE {PAYROLL_TABLE} SET 日期 = ? WHERE rowid = ?",
                                (source_date, rowid)
                            )
            else:
                prev_row = row

    if not dry_run:
        conn.commit()

    return filled_records


def print_filled_records(columns: list, filled_records: list):
    """打印每条填充记录及其来源行，并汇总。"""
    date_idx = columns.index("日期")
    print("\n" + "=" * 120)
    print(f"共填充 {len(filled_records)} 条记录:")
    print("=" * 120)

    prev_source_rowid = None
    source_count = 0
    for filled_row, source_row in filled_records:
        source_rowid = source_row[0]
        source_date = source_row[date_idx]
        if source_rowid != prev_source_rowid:
            print(f"\n--- 来源行 (rowid={source_rowid}), 日期='{source_date}' ---")
            print(",".join(str(v) for v in source_row[1:]))
            source_count += 1

        print(f"  -> 目标行 (rowid={filled_row[0]}) <- 日期将设为 '{source_date}'")
        print(f"      {','.join(str(v) for v in filled_row[1:])}")
        prev_source_rowid = source_rowid

    print("\n" + "=" * 120)
    print("汇总:")
    print(f"  - 填充记录总数: {len(filled_records)}")
    print(f"  - 来源行总数: {source_count}")
    print("=" * 120)


def main():
    parser = argparse.ArgumentParser(
        description="清洁工资数据库中的异常日期记录 - Step 2 (前向填充)"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="仅预览填充结果，不执行更新"
    )
    args = parser.parse_args()

    if not DB_PATH.exists():
        print(f"错误: 数据库文件不存在: {DB_PATH}")
        sys.exit(1)

    conn = sqlite3.connect(str(DB_PATH))

    cursor = conn.execute(f"SELECT rowid, * FROM {PAYROLL_TABLE} WHERE 日期 IS NULL OR 日期 = ''")
    columns = [description[0] for description in cursor.description]
    rows = cursor.fetchall()
    blank_count = len(rows)

    print(f"找到 {blank_count} 条日期为空的记录。")

    if blank_count == 0:
        conn.close()
        sys.exit(0)

    filled_records = fill_blank_dates(conn, dry_run=args.dry_run)
    filled_count = len(filled_records)

    if args.dry_run:
        print(f"\n[DRY-RUN 模式] 预览: 本次可填充 {filled_count} 条记录。")
    else:
        print(f"\n已成功填充 {filled_count} 条记录。")

    if filled_count > 0:
        print_filled_records(columns, filled_records)

    conn.close()


if __name__ == "__main__":
    main()
