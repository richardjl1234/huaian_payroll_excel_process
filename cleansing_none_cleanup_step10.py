#!/usr/bin/env python3
"""
清洁工资数据库 Step 10 - 清理 'None' 字面量占位符 (cleansing pipeline 最后一步)。

用途: 清理 pandas `astype(str)` 在导入时把空单元格写成的字面字符串 'None'。
前提: Step 0-9 已完成,日期列已规范化;本步骤不依赖也不修改日期。

占位符来源:
  excel_processor/sheet_processor.py:523 `df[col] = df[col].astype(str)`
  把空 pandas 单元格 (Python None) 转成字符串 'None' 后写入 DB。

目标列 (6 个,均为文本列):
  代码, 客户名称, 备注, 工序, 型号, 工序全名

占位符值 (经 dry-run 确认当前仅 'None' 存在,其他为前向兼容):
  'None', 'none', 'null', 'NULL', 'nan', 'NaN'
  -> 全部统一替换为 '' (空串)

不动:
  - '文件名' / 'sheet名' / '职员全名' / '日期'  (经核查,这些列无占位符)
  - 已存在的 '' 空串 (用户已确认不动)

行为: 6 个 UPDATE 在一个事务内执行,任一失败整体 rollback。
  与 reconcile_excel_vs_db.py 的 normalize_value() 行为保持一致。

用法:
    python3 cleansing_none_cleanup_step10.py [--dry-run]

参数:
    --dry-run  仅预览并导出HTML报告,不执行 UPDATE
    无参数    提示输入 'yes' 确认,执行 UPDATE

输出:
    none_cleanup_step10_output.html (同目录)
"""

import sqlite3
import sys
import argparse
from pathlib import Path

DB_PATH = Path(__file__).parent.parent / "payroll_database.db"
OUTPUT_PATH = Path(__file__).parent / "none_cleanup_step10_output.html"
PAYROLL_TABLE = "payroll_details"

# 6 个目标列 (经 dry-run 确认含占位符的列)
TARGET_COLUMNS = ['代码', '客户名称', '备注', '工序', '型号', '工序全名']

# 占位符值集合 - 当前生产 DB 中只有 'None' 实际存在
# 其他变体为前向兼容, 与 reconcile_excel_vs_db.py normalize_value() 一致
PLACEHOLDERS = ['None', 'none', 'null', 'NULL', 'nan', 'NaN']
PLACEHOLDER_IN = ", ".join(f"'{p}'" for p in PLACEHOLDERS)

# 4 个对照列 (预期干净)
SANITY_COLUMNS = ['文件名', 'sheet名', '职员全名', '日期']

SAMPLE_LIMIT = 50


def build_placeholder_where(col: str) -> str:
    return f"{col} IN ({PLACEHOLDER_IN})"


def collect_stats(conn) -> dict:
    """
    对每个目标列:
      - total: 占位符总数
      - per_variant: 每种占位符值的计数
      - sample: 最多 50 条 rowid + 全列, 用于 HTML 报告
    """
    stats = {}
    for col in TARGET_COLUMNS:
        where = build_placeholder_where(col)

        total = conn.execute(
            f"SELECT COUNT(*) FROM {PAYROLL_TABLE} WHERE {where}"
        ).fetchone()[0]

        per_variant = {}
        for p in PLACEHOLDERS:
            cnt = conn.execute(
                f"SELECT COUNT(*) FROM {PAYROLL_TABLE} WHERE {col} = ?", (p,)
            ).fetchone()[0]
            if cnt:
                per_variant[p] = cnt

        sample = []
        if total:
            cur = conn.execute(
                f"SELECT rowid, * FROM {PAYROLL_TABLE} WHERE {where} LIMIT {SAMPLE_LIMIT}"
            )
            sample = cur.fetchall()

        stats[col] = {
            'total': total,
            'per_variant': per_variant,
            'sample': sample,
        }
    return stats


def check_clean_columns(conn) -> dict:
    """4 个对照列(预期干净)的占位符计数,任一 > 0 都是异常需要排查"""
    result = {}
    for col in SANITY_COLUMNS:
        where = build_placeholder_where(col)
        cnt = conn.execute(
            f"SELECT COUNT(*) FROM {PAYROLL_TABLE} WHERE {where}"
        ).fetchone()[0]
        result[col] = cnt
    return result


def export_to_html(conn, stats: dict, sanity: dict, output_path: Path):
    """导出处理结果到 HTML"""
    total_rows = sum(s['total'] for s in stats.values())

    # 取一个样本的列名 (所有 sample 共享相同列结构)
    sample_row = next(
        (s['sample'][0] for s in stats.values() if s['sample']), None
    )

    html = f'''<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Step 10 - 'None' 占位符清理结果</title>
<style>
body {{ font-family: sans-serif; margin: 20px; }}
table {{ border-collapse: collapse; font-size: 12px; margin-bottom: 30px; }}
th, td {{ border: 1px solid #ddd; padding: 6px; }}
th {{ background-color: #4472C4; color: white; position: sticky; top: 0; }}
tr:nth-child(even) {{ background-color: #f2f2f2; }}
tr:hover {{ background-color: #ffeb99; }}
td {{ max-width: 250px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
.section {{ font-size: 18px; font-weight: bold; margin-top: 30px; margin-bottom: 10px; }}
.updated {{ background-color: #ffffcc; }}
.summary {{ background-color: #e6f3ff; font-weight: bold; }}
.zero {{ color: #999; }}
</style>
</head>
<body>
<h1>Step 10 - 'None' 占位符清理结果</h1>

<div class="section">一、汇总</div>
<table>
<tr><th>列名</th><th>占位符总数</th><th>占位符类型分布</th></tr>
'''
    for col in TARGET_COLUMNS:
        s = stats[col]
        if s['total'] == 0:
            dist = '<span class="zero">(无)</span>'
        else:
            dist = ', '.join(f"{k}={v}" for k, v in sorted(s['per_variant'].items()))
        html += (
            f'<tr class="{"summary" if s["total"] else "zero"}">'
            f'<td>{col}</td>'
            f'<td>{s["total"]}</td>'
            f'<td>{dist}</td>'
            f'</tr>\n'
        )
    html += f'''
<tr class="summary"><td>合计</td><td>{total_rows}</td><td>-</td></tr>
</table>

<div class="section">二、对照列(预期干净,应为 0)</div>
<table>
<tr><th>列名</th><th>占位符计数</th><th>状态</th></tr>
'''
    for col, cnt in sanity.items():
        status = '✅ 干净' if cnt == 0 else f'❌ 异常! {cnt} 个'
        html += (
            f'<tr class="{"zero" if cnt == 0 else "updated"}">'
            f'<td>{col}</td><td>{cnt}</td><td>{status}</td></tr>\n'
        )
    html += '</table>\n'

    # 每个目标列 1 个 section
    if sample_row:
        sample_cols = [d[0] for d in conn.execute(
            f"SELECT rowid, * FROM {PAYROLL_TABLE} LIMIT 1"
        ).description]
        display_cols = [c for c in sample_cols if c != 'rowid']

        for col in TARGET_COLUMNS:
            s = stats[col]
            html += f'<div class="section">{TARGET_COLUMNS.index(col)+3}、{col} 列样本 ({s["total"]} 条, 前 {min(SAMPLE_LIMIT, s["total"])} 条展示)</div>\n'
            if not s['sample']:
                html += '<p class="zero">(无占位符记录,无需处理)</p>\n'
                continue
            html += '<table>\n<tr><th>rowid</th>' + ''.join(f'<th>{c}</th>' for c in display_cols) + '</tr>\n'
            for row in s['sample']:
                # rowid is first (index 0)
                cells = ''.join(
                    f'<td>{str(v) if v is not None else ""}</td>'
                    for v in row[1:]
                )
                html += f'<tr class="updated"><td>{row[0]}</td>{cells}</tr>\n'
            html += '</table>\n'

    html += '</body></html>'

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)


def print_summary(stats: dict, sanity: dict):
    """打印文本汇总到 stdout"""
    total_rows = sum(s['total'] for s in stats.values())
    print(f"\n【目标列占位符统计】")
    for col in TARGET_COLUMNS:
        s = stats[col]
        if s['total']:
            dist = ', '.join(f"{k}={v}" for k, v in sorted(s['per_variant'].items()))
            print(f"  {col:<10s} {s['total']:>7d}  ({dist})")
        else:
            print(f"  {col:<10s} {0:>7d}  (无)")
    print(f"  {'合计':<10s} {total_rows:>7d}")

    print(f"\n【对照列(预期干净)】")
    any_dirty = False
    for col, cnt in sanity.items():
        flag = '❌ 异常!' if cnt else '✅'
        if cnt:
            any_dirty = True
        print(f"  {col:<10s} {cnt:>7d}  {flag}")
    if any_dirty:
        print("\n警告: 4 个对照列中至少 1 个有占位符,需要先排查! 脚本仍会继续但请关注。")


def main():
    parser = argparse.ArgumentParser(
        description="清洁工资数据库 Step 10 - 清理 'None' 字面量占位符 (cleansing 最后一步)"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="仅预览并导出HTML报告, 不执行 UPDATE"
    )
    args = parser.parse_args()

    if not DB_PATH.exists():
        print(f"错误: 数据库文件不存在: {DB_PATH}")
        sys.exit(1)

    conn = sqlite3.connect(str(DB_PATH))

    # 总行数 (用于 sanity 报告)
    total = conn.execute(f"SELECT COUNT(*) FROM {PAYROLL_TABLE}").fetchone()[0]
    print(f"数据库 {DB_PATH} 共 {total} 条记录")

    stats = collect_stats(conn)
    sanity = check_clean_columns(conn)
    print_summary(stats, sanity)

    # 导出 HTML (dry-run 和正式模式都生成,方便事后查阅)
    export_to_html(conn, stats, sanity, OUTPUT_PATH)
    print(f"\n已导出 HTML 报告: {OUTPUT_PATH}")

    if args.dry_run:
        print(f"\n[DRY-RUN 模式] 未执行任何 UPDATE。")
        conn.close()
        return

    # 二次确认
    total_updates = sum(s['total'] for s in stats.values())
    if total_updates == 0:
        print("\n无占位符需要清理, 直接退出。")
        conn.close()
        return

    confirm = input(
        f"\n[确认] 即将对 {len([c for c in TARGET_COLUMNS if stats[c]['total']])} 个列 "
        f"执行 UPDATE, 共 {total_updates} 行 'None' → ''.\n"
        f"请输入 'yes' 确认: "
    )
    if confirm.strip().lower() != "yes":
        print("已取消操作。")
        conn.close()
        sys.exit(0)

    # 单事务执行 6 条 UPDATE
    print("\n[执行] 开始 UPDATE ...")
    try:
        for col in TARGET_COLUMNS:
            cnt = stats[col]['total']
            if cnt == 0:
                print(f"  {col}: 跳过 (0 行)")
                continue
            cur = conn.execute(
                f"UPDATE {PAYROLL_TABLE} SET {col} = '' WHERE {build_placeholder_where(col)}"
            )
            print(f"  {col}: {cur.rowcount} 行已更新")
        conn.commit()
    except Exception as e:
        conn.rollback()
        print(f"\n[错误] UPDATE 失败, 已 rollback: {e}")
        conn.close()
        sys.exit(1)

    # 验证残留
    print("\n[验证] 检查占位符残留 ...")
    residual = 0
    for col in TARGET_COLUMNS:
        cnt = conn.execute(
            f"SELECT COUNT(*) FROM {PAYROLL_TABLE} WHERE {build_placeholder_where(col)}"
        ).fetchone()[0]
        residual += cnt
        if cnt:
            print(f"  ❌ {col}: 仍有 {cnt} 个占位符残留!")
        else:
            print(f"  ✅ {col}: 0 残留")
    if residual == 0:
        print("\n验证通过: 6 个目标列占位符全部清理完成。")
    else:
        print(f"\n警告: 仍有 {residual} 个占位符残留, 请排查。")
        conn.close()
        sys.exit(1)

    # 最终行数 (理论上应与开始时一致, 没删除/插入)
    final = conn.execute(f"SELECT COUNT(*) FROM {PAYROLL_TABLE}").fetchone()[0]
    print(f"\n数据库当前总行数: {final} (修复前 {total})")
    conn.close()


if __name__ == "__main__":
    main()
