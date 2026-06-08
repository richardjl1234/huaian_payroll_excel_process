"""
reconcile_excel_vs_db.py — 一次性脚本：全量 Excel vs DB 逐行 reconcile

⚠️ 强约束：本脚本对 DB **零写入**，全程 SELECT。
- 启动时 PRAGMA query_only = ON
- 不创建新表，不插入任何数据

算法：
1. Excel 端：复现生产 pipeline (sheet_gen + df_gen + special_logic_preprocess_df)
2. DB 端：读 payroll_details + load_log
3. 每行 row_hash = SHA256(各字段用 \x1f 分隔，Decimal 量化，None 用 \x00 哨兵)
4. 集合差集：漏行（excel 独有）/ 多行（db 独有）
5. 跨文件 hash 碰撞：识别重复入库
6. UNREADABLE 文件：标记 (new_payroll/202003.xls, 202109.xls)
7. DB 一致性诊断：装配误识别 / 系数异常 / 定额异常 / 小计残留 / 黄陈溯源

输出：reconcile_report.html + reconcile_report.csv
"""
import argparse
import hashlib
import os
import re
import sqlite3
import sys
import time
from collections import OrderedDict, defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pandas as pd

# 路径
PROJECT_ROOT = Path(__file__).resolve().parent.parent
DB_PATH = PROJECT_ROOT.parent / "payroll_database.db"

# 损坏文件清单（已知）
UNREADABLE_FILES = {"202003.xls", "202109.xls"}

# 期望列（hash 字段）
# 注意：日期列被排除 —— Excel 端是原始 `3·4·5` (中文点号) 格式，
# DB 端是 Step 0-9 清洗后的 `3,4,5` (英文逗号) 格式，
# 两端格式不一致会导致 hash 必不匹配。在报告中单独看日期差异。
EXPECTED_COLUMNS = ['职员全名', '客户名称', '型号', '工序全名', '工序',
                    '计件数量', '系数', '定额', '金额', '备注', '代码']

# 数值列：空字符串/None/NaN 归一为 0.00（与 production pipeline 的
# pd.to_numeric(errors='coerce').fillna(0.0) 行为一致）
NUMERIC_COLUMNS = {'计件数量', '系数', '定额', '金额'}
ZERO_QUANTIZED = str(Decimal('0').quantize(Decimal('0.01')))

# 分隔符
SEP = "\x1f"
NULL_SENTINEL = "\x00"

# 关键诊断规则
DISCARD_PHRASES = ['下料', '铣底脚：', '铣：', '校平衡', '车转子', '压：', '磨：']


# L19 检测：检查值尾部是否被 L19 附加了 " (L19字符串)" 格式
# 格式：原值 + " (" + L19字符串 + ")"  → "校正 (1.5H)" / "校平 (8H)" / "(2套)"
# 内部字符串（非数字含数字）："1.5H"/"8H"/"24H"/"2套"/"3人8H"/"1.5X"/"220V60HZ" 等
# 排除：括号内是纯数字字符串（"1.0"/"1"/"1.5"）—— 这种直接被 to_numeric 转走，不会触发 L19

_L19_PAREN_RE = re.compile(r'\(([^)]+)\)\s*$')


def has_l19_suffix(value):
    """L19 检测：value 列尾部是否被 L19 附加了 " (L19字符串)" 格式
    - "校正 (1.5H)" → True
    - "校平 (8H)" → True
    - "(2套)" → True（无前缀也行）
    - "校正 (1.0)" → False（"1.0" 是纯数字）
    - "校正" → False
    - "校正 1.5H" → False（必须带括号）
    - "" → False
    - None → False
    """
    if not value:
        return False
    s = str(value)
    m = _L19_PAREN_RE.search(s)
    if not m:
        return False
    inner = m.group(1).strip()
    if not inner:
        return False
    if not re.search(r'\d', inner):
        return False
    try:
        float(inner)
        return False  # 纯数字字符串，不是 L19
    except ValueError:
        return True


def _is_zero(v):
    """判断值是否为 0（处理 Decimal/float/int 各种类型）"""
    if v is None:
        return False
    if isinstance(v, float):
        if pd.isna(v):
            return False
        return abs(v) < 1e-9
    if isinstance(v, Decimal):
        return v == 0
    if isinstance(v, int):
        return v == 0
    s = str(v).strip()
    if not s:
        return False
    try:
        return float(s) == 0
    except (ValueError, InvalidOperation):
        return False


def normalize_value(v, col_name=None):
    """统一字段值：
    - 数值列（col_name in NUMERIC_COLUMNS）：
        * None/空字符串/NaN/无法解析的字符串 → "0.00"（与 production 的
          pd.to_numeric(errors='coerce').fillna(0.0) 一致）
    - 其他列：
        * None/空字符串 → NULL_SENTINEL
        * 能转数字的 → Decimal 量化
        * 无法解析的字符串保持原样
    """
    is_numeric = col_name in NUMERIC_COLUMNS
    if v is None:
        return ZERO_QUANTIZED if is_numeric else NULL_SENTINEL
    if isinstance(v, float):
        if pd.isna(v):
            return ZERO_QUANTIZED if is_numeric else NULL_SENTINEL
        try:
            return str(Decimal(str(v)).quantize(Decimal('0.01')))
        except (InvalidOperation, ValueError):
            return str(v)
    if isinstance(v, int):
        return str(Decimal(v).quantize(Decimal('0.01')))
    if isinstance(v, Decimal):
        return str(v.quantize(Decimal('0.01')))
    s = str(v).strip()
    if s == "" or s.lower() in ("nan", "none", "null"):
        return ZERO_QUANTIZED if is_numeric else NULL_SENTINEL
    # 尝试转 Decimal
    try:
        return str(Decimal(s).quantize(Decimal('0.01')))
    except (InvalidOperation, ValueError):
        # 数值列的不可解析字符串（"2套"、"3人8H"、"1台3H"、"200，") → 0
        # 与 production 行为一致；信息已无法在数值列保留
        if is_numeric:
            return ZERO_QUANTIZED
        return s  # 字符串列保留原值（如 "1.5H" 已挪到 工序；此列不会出现）


def row_hash(row_dict):
    """算单行 hash：字段 repr() 包裹 + \x1f 分隔 + SHA256"""
    parts = []
    for k in EXPECTED_COLUMNS:
        v = row_dict.get(k)
        normalized = normalize_value(v, k)
        parts.append(repr(normalized))
    blob = SEP.join(parts).encode("utf-8")
    return hashlib.sha256(blob).hexdigest()


def row_content_hash(row_dict):
    """排除 file_name/sheet_name 的内容 hash，用于跨文件碰撞检测"""
    parts = []
    for k in EXPECTED_COLUMNS:
        v = row_dict.get(k)
        normalized = normalize_value(v, k)
        parts.append(repr(normalized))
    blob = SEP.join(parts).encode("utf-8")
    return hashlib.sha256(blob).hexdigest()


def excel_pipeline_hashes():
    """跑完整生产 pipeline，对每行算 hash
    返回: (results, skipped_files, hour_info_rows)
        - results: dict[(file_name, sheet_name)] -> list of dicts
        - skipped_files: 跳过的文件列表
        - hour_info_rows: 工序全名/工序 列尾部被 L19 附加了 " (L19字符串)" 格式的行，list of dict {文件, sheet, 工序, 工序列, 计件数量, 姓名, 客户名称, 型号}
    """
    sys.path.insert(0, str(PROJECT_ROOT))
    from excel_processor.sheet_gen import get_excel_files, sheet_gen
    from excel_processor.df_gen import df_gen
    from excel_processor.special_logic import special_logic_preprocess_df

    files = get_excel_files()
    print(f"📁 Excel 文件: {len(files)} 个")

    results = defaultdict(list)  # (file, sheet) -> list of dicts
    skipped_files = []
    hour_info_rows = []  # L19 保留的工时信息行
    processed_count = 0

    for sc in sheet_gen(files):
        try:
            for sdf in df_gen(sc):
                processed_count += 1
                try:
                    processed_df, new_sheet, _ = special_logic_preprocess_df(
                        sdf.split_df, sc.sheet_name, sc.file_name, sdf.table_index
                    )
                    if processed_df is None or processed_df.empty:
                        continue
                    # 计算每行 hash
                    for _, row in processed_df.iterrows():
                        row_dict = {col: row.get(col) for col in EXPECTED_COLUMNS if col in row.index}
                        if not row_dict:
                            continue
                        results[(sc.file_name, new_sheet)].append({
                            "row_hash": row_hash(row_dict),
                            "row_content_hash": row_content_hash(row_dict),
                            "row_dict": row_dict,
                        })
                        # L19 检测：检查 工序全名 / 工序 列是否被 L19 附加了 " (L19字符串)" 格式
                        # 关键修正：L19 应用时 计件数量 必为 0（被强制转 0），
                        # 所以要同时检查 末段匹配 L19 模式 AND 计件数量=0
                        # 检查顺序：先 工序全名，再 工序（与生产 L19 一致）
                        qty = row_dict.get('计件数量', 0)
                        gx_name = row_dict.get('工序全名')
                        gx = row_dict.get('工序')
                        l19_target_value = None
                        if has_l19_suffix(gx_name):
                            l19_target_value = gx_name
                            l19_target_col = '工序全名'
                        elif has_l19_suffix(gx):
                            l19_target_value = gx
                            l19_target_col = '工序'
                        if l19_target_value is not None and _is_zero(qty):
                            hour_info_rows.append({
                                "文件": sc.file_name,
                                "sheet": new_sheet,
                                "姓名": row_dict.get('职员全名', ''),
                                "客户名称": row_dict.get('客户名称', ''),
                                "型号": str(row_dict.get('型号') or ''),
                                "工序": str(l19_target_value),
                                "工序列": l19_target_col,
                                "计件数量": qty,
                            })
                except Exception as e:
                    print(f"⚠️ special_logic 失败: {sc.file_name}::{sc.sheet_name} table={sdf.table_index}: {e}")
        except Exception as e:
            print(f"⚠️ df_gen 失败: {sc.file_name}::{sc.sheet_name}: {e}")

    print(f"✅ Excel pipeline 处理: {processed_count} 个 split DataFrame")
    print(f"✅ L19 工时/单位信息保留: {len(hour_info_rows)} 行 (含 '1.5H'/'8H'/'24H'/'2套'/'3人8H'/'1台3H' 等非数字字符串)")
    return results, skipped_files, hour_info_rows


def db_hashes(conn):
    """读 payroll_details，按 (file_name, sheet_name) 分组算 hash"""
    cur = conn.execute("SELECT * FROM payroll_details")
    cols = [d[0] for d in cur.description]
    print(f"📊 DB 列: {cols}")

    results = defaultdict(list)  # (file_name, sheet_name) -> list of dicts
    total = 0
    for row in cur.fetchall():
        total += 1
        row_dict = dict(zip(cols, row))
        # 关键列存在性
        key = (row_dict.get("文件名"), row_dict.get("sheet名"))
        results[key].append({
            "row_hash": row_hash(row_dict),
            "row_content_hash": row_content_hash(row_dict),
            "row_dict": row_dict,
        })

    print(f"✅ DB 读取: {total} 行")
    return results


def audit_db_consistency(conn):
    """任务 C：DB 内部一致性诊断"""
    findings = []

    # 1. `装配` 误识别为员工
    cur = conn.execute("SELECT COUNT(*) FROM payroll_details WHERE 职员全名 = '装配'")
    n = cur.fetchone()[0]
    if n > 0:
        findings.append({
            "类别": "职员误识别",
            "项目": "`装配` 被识别为职员全名",
            "数量": n,
            "说明": "L1 规则把首列 `装配` 改名为 `职员全名`，但 `装配` 是工序名不是员工名",
            "影响": "按 职员全名 聚合统计时混入虚假员工",
        })

    # 2. 系数异常 (>2 或 <0)
    cur = conn.execute("SELECT COUNT(*) FROM payroll_details WHERE CAST(系数 AS REAL) > 2 OR CAST(系数 AS REAL) < 0")
    n = cur.fetchone()[0]
    if n > 0:
        cur2 = conn.execute("SELECT MIN(CAST(系数 AS REAL)), MAX(CAST(系数 AS REAL)) FROM payroll_details WHERE CAST(系数 AS REAL) > 2 OR CAST(系数 AS REAL) < 0")
        mn, mx = cur2.fetchone()
        findings.append({
            "类别": "数值异常",
            "项目": "系数 > 2 或 < 0",
            "数量": n,
            "说明": f"系数正常范围 0-2；实际范围 {mn:.2f} ~ {mx:.2f}",
            "影响": "可能误录入或列错位",
        })

    # 3. 定额异常 (>100)
    cur = conn.execute("SELECT COUNT(*) FROM payroll_details WHERE CAST(定额 AS REAL) > 100")
    n = cur.fetchone()[0]
    if n > 0:
        cur2 = conn.execute("SELECT MAX(CAST(定额 AS REAL)) FROM payroll_details WHERE CAST(定额 AS REAL) > 100")
        mx = cur2.fetchone()[0]
        findings.append({
            "类别": "数值异常",
            "项目": "定额 > 100",
            "数量": n,
            "说明": f"定额最大值 {mx:.2f}；可能混入其他数值列",
            "影响": "影响金额计算正确性",
        })

    # 4. 小计行残留（"袁崇雷合计"/"姜浩合计"等）
    cur = conn.execute("SELECT 职员全名, COUNT(*) FROM payroll_details WHERE 职员全名 LIKE '%合计%' OR 职员全名 LIKE '%小计%' GROUP BY 职员全名")
    rows = cur.fetchall()
    if rows:
        total = sum(c for _, c in rows)
        findings.append({
            "类别": "数据污染",
            "项目": "小计/合计行残留",
            "数量": total,
            "说明": f"含合计关键字的职员名: {dict(rows)}",
            "影响": "聚合统计时重复计算",
        })

    # 5. 黄志梅 == 陈会清 行数和金额
    cur = conn.execute("SELECT 职员全名, COUNT(*), SUM(CAST(金额 AS REAL)) FROM payroll_details WHERE 职员全名 IN ('黄志梅', '陈会清') GROUP BY 职员全名")
    rows = cur.fetchall()
    if len(rows) == 2:
        a, b = rows
        if a[1] == b[1] and abs(a[2] - b[2]) < 0.01:
            findings.append({
                "类别": "L14 拆分",
                "项目": "黄志梅 == 陈会清 数据完全相同",
                "数量": a[1],
                "说明": f"两人各 {a[1]} 行，总金额 {a[2]:.2f} / {b[2]:.2f}；来自 special_logic L14 `前装` 拆 2 行",
                "影响": "业务合理（按工作量平分），不修",
            })

    # 6. 金额 = 0 但计件数量 ≠ 0
    cur = conn.execute("SELECT COUNT(*) FROM payroll_details WHERE CAST(金额 AS REAL) = 0 AND CAST(计件数量 AS REAL) != 0")
    n = cur.fetchone()[0]
    if n > 0:
        findings.append({
            "类别": "数据异常",
            "项目": "金额=0 但计件数量≠0",
            "数量": n,
            "说明": "可能 #VALUE! 错误被 xlrd 吞掉变 0",
            "影响": "工资计算结果错误",
        })

    return findings


def reconcile_excel_vs_db(excel_data, db_data):
    """对每 (file, sheet, table) 做集合差集"""
    report = []
    all_keys = set(excel_data.keys()) | set(db_data.keys())

    for key in sorted(all_keys):
        file_name, sheet_name = key[0], key[1]
        excel_rows = excel_data.get(key, [])
        db_rows = db_data.get(key, [])

        excel_hashes = {r["row_hash"] for r in excel_rows}
        db_hashes_set = {r["row_hash"] for r in db_rows}

        missing_in_db = excel_hashes - db_hashes_set  # 漏行
        extra_in_db = db_hashes_set - excel_hashes    # 多行

        status = "OK"
        if file_name in UNREADABLE_FILES:
            status = "UNREADABLE"
        elif missing_in_db or extra_in_db:
            status = "WARN"

        report.append({
            "file_name": file_name,
            "sheet_name": sheet_name,
            "excel_rows": len(excel_rows),
            "db_rows": len(db_rows),
            "missing_in_db": len(missing_in_db),
            "extra_in_db": len(extra_in_db),
            "status": status,
            "samples_missing": [
                next((r["row_dict"] for r in excel_rows if r["row_hash"] == h), None)
                for h in list(missing_in_db)[:3]
            ],
            "samples_extra": [
                next((r["row_dict"] for r in db_rows if r["row_hash"] == h), None)
                for h in list(extra_in_db)[:3]
            ],
        })

    return report


def cross_file_duplicate_check(db_data):
    """跨文件 hash 碰撞检测"""
    content_hash_to_files = defaultdict(set)
    for (file_name, sheet_name), rows in db_data.items():
        for r in rows:
            content_hash_to_files[r["row_content_hash"]].add(file_name)

    duplicates = []
    for ch, files in content_hash_to_files.items():
        if len(files) > 1:
            duplicates.append({
                "content_hash": ch,
                "file_count": len(files),
                "files": sorted(files),
            })
    return duplicates


def load_log_reconcile(conn):
    """检查 load_log 是否有 FAILED 记录（UNREADABLE 文件应有）"""
    cur = conn.execute("SELECT * FROM load_log")
    cols = [d[0] for d in cur.description]
    rows = [dict(zip(cols, r)) for r in cur.fetchall()]
    print(f"📋 load_log 记录: {len(rows)} 条")
    return rows


def check_db_l19_applied(conn):
    """检查 DB 中 L19 是否已应用：
    L19 特征：工序全名 或 工序 列尾部含 " (L19字符串)" 格式 + 计件数量=0

    SQL 粗筛（行数级）→ Python 精筛（has_l19_suffix + qty=0）
    """
    cur = conn.execute("""
        SELECT 工序全名, 工序, 计件数量 FROM payroll_details
        WHERE 计件数量 = 0
          AND (工序全名 LIKE '%(%' OR 工序全名 LIKE '%（%'
               OR 工序 LIKE '%(%' OR 工序 LIKE '%（%')
    """)
    n = 0
    for gx_name, gx, qty in cur:
        if has_l19_suffix(gx_name) or has_l19_suffix(gx):
            n += 1
    return n


def generate_html_report(reconcile_report, duplicates, audit_findings, load_log_rows, hour_info_rows, db_l19_count, output_path):
    """生成 HTML 报告"""
    # 汇总指标
    total_files = len(set(r["file_name"] for r in reconcile_report))
    ok_count = sum(1 for r in reconcile_report if r["status"] == "OK")
    warn_count = sum(1 for r in reconcile_report if r["status"] == "WARN")
    unrd_count = sum(1 for r in reconcile_report if r["status"] == "UNREADABLE")
    total_missing = sum(r["missing_in_db"] for r in reconcile_report)
    total_extra = sum(r["extra_in_db"] for r in reconcile_report)

    html = [
        "<!DOCTYPE html><html><head><meta charset='utf-8'>",
        "<title>Reconcile Report - Excel vs DB</title>",
        "<style>",
        "body{font-family:sans-serif;margin:20px;}",
        "h1{color:#06c;}",
        "h2{background:#eef;padding:8px;border-left:4px solid #06c;margin-top:30px;}",
        "table{border-collapse:collapse;margin:10px 0;font-size:13px;}",
        "th,td{border:1px solid #ccc;padding:4px 8px;text-align:left;}",
        "th{background:#f5f5f5;}",
        ".OK{color:green;font-weight:bold;}",
        ".WARN{color:#c80;font-weight:bold;}",
        ".UNREADABLE{color:red;font-weight:bold;}",
        ".summary{background:#f9f9f9;padding:15px;border-radius:4px;margin:15px 0;}",
        "</style></head><body>",
        f"<h1>Reconcile 报告 - Excel vs SQLite payroll_details</h1>",
        f"<p>生成时间: {time.strftime('%Y-%m-%d %H:%M:%S')}</p>",
        f"<div class='summary'>",
        f"<h3>📊 汇总指标</h3>",
        f"<ul>",
        f"<li>涉及文件: <b>{total_files}</b></li>",
        f"<li>OK (无差异): <b class='OK'>{ok_count}</b></li>",
        f"<li>WARN (有差异): <b class='WARN'>{warn_count}</b></li>",
        f"<li>UNREADABLE (损坏文件): <b class='UNREADABLE'>{unrd_count}</b></li>",
        f"<li>总漏行 (Excel 有 / DB 无): <b>{total_missing}</b></li>",
        f"<li>总多行 (DB 有 / Excel 无): <b>{total_extra}</b></li>",
        f"<li>跨文件重复 (content_hash 相同): <b>{len(duplicates)}</b></li>",
        f"</ul>",
        f"</div>",
    ]

    # 任务 C：DB 一致性诊断
    html.append("<h2>DB 内部一致性诊断 (Task C)</h2>")
    if audit_findings:
        html.append("<table><tr><th>类别</th><th>项目</th><th>数量</th><th>说明</th><th>影响</th></tr>")
        for f in audit_findings:
            html.append(
                f"<tr><td>{f['类别']}</td><td>{f['项目']}</td>"
                f"<td><b>{f['数量']}</b></td><td>{f['说明']}</td>"
                f"<td>{f['影响']}</td></tr>"
            )
        html.append("</table>")
    else:
        html.append("<p>✅ 无问题</p>")

    # L19 工时/单位信息保留诊断
    html.append("<h2>工时/单位信息保留 (L19) 诊断</h2>")
    html.append(f"<p>Excel pipeline 输出中，<b>工序全名 / 工序</b> 列尾部被 L19 附加了 <code>\" (L19字符串)\"</code> 格式"
                f"（工时/单位信息，如 \" (1.5H)\"/\" (8H)\"/\" (24H)\"/\" (2套)\"/\" (3人8H)\"/\" (1台3H)\"）的行共 <b>{len(hour_info_rows)}</b> 行。</p>")
    if hour_info_rows:
        # 按文件分组
        file_groups = OrderedDict()
        for r in hour_info_rows:
            k = r["文件"]
            file_groups.setdefault(k, []).append(r)
        html.append(f"<p>涉及 <b>{len(file_groups)}</b> 个文件。下表展示前 20 行：</p>")
        html.append("<table><tr><th>文件</th><th>sheet</th><th>姓名</th><th>客户</th><th>型号</th><th>工序列</th><th>工序 (含 L19 后缀)</th><th>计件数量</th></tr>")
        shown = 0
        for fname, rows in file_groups.items():
            for r in rows:
                if shown >= 20:
                    break
                qty = r['计件数量']
                qty_str = f"{qty:.2f}" if isinstance(qty, (int, float)) else str(qty)
                html.append(
                    f"<tr><td>{fname}</td><td>{r['sheet']}</td>"
                    f"<td>{r['姓名'] or ''}</td><td>{r['客户名称'] or ''}</td><td>{r.get('型号','') or ''}</td>"
                    f"<td>{r.get('工序列','') or ''}</td><td><code>{r['工序']}</code></td><td>{qty_str}</td></tr>"
                )
                shown += 1
            if shown >= 20:
                break
        if len(hour_info_rows) > 20:
            html.append(f"<tr><td colspan=8>... 还有 {len(hour_info_rows) - 20} 行 ...</td></tr>")
        html.append("</table>")
        if db_l19_count > 0:
            if len(hour_info_rows) > db_l19_count:
                html.append(
                    f"<p>✅ DB 中已检测到 <b>{db_l19_count}</b> 行 工序全名/工序 列含 L19 后缀（L19 已生效）。"
                    f"但 Excel pipeline 输出 <b>{len(hour_info_rows)}</b> 行，比 DB 多 <b>{len(hour_info_rows) - db_l19_count}</b> 行"
                    f"——这些行可能在 L19 重写之前已入库。"
                    f"需要再跑 <code>./sqlite_payroll_details_refresh.sh</code> 让新版 L19（合并到 工序全名/工序 列）生效。</p>"
                )
            elif len(hour_info_rows) < db_l19_count:
                html.append(
                    f"<p>✅ DB 中已检测到 <b>{db_l19_count}</b> 行，Excel pipeline 输出 <b>{len(hour_info_rows)}</b> 行。"
                    f"DB 比 Excel 多 <b>{db_l19_count - len(hour_info_rows)}</b> 行（可能有重复入库或 Excel pipeline 处理差异）。</p>"
                )
            else:
                html.append(
                    f"<p>✅ DB 和 Excel pipeline 各 <b>{db_l19_count}</b> 行 工序全名/工序 列含 L19 后缀，L19 完整生效。</p>"
                )
        else:
            html.append(
                "<p>⚠️ DB 中未检测到 L19 应用的痕迹（工序全名/工序 列无 \" (L19字符串)\" 括号后缀），"
                "说明当前 DB 是用 L19 之前的 production pipeline 生成的。"
                "修复方法：重新跑 <code>./sqlite_payroll_details_refresh.sh</code>，"
                "用 L19 之后的 pipeline 重新生成 DB，这些行将匹配。</p>"
            )
    else:
        html.append("<p>✅ Excel pipeline 输出中未发现工时/单位信息行</p>")

    # 跨文件重复检测
    html.append("<h2>跨文件重复入库检测</h2>")
    if duplicates:
        html.append(f"<p>⚠️ 发现 <b>{len(duplicates)}</b> 组 content_hash 出现在多个文件 (疑似重复入库)</p>")
        html.append("<table><tr><th>content_hash</th><th>涉及文件数</th><th>文件列表</th></tr>")
        for d in duplicates[:20]:
            html.append(
                f"<tr><td><code>{d['content_hash'][:16]}...</code></td>"
                f"<td>{d['file_count']}</td><td>{', '.join(d['files'][:5])}{'...' if len(d['files']) > 5 else ''}</td></tr>"
            )
        html.append("</table>")
    else:
        html.append("<p>✅ 未发现跨文件重复</p>")

    # load_log reconcile
    html.append("<h2>load_log 状态</h2>")
    html.append(f"<p>load_log 共 <b>{len(load_log_rows)}</b> 条记录</p>")
    if load_log_rows:
        cols = list(load_log_rows[0].keys())
        html.append("<table><tr>" + "".join(f"<th>{c}</th>" for c in cols) + "</tr>")
        for r in load_log_rows[:30]:
            html.append("<tr>" + "".join(f"<td>{str(r.get(c, ''))[:50]}</td>" for c in cols) + "</tr>")
        html.append("</table>")

    # 按文件 reconcile 详情
    html.append("<h2>文件级 Reconcile 详情</h2>")
    html.append("<table><tr><th>文件</th><th>Sheet</th><th>Excel 行数</th><th>DB 行数</th><th>漏行</th><th>多行</th><th>状态</th></tr>")
    for r in reconcile_report:
        status_class = r["status"]
        html.append(
            f"<tr><td>{r['file_name']}</td><td>{r['sheet_name']}</td>"
            f"<td>{r['excel_rows']}</td><td>{r['db_rows']}</td>"
            f"<td>{r['missing_in_db']}</td><td>{r['extra_in_db']}</td>"
            f"<td class='{status_class}'>{r['status']}</td></tr>"
        )
    html.append("</table>")

    # WARN 详情
    warn_rows = [r for r in reconcile_report if r["status"] == "WARN"]
    if warn_rows:
        html.append("<h2>WARN 详情 (有差异的行)</h2>")
        for r in warn_rows[:10]:
            html.append(f"<h3>{r['file_name']} :: {r['sheet_name']}</h3>")
            html.append(f"<p>漏行: {r['missing_in_db']}, 多行: {r['extra_in_db']}</p>")
            if r["samples_missing"]:
                html.append("<h4>漏行样本</h4><pre>")
                for s in r["samples_missing"]:
                    if s:
                        html.append(f"  {s}\n")
                html.append("</pre>")
            if r["samples_extra"]:
                html.append("<h4>多行样本</h4><pre>")
                for s in r["samples_extra"]:
                    if s:
                        html.append(f"  {s}\n")
                html.append("</pre>")

    html.append("</body></html>")
    Path(output_path).write_text("\n".join(html), encoding="utf-8")
    print(f"✅ 报告已写入: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="Reconcile Excel vs DB")
    parser.add_argument("--audit-only", action="store_true", help="仅跑任务 C 诊断")
    parser.add_argument("--output", default="reconcile_report.html", help="报告路径")
    args = parser.parse_args()

    if not DB_PATH.exists():
        print(f"❌ DB 不存在: {DB_PATH}")
        sys.exit(1)

    # 只读模式
    conn = sqlite3.connect(str(DB_PATH))
    conn.execute("PRAGMA query_only = ON")
    print("✅ DB 已切换到 query_only 模式")

    # 任务 C：DB 一致性诊断（可单独跑）
    print("\n=== 任务 C：DB 一致性诊断 ===")
    audit_findings = audit_db_consistency(conn)
    for f in audit_findings:
        print(f"  [{f['类别']}] {f['项目']}: {f['数量']}")

    if args.audit_only:
        # 生成简版报告
        html = ["<html><body><h1>DB 一致性诊断报告 (--audit-only)</h1>"]
        html.append("<table><tr><th>类别</th><th>项目</th><th>数量</th><th>说明</th><th>影响</th></tr>")
        for f in audit_findings:
            html.append(f"<tr><td>{f['类别']}</td><td>{f['项目']}</td><td>{f['数量']}</td><td>{f['说明']}</td><td>{f['影响']}</td></tr>")
        html.append("</table></body></html>")
        Path(args.output).write_text("\n".join(html), encoding="utf-8")
        print(f"✅ 简版报告: {args.output}")
        conn.close()
        return

    # 完整 reconcile
    print("\n=== 任务 B：Excel vs DB 全量 Reconcile ===")
    print("\n--- 阶段 1: Excel 端复现生产 pipeline ---")
    excel_data, _, hour_info_rows = excel_pipeline_hashes()

    print("\n--- 阶段 2: DB 端读取 ---")
    db_data = db_hashes(conn)

    print("\n--- 阶段 3: 集合差集 ---")
    reconcile_report = reconcile_excel_vs_db(excel_data, db_data)
    ok_count = sum(1 for r in reconcile_report if r["status"] == "OK")
    warn_count = sum(1 for r in reconcile_report if r["status"] == "WARN")
    print(f"  OK: {ok_count}, WARN: {warn_count}")

    print("\n--- 阶段 4: 跨文件重复检测 ---")
    duplicates = cross_file_duplicate_check(db_data)
    print(f"  发现 {len(duplicates)} 组跨文件重复")

    print("\n--- 阶段 5: load_log reconcile ---")
    load_log_rows = load_log_reconcile(conn)

    print("\n--- 阶段 5.5: 检查 DB L19 是否已应用 ---")
    db_l19_count = check_db_l19_applied(conn)
    print(f"  DB 中 工序全名/工序 列含 L19 后缀的行数: {db_l19_count}")

    print("\n--- 阶段 6: 生成报告 ---")
    generate_html_report(reconcile_report, duplicates, audit_findings, load_log_rows, hour_info_rows, db_l19_count, args.output)

    conn.close()
    print("\n✅ 全部完成")


if __name__ == "__main__":
    main()
