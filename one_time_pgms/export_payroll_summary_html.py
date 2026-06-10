"""导出 payroll_summary 表到单文件 HTML（方便浏览全部 2441 行）。

用法:
    python one_time_pgms/export_payroll_summary_html.py
    python one_time_pgms/export_payroll_summary_html.py --output summary.html
    python one_time_pgms/export_payroll_summary_html.py --db /path/to/payroll_database.db
"""
import argparse
import os
import sqlite3
import sys
import html
from collections import defaultdict
from pathlib import Path


COLUMNS = [
    ("id", "ID"),
    ("文件名", "文件"),
    ("sheet名", "Sheet"),
    ("汇总行索引", "行"),
    ("车间", "车间"),
    ("工序", "工序"),
    ("总件数", "总件数"),
    ("姓名", "姓名"),
    ("工作日", "工作日"),
    ("事假", "事假"),
    ("上月件数", "上月件数"),
    ("累计件数", "累计件数"),
]


def fetch_rows(conn):
    """读全部 payroll_summary 行，按 (文件名, sheet名, 汇总行索引) 排序"""
    cols = [c[0] for c in COLUMNS]
    cur = conn.execute(
        f"SELECT {','.join(cols)} FROM payroll_summary "
        f"ORDER BY 文件名, sheet名, 汇总行索引, id"
    )
    return [dict(zip(cols, r)) for r in cur.fetchall()]


def render_html(rows, output_path, db_path):
    n_total = len(rows)
    files = sorted({r["文件名"] for r in rows})
    sheets_per_file = defaultdict(set)
    for r in rows:
        sheets_per_file[r["文件名"]].add(r["sheet名"])
    n_files = len(files)

    # 渲染行（HTML escape + 数字右对齐样式）
    def cell(r, col, is_num):
        v = r.get(col)
        if v is None or v == "":
            return "<td class='null'>-</td>"
        s = str(v) if not isinstance(v, float) else f"{v:g}"
        cls = " class='num'" if is_num else ""
        return f"<td{cls}>{html.escape(s)}</td>"

    body_rows = []
    for r in rows:
        tds = "".join([
            cell(r, "id", True),
            cell(r, "文件名", False),
            cell(r, "sheet名", False),
            cell(r, "汇总行索引", True),
            cell(r, "车间", False),
            cell(r, "工序", False),
            cell(r, "总件数", True),
            cell(r, "姓名", False),
            cell(r, "工作日", True),
            cell(r, "事假", True),
            cell(r, "上月件数", True),
            cell(r, "累计件数", True),
        ])
        body_rows.append(f"<tr data-file='{html.escape(r['文件名'])}'>{tds}</tr>")

    # 文件过滤器选项
    file_options = "".join(
        f"<option value='{html.escape(f)}'>{html.escape(f)}</option>" for f in files
    )

    html_doc = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>payroll_summary 全部记录 ({n_total} 行 / {n_files} 文件)</title>
<style>
  body {{ font-family: -apple-system, "PingFang SC", "Microsoft YaHei", sans-serif; margin: 16px; color: #222; }}
  h1 {{ font-size: 20px; margin: 0 0 8px 0; }}
  .summary {{ color: #666; font-size: 13px; margin-bottom: 12px; }}
  .controls {{ position: sticky; top: 0; background: white; padding: 8px 0;
              border-bottom: 1px solid #ddd; margin-bottom: 8px; z-index: 10; }}
  .controls input, .controls select {{ padding: 4px 8px; font-size: 13px; margin-right: 8px; }}
  table {{ border-collapse: collapse; width: 100%; font-size: 13px; }}
  th, td {{ border: 1px solid #ddd; padding: 4px 8px; text-align: left; }}
  th {{ background: #f5f5f5; position: sticky; top: 41px; z-index: 5; }}
  td.num {{ text-align: right; font-variant-numeric: tabular-nums; }}
  td.null {{ color: #999; text-align: center; }}
  tr:hover {{ background: #fffbe6; }}
  .file-group {{ background: #f0f7ff; font-weight: bold; }}
  #counter {{ color: #888; font-size: 12px; margin-left: 12px; }}
  #toTop {{ position: fixed; right: 20px; bottom: 20px; padding: 8px 12px;
           background: #0066cc; color: white; border: none; border-radius: 4px;
           cursor: pointer; display: none; }}
</style>
</head>
<body>
<h1>payroll_summary 全部记录</h1>
<div class="summary">
  共 <b>{n_total}</b> 行，<b>{n_files}</b> 个文件。DB: <code>{html.escape(str(db_path))}</code>
</div>
<div class="controls">
  <input type="text" id="search" placeholder="按 姓名/工序/车间 过滤..." size="20">
  <select id="fileFilter">
    <option value="">-- 全部文件 --</option>
    {file_options}
  </select>
  <span id="counter"></span>
</div>
<table id="data">
  <thead>
    <tr>
      <th>ID</th><th>文件</th><th>Sheet</th><th>行</th>
      <th>车间</th><th>工序</th><th>总件数</th><th>姓名</th>
      <th>工作日</th><th>事假</th><th>上月件数</th><th>累计件数</th>
    </tr>
  </thead>
  <tbody>
    {"".join(body_rows)}
  </tbody>
</table>
<button id="toTop" onclick="window.scrollTo(0,0)">↑ 顶部</button>
<script>
  const search = document.getElementById('search');
  const fileFilter = document.getElementById('fileFilter');
  const counter = document.getElementById('counter');
  const toTop = document.getElementById('toTop');
  const rows = document.querySelectorAll('#data tbody tr');

  function applyFilter() {{
    const q = search.value.toLowerCase().trim();
    const f = fileFilter.value;
    let shown = 0;
    rows.forEach(r => {{
      const text = r.textContent.toLowerCase();
      const matchFile = !f || r.dataset.file === f;
      const matchText = !q || text.includes(q);
      const show = matchFile && matchText;
      r.style.display = show ? '' : 'none';
      if (show) shown++;
    }});
    counter.textContent = `显示 ${{shown}} / ${{rows.length}} 行`;
  }}

  search.addEventListener('input', applyFilter);
  fileFilter.addEventListener('change', applyFilter);
  applyFilter();

  window.addEventListener('scroll', () => {{
    toTop.style.display = window.scrollY > 300 ? 'block' : 'none';
  }});
</script>
</body>
</html>
"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_doc)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--db", default=os.environ.get("SQLITE_DB_PATH",
        str(Path(__file__).resolve().parent.parent.parent / "payroll_database.db")))
    parser.add_argument("--output", default="payroll_summary.html")
    args = parser.parse_args()

    if not os.path.exists(args.db):
        print(f"[ERROR] DB not found: {args.db}")
        sys.exit(1)

    conn = sqlite3.connect(args.db)
    print(f"Reading payroll_summary from {args.db} ...")
    rows = fetch_rows(conn)
    conn.close()

    out = Path(args.output).resolve()
    render_html(rows, str(out), args.db)
    print(f"[OK] {len(rows):,} rows → {out}")


if __name__ == "__main__":
    main()
