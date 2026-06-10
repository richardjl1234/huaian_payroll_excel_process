"""从 reconcile_report.html 提取 WARN 详情段为纯文本，方便滚动浏览/grep。

用法:
    python one_time_pgms/extract_warn_full.py
    python one_time_pgms/extract_warn_full.py --input reconcile_report.html --output /tmp/warn_full.txt
    python one_time_pgms/extract_warn_full.py --tolerance   # 提取 TOL 段而非 WARN
"""
import argparse
import re
import html
import os

DEFAULT_INPUT = "reconcile_report.html"
DEFAULT_OUTPUT = "/tmp/warn_full.txt"
DEFAULT_SECTION = "WARN 详情"  # 或 "Tolerance 匹配明细" 等


def extract(input_path, output_path, section_header):
    with open(input_path) as f:
        s = f.read()
    # 找段
    m = re.search(rf'<h2>{re.escape(section_header)}.*$', s, re.DOTALL)
    if not m:
        # 尝试下一个 h2 作为终止
        rest = s[m.end() if m else 0:]
        next_h2 = re.search(r'<h2>', rest)
        if next_h2:
            warn_html = s[m.start():m.start() + next_h2.start()]
        else:
            print(f"[ERROR] 段 '{section_header}' not found in {input_path}")
            return 1
    else:
        warn_html = m.group(0)
        # 截断到下一个 <h2>
        rest = s[m.end():]
        next_h2 = re.search(r'<h2>', rest)
        if next_h2:
            warn_html = warn_html[:-(len(rest) - next_h2.start())]
    # 标签替换
    text = warn_html
    for tag in ['h3', 'h4', 'p', 'pre']:
        text = re.sub(f'</?{tag}>', '\n', text)
    text = re.sub(r'<[^>]+>', '', text)
    text = html.unescape(text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    with open(output_path, 'w') as f:
        f.write(text)
    n_sections = len(re.findall(r'<h3>', warn_html))
    print(f"[OK] {len(text):,} bytes → {output_path} ({n_sections} 段)")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default=DEFAULT_INPUT)
    parser.add_argument("--output", default=DEFAULT_OUTPUT)
    parser.add_argument("--section", default=DEFAULT_SECTION,
                        help="HTML 报告里的 <h2> 段标题（默认 'WARN 详情'）")
    args = parser.parse_args()
    if not os.path.exists(args.input):
        print(f"[ERROR] {args.input} not found. Run reconcile first.")
        exit(1)
    extract(args.input, args.output, args.section)
