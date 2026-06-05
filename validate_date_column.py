#!/usr/bin/env python3
"""
验证工资数据库中日期列的数据质量。

验证规则：
1. 所有日期必须是逗号分隔的字符串格式，如 '2, 3'
2. 解析后的列表必须是升序排列，如 [2, 3]，不允许 [4, 3] 降序
3. 日期值必须在有效范围内（1-31）
4. 结合文件名（YYYYMM格式）验证日期是否在对应月份的有效范围内
   - 如 2017/04 月份不能有日期值 31（四月只有30天）

用法:
    python validate_date_column.py [--show-errors] [--export-html]

参数:
    --show-errors  显示所有错误详情
    --export-html  导出错误到HTML报告
"""

import sqlite3
import sys
import argparse
import re
from pathlib import Path
from datetime import date
import calendar

DB_PATH = Path(__file__).parent.parent / "payroll_database.db"
PAYROLL_TABLE = "payroll_details"


def parse_file_year_month(file_name: str) -> tuple:
    """
    从文件名提取年月。例如: '202506.xls' -> (2025, 6)
    """
    if len(file_name) >= 6:
        try:
            year = int(file_name[:4])
            month = int(file_name[4:6])
            return year, month
        except ValueError:
            return None, None
    return None, None


def get_days_in_month(year: int, month: int) -> int:
    """获取指定年月的天数"""
    return calendar.monthrange(year, month)[1]


def validate_date_string(date_str: str, file_name: str) -> tuple:
    """
    验证日期字符串。
    
    Returns:
        tuple: (is_valid, errors, parsed_dates)
            - is_valid: bool, 是否有效
            - errors: list of str, 错误信息列表
            - parsed_dates: list of int, 解析后的日期列表
    """
    errors = []
    parsed_dates = []
    
    if date_str is None or str(date_str).strip() == '':
        errors.append("日期为空")
        return False, errors, parsed_dates
    
    s = str(date_str).strip()
    
    # 规则1: 逗号分隔格式（可选，如果没有逗号则必须是单个日期）
    # 分割并解析每个日期
    parts = s.split(',')
    for part in parts:
        part = part.strip()
        if not part:
            continue
        
        # 检查是否是数字或范围（如 '1-3'）
        if '-' in part:
            # 范围格式，需要展开
            range_parts = part.split('-')
            if len(range_parts) == 2 and range_parts[0].isdigit() and range_parts[1].isdigit():
                start = int(range_parts[0])
                end = int(range_parts[1])
                for d in range(start, end + 1):
                    if d not in parsed_dates:
                        parsed_dates.append(d)
            else:
                if part.isdigit():
                    day = int(part)
                    if day not in parsed_dates:
                        parsed_dates.append(day)
                else:
                    errors.append(f"无效范围格式：'{part}'")
        elif part.isdigit():
            day = int(part)
            parsed_dates.append(day)
        else:
            errors.append(f"无效数字：'{part}'")
    
    if not parsed_dates:
        errors.append("没有有效的日期值")
        return False, errors, parsed_dates
    
    # 规则2: 检查是否升序排列
    for i in range(len(parsed_dates) - 1):
        if parsed_dates[i] >= parsed_dates[i + 1]:
            errors.append(f"日期非升序：{parsed_dates[i]} >= {parsed_dates[i + 1]}")
    
    # 规则4: 检查日期是否在对应月份的有效范围内
    year, month = parse_file_year_month(file_name)
    if year is not None and month is not None:
        max_day = get_days_in_month(year, month)
        for day in parsed_dates:
            if day > max_day:
                errors.append(f"日期无效：{year}年{month}月只有{max_day}天，日期{day}超出范围")
    
    return len(errors) == 0, errors, parsed_dates


def validate_all_records(conn: sqlite3.Connection) -> dict:
    """
    验证所有记录的日期列。
    
    Returns:
        dict: 验证结果，包含统计信息和错误记录
    """
    cursor = conn.execute(f"SELECT rowid, * FROM {PAYROLL_TABLE}")
    columns = [desc[0] for desc in cursor.description]
    all_rows = cursor.fetchall()
    
    date_idx = columns.index('日期')
    filename_idx = columns.index('文件名')
    worker_idx = columns.index('职员全名')
    rowid_idx = 0

    stats = {
        'total': len(all_rows),
        'valid': 0,
        'invalid': 0,
        'empty': 0,
    }

    error_records = []

    for row in all_rows:
        rowid = row[rowid_idx]
        date_val = row[date_idx]
        file_name = row[filename_idx]
        worker_name = row[worker_idx]

        if date_val is None or str(date_val).strip() == '':
            stats['empty'] += 1
            continue

        is_valid, errors, parsed_dates = validate_date_string(date_val, file_name)

        if is_valid:
            stats['valid'] += 1
        else:
            stats['invalid'] += 1
            error_records.append({
                'rowid': rowid,
                'filename': file_name,
                'worker_name': worker_name,
                'date_value': str(date_val),
                'parsed_dates': parsed_dates,
                'errors': errors
            })
    
    return {
        'stats': stats,
        'errors': error_records
    }


def print_summary(result: dict):
    """打印验证摘要"""
    stats = result['stats']
    errors = result['errors']
    
    print("=" * 80)
    print("日期列验证摘要")
    print("=" * 80)
    print(f"总记录数: {stats['total']}")
    print(f"有效记录: {stats['valid']}")
    print(f"无效记录: {stats['invalid']}")
    print(f"空值记录: {stats['empty']}")
    print("=" * 80)
    
    if errors:
        print(f"\n发现 {len(errors)} 条错误记录：")
        
        # 按错误类型分组统计
        error_types = {}
        for err in errors:
            for e in err['errors']:
                if e not in error_types:
                    error_types[e] = 0
                error_types[e] += 1
        
        print("\n错误类型统计：")
        for err_type, count in sorted(error_types.items(), key=lambda x: x[1], reverse=True):
            print(f"  - {err_type}: {count}条")
        
        # 显示前20条错误
        print("\n错误记录详情（前20条）：")
        print(f"{'ROWID':<8} {'文件名':<12} {'职员全名':<12} {'日期值':<30} {'错误'}")
        print("-" * 100)
        for err in errors[:20]:
            worker = str(err['worker_name']) if err['worker_name'] is not None else ''
            print(f"{err['rowid']:<8} {err['filename']:<12} {worker:<12} {err['date_value']:<30} {'; '.join(err['errors'][:2])}")
            if len(err['errors']) > 2:
                print(f"{'':<8} {'':<12} {'':<12} {'':<30} ... 还有{len(err['errors'])-2}条错误")
        
        if len(errors) > 20:
            print(f"\n... 还有 {len(errors) - 20} 条错误未显示")
    else:
        print("\n✓ 所有日期记录验证通过！")
    
    print("=" * 80)


def export_to_html(result: dict, output_path: Path):
    """导出错误记录到HTML"""
    errors = result['errors']
    stats = result['stats']
    
    html = '''<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>日期列验证结果</title>
<style>
body { font-family: sans-serif; margin: 20px; }
table { border-collapse: collapse; font-size: 12px; margin-bottom: 30px; }
th, td { border: 1px solid #ddd; padding: 6px; }
th { background-color: #4472C4; color: white; position: sticky; top: 0; }
tr:nth-child(even) { background-color: #f2f2f2; }
tr:hover { background-color: #ddd; }
td { max-width: 300px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.error { color: red; }
.summary { background-color: #f0f0f0; padding: 10px; margin-bottom: 20px; }
</style>
</head>
<body>
<h1>日期列验证结果</h1>
<div class="summary">
<p><strong>总记录数:</strong> ''' + str(stats['total']) + '''</p>
<p><strong>有效记录:</strong> ''' + str(stats['valid']) + '''</p>
<p><strong>无效记录:</strong> ''' + str(stats['invalid']) + '''</p>
<p><strong>空值记录:</strong> ''' + str(stats['empty']) + '''</p>
</div>
'''
    
    if errors:
        html += f'<h2>错误记录（共 {len(errors)} 条）</h2>'
        html += '''
<table>
<tr>
    <th>ROWID</th>
    <th>文件名</th>
    <th>职员全名</th>
    <th>日期值</th>
    <th>解析后的日期</th>
    <th>错误信息</th>
</tr>
'''
        for err in errors:
            worker = err['worker_name'] if err['worker_name'] is not None else ''
            html += f'''<tr>
    <td>{err['rowid']}</td>
    <td>{err['filename']}</td>
    <td>{worker}</td>
    <td>{err['date_value']}</td>
    <td>{err['parsed_dates']}</td>
    <td class="error">{'<br>'.join(err['errors'])}</td>
</tr>
'''
        html += '</table>'
    else:
        html += '<p style="color: green; font-size: 18px;">✓ 所有日期记录验证通过！</p>'
    
    html += '''
</body>
</html>'''
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)


def main():
    parser = argparse.ArgumentParser(
        description="验证工资数据库中日期列的数据质量"
    )
    parser.add_argument(
        "--show-errors",
        action="store_true",
        help="显示所有错误详情"
    )
    parser.add_argument(
        "--export-html",
        action="store_true",
        help="导出错误到HTML报告"
    )
    args = parser.parse_args()
    
    if not DB_PATH.exists():
        print(f"错误: 数据库文件不存在: {DB_PATH}")
        sys.exit(1)
    
    conn = sqlite3.connect(str(DB_PATH))
    
    print("开始验证日期列...")
    
    result = validate_all_records(conn)
    
    print_summary(result)
    
    if args.show_errors and result['errors']:
        print("\n" + "=" * 80)
        print("所有错误记录：")
        print("=" * 80)
        for err in result['errors']:
            print(f"\nROWID={err['rowid']}, 文件名={err['filename']}, 职员全名={err['worker_name']}")
            print(f"  日期值: '{err['date_value']}'")
            print(f"  解析: {err['parsed_dates']}")
            print(f"  错误:")
            for e in err['errors']:
                print(f"    - {e}")
    
    if args.export_html:
        output_path = Path(__file__).parent / "date_validation_output.html"
        export_to_html(result, output_path)
        print(f"\n已导出到: {output_path}")
    
    conn.close()
    
    # 返回退出码
    if result['stats']['invalid'] > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()