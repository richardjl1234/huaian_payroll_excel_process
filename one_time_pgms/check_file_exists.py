import os
import time

# 检查文件是否存在并显示基本信息
csv_file = "excel_headers_summary.csv"

if os.path.exists(csv_file):
    file_size = os.path.getsize(csv_file) / 1024  # KB
    mod_time = time.ctime(os.path.getmtime(csv_file))
    
    print(f"文件 {csv_file} 存在！")
    print(f"文件大小: {file_size:.2f} KB")
    print(f"最后修改时间: {mod_time}")
    
    # 尝试打开文件并读取少量内容
    try:
        with open(csv_file, 'r', encoding='utf-8') as f:
            first_lines = [next(f) for _ in range(min(3, sum(1 for _ in f)) + 1)]
            print("\n文件内容预览:")
            for i, line in enumerate(first_lines):
                print(f"第{i+1}行: {line.strip()[:100]}..." if len(line) > 100 else f"第{i+1}行: {line.strip()}")
    except Exception as e:
        print(f"读取文件内容时出错: {str(e)}")
else:
    print(f"错误: 文件 {csv_file} 不存在！")