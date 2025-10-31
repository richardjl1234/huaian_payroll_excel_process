import csv
import os

# 直接查看CSV文件内容
def simple_view_csv(csv_file):
    if not os.path.exists(csv_file):
        print(f"文件 {csv_file} 不存在")
        return
    
    print(f"\n=== {csv_file} 文件内容 ===")
    
    try:
        with open(csv_file, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            headers = next(reader)  # 获取标题行
            
            print(f"列名: {', '.join(headers)}")
            print("\n前5行数据:")
            
            # 显示前5行数据
            for i, row in enumerate(reader):
                if i < 5:
                    print(f"行 {i+1}: {row}")
                else:
                    break
            
    except Exception as e:
        print(f"读取文件时出错: {str(e)}")

if __name__ == "__main__":
    csv_file = "excel_headers_summary.csv"
    simple_view_csv(csv_file)
    
    # 检查文件大小
    if os.path.exists(csv_file):
        file_size = os.path.getsize(csv_file) / 1024  # KB
        print(f"\n文件大小: {file_size:.2f} KB")