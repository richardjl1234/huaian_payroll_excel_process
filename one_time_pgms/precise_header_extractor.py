import os
import xlrd
from datetime import datetime

# 精确的Excel标题提取器
def extract_headers_from_excel(file_path):
    """
    精确提取Excel文件中的表格标题行
    """
    # 检查文件是否存在
    if not os.path.exists(file_path):
        print(f"错误：文件不存在 - {file_path}")
        return False
    
    print(f"\n=== 开始处理文件：{os.path.basename(file_path)} ===")
    print(f"文件路径：{file_path}")
    print(f"处理时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    try:
        # 打开Excel文件
        workbook = xlrd.open_workbook(file_path)
        print(f"文件类型：.xls")
        print(f"工作表总数：{workbook.nsheets}")
        
        total_headers = 0
        
        # 遍历所有工作表
        for sheet_idx in range(workbook.nsheets):
            sheet_name = workbook.sheet_names()[sheet_idx]
            sheet = workbook.sheet_by_index(sheet_idx)
            
            # 忽略名为'汇总'的工作表
            if sheet_name == '汇总':
                print(f"\n忽略工作表：'汇总'")
                continue
            
            # 检查是否为空表
            if sheet.nrows <= 1 or sheet.ncols <= 1:
                print(f"\n忽略空工作表：'{sheet_name}'")
                continue
            
            print(f"\n处理工作表：{sheet_name}")
            print(f"工作表尺寸：{sheet.nrows}行 x {sheet.ncols}列")
            
            # 尝试在第一行找到标题
            if sheet.nrows > 0:
                headers = []
                has_content = False
                
                # 提取第一行作为标题行（根据我们看到的文件结构，标题通常在第一行）
                for col_idx in range(sheet.ncols):
                    cell_value = sheet.cell_value(0, col_idx)
                    if cell_value is not None and str(cell_value).strip() != '':
                        header_text = str(cell_value).strip()
                        headers.append(header_text)
                        has_content = True
                
                if has_content:
                    total_headers += 1
                    print(f"找到标题行（第1行）：")
                    print(f"标题: {', '.join(headers)}")
                
                # 显示前3行数据作为参考
                print(f"\n前3行数据示例：")
                max_rows = min(3, sheet.nrows)
                for row_idx in range(max_rows):
                    row_data = []
                    # 只显示前5列
                    max_cols = min(5, sheet.ncols)
                    for col_idx in range(max_cols):
                        cell_value = sheet.cell_value(row_idx, col_idx)
                        cell_text = "空" if cell_value == '' or cell_value is None else str(cell_value).strip()
                        row_data.append(cell_text)
                    print(f"行{row_idx+1}: {', '.join(row_data)}")
        
        print(f"\n=== 处理完成 ===")
        print(f"总标题行数量：{total_headers}")
        return True
    except Exception as e:
        print(f"处理文件时出错：{str(e)}")
        return False

if __name__ == '__main__':
    # 获取当前工作目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 处理202001文件
    file_path = os.path.join(current_dir, 'old_payroll', '202001.xls')
    extract_headers_from_excel(file_path)
    
    # 如果需要，也可以处理new_payroll中的202001文件（如果存在）
    new_file_path = os.path.join(current_dir, 'new_payroll', '202001.xls')
    if os.path.exists(new_file_path) and new_file_path != file_path:
        extract_headers_from_excel(new_file_path)