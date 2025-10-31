import os
import xlrd
import pandas as pd
from datetime import datetime
import logging

# 设置日志
def setup_logger():
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    
    # 清除现有的处理器
    if logger.handlers:
        logger.handlers.clear()
    
    # 添加控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter('%(message)s')
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)
    
    return logger

class SingleExcelFileProcessor:
    """
    单个Excel文件处理器，精确提取标题行并输出结果
    """
    def __init__(self, logger):
        self.logger = logger
        self.results = []
    
    def is_valid_sheet(self, sheet_name):
        """
        检查工作表是否有效（非'汇总'表）
        """
        return sheet_name != '汇总'
    
    def process_xls_file(self, file_path):
        """
        处理.xls格式的Excel文件，提取标题行
        """
        try:
            # 打开文件
            workbook = xlrd.open_workbook(file_path)
            file_name = os.path.basename(file_path)
            
            self.logger.info(f"\n=== 处理文件: {file_name} ===")
            self.logger.info(f"文件路径: {file_path}")
            self.logger.info(f"工作表总数: {workbook.nsheets}")
            
            valid_sheet_names = []
            all_headers = set()
            
            # 遍历所有工作表
            for sheet_idx in range(workbook.nsheets):
                sheet_name = workbook.sheet_names()[sheet_idx]
                sheet = workbook.sheet_by_index(sheet_idx)
                
                if not self.is_valid_sheet(sheet_name):
                    self.logger.info(f"忽略工作表: '{sheet_name}'")
                    continue
                
                # 检查工作表是否有数据
                if sheet.nrows <= 0 or sheet.ncols <= 0:
                    self.logger.info(f"忽略空工作表: '{sheet_name}'")
                    continue
                
                valid_sheet_names.append(sheet_name)
                
                # 提取标题行（位于第1行）
                headers = []
                for col_idx in range(sheet.ncols):
                    cell_value = sheet.cell_value(0, col_idx)
                    if cell_value is not None and str(cell_value).strip() != '':
                        header_text = str(cell_value).strip()
                        headers.append(header_text)
                        all_headers.add(header_text)
                
                if headers:
                    self.logger.info(f"\n工作表: {sheet_name}")
                    self.logger.info(f"尺寸: {sheet.nrows}行 x {sheet.ncols}列")
                    self.logger.info(f"标题行: {', '.join(headers)}")
                    
                    # 显示前2行数据作为参考
                    self.logger.info(f"数据示例（前2行）:")
                    max_rows = min(3, sheet.nrows)  # 包括标题行
                    for row_idx in range(1, max_rows):  # 从第2行开始（数据行）
                        row_data = []
                        # 只显示前5列
                        max_cols = min(5, sheet.ncols)
                        for col_idx in range(max_cols):
                            cell_value = sheet.cell_value(row_idx, col_idx)
                            cell_text = "空" if cell_value == '' or cell_value is None else str(cell_value).strip()
                            row_data.append(cell_text)
                        self.logger.info(f"行{row_idx+1}: {', '.join(row_data)}")
            
            # 保存结果
            self.results.append({
                'file_name': file_name,
                'sheet_names': valid_sheet_names,
                'headers_set': all_headers
            })
            
            return valid_sheet_names, all_headers
        except Exception as e:
            self.logger.error(f"处理文件时出错: {str(e)}")
            return [], set()
    
    def save_results_to_csv(self, output_file='single_file_results.csv'):
        """
        将结果保存到CSV文件
        """
        if not self.results:
            self.logger.warning("没有可保存的结果")
            return False
        
        try:
            # 准备数据
            csv_data = []
            for result in self.results:
                csv_data.append({
                    'file_name': result['file_name'],
                    'sheet_names': ', '.join(result['sheet_names']),
                    'headers_count': len(result['headers_set']),
                    'headers_sample': ', '.join(list(result['headers_set'])[:5]) + ('...' if len(result['headers_set']) > 5 else '')
                })
            
            # 创建DataFrame并保存
            df = pd.DataFrame(csv_data)
            df.to_csv(output_file, index=False, encoding='utf-8-sig')
            self.logger.info(f"\n结果已保存到: {output_file}")
            return True
        except Exception as e:
            self.logger.error(f"保存结果时出错: {str(e)}")
            return False

if __name__ == '__main__':
    start_time = datetime.now()
    
    # 设置日志
    logger = setup_logger()
    
    # 创建处理器
    processor = SingleExcelFileProcessor(logger)
    
    # 获取当前工作目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 用户可以修改这个文件名来处理不同的文件
    target_file = "202001.xls"
    
    # 查找并处理目标文件
    found_files = []
    for root, dirs, files in os.walk(current_dir):
        for file in files:
            if file.lower() == target_file.lower() and file.lower().endswith(('.xls', '.xlsx')):
                file_path = os.path.join(root, file)
                found_files.append(file_path)
    
    if not found_files:
        logger.error(f"未找到文件: {target_file}")
    else:
        logger.info(f"找到 {len(found_files)} 个匹配的文件")
        
        for file_path in found_files:
            valid_sheets, headers = processor.process_xls_file(file_path)
            logger.info(f"处理完成: {len(valid_sheets)}个有效工作表, {len(headers)}个唯一标题")
        
        # 保存结果
        processor.save_results_to_csv()
    
    end_time = datetime.now()
    logger.info(f"\n总处理时间: {end_time - start_time}")