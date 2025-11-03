import pandas as pd
import logging
import os
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP

def special_logic_preprocess_df(df: pd.DataFrame, sheet_name: str, file_name: str, table_index: int) -> tuple:
    """
    特殊逻辑预处理函数 - 在将DataFrame加载到SQLite数据库之前应用特殊逻辑
    
    Parameters:
        df (pd.DataFrame): 从Excel文件提取的原始数据框
        sheet_name (str): 工作表名称
        file_name (str): 文件名
        table_index (int): 表索引
        
    Returns:
        tuple: (processed_df, updated_sheet_name, updated_file_name) - 应用特殊逻辑后的数据框和更新的工作表名称、文件名
    """
    # 设置日志文件
    log_file = "special_logic_applied.log"
    
    def log_logic(description: str):
        """记录特殊逻辑应用的日志"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"{timestamp} | {file_name} | {sheet_name} | {table_index} | {description}\n"
        
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(log_entry)
    
    # 工作表名称映射逻辑
    original_sheet_name = sheet_name
    sheet_name_mappings = {
        '14年6月精加工': '精加工',
        '14年6月装配 喷漆': '装配喷漆',
        '14年6月绕嵌排': '绕嵌排',
        '装配 喷漆': '装配喷漆',
        '喷漆装配': '装配喷漆',
        '金加工': '精加工'
    }
    
    # 应用工作表名称映射
    if sheet_name in sheet_name_mappings:
        new_sheet_name = sheet_name_mappings[sheet_name]
        log_logic(f"工作表名称映射: '{sheet_name}' -> '{new_sheet_name}'")
        sheet_name = new_sheet_name
    
    # 如果数据框为空，直接返回
    if df.empty or len(df.columns) == 0:
        return df
    
    # 初始化操作计数器
    operation_counts = {
        '前装拆分': 0,
        '中装替换': 0,
        '后装替换': 0
    }
    
    # 逻辑1: 当工作表名称（去除空格）是"喷漆装配"时，如果第一列是"前装"、"中装"、"后装"或"刘雷", "装配"，则替换列名为"职员全名"
    clean_sheet_name = sheet_name.replace(' ', '')
    if clean_sheet_name == "喷漆装配" or clean_sheet_name == "装配喷漆":
        first_col = df.columns[0]
        if first_col in ["前装", "中装", "后装", "刘雷", "装配"]:
            df = df.rename(columns={first_col: "职员全名"})
            log_logic(f"将列名 '{first_col}' 替换为 '职员全名'")
    
    # 逻辑2: 当工作表名称（去除空格）是"喷漆装配"时，如果第一列是"后装曾大军"，则替换列名为"职员全名"，并将该列所有值设为"曾大军"
    if clean_sheet_name == "喷漆装配" or clean_sheet_name == "装配喷漆":
        first_col = df.columns[0]
        if first_col == "后装曾大军":
            df = df.rename(columns={first_col: "职员全名"})
            df["职员全名"] = "曾大军"
            log_logic(f"将列名 '{first_col}' 替换为 '职员全名' 并将所有值设为 '曾大军'")
    
    # 逻辑3: 无论工作表名称是什么，如果第一列是"姓名"，则替换为"职员全名"
    first_col = df.columns[0]
    if first_col == "姓名":
        df = df.rename(columns={first_col: "职员全名"})
        log_logic(f"将列名 '{first_col}' 替换为 '职员全名'")
    
    # 逻辑4: 当工作表名称是"绕嵌排"时，在"型号"列之后的列如果是"嵌线"或"排线"，则替换为"工序全名"
    if sheet_name == "绕嵌排":
        if "型号" in df.columns:
            # 找到"型号"列的位置
            model_col_index = list(df.columns).index("型号")
            # 检查"型号"列之后的列
            if model_col_index + 1 < len(df.columns):
                next_col = df.columns[model_col_index + 1]
                if next_col in ["嵌线", "排线"]:
                    df = df.rename(columns={next_col: "工序全名"})
                    log_logic(f"将列名 '{next_col}' 替换为 '工序全名'")
    
    # 逻辑5: 当工作表名称是"绕嵌排"时，在"型号"列之后的列如果是"工序名称"，则替换为"工序全名"
    if sheet_name == "绕嵌排":
        if "型号" in df.columns:
            # 找到"型号"列的位置
            model_col_index = list(df.columns).index("型号")
            # 检查"型号"列之后的列
            if model_col_index + 1 < len(df.columns):
                next_col = df.columns[model_col_index + 1]
                if next_col == "工序名称":
                    df = df.rename(columns={next_col: "工序全名"})
                    log_logic(f"将列名 '{next_col}' 替换为 '工序全名'")
    
    # 逻辑6: 如果存在列名为'数量'且同时存在列名为'职工全名'，则将'数量'改为'计件数量'
    if '数量' in df.columns and '职员全名' in df.columns:
        df = df.rename(columns={'数量': '计件数量'})
        log_logic(f"将列名 '数量' 替换为 '计件数量'")
    
    # 逻辑7: 如果存在列名为'加工型号'，则将'加工型号'改为'型号'
    if '加工型号' in df.columns:
        df = df.rename(columns={'加工型号': '型号'})
        log_logic(f"将列名 '加工型号' 替换为 '型号'")
    
    # 逻辑8: 在df.columns中，当'计件数量'包含在列名中时，将该列替换为'计件数量'
    for col in df.columns:
        if '计件数量' in col and col != '计件数量':
            df = df.rename(columns={col: '计件数量'})
            log_logic(f"将包含'计件数量'的列名 '{col}' 替换为 '计件数量'")
            break  # 只替换第一个匹配的列
    
    # 逻辑9: 在df.columns中，将'单位工资'替换为'定额'
    if '单位工资' in df.columns:
        df = df.rename(columns={'单位工资': '定额'})
        log_logic(f"将列名 '单位工资' 替换为 '定额'")
    
    # 逻辑10: 在df.columns中，将'合计金额'替换为'金额'
    if '合计金额' in df.columns:
        df = df.rename(columns={'合计金额': '金额'})
        log_logic(f"将列名 '合计金额' 替换为 '金额'")
    
    # 逻辑11: 在df.columns中，将'规格'替换为'型号'
    if '规格' in df.columns:
        df = df.rename(columns={'规格': '型号'})
        log_logic(f"将列名 '规格' 替换为 '型号'")
    
    # 逻辑12: 当存在'定额'列且其后有'合计'列时，将'合计'替换为'金额'
    if '定额' in df.columns and '合计' in df.columns:
        # 找到'定额'列的位置
        quota_col_index = list(df.columns).index('定额')
        # 检查'定额'列之后的列
        if quota_col_index + 1 < len(df.columns):
            next_col = df.columns[quota_col_index + 1]
            if next_col == '合计':
                df = df.rename(columns={'合计': '金额'})
                log_logic(f"将列名 '合计' 替换为 '金额' (在'定额'列之后)")
    
    # 逻辑13: 在df.columns中，将'任务名称'替换为'客户名称'
    if '任务名称' in df.columns:
        df = df.rename(columns={'任务名称': '客户名称'})
        log_logic(f"将列名 '任务名称' 替换为 '客户名称'")
    
    # 逻辑14: 当'职员全名'是'前装'或'前装人员'时，将记录拆分为2行
    if '职员全名' in df.columns and '计件数量' in df.columns and '金额' in df.columns:
        rows_to_add = []
        rows_to_remove = []
        
        for idx, row in df.iterrows():
            if row['职员全名'].startswith('前装'):
                # 创建第一行：黄志梅
                row1 = row.copy()
                row1['职员全名'] = '黄志梅'
                if pd.notna(row1['计件数量']) and row1['计件数量'] != '':
                    try:
                        # 使用pd.to_numeric安全转换，处理无效数值
                        numeric_val = pd.to_numeric([row1['计件数量']], errors='coerce')[0]
                        if pd.notna(numeric_val):
                            row1['计件数量'] = (Decimal(str(numeric_val)) / Decimal('2')).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                        else:
                            # 记录无效数值到日志，包括定额和金额的值
                            quota_value = row1.get('定额', 'N/A')
                            amount_value = row1.get('金额', 'N/A')
                            log_logic(f"无效的计件数量值 '{row1['计件数量']}' 在行 {idx}，使用默认值0, (定额的值是：{quota_value}, 金额的值是：{amount_value})")
                            row1['计件数量'] = Decimal('0')
                    except Exception as e:
                        # 记录转换错误到日志
                        log_logic(f"计件数量转换错误 '{row1['计件数量']}' 在行 {idx}: {str(e)}，使用默认值0")
                        row1['计件数量'] = Decimal('0')
                if pd.notna(row1['金额']) and row1['金额'] != '':
                    try:
                        # 使用pd.to_numeric安全转换，处理无效数值
                        numeric_val = pd.to_numeric([row1['金额']], errors='coerce')[0]
                        if pd.notna(numeric_val):
                            row1['金额'] = (Decimal(str(numeric_val)) / Decimal('2')).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                        else:
                            # 记录无效数值到日志
                            log_logic(f"无效的金额值 '{row1['金额']}' 在行 {idx}，使用默认值0")
                            row1['金额'] = Decimal('0')
                    except Exception as e:
                        # 记录转换错误到日志
                        log_logic(f"金额转换错误 '{row1['金额']}' 在行 {idx}: {str(e)}，使用默认值0")
                        row1['金额'] = Decimal('0')
                rows_to_add.append(row1)
                
                # 创建第二行：陈会清
                row2 = row.copy()
                row2['职员全名'] = '陈会清'
                if pd.notna(row2['计件数量']) and row2['计件数量'] != '':
                    try:
                        # 使用pd.to_numeric安全转换，处理无效数值
                        numeric_val = pd.to_numeric([row2['计件数量']], errors='coerce')[0]
                        if pd.notna(numeric_val):
                            row2['计件数量'] = (Decimal(str(numeric_val)) / Decimal('2')).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                        else:
                            # 记录无效数值到日志，包括定额和金额的值
                            quota_value = row2.get('定额', 'N/A')
                            amount_value = row2.get('金额', 'N/A')
                            log_logic(f"无效的计件数量值 '{row2['计件数量']}' 在行 {idx}，使用默认值0, (定额的值是：{quota_value}, 金额的值是：{amount_value})")
                            row2['计件数量'] = Decimal('0')
                    except Exception as e:
                        # 记录转换错误到日志
                        log_logic(f"计件数量转换错误 '{row2['计件数量']}' 在行 {idx}: {str(e)}，使用默认值0")
                        row2['计件数量'] = Decimal('0')
                if pd.notna(row2['金额']) and row2['金额'] != '':
                    try:
                        # 使用pd.to_numeric安全转换，处理无效数值
                        numeric_val = pd.to_numeric([row2['金额']], errors='coerce')[0]
                        if pd.notna(numeric_val):
                            row2['金额'] = (Decimal(str(numeric_val)) / Decimal('2')).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                        else:
                            # 记录无效数值到日志
                            log_logic(f"无效的金额值 '{row2['金额']}' 在行 {idx}，使用默认值0")
                            row2['金额'] = Decimal('0')
                    except Exception as e:
                        # 记录转换错误到日志
                        log_logic(f"金额转换错误 '{row2['金额']}' 在行 {idx}: {str(e)}，使用默认值0")
                        row2['金额'] = Decimal('0')
                rows_to_add.append(row2)
                
                # 标记原始行需要删除
                rows_to_remove.append(idx)
                operation_counts['前装拆分'] += 1
        
        # 删除原始行并添加新行
        if rows_to_remove:
            df = df.drop(rows_to_remove)
            new_rows_df = pd.DataFrame(rows_to_add)
            df = pd.concat([df, new_rows_df], ignore_index=True)
    
    # 逻辑15: 当'职员全名'是'中装'或'中装人员'时，将值改为'李兆军'
    if '职员全名' in df.columns:
        mask = df['职员全名'].str.startswith('中装', na=False)
        if mask.any():
            df.loc[mask, '职员全名'] = '李兆军'
            operation_counts['中装替换'] = mask.sum()
    
    # 逻辑16: 当'职员全名'是'后装'或'后装人员'时，将值改为'汤雅林'
    if '职员全名' in df.columns:
        mask = df['职员全名'].str.startswith('后装', na=False)
        if mask.any():
            df.loc[mask, '职员全名'] = '汤雅林'
            operation_counts['后装替换'] = mask.sum()
    
    # 记录汇总日志
    if operation_counts['前装拆分'] > 0:
        log_logic(f"将'前装'记录拆分为2行: 黄志梅 和 陈会清 共{operation_counts['前装拆分']}次")
    if operation_counts['中装替换'] > 0:
        log_logic(f"将'中装'改为'李兆军' 共{operation_counts['中装替换']}次")
    if operation_counts['后装替换'] > 0:
        log_logic(f"将'后装'改为'汤雅林' 共{operation_counts['后装替换']}次")
    
    # 添加文件名和工作表名列到DataFrame
    if not df.empty:
        df = df.copy()
        df.loc[:, '文件名'] = file_name
        df.loc[:, 'sheet名'] = sheet_name
    
    return df, sheet_name, file_name
