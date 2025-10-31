import os
import re
from datetime import datetime

# 定义路径
base_dir = os.path.dirname(os.path.abspath(__file__))
new_payroll_path = os.path.join(base_dir, 'new_payroll')
old_payroll_path = os.path.join(base_dir, 'old_payroll')

# 定义日志文件路径
log_file_path = os.path.join(base_dir, 'remove_suffixes_log.txt')

# 初始化计数器
total_processed = 0
renamed_successfully = 0
conflicts_resolved = 0
skipped_files = 0
no_suffix_files = 0

# 打印并写入日志函数
def log_message(message):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    formatted_message = f'[{timestamp}] {message}'
    print(formatted_message)  # 输出到终端
    
    # 写入日志文件
    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        log_file.write(f'{formatted_message}\n')

# 处理单个文件
def process_file(file_path, folder_path):
    global renamed_successfully, conflicts_resolved, skipped_files, no_suffix_files
    
    file_name = os.path.basename(file_path)
    base_name, extension = os.path.splitext(file_name)
    
    # 检查文件名是否符合 yyyymm_xxx 格式（xxx为任意后缀）
    pattern = r'^(\d{6})(_.*)$'
    match = re.match(pattern, base_name)
    
    if match:
        # 提取基本文件名（没有后缀部分）
        clean_base_name = match.group(1)
        suffix = match.group(2)
        
        # 构建新文件名
        new_file_name = f'{clean_base_name}{extension}'
        new_file_path = os.path.join(folder_path, new_file_name)
        
        # 检查是否存在同名文件
        if not os.path.exists(new_file_path):
            # 不存在同名文件，可以直接重命名
            try:
                os.rename(file_path, new_file_path)
                log_message(f'重命名文件: {file_name} -> {new_file_name}')
                renamed_successfully += 1
                return True
            except Exception as e:
                log_message(f'重命名文件失败: {file_name}, 错误: {str(e)}')
                skipped_files += 1
                return False
        else:
            # 存在同名文件，需要添加数字后缀
            counter = 1
            while True:
                conflict_resolved_name = f'{clean_base_name}_{counter}{extension}'
                conflict_resolved_path = os.path.join(folder_path, conflict_resolved_name)
                
                if not os.path.exists(conflict_resolved_path):
                    try:
                        os.rename(file_path, conflict_resolved_path)
                        log_message(f'重命名冲突文件: {file_name} -> {conflict_resolved_name}')
                        conflicts_resolved += 1
                        return True
                    except Exception as e:
                        log_message(f'重命名冲突文件失败: {file_name}, 错误: {str(e)}')
                        skipped_files += 1
                        return False
                
                counter += 1
    else:
        # 文件名不符合 yyyymm_xxx 格式，不需要重命名
        log_message(f'文件无后缀，无需重命名: {file_name}')
        no_suffix_files += 1
        return True

# 处理指定文件夹
def process_folder(folder_path, folder_name):
    global total_processed
    
    log_message(f'开始处理文件夹: {folder_name}')
    log_message(f'文件夹路径: {folder_path}')
    
    # 检查文件夹是否存在
    if not os.path.exists(folder_path):
        log_message(f'错误: 文件夹 {folder_path} 不存在')
        return False
    
    # 获取文件夹中的所有文件
    files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    total_files = len(files)
    log_message(f'找到 {total_files} 个文件需要处理')
    
    # 处理每个文件
    for i, file_name in enumerate(files, 1):
        file_path = os.path.join(folder_path, file_name)
        log_message(f'处理文件 {i}/{total_files}: {file_name}')
        
        # 处理文件
        process_file(file_path, folder_path)
        total_processed += 1
    
    return True

# 主函数
def main():
    # 清空日志文件
    with open(log_file_path, 'w', encoding='utf-8') as log_file:
        log_file.write('文件名后缀清理操作日志\n')
        log_file.write('=' * 70 + '\n')
    
    log_message('=' * 70)
    log_message('开始执行文件名后缀清理操作')
    log_message('=' * 70)
    
    # 处理 new_payroll 文件夹
    process_folder(new_payroll_path, 'new_payroll')
    log_message('-' * 70)
    
    # 处理 old_payroll 文件夹
    process_folder(old_payroll_path, 'old_payroll')
    log_message('=' * 70)
    
    # 输出总结
    log_message('操作总结:')
    log_message(f'总处理文件数: {total_processed}')
    log_message(f'成功移除后缀: {renamed_successfully}')
    log_message(f'解决文件冲突: {conflicts_resolved}')
    log_message(f'无后缀文件: {no_suffix_files}')
    log_message(f'跳过文件数: {skipped_files}')
    log_message(f'日志文件已保存至: {log_file_path}')
    log_message('操作完成')
    log_message('=' * 70)

if __name__ == '__main__':
    main()