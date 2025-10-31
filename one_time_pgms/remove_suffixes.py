import os
import re
from datetime import datetime

# 定义路径
base_dir = os.path.dirname(os.path.abspath(__file__))
new_payroll_path = os.path.join(base_dir, 'new_payroll')
old_payroll_path = os.path.join(base_dir, 'old_payroll')

# 初始化计数器
total_processed = 0
renamed_successfully = 0
conflicts_resolved = 0
skipped_files = 0
no_suffix_files = 0

# 打印日志函数
def print_log(message):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f'[{timestamp}] {message}')

# 处理单个文件
def process_file(file_path, folder_path):
    global renamed_successfully, conflicts_resolved, skipped_files, no_suffix_files
    
    file_name = os.path.basename(file_path)
    base_name, extension = os.path.splitext(file_name)
    
    # 检查文件名是否符合 yyyymm_xxx 格式（xxx为任意后缀）
    # 这里使用正则表达式匹配 yyyymm 开头，后面跟着下划线和任意字符的模式
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
                print_log(f'重命名文件: {file_name} -> {new_file_name}')
                renamed_successfully += 1
                return True
            except Exception as e:
                print_log(f'重命名文件失败: {file_name}, 错误: {str(e)}')
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
                        print_log(f'重命名冲突文件: {file_name} -> {conflict_resolved_name}')
                        conflicts_resolved += 1
                        return True
                    except Exception as e:
                        print_log(f'重命名冲突文件失败: {file_name}, 错误: {str(e)}')
                        skipped_files += 1
                        return False
                
                counter += 1
    else:
        # 文件名不符合 yyyymm_xxx 格式，不需要重命名
        no_suffix_files += 1
        return True

# 处理指定文件夹
def process_folder(folder_path, folder_name):
    global total_processed
    
    print_log(f'开始处理文件夹: {folder_name}')
    print_log(f'文件夹路径: {folder_path}')
    
    # 检查文件夹是否存在
    if not os.path.exists(folder_path):
        print_log(f'错误: 文件夹 {folder_path} 不存在')
        return False
    
    # 获取文件夹中的所有文件
    files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    total_files = len(files)
    print_log(f'找到 {total_files} 个文件需要处理')
    
    # 处理每个文件
    for i, file_name in enumerate(files, 1):
        file_path = os.path.join(folder_path, file_name)
        print_log(f'处理文件 {i}/{total_files}: {file_name}')
        
        # 处理文件
        process_file(file_path, folder_path)
        total_processed += 1
    
    return True

# 主函数
def main():
    print_log('=' * 70)
    print_log('开始执行文件名后缀清理操作')
    print_log('=' * 70)
    
    # 处理 new_payroll 文件夹
    process_folder(new_payroll_path, 'new_payroll')
    print_log('-' * 70)
    
    # 处理 old_payroll 文件夹
    process_folder(old_payroll_path, 'old_payroll')
    print_log('=' * 70)
    
    # 输出总结
    print_log('操作总结:')
    print_log(f'总处理文件数: {total_processed}')
    print_log(f'成功移除后缀: {renamed_successfully}')
    print_log(f'解决文件冲突: {conflicts_resolved}')
    print_log(f'无后缀文件: {no_suffix_files}')
    print_log(f'跳过文件数: {skipped_files}')
    print_log('操作完成')
    print_log('=' * 70)

if __name__ == '__main__':
    main()