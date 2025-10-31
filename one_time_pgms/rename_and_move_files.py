import os
import re
import shutil
from datetime import datetime

# 定义路径
base_dir = os.path.dirname(os.path.abspath(__file__))
new_payroll_path = os.path.join(base_dir, 'new_payroll')
outliers_path = os.path.join(base_dir, 'outliers')

# 定义日志文件路径
log_file_path = os.path.join(base_dir, 'rename_move_log.txt')

# 初始化计数器
renamed_count = 0
moved_count = 0
skipped_count = 0

# 创建日志文件并写入开始信息
def write_log(message):
    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_file.write(f'[{timestamp}] {message}\n')
    print(message)  # 同时在控制台输出

# 检查并创建outliers目录
def ensure_outliers_directory():
    if not os.path.exists(outliers_path):
        os.makedirs(outliers_path)
        write_log(f'创建outliers目录: {outliers_path}')

# 重命名符合格式的文件
def rename_file(file_name):
    global renamed_count
    # 匹配格式: yyyy.m月.xls 或 yyyy.m.xls
    pattern = r'^(\d{4})\.(\d{1,2})(?:月)?\.xls(x)?$'
    match = re.match(pattern, file_name, re.IGNORECASE)
    
    if match:
        year = match.group(1)
        month = match.group(2).zfill(2)  # 确保月份是两位数
        extension = match.group(3) or ''  # 处理可能的xlsx扩展名
        
        # 构建新文件名
        new_file_name = f'{year}{month}.xls{extension}'
        
        # 构建完整路径
        old_file_path = os.path.join(new_payroll_path, file_name)
        new_file_path = os.path.join(new_payroll_path, new_file_name)
        
        # 检查新文件名是否已存在
        if os.path.exists(new_file_path):
            # 如果存在，添加一个数字后缀
            counter = 1
            while os.path.exists(new_file_path):
                new_file_path = os.path.join(new_payroll_path, f'{year}{month}_{counter}.xls{extension}')
                counter += 1
            new_file_name = os.path.basename(new_file_path)
        
        try:
            os.rename(old_file_path, new_file_path)
            write_log(f'重命名文件: {file_name} -> {new_file_name}')
            renamed_count += 1
            return True
        except Exception as e:
            write_log(f'重命名文件失败: {file_name}, 错误: {str(e)}')
            return False
    
    return False

# 移动不符合格式的文件到outliers目录
def move_to_outliers(file_name):
    global moved_count
    try:
        source_path = os.path.join(new_payroll_path, file_name)
        target_path = os.path.join(outliers_path, file_name)
        
        # 检查目标文件是否已存在
        if os.path.exists(target_path):
            # 如果存在，添加一个数字后缀
            base_name, extension = os.path.splitext(file_name)
            counter = 1
            while os.path.exists(target_path):
                target_path = os.path.join(outliers_path, f'{base_name}_{counter}{extension}')
                counter += 1
        
        shutil.move(source_path, target_path)
        write_log(f'移动文件到outliers: {file_name}')
        moved_count += 1
        return True
    except Exception as e:
        write_log(f'移动文件失败: {file_name}, 错误: {str(e)}')
        return False

# 主函数
def main():
    global skipped_count
    # 清空日志文件
    with open(log_file_path, 'w', encoding='utf-8') as log_file:
        log_file.write('文件重命名和移动操作日志\n')
        log_file.write('=' * 50 + '\n')
    
    write_log('开始执行文件重命名和移动操作')
    write_log(f'处理目录: {new_payroll_path}')
    
    # 确保outliers目录存在
    ensure_outliers_directory()
    
    # 获取new_payroll目录中的所有文件
    if not os.path.exists(new_payroll_path):
        write_log(f'错误: 目录 {new_payroll_path} 不存在')
        return
    
    files = [f for f in os.listdir(new_payroll_path) if os.path.isfile(os.path.join(new_payroll_path, f))]
    total_files = len(files)
    write_log(f'找到 {total_files} 个文件需要处理')
    
    # 处理每个文件
    for i, file_name in enumerate(files, 1):
        write_log(f'处理文件 {i}/{total_files}: {file_name}')
        
        # 尝试重命名文件
        if not rename_file(file_name):
            # 如果重命名失败，尝试移动到outliers目录
            if not move_to_outliers(file_name):
                write_log(f'跳过文件: {file_name} (无法重命名且无法移动)')
                skipped_count += 1
    
    # 输出总结
    write_log('\n操作总结:')
    write_log(f'总文件数: {total_files}')
    write_log(f'成功重命名: {renamed_count}')
    write_log(f'移动到outliers: {moved_count}')
    write_log(f'跳过: {skipped_count}')
    write_log(f'日志文件: {log_file_path}')
    write_log('操作完成')

if __name__ == '__main__':
    main()