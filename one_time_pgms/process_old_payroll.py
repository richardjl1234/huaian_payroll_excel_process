import os
import re
import shutil
from datetime import datetime

# 定义路径
base_dir = os.path.dirname(os.path.abspath(__file__))
old_payroll_path = os.path.join(base_dir, 'old_payroll')
old_outliers_path = os.path.join(base_dir, 'old_outliers')

# 定义日志文件路径
log_file_path = os.path.join(base_dir, 'process_old_payroll_log.txt')

# 初始化计数器
processed_count = 0
renamed_count = 0
moved_to_outliers_count = 0
skipped_count = 0
removed_dirs_count = 0

# 创建日志文件并写入开始信息
def write_log(message):
    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_file.write(f'[{timestamp}] {message}\n')
    print(message)  # 同时在控制台输出

# 检查并创建old_outliers目录
def ensure_old_outliers_directory():
    if not os.path.exists(old_outliers_path):
        os.makedirs(old_outliers_path)
        write_log(f'创建old_outliers目录: {old_outliers_path}')

# 清理文件名中的特殊字符
def clean_filename(filename):
    # 替换全角字符为半角字符
    full_width_to_half_width = {
        '。': '.',  # 全角句号
        '，': ',',  # 全角逗号
        '：': ':',  # 全角冒号
        '；': ';',  # 全角分号
        '！': '!',  # 全角感叹号
        '？': '?',  # 全角问号
        '（': '(',  # 全角左括号
        '）': ')',  # 全角右括号
        '【': '[',  # 全角左方括号
        '】': ']',  # 全角右方括号
        '“': '"',  # 全角双引号
        '”': '"',  # 全角双引号
        '‘': "'",  # 全角单引号
        '’': "'",  # 全角单引号
        '、': ',',  # 全角顿号
    }
    
    # 替换全角字符
    for full, half in full_width_to_half_width.items():
        filename = filename.replace(full, half)
    
    # 去除多余的空格
    filename = ' '.join(filename.split())
    
    return filename

# 尝试从文件名中提取年份和月份并重命名
def extract_date_and_rename(file_path, filename):
    global renamed_count, moved_to_outliers_count, skipped_count
    
    # 清理文件名
    cleaned_filename = clean_filename(filename)
    base_name, extension = os.path.splitext(cleaned_filename)
    
    # 尝试匹配各种可能的日期格式
    # 格式1: yyyy.月份.xls (例如: 2015.8月.xls, 2018.10.xls)
    pattern1 = r'^(\d{4})\.(\d{1,2})(?:月)?(.*?)\.(xls[x]?)$'
    # 格式2: yyyy年mm月.xls (例如: 2015年8月.xls)
    pattern2 = r'^(\d{4})年(\d{1,2})月(.*?)\.(xls[x]?)$'
    # 格式3: yyyy.mm.xls (例如: 2019.06月.xls)
    pattern3 = r'^(\d{4})\.(\d{2})月(.*?)\.(xls[x]?)$'
    # 格式4: yyyy年m月.xls (例如: 2015年8月.xlsx)
    pattern4 = r'^(\d{4})年(\d{1})月(.*?)\.(xls[x]?)$'
    
    # 尝试匹配各种格式
    match = re.match(pattern1, cleaned_filename, re.IGNORECASE)
    if not match:
        match = re.match(pattern2, cleaned_filename, re.IGNORECASE)
    if not match:
        match = re.match(pattern3, cleaned_filename, re.IGNORECASE)
    if not match:
        match = re.match(pattern4, cleaned_filename, re.IGNORECASE)
    
    if match:
        year = match.group(1)
        month = match.group(2).zfill(2)  # 确保月份是两位数
        suffix = match.group(3).strip()
        ext = match.group(4).lower()
        
        # 构建新文件名
        if suffix:
            new_file_name = f'{year}{month}_{suffix}.{ext}'
        else:
            new_file_name = f'{year}{month}.{ext}'
        
        # 构建目标路径
        target_path = os.path.join(old_payroll_path, new_file_name)
        
        # 检查新文件名是否已存在
        counter = 1
        base_new_name, base_ext = os.path.splitext(new_file_name)
        while os.path.exists(target_path):
            target_path = os.path.join(old_payroll_path, f'{base_new_name}_{counter}{base_ext}')
            counter += 1
        
        try:
            # 移动并重命名文件
            shutil.move(file_path, target_path)
            write_log(f'重命名并移动文件: {filename} -> {os.path.basename(target_path)}')
            renamed_count += 1
            return True
        except Exception as e:
            write_log(f'重命名并移动文件失败: {filename}, 错误: {str(e)}')
            return False
    else:
        # 文件不符合日期格式，移动到outliers目录
        try:
            outliers_target_path = os.path.join(old_outliers_path, filename)
            
            # 检查目标文件是否已存在
            counter = 1
            base_name, ext = os.path.splitext(filename)
            while os.path.exists(outliers_target_path):
                outliers_target_path = os.path.join(old_outliers_path, f'{base_name}_{counter}{ext}')
                counter += 1
            
            shutil.move(file_path, outliers_target_path)
            write_log(f'移动文件到old_outliers: {filename}')
            moved_to_outliers_count += 1
            return True
        except Exception as e:
            write_log(f'移动文件失败: {filename}, 错误: {str(e)}')
            skipped_count += 1
            return False

# 处理子文件夹中的文件
def process_subdirectories():
    global processed_count, removed_dirs_count
    
    # 获取old_payroll下的所有子文件夹
    subdirs = [d for d in os.listdir(old_payroll_path) 
              if os.path.isdir(os.path.join(old_payroll_path, d))]
    
    total_dirs = len(subdirs)
    write_log(f'找到 {total_dirs} 个子文件夹需要处理')
    
    for dir_idx, subdir_name in enumerate(subdirs, 1):
        subdir_path = os.path.join(old_payroll_path, subdir_name)
        write_log(f'处理文件夹 {dir_idx}/{total_dirs}: {subdir_name}')
        
        # 获取子文件夹中的所有文件
        files = [f for f in os.listdir(subdir_path) 
                if os.path.isfile(os.path.join(subdir_path, f))]
        
        total_files_in_dir = len(files)
        write_log(f'  文件夹内有 {total_files_in_dir} 个文件')
        
        for file_idx, filename in enumerate(files, 1):
            file_path = os.path.join(subdir_path, filename)
            write_log(f'  处理文件 {file_idx}/{total_files_in_dir}: {filename}')
            
            # 处理文件
            extract_date_and_rename(file_path, filename)
            processed_count += 1
        
        # 尝试删除空的子文件夹
        try:
            os.rmdir(subdir_path)
            write_log(f'  删除空文件夹: {subdir_name}')
            removed_dirs_count += 1
        except Exception as e:
            write_log(f'  删除文件夹失败: {subdir_name}, 错误: {str(e)}')

# 主函数
def main():
    # 清空日志文件
    with open(log_file_path, 'w', encoding='utf-8') as log_file:
        log_file.write('处理old_payroll文件夹操作日志\n')
        log_file.write('=' * 50 + '\n')
    
    write_log('开始执行old_payroll文件夹处理操作')
    write_log(f'处理根目录: {old_payroll_path}')
    
    # 确保old_outliers目录存在
    ensure_old_outliers_directory()
    
    # 检查old_payroll目录是否存在
    if not os.path.exists(old_payroll_path):
        write_log(f'错误: 目录 {old_payroll_path} 不存在')
        return
    
    # 处理所有子文件夹
    process_subdirectories()
    
    # 输出总结
    write_log('\n操作总结:')
    write_log(f'总处理文件数: {processed_count}')
    write_log(f'成功重命名并移动到old_payroll根目录: {renamed_count}')
    write_log(f'移动到old_outliers: {moved_to_outliers_count}')
    write_log(f'跳过: {skipped_count}')
    write_log(f'删除的子文件夹数: {removed_dirs_count}')
    write_log(f'日志文件: {log_file_path}')
    write_log('操作完成')

if __name__ == '__main__':
    main()