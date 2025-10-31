# 文件名格式处理脚本

# 定义路径
$newPayrollPath = ".\new_payroll"
$outliersPath = ".\outliers"

# 创建outliers文件夹（如果不存在）
if (-not (Test-Path -Path $outliersPath -PathType Container)) {
    New-Item -Path $outliersPath -ItemType Directory
    Write-Host "已创建outliers文件夹"
}

# 获取new_payroll文件夹中的所有文件
$files = Get-ChildItem -Path $newPayrollPath -File

# 初始化计数器
$renamedCount = 0
$movedCount = 0
$skippedCount = 0

# 遍历所有文件
foreach ($file in $files) {
    $fileName = $file.Name
    Write-Host "处理文件: $fileName"
    
    # 检查文件名是否匹配yyyy.m月格式
    if ($fileName -match '^(\d{4})\.(\d{1,2})月\.(xls|xlsx)$') {
        $year = $Matches[1]
        $month = $Matches[2]
        $extension = $Matches[3]
        
        # 确保月份是两位数
        if ($month.Length -eq 1) {
            $month = "0$month"
        }
        
        # 构建新的文件名
        $newFileName = "${year}${month}.${extension}"
        
        # 重命名文件
        try {
            Rename-Item -Path $file.FullName -NewName $newFileName -Force
            Write-Host "  ✓ 已重命名为: $newFileName"
            $renamedCount++
        } catch {
            Write-Host "  ✗ 重命名失败: $_"
            $skippedCount++
        }
    } else {
        # 移动到outliers文件夹
        try {
            $destinationPath = Join-Path -Path $outliersPath -ChildPath $fileName
            Move-Item -Path $file.FullName -Destination $destinationPath -Force
            Write-Host "  ✓ 已移动到outliers文件夹"
            $movedCount++
        } catch {
            Write-Host "  ✗ 移动失败: $_"
            $skippedCount++
        }
    }
}

# 输出汇总信息
Write-Host "`n==== 处理汇总 ===="
Write-Host "重命名成功: $renamedCount 个文件"
Write-Host "移动到outliers: $movedCount 个文件"
Write-Host "处理失败: $skippedCount 个文件"
Write-Host "总计处理: $($files.Count) 个文件"