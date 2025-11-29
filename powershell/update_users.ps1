param([switch]$DryRun)

# 导入Active Directory模块
Import-Module ActiveDirectory

# 设置输出编码为 UTF-8（支持中英文版 Windows）
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 加载配置文件（使用 UTF-8 编码）
$configPath = "$env:USERPROFILE\dc_config.ps1"
if (Test-Path $configPath) {
    $configContent = [System.IO.File]::ReadAllText($configPath, [System.Text.Encoding]::UTF8)
    Invoke-Expression $configContent
} else {
    Write-Error "配置文件不存在: $configPath"
    exit 1
}

# CSV文件路径
$csvPath = "$env:USERPROFILE\ad_update_accounts.csv"
$failuresCsvPath = "$env:USERPROFILE\UpdateAccountFailures.csv"

if ($DryRun) {
    Write-Host "`n===== 更新用户 [DRY-RUN] =====" -ForegroundColor Cyan
    Write-Host "预览模式 - 不会实际更新用户"
} else {
    Write-Host "`n===== 更新用户 =====" -ForegroundColor Green
}

Write-Host "开始处理用户..."

$successCount = 0
$failureCount = 0
$skippedCount = 0
$failures = @()

try {
    $users = Import-Csv -Path $csvPath -Encoding Default
    
    foreach ($user in $users) {
        # 确保所有属性都是字符串类型
        $samAccountName = [string]$user.SamAccountName
        $displayName = [string]$user.DisplayName
        $email = [string]$user.EmailAddress
        $employeeID = [string]$user.EmployeeID
        $employeeNumber = [string]$user.EmployeeNumber
        
        try {
            # 获取现有用户
            $adUser = Get-ADUser -Identity $samAccountName -Properties DisplayName, EmailAddress, EmployeeID, EmployeeNumber -ErrorAction Stop
            
            # 准备更新参数
            $updateParams = @{Identity = $adUser.DistinguishedName}
            $changes = @()
            
            # 检查 DisplayName
            if ($displayName -and $displayName.Trim() -ne "" -and $adUser.DisplayName -ne $displayName) {
                $updateParams.Add("DisplayName", $displayName)
                $changes += "DisplayName: '$($adUser.DisplayName)' -> '$displayName'"
            }
            
            # 检查 EmailAddress
            if ($email -and $email.Trim() -ne "" -and $adUser.EmailAddress -ne $email) {
                $updateParams.Add("EmailAddress", $email)
                $changes += "EmailAddress: '$($adUser.EmailAddress)' -> '$email'"
            }
            
            # 检查 EmployeeID
            if ($employeeID -and $employeeID.Trim() -ne "" -and $adUser.EmployeeID -ne $employeeID) {
                $updateParams.Add("EmployeeID", $employeeID)
                $changes += "EmployeeID: '$($adUser.EmployeeID)' -> '$employeeID'"
            }
            
            # 检查 EmployeeNumber
            if ($employeeNumber -and $employeeNumber.Trim() -ne "" -and $adUser.EmployeeNumber -ne $employeeNumber) {
                $updateParams.Add("EmployeeNumber", $employeeNumber)
                $changes += "EmployeeNumber: '$($adUser.EmployeeNumber)' -> '$employeeNumber'"
            }
            
            # 如果有参数需要更新
            if ($updateParams.Count -gt 1) {
                $changesList = $changes -join ', '
                if ($DryRun) {
                    Write-Host "[DRY-RUN] 将更新: $samAccountName $displayName ($changesList)" -ForegroundColor Yellow
                    $successCount++
                } else {
                    Set-ADUser @updateParams
                    Write-Host "✓ 更新成功: $samAccountName $displayName ($changesList)" -ForegroundColor Green
                    $successCount++
                }
            } else {
                if ($DryRun) {
                    # Dry-run 模式不显示跳过信息，避免输出过多
                } else {
                    # 正式执行时也不显示跳过，只统计
                }
                $skippedCount++
            }
        }
        catch {
            Write-Host "✗ 更新失败: $samAccountName - $_" -ForegroundColor Red
            $failureCount++
            
            # 记录失败详情
            $failures += [PSCustomObject]@{
                SamAccountName = $samAccountName
                DisplayName = $displayName
                ErrorMessage = $_.Exception.Message
            }
        }
    }
    
    # 导出失败列表
    if (-not $DryRun -and $failures.Count -gt 0) {
        $failures | Export-Csv -Path $failuresCsvPath -NoTypeInformation -Encoding Default
        Write-Host "`n失败详情已导出至: $failuresCsvPath" -ForegroundColor Red
    }
    
    Write-Host "`n===== 处理完成 =====" -ForegroundColor Cyan
    Write-Host ""
    if ($DryRun) {
        Write-Host "将更新: $successCount 个用户"
    } else {
        Write-Host "成功: $successCount 个用户" -ForegroundColor Green
    }
    if ($failureCount -gt 0) {
        Write-Host "处理失败: $failureCount 个用户" -ForegroundColor Red
    }
}
catch {
    Write-Host "错误: $_" -ForegroundColor Red
    exit 1
}
