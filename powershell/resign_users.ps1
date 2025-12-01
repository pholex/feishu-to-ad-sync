# 禁用用户并移动到离职员工 OU
param(
    [string]$CsvPath = "$env:USERPROFILE\ad_resign_users.csv"
)

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

# 使用配置文件中的变量
$ResignedOU = $DC_RESIGNED_OU

$successCount = 0
$failedCount = 0

try {
    $users = Import-Csv -Path $CsvPath -Encoding UTF8
    
    Write-Host "`n===== 处理离职用户 =====" -ForegroundColor Cyan
    Write-Host "目标 OU: $ResignedOU`n"
    
    foreach ($user in $users) {
        $samAccountName = $user.SamAccountName
        $displayName = $user.DisplayName
        
        try {
            # 获取用户对象
            $adUser = Get-ADUser -Identity $samAccountName -Properties DistinguishedName
            
            if ($adUser) {
                # 禁用用户
                Disable-ADAccount -Identity $samAccountName
                
                # 移动到离职员工 OU
                Move-ADObject -Identity $adUser.DistinguishedName -TargetPath $ResignedOU
                
                Write-Host "已处理: $samAccountName $displayName" -ForegroundColor Green
                $successCount++
            }
        }
        catch {
            Write-Host "✗ 处理失败: $samAccountName $displayName - $_" -ForegroundColor Red
            $failedCount++
        }
    }
    
    Write-Host "`n===== 处理完成 =====" -ForegroundColor Cyan
    Write-Host "成功: $successCount 个用户" -ForegroundColor Green
    Write-Host "失败: $failedCount 个用户" -ForegroundColor Red
}
catch {
    Write-Host "错误: $_" -ForegroundColor Red
}
