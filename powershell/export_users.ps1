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

# 使用配置文件中的变量
$BaseOU = $DC_BASE_OU
$ExcludeOU = $DC_EXCLUDE_OU

# 设置导出的CSV文件路径
$csvFilePath = "$env:USERPROFILE\ExportedUsers.csv"

try {
    # 获取所有域用户（包括禁用的）
    $users = Get-ADUser -Filter * -SearchBase $BaseOU -Properties SamAccountName, EmailAddress, Mobile, EmployeeID, EmployeeNumber, info, DisplayName, Enabled, DistinguishedName
    
    # 排除指定 OU，但离职员工 OU 总是包含
    $filteredUsers = $users | Where-Object {
        ($_.DistinguishedName -notlike "*$ExcludeOU") -or ($_.DistinguishedName -like "*$DC_RESIGNED_OU*")
    }
    
    if ($filteredUsers) {
        # 导出到CSV
        $filteredUsers | Select-Object SamAccountName, EmailAddress, Mobile, EmployeeID, EmployeeNumber, info, DisplayName, Enabled, DistinguishedName | 
            Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8
        
        Write-Host "成功导出 $($filteredUsers.Count) 个用户（已排除指定 OU）" -ForegroundColor Green
    } else {
        Write-Host "未找到用户" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "错误: $_" -ForegroundColor Red
}

