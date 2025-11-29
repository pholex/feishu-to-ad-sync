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

# 使用配置文件中的变量
$BaseOU = $DC_BASE_OU
$Domain = $DC_DOMAIN

# CSV文件路径
$csvPath = "$env:USERPROFILE\ad_new_accounts.csv"
$passwordCsvPath = "$env:USERPROFILE\GeneratedPasswords.csv"
$failuresCsvPath = "$env:USERPROFILE\CreateAccountFailures.csv"

if ($DryRun) {
    Write-Host "`n===== 创建用户 [DRY-RUN] =====" -ForegroundColor Cyan
    Write-Host "预览模式 - 不会实际创建用户"
} else {
    Write-Host "`n===== 创建用户 =====" -ForegroundColor Green
}

Write-Host "开始处理用户..."

$successCount = 0
$failureCount = 0
$passwords = @()
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
        $deptName = [string]$user.DepartmentName
        
        # 构建 OU 路径
        $ouPath = $BaseOU
        if ($deptName) {
            $deptParts = $deptName -split '\\'
            [array]::Reverse($deptParts)
            foreach ($part in $deptParts) {
                $ouPath = "OU=$part,$ouPath"
            }
        }
        
        # 生成随机密码（包含大小写字母、数字和特殊字符）
        $upperCase = [char[]]'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        $lowerCase = [char[]]'abcdefghijklmnopqrstuvwxyz'
        $numbers = [char[]]'0123456789'
        $specialChars = [char[]]'!@#$%^&*()-_=+[]{}|;:",.<>/?'
        
        $passwordBuilder = [System.Text.StringBuilder]::new()
        $passwordBuilder.Append(($upperCase | Get-Random)) | Out-Null
        $passwordBuilder.Append(($lowerCase | Get-Random)) | Out-Null
        $passwordBuilder.Append(($numbers | Get-Random)) | Out-Null
        $passwordBuilder.Append(($specialChars | Get-Random)) | Out-Null
        
        $allChars = $upperCase + $lowerCase + $numbers + $specialChars
        for ($i = 4; $i -lt 12; $i++) {
            $passwordBuilder.Append(($allChars | Get-Random)) | Out-Null
        }
        
        $passwordChars = $passwordBuilder.ToString().ToCharArray()
        $password = -join ($passwordChars | Sort-Object {Get-Random})
        
        # 确保密码是字符串类型
        if ($password -isnot [string]) {
            $password = [string]$password
        }
        
        $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
        
        # UPN
        $upn = "$samAccountName@$Domain"
        
        if ($DryRun) {
            Write-Host "[DRY-RUN] 将创建: $samAccountName $displayName (工号=$employeeID, 邮箱=$email, 部门=$deptName)" -ForegroundColor Yellow
            $successCount++
        } else {
            try {
                $newUserParams = @{
                    Name = $displayName
                    SamAccountName = $samAccountName
                    UserPrincipalName = $upn
                    DisplayName = $displayName
                    EmailAddress = $email
                    EmployeeID = $employeeID
                    EmployeeNumber = $employeeNumber
                    AccountPassword = $securePassword
                    Enabled = $true
                    Path = $ouPath
                    PasswordNeverExpires = $false
                    ChangePasswordAtLogon = $true
                }
                
                New-ADUser @newUserParams
                
                Write-Host "✓ 创建成功: $samAccountName $displayName (工号=$employeeID, 邮箱=$email, 部门=$deptName)" -ForegroundColor Green
                $successCount++
                
                # 记录密码
                $passwords += [PSCustomObject]@{
                    SamAccountName = $samAccountName
                    DisplayName = $displayName
                    EmailAddress = $email
                    Password = $password
                    Department = $deptName
                }
            }
            catch {
                Write-Host "✗ 创建失败: $samAccountName - $_" -ForegroundColor Red
                $failureCount++
                
                # 记录失败详情
                $failures += [PSCustomObject]@{
                    SamAccountName = $samAccountName
                    DisplayName = $displayName
                    EmailAddress = $email
                    ErrorMessage = $_.Exception.Message
                }
            }
        }
    }
    
    # 导出密码
    if (-not $DryRun -and $passwords.Count -gt 0) {
        $passwords | Export-Csv -Path $passwordCsvPath -NoTypeInformation -Encoding Default
        Write-Host "`n密码已导出至: $passwordCsvPath" -ForegroundColor Yellow
        Write-Host "警告: 密码文件包含明文密码，请妥善保管并在使用后删除!" -ForegroundColor Red
    }
    
    # 导出失败列表
    if (-not $DryRun -and $failures.Count -gt 0) {
        $failures | Export-Csv -Path $failuresCsvPath -NoTypeInformation -Encoding Default
        Write-Host "失败详情已导出至: $failuresCsvPath" -ForegroundColor Red
    }
    
    Write-Host "`n===== 处理完成 =====" -ForegroundColor Cyan
    Write-Host ""
    if ($DryRun) {
        Write-Host "将创建: $successCount 个用户"
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
