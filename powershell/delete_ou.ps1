# 删除多余的 OU
param(
    [string]$CsvPath = "$env:USERPROFILE\ad_delete_ous.csv"
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
$BaseOU = $DC_BASE_OU

$successCount = 0
$failedCount = 0
$skippedCount = 0

try {
    $ous = Import-Csv -Path $CsvPath -Encoding Default
    
    # 获取所有 OU 的完整信息并按层级深度排序（从深到浅）
    $ouList = @()
    foreach ($ou in $ous) {
        $ouName = $ou.'OU名称'
        $adOU = Get-ADOrganizationalUnit -Filter {Name -eq $ouName} -SearchBase $BaseOU -Properties DistinguishedName -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($adOU) {
            $depth = ($adOU.DistinguishedName -split ',').Count
            $ouList += [PSCustomObject]@{
                Name = $ouName
                DN = $adOU.DistinguishedName
                Depth = $depth
            }
        }
    }
    
    # 按深度降序排序（深层级先删除）
    $ouList = $ouList | Sort-Object -Property Depth -Descending
    
    Write-Host "`n===== 删除多余的 OU =====" -ForegroundColor Cyan
    Write-Host "基础 OU: $BaseOU"
    Write-Host "待处理: $($ouList.Count) 个OU`n"
    
    foreach ($ouItem in $ouList) {
        $ouName = $ouItem.Name
        $ouDN = $ouItem.DN
        
        try {
            # 检查是否是基础OU本身
            if ($ouDN -eq $BaseOU) {
                Write-Host "⚠ 跳过: $ouName (不能删除基础OU)" -ForegroundColor Yellow
                $skippedCount++
                continue
            }
            
            # 检查 OU 是否为空（没有用户和子 OU）
            $users = Get-ADUser -Filter * -SearchBase $ouDN -SearchScope OneLevel
            $childOUs = Get-ADOrganizationalUnit -Filter * -SearchBase $ouDN -SearchScope OneLevel
            
            if ($users.Count -eq 0 -and $childOUs.Count -eq 0) {
                # OU 为空，可以删除
                # 先取消防止意外删除保护
                Set-ADOrganizationalUnit -Identity $ouDN -ProtectedFromAccidentalDeletion $false
                Remove-ADOrganizationalUnit -Identity $ouDN -Confirm:$false
                Write-Host "✓ 已删除: $ouName" -ForegroundColor Green
                $successCount++
            } else {
                Write-Host "⚠ 跳过: $ouName (OU 不为空，包含 $($users.Count) 个用户和 $($childOUs.Count) 个子 OU)" -ForegroundColor Yellow
                $skippedCount++
            }
        }
        catch {
            Write-Host "✗ 删除失败: $ouName - $_" -ForegroundColor Red
            $failedCount++
        }
    }
    
    Write-Host "`n===== 处理完成 =====" -ForegroundColor Cyan
    Write-Host "成功删除: $successCount 个 OU" -ForegroundColor Green
    Write-Host "跳过: $skippedCount 个 OU" -ForegroundColor Yellow
    Write-Host "失败: $failedCount 个 OU" -ForegroundColor Red
}
catch {
    Write-Host "错误: $_" -ForegroundColor Red
}
