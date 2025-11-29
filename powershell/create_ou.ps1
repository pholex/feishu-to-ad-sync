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

# 设置CSV文件路径
$csvPath = "$env:USERPROFILE\feishu_departments.csv"

if ($DryRun) {
    Write-Host "`n===== AD 域部门 OU 同步 [DRY-RUN] =====" -ForegroundColor Cyan
    Write-Host "预览模式 - 不会实际创建 OU" -ForegroundColor Yellow
} else {
    Write-Host "`n===== AD 域部门 OU 同步 =====" -ForegroundColor Cyan
}
Write-Host "基础 OU: $BaseOU" -ForegroundColor Yellow
Write-Host ""

# 递归函数：根据部门ID获取OU路径
function Get-OUPath {
    param($DepartmentId, $Departments, $BaseOU)
    
    $dept = $Departments | Where-Object {$_.dept_id -eq $DepartmentId}
    if (-not $dept) { return $BaseOU }
    
    $parentId = $dept.parent_dept_id
    if ($parentId -eq "0") {
        return "OU=$($dept.dept_name),$BaseOU"
    } else {
        $parentPath = Get-OUPath -DepartmentId $parentId -Departments $Departments -BaseOU $BaseOU
        return "OU=$($dept.dept_name),$parentPath"
    }
}

# 读取部门CSV
try {
    $departments = Import-Csv -Path $csvPath -Encoding Default -ErrorAction Stop
    
    # 按层级排序
    $departments = $departments | Sort-Object {[int]$_.层级}
    
    # 记录统计
    $successCount = 0
    $updatedCount = 0
    $renamedCount = 0
    $movedCount = 0
    $skippedCount = 0
    $failureCount = 0
    
    Write-Host "开始处理部门 OU..." -ForegroundColor Cyan
    
    foreach ($dept in $departments) {
        try {
            $deptName = $dept.dept_name
            $deptId = $dept.dept_id
            $parentId = $dept.parent_dept_id
            $level = [int]$dept.层级
            
            # 确定父OU路径
            if ($parentId -eq "0") {
                $parentPath = $BaseOU
            } else {
                $parentDept = $departments | Where-Object {$_.dept_id -eq $parentId}
                if ($parentDept) {
                    $parentPath = Get-OUPath -DepartmentId $parentId -Departments $departments -BaseOU $BaseOU
                } else {
                    $parentPath = $BaseOU
                }
            }
            
            # 期望的 DN 和 Description
            $expectedDN = "OU=$deptName,$parentPath"
            $expectedDesc = $deptId
            
            # 先通过部门ID查找（可能已改名或移动）
            $existingOUByID = Get-ADOrganizationalUnit -Filter {Description -eq $expectedDesc} -SearchBase $BaseOU -Properties Description,DistinguishedName,Name -ErrorAction SilentlyContinue
            
            if ($existingOUByID) {
                # 找到具有相同部门ID的OU
                $currentDN = $existingOUByID.DistinguishedName
                $currentName = $existingOUByID.Name
                
                if ($currentDN -eq $expectedDN) {
                    # 位置和名称都正确
                    $skippedCount++
                } else {
                    # 需要重命名或移动
                    $currentParent = ($currentDN -split ',',2)[1]
                    $needRename = ($currentName -ne $deptName)
                    $needMove = ($currentParent -ne $parentPath)
                    
                    if ($DryRun) {
                        if ($needRename -and $needMove) {
                            Write-Host "[DRY-RUN] 将重命名并移动 - $currentName -> $deptName, 移动到 $parentPath" -ForegroundColor Cyan
                            $renamedCount++
                            $movedCount++
                        } elseif ($needRename) {
                            Write-Host "[DRY-RUN] 将重命名 - $currentName -> $deptName" -ForegroundColor Cyan
                            $renamedCount++
                        } elseif ($needMove) {
                            Write-Host "[DRY-RUN] 将移动 - $deptName 到 $parentPath" -ForegroundColor Cyan
                            $movedCount++
                        }
                    } else {
                        # 先移动再重命名
                        if ($needMove) {
                            Move-ADObject -Identity $currentDN -TargetPath $parentPath
                            Write-Host "✓ 已移动 - $deptName 到 $parentPath" -ForegroundColor Cyan
                            $movedCount++
                            $currentDN = "OU=$currentName,$parentPath"
                        }
                        if ($needRename) {
                            Rename-ADObject -Identity $currentDN -NewName $deptName
                            Write-Host "✓ 已重命名 - $currentName -> $deptName" -ForegroundColor Cyan
                            $renamedCount++
                        }
                    }
                    $updatedCount++
                }
            } else {
                # 通过部门ID没找到，检查目标位置是否有同名OU
                $existingOUByDN = Get-ADOrganizationalUnit -Filter {DistinguishedName -eq $expectedDN} -Properties Description -ErrorAction SilentlyContinue
                
                if ($existingOUByDN) {
                    # 目标位置有同名OU，更新Description
                    if ($existingOUByDN.Description -ne $expectedDesc) {
                        if ($DryRun) {
                            Write-Host "[DRY-RUN] 将更新 - $deptName (Description: $($existingOUByDN.Description) -> $expectedDesc)" -ForegroundColor Yellow
                        } else {
                            Set-ADOrganizationalUnit -Identity $expectedDN -Description $expectedDesc
                            Write-Host "✓ 已更新 - $deptName (Description: $expectedDesc)" -ForegroundColor Yellow
                        }
                        $updatedCount++
                    } else {
                        $skippedCount++
                    }
                } else {
                    # 完全不存在，创建新OU
                    if ($DryRun) {
                        Write-Host "[DRY-RUN] 将创建 - $deptName (层级$level, Description=$deptId) -> $parentPath" -ForegroundColor Green
                        $successCount++
                    } else {
                        New-ADOrganizationalUnit -Name $deptName -Path $parentPath -Description $expectedDesc -ErrorAction Stop
                        Write-Host "✓ 成功创建 - $deptName (层级$level, Description=$deptId) -> $parentPath" -ForegroundColor Green
                        $successCount++
                    }
                }
            }
        }
        catch {
            Write-Host "✗ 处理失败 - $deptName : $_" -ForegroundColor Red
            $failureCount++
        }
    }
    
    # 输出汇总
    Write-Host "`n===== 处理完成 =====`n" -ForegroundColor Cyan
    if ($DryRun) {
        if ($successCount -gt 0) { Write-Host "将创建: $successCount 个OU" -ForegroundColor Green }
        if ($renamedCount -gt 0) { Write-Host "将重命名: $renamedCount 个OU" -ForegroundColor Cyan }
        if ($movedCount -gt 0) { Write-Host "将移动: $movedCount 个OU" -ForegroundColor Cyan }
        if ($updatedCount -gt 0) { Write-Host "将更新: $updatedCount 个OU" -ForegroundColor Yellow }
    } else {
        if ($successCount -gt 0) { Write-Host "成功创建: $successCount 个OU" -ForegroundColor Green }
        if ($renamedCount -gt 0) { Write-Host "已重命名: $renamedCount 个OU" -ForegroundColor Cyan }
        if ($movedCount -gt 0) { Write-Host "已移动: $movedCount 个OU" -ForegroundColor Cyan }
        if ($updatedCount -gt 0) { Write-Host "已更新: $updatedCount 个OU" -ForegroundColor Yellow }
    }
    if ($skippedCount -gt 0) { Write-Host "无需变更: $skippedCount 个OU" -ForegroundColor Gray }
    if ($failureCount -gt 0) { Write-Host "处理失败: $failureCount 个OU" -ForegroundColor Red }
}
catch {
    Write-Host "错误: $_" -ForegroundColor Red
}
