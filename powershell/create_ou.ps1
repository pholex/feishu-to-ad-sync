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

# 辅助函数：检查OU是否存在（通过DN）
function Test-OUExists {
    param($DN)
    try {
        $null = Get-ADOrganizationalUnit -Identity $DN -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

# 拓扑排序函数：确保父部门在子部门之前
function Sort-DepartmentsByDependency {
    param($Departments)
    
    $sorted = New-Object System.Collections.ArrayList
    $processed = @{}
    
    function Add-DepartmentRecursive {
        param($DeptId, $Depts, $SortedList, $ProcessedHash)
        
        if ($ProcessedHash.ContainsKey($DeptId)) { return }
        
        $dept = $Depts | Where-Object {$_.dept_id -eq $DeptId}
        if (-not $dept) { return }
        
        # 先处理父部门
        if ($dept.parent_dept_id -ne "0") {
            Add-DepartmentRecursive -DeptId $dept.parent_dept_id -Depts $Depts -SortedList $SortedList -ProcessedHash $ProcessedHash
        }
        
        # 再添加当前部门
        if (-not $ProcessedHash.ContainsKey($DeptId)) {
            [void]$SortedList.Add($dept)
            $ProcessedHash[$DeptId] = $true
        }
    }
    
    # 遍历所有部门
    foreach ($dept in $Departments) {
        Add-DepartmentRecursive -DeptId $dept.dept_id -Depts $Departments -SortedList $sorted -ProcessedHash $processed
    }
    
    return $sorted
}

# 读取部门CSV
try {
    $departments = Import-Csv -Path $csvPath -Encoding UTF8 -ErrorAction Stop
    
    # 拓扑排序：确保父OU在子OU之前创建
    $departments = Sort-DepartmentsByDependency -Departments $departments
    
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
            $level = [int]$dept.level
            
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
            
            # 调试：输出父路径
            # Write-Host "  调试: $deptName 的父路径: $parentPath" -ForegroundColor Gray
            
            # 期望的 DN 和 Description
            $expectedDN = "OU=$deptName,$parentPath"
            $expectedDesc = $deptId
            
            # 检查父OU是否存在，如果不存在则尝试创建
            if (-not (Test-OUExists -DN $parentPath)) {
                # 尝试递归创建父OU（从根到叶）
                $parentParts = $parentPath -split ',' | Where-Object { $_ -like 'OU=*' }
                [array]::Reverse($parentParts)  # 反转数组，从根部门开始
                $dcParts = ($parentPath -split ',' | Where-Object { $_ -like 'DC=*' }) -join ','
                $currentPath = $dcParts
                $createFailed = $false
                
                foreach ($part in $parentParts) {
                    $ouName = $part -replace '^OU=', ''
                    $testPath = "$part,$currentPath"
                    
                    if (-not (Test-OUExists -DN $testPath)) {
                        try {
                            if (-not $DryRun) {
                                New-ADOrganizationalUnit -Name $ouName -Path $currentPath -ErrorAction Stop
                                Write-Host "  ✓ 自动创建父OU - $ouName" -ForegroundColor Green
                            }
                        } catch {
                            Write-Host "✗ 跳过 - $deptName : 无法创建父OU ($ouName): $($_.Exception.Message)" -ForegroundColor Red
                            $failureCount++
                            $createFailed = $true
                            break  # 跳出 foreach ($part)
                        }
                    }
                    $currentPath = $testPath
                }
                
                # 如果创建失败，跳过当前部门
                if ($createFailed) {
                    continue  # 跳出 foreach ($dept)
                }
                
                # 再次检查父OU
                if (-not (Test-OUExists -DN $parentPath)) {
                    Write-Host "✗ 跳过 - $deptName : 父OU仍不存在 ($parentPath)" -ForegroundColor Red
                    $failureCount++
                    continue
                }
            }
            
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
                            try {
                                Move-ADObject -Identity $currentDN -TargetPath $parentPath -ErrorAction Stop
                                Write-Host "✓ 已移动 - $deptName 到 $parentPath" -ForegroundColor Cyan
                                $movedCount++
                                $currentDN = "OU=$currentName,$parentPath"
                            } catch {
                                # 移动失败，记录详细错误信息
                                Write-Host "✗ OU移动失败 - $deptName" -ForegroundColor Red
                                Write-Host "  错误: $($_.Exception.Message)" -ForegroundColor Red
                                Write-Host "  错误类型: $($_.Exception.GetType().FullName)" -ForegroundColor Red
                                Write-Host "  当前位置: $currentDN" -ForegroundColor Red
                                Write-Host "  目标位置: OU=$deptName,$parentPath" -ForegroundColor Red
                                Write-Host "  部门ID: $expectedDesc" -ForegroundColor Red
                                
                                # 检查目标父OU是否真的存在
                                if (Test-OUExists -DN $parentPath) {
                                    Write-Host "  目标父OU: 存在" -ForegroundColor Yellow
                                } else {
                                    Write-Host "  目标父OU: 不存在！" -ForegroundColor Red
                                }
                                
                                # 检查当前OU是否有保护
                                $ouObj = Get-ADOrganizationalUnit -Identity $currentDN -Properties ProtectedFromAccidentalDeletion -ErrorAction SilentlyContinue
                                if ($ouObj) {
                                    Write-Host "  防止意外删除: $($ouObj.ProtectedFromAccidentalDeletion)" -ForegroundColor Yellow
                                }
                                
                                # 检查OU中是否有对象
                                $childCount = (Get-ADObject -Filter * -SearchBase $currentDN -SearchScope OneLevel -ErrorAction SilentlyContinue | Measure-Object).Count
                                Write-Host "  OU中对象数量: $childCount" -ForegroundColor Yellow
                                
                                Write-Host "  ⚠ 这是同一个OU（部门ID匹配），必须移动，不能创建新OU" -ForegroundColor Red
                                Write-Host "  请手动检查权限或手动移动此OU" -ForegroundColor Red
                                $failureCount++
                                continue
                            }
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
                    # 目标位置没有，检查是否有同名OU在其他位置
                    $existingOUByName = Get-ADOrganizationalUnit -Filter {Name -eq $deptName} -SearchBase $BaseOU -Properties Description,DistinguishedName -ErrorAction SilentlyContinue | Select-Object -First 1
                    
                    if ($existingOUByName -and $existingOUByName.DistinguishedName -ne $expectedDN) {
                        # 找到同名OU在其他位置
                        $oldDN = $existingOUByName.DistinguishedName
                        $oldDesc = $existingOUByName.Description
                        
                        if ($oldDesc -eq $expectedDesc) {
                            # Description匹配，说明是同一个部门，移动到正确位置
                            if ($DryRun) {
                                Write-Host "[DRY-RUN] 将移动 - $deptName 从 $oldDN 到 $parentPath" -ForegroundColor Cyan
                                $movedCount++
                            } else {
                                try {
                                    Move-ADObject -Identity $oldDN -TargetPath $parentPath -ErrorAction Stop
                                    Write-Host "✓ 已移动 - $deptName 从其他位置到 $parentPath" -ForegroundColor Cyan
                                    $movedCount++
                                } catch {
                                    Write-Host "✗ 移动失败 - $deptName : $($_.Exception.Message)" -ForegroundColor Red
                                    $failureCount++
                                    continue
                                }
                            }
                            $updatedCount++
                        } else {
                            # Description不匹配或为空，说明是不同的部门，在正确位置创建新OU
                            Write-Host "  ⚠ 注意: 存在同名OU在其他位置 ($oldDN)，将在正确位置创建新OU" -ForegroundColor Yellow
                            if ($DryRun) {
                                Write-Host "[DRY-RUN] 将创建 - $deptName (层级$level, Description=$deptId) -> $parentPath" -ForegroundColor Cyan
                                $successCount++
                            } else {
                                try {
                                    New-ADOrganizationalUnit -Name $deptName -Path $parentPath -Description $expectedDesc -ErrorAction Stop
                                    Write-Host "✓ 成功创建 - $deptName (层级$level, Description=$deptId) -> $parentPath" -ForegroundColor Green
                                    $successCount++
                                } catch {
                                    Write-Host "✗ 创建失败 - $deptName : $($_.Exception.Message)" -ForegroundColor Red
                                    $failureCount++
                                    continue
                                }
                            }
                        }
                    } else {
                        # 完全不存在，直接创建新OU
                        if ($DryRun) {
                            Write-Host "[DRY-RUN] 将创建 - $deptName (层级$level, Description=$deptId) -> $parentPath" -ForegroundColor Cyan
                            $successCount++
                        } else {
                            try {
                                New-ADOrganizationalUnit -Name $deptName -Path $parentPath -Description $expectedDesc -ErrorAction Stop
                                Write-Host "✓ 成功创建 - $deptName (层级$level, Description=$deptId) -> $parentPath" -ForegroundColor Green
                                $successCount++
                            } catch {
                                Write-Host "✗ 创建失败 - $deptName : $($_.Exception.Message)" -ForegroundColor Red
                                $failureCount++
                            }
                        }
                    }
                }
            }
        }
        catch {
            Write-Host "✗ 处理失败 - $deptName : $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "  错误位置: $($_.InvocationInfo.ScriptLineNumber) 行" -ForegroundColor Gray
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
