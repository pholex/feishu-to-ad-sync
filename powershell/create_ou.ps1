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

# 加载 DN 转义/比较工具（与配置一同上传到域控，见 DnUtils.ps1）
$dnUtilsPath = "$env:USERPROFILE\DnUtils.ps1"
if (-not (Test-Path $dnUtilsPath)) {
    Write-Error "DN 工具不存在: $dnUtilsPath"
    exit 1
}
. $dnUtilsPath

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
        return "OU=$(ConvertTo-RdnValue $dept.dept_name),$BaseOU"
    } else {
        $parentPath = Get-OUPath -DepartmentId $parentId -Departments $Departments -BaseOU $BaseOU
        return "OU=$(ConvertTo-RdnValue $dept.dept_name),$parentPath"
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


# ===== 同名 OU 冲突处理：把目标位置的同名 OU 合并进源 OU =====
# 部门 OU 移动时，目标位置可能已有同名 OU（旧版脚本补建的空壳、或人工预建的结构）。
# 保留源 OU（Description=飞书部门ID，GUID 不变，GPO 链接/委派权限不受影响），
# 把目标 OU 的内容并入源 OU、删除已清空的目标 OU，然后再把源 OU 移过去。
# 无人值守运行，遇到以下情况一律不动、列出问题让人工处理：
#   - 要删除的 OU 链接了 GPO 或设置了阻止继承（删除会丢失）
#   - 存在用户以外的对象（计算机/组/联系人等：换 OU 可能改变适用的 GPO 或外部引用）
#   - 同名 OU 的部门ID不一致且无法按飞书判定（见 Resolve-DeptIdConflict）
#   - 要移入的对象在源 OU 里已有同名对象
#   - 整体移入的 OU 带有飞书中存在的部门ID，且该ID在别处已有 OU（会导致按ID查找到两个 OU）

# 飞书部门表：dept_id -> 部门行（主流程加载 CSV 后填充），用于判定部门ID是否仍有效
$script:FeishuDepts = @{}

# OU 的上级部门ID：父 OU 为基础 OU 时为 "0"（与飞书 parent_dept_id 约定一致）
function Get-ParentDeptId {
    param($DN)
    $parentDN = Get-ParentDN $DN
    if (Test-SameDN $parentDN $BaseOU) { return '0' }
    return [string](Get-ADObject -Identity $parentDN -Properties Description -ErrorAction Stop).Description
}

# 同名 OU 部门ID不一致时，按飞书判定保留哪个ID：
#   same        目标无ID或两者相同            → 正常合并
#   keepSource  目标ID在飞书中已不存在（残留）→ 合并，丢弃目标ID
#   adoptTarget 源ID已不存在、目标ID仍有效，且飞书中该部门的名称/上级与此位置一致
#               → 合并，并把源 OU 的 Description 改为目标ID（源 OU 对象/GUID 保留）
#   conflict    两个ID在飞书中都存在，或目标ID在飞书中属于别处 → 不能自动合并
# ParentDeptId：合并完成后源 OU 的上级部门ID（上级本身可能刚被 adoptTarget 改过ID）
function Resolve-DeptIdConflict {
    param($T, $S, $ParentDeptId)
    $tId = [string]$T.Description
    $sId = [string]$S.Description
    if (-not $tId -or $tId -eq $sId) { return @{ Action = 'same'; EffectiveId = $sId } }

    $tLive = $script:FeishuDepts.ContainsKey($tId)
    $sLive = $sId -and $script:FeishuDepts.ContainsKey($sId)
    if ($tLive -and $sLive) {
        return @{ Action = 'conflict'; Reason = "两个部门ID在飞书中都存在，是两个不同的部门: $($T.DistinguishedName) (Description='$tId') vs 源 $($S.DistinguishedName) (Description='$sId')" }
    }
    if (-not $tLive) {
        return @{ Action = 'keepSource'; EffectiveId = $sId; Note = "目标OU的部门ID '$tId' 在飞书中已不存在，保留源OU的部门ID '$sId'" }
    }
    $row = $script:FeishuDepts[$tId]
    if ($row.dept_name -ne $S.Name -or [string]$row.parent_dept_id -ne [string]$ParentDeptId) {
        return @{ Action = 'conflict'; Reason = "目标OU的部门ID '$tId' 在飞书中是「$($row.dept_name)」(上级ID $($row.parent_dept_id))，与合并后的位置不符: $($T.DistinguishedName)" }
    }
    return @{ Action = 'adoptTarget'; EffectiveId = $tId; Note = "源OU的部门ID '$sId' 在飞书中已不存在，采用目标OU的部门ID '$tId'" }
}

function Test-ADObjectExists {
    param($DN)
    try {
        $null = Get-ADObject -Identity $DN -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Test-HasGPLink {
    param($Obj)
    return ([string]$Obj.gPLink).Trim() -ne ''
}

# 返回让该 OU 不能算作"空壳"的对象：非 OU 对象、带 Description（部门ID）的 OU、链接了 GPO 或设置了阻止继承的 OU
# 空壳典型来源：旧版脚本在父 OU 移动失败后，为子部门自动补建的无 ID 父 OU
function Get-ShellBlockers {
    param($DN)
    $objs = @(Get-ADObject -SearchBase $DN -SearchScope Subtree -Filter * -Properties Description, gPLink, gPOptions -ErrorAction Stop)
    return @($objs | Where-Object { $_.ObjectClass -ne 'organizationalUnit' -or $_.Description -or (Test-HasGPLink $_) -or $_.gPOptions -eq 1 })
}

function Test-IsShellOU {
    param($DN)
    return (@(Get-ShellBlockers -DN $DN).Count -eq 0)
}

# 删除空壳 OU（含其下的空壳子 OU），先解除"防止意外删除"
function Remove-ShellOU {
    param($DN)
    Get-ADOrganizationalUnit -SearchBase $DN -SearchScope Subtree -Filter * -ErrorAction Stop |
        ForEach-Object { Set-ADOrganizationalUnit -Identity $_.ObjectGUID -ProtectedFromAccidentalDeletion $false -ErrorAction Stop }
    Remove-ADOrganizationalUnit -Identity $DN -Recursive -Confirm:$false -ErrorAction Stop
}

# 移动 AD 对象（OU 或用户）：开启"防止意外删除"的对象会被 Move-ADObject 拒绝（Access is denied），
# 因为移动=从原位置删除+挂到新位置。先临时关闭保护，移动后恢复（子对象随父 OU 一起移动）
function Move-ObjectSafely {
    param($DN, $TargetPath)
    $obj = Get-ADObject -Identity $DN -Properties ProtectedFromAccidentalDeletion -ErrorAction Stop
    $guid = $obj.ObjectGUID
    $wasProtected = $obj.ProtectedFromAccidentalDeletion
    if ($wasProtected) {
        Set-ADObject -Identity $guid -ProtectedFromAccidentalDeletion $false -ErrorAction Stop
    }
    try {
        Move-ADObject -Identity $guid -TargetPath $TargetPath -ErrorAction Stop
    } finally {
        if ($wasProtected) {
            Set-ADObject -Identity $guid -ProtectedFromAccidentalDeletion $true -ErrorAction SilentlyContinue
        }
    }
}

# 预检（只读）：检查把整棵 OU 子树原样移入源 OU 是否安全，输出问题描述
function Get-MoveInProblems {
    param($DN)
    $objs = @(Get-ADObject -SearchBase $DN -SearchScope Subtree -Filter * -Properties Description -ErrorAction Stop)
    foreach ($o in $objs) {
        if ($o.ObjectClass -eq 'organizationalUnit') {
            # 只有飞书中仍存在的部门ID才会被按ID查找，残留ID重复无影响
            if ($o.Description -and $script:FeishuDepts.ContainsKey([string]$o.Description)) {
                $d = $o.Description
                $dupes = @(Get-ADOrganizationalUnit -Filter { Description -eq $d } -SearchBase $BaseOU -ErrorAction Stop |
                    Where-Object { $_.ObjectGUID -ne $o.ObjectGUID })
                if ($dupes.Count -gt 0) {
                    "部门ID重复，移入后按ID会查到两个OU: $($o.DistinguishedName) 与 $($dupes[0].DistinguishedName) (Description='$d')"
                }
            }
        } elseif ($o.ObjectClass -ne 'user') {
            "非用户对象 [$($o.ObjectClass)]，需人工处理: $($o.DistinguishedName)"
        }
    }
}

# 预检（只读）：检查能否把 TargetDN 合并进 SourceDN，输出问题描述（无输出=可以合并）
function Get-MergeProblems {
    param($TargetDN, $SourceDN, $ParentDeptId)
    $t = Get-ADObject -Identity $TargetDN -Properties Description, gPLink, gPOptions -ErrorAction Stop
    $s = Get-ADObject -Identity $SourceDN -Properties Description -ErrorAction Stop
    if (Test-HasGPLink $t) { "链接了 GPO，合并后删除会丢失链接: $TargetDN" }
    if ($t.gPOptions -eq 1) { "设置了阻止继承，合并后删除会丢失该设置: $TargetDN" }
    if ($null -eq $ParentDeptId) { $ParentDeptId = Get-ParentDeptId $SourceDN }
    $res = Resolve-DeptIdConflict -T $t -S $s -ParentDeptId $ParentDeptId
    if ($res.Action -eq 'conflict') { $res.Reason; return }    # ID 冲突时不再深入检查子对象

    $children = @(Get-ADObject -SearchBase $TargetDN -SearchScope OneLevel -Filter * -ErrorAction Stop)
    foreach ($c in $children) {
        $destDN = "$(Get-RDN $c.DistinguishedName),$SourceDN"
        if ($c.ObjectClass -eq 'organizationalUnit') {
            if (Test-ADObjectExists $destDN) {
                Get-MergeProblems -TargetDN $c.DistinguishedName -SourceDN $destDN -ParentDeptId $res.EffectiveId    # 两边都有：递归合并
            } elseif (-not (Test-IsShellOU -DN $c.DistinguishedName)) {
                Get-MoveInProblems -DN $c.DistinguishedName                            # 源里没有：整体移入
            }                                                                          # 源里没有且是空壳：直接删除，无需检查
        } elseif ($c.ObjectClass -eq 'user') {
            if (Test-ADObjectExists $destDN) { "源OU中已有同名对象，无法移入: $($c.DistinguishedName) -> $destDN" }
        } else {
            "非用户对象 [$($c.ObjectClass)]，需人工处理: $($c.DistinguishedName)"
        }
    }
}

# 执行合并（须先通过 Get-MergeProblems 预检）：TargetDN 的内容并入 SourceDN，然后删除已清空的 TargetDN
function Merge-OUInto {
    param($TargetDN, $SourceDN, $ParentDeptId, [switch]$DryRun)
    $prefix = if ($DryRun) { "[DRY-RUN]   将" } else { "    ✓ 已" }

    # 部门ID按飞书判定（预检已排除 conflict）
    $t = Get-ADObject -Identity $TargetDN -Properties Description -ErrorAction Stop
    $s = Get-ADObject -Identity $SourceDN -Properties Description -ErrorAction Stop
    if ($null -eq $ParentDeptId) { $ParentDeptId = Get-ParentDeptId $SourceDN }
    $res = Resolve-DeptIdConflict -T $t -S $s -ParentDeptId $ParentDeptId
    if ($res.Action -eq 'conflict') { throw $res.Reason }
    if ($res.Note) { Write-Host "    ⓘ $($res.Note): $SourceDN" -ForegroundColor Yellow }
    # adoptTarget 的部门ID改写放到最后（目标 OU 删除之后）：中途任何一步失败时，
    # 目标仍持有该ID、源仍是旧ID，不会出现两个 OU 带同一个有效ID

    $children = @(Get-ADObject -SearchBase $TargetDN -SearchScope OneLevel -Filter * -ErrorAction Stop)
    foreach ($c in $children) {
        $destDN = "$(Get-RDN $c.DistinguishedName),$SourceDN"
        if ($c.ObjectClass -eq 'organizationalUnit' -and (Test-ADObjectExists $destDN)) {
            Merge-OUInto -TargetDN $c.DistinguishedName -SourceDN $destDN -ParentDeptId $res.EffectiveId -DryRun:$DryRun
        } elseif ($c.ObjectClass -eq 'organizationalUnit' -and (Test-IsShellOU -DN $c.DistinguishedName)) {
            if (-not $DryRun) { Remove-ShellOU -DN $c.DistinguishedName }
            Write-Host "${prefix}删除空壳OU: $($c.DistinguishedName)" -ForegroundColor Yellow
        } else {
            $fromDN = $c.DistinguishedName
            if (-not $DryRun) { Move-ObjectSafely -DN $fromDN -TargetPath $SourceDN }
            Write-Host "${prefix}移入 [$($c.ObjectClass)] $($c.Name): $fromDN -> $SourceDN" -ForegroundColor Yellow
        }
    }
    if (-not $DryRun) {
        # 已清空；不加 -Recursive，万一还有残留对象则报错而不是连带删除
        Set-ADOrganizationalUnit -Identity $TargetDN -ProtectedFromAccidentalDeletion $false -ErrorAction Stop
        Remove-ADOrganizationalUnit -Identity $TargetDN -Confirm:$false -ErrorAction Stop
    }
    Write-Host "${prefix}删除已清空的OU: $TargetDN" -ForegroundColor Yellow

    if ($res.Action -eq 'adoptTarget') {
        $oldId = [string]$s.Description
        if (-not $DryRun) { Set-ADOrganizationalUnit -Identity $s.ObjectGUID -Description $res.EffectiveId -ErrorAction Stop }
        Write-Host "${prefix}更新源OU的部门ID: $SourceDN Description '$oldId' -> '$($res.EffectiveId)'" -ForegroundColor Yellow
    }
}

# 移动前处理目标位置的同名 OU：空壳直接删除；否则预检通过后合并进源 OU；预检不过则报错并列出问题
function Clear-MoveTarget {
    param($TargetDN, $SourceDN, [switch]$DryRun)
    if (Test-SameDN $TargetDN $SourceDN) { return }    # 目标就是源自身（如仅大小写/转义写法不同），不能把自己并进自己
    if (-not (Test-OUExists -DN $TargetDN)) { return }

    if (Test-IsShellOU -DN $TargetDN) {
        if ($DryRun) {
            Write-Host "[DRY-RUN] 将删除目标位置的空壳OU（无部门ID、无用户）- $TargetDN" -ForegroundColor Cyan
        } else {
            Remove-ShellOU -DN $TargetDN
            Write-Host "  ✓ 已删除目标位置的空壳OU（无部门ID、无用户）- $TargetDN" -ForegroundColor Yellow
        }
        return
    }

    $problems = @(Get-MergeProblems -TargetDN $TargetDN -SourceDN $SourceDN)
    if ($problems.Count -gt 0) {
        $maxShow = 20
        $lines = @($problems | Select-Object -First $maxShow | ForEach-Object { "      - $_" })
        if ($problems.Count -gt $maxShow) { $lines += "      ... 共 $($problems.Count) 个，仅显示前 $maxShow 个" }
        throw ("目标位置已存在同名OU，无法自动合并: $TargetDN`n    问题 ($($problems.Count) 个):`n" + ($lines -join "`n"))
    }

    $head = if ($DryRun) { "[DRY-RUN] 将合并" } else { "  合并" }
    Write-Host "$head 目标位置的同名OU到源OU: $TargetDN -> $SourceDN" -ForegroundColor Cyan
    Merge-OUInto -TargetDN $TargetDN -SourceDN $SourceDN -DryRun:$DryRun
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
    foreach ($d in $departments) { $script:FeishuDepts[[string]$d.dept_id] = $d }
    
    # 记录统计
    $successCount = 0
    $updatedCount = 0
    $renamedCount = 0
    $movedCount = 0
    $skippedCount = 0
    $failureCount = 0
    
    Write-Host "开始处理部门 OU..." -ForegroundColor Cyan
    
    # 处理失败的部门ID：其所有子部门直接跳过，避免把子 OU 单独拆散到空壳父 OU 下
    $failedDeptIds = @{}

    foreach ($dept in $departments) {
        $deptId = $dept.dept_id
        $failBefore = $failureCount
        try {
            $deptName = $dept.dept_name
            $parentId = $dept.parent_dept_id
            $level = [int]$dept.level

            if ($failedDeptIds.ContainsKey($parentId)) {
                Write-Host "✗ 跳过 - $deptName : 上级部门处理失败，待上级修复后自动处理" -ForegroundColor DarkYellow
                $failureCount++
                continue
            }

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
            $deptRdn = ConvertTo-RdnValue $deptName
            $expectedDN = "OU=$deptRdn,$parentPath"
            $expectedDesc = $deptId
            
            # 检查父OU是否存在。部门已按拓扑序处理，父部门总是先于子部门被创建/移动，
            # 父OU不存在说明父部门没处理好——不能在这里补建无部门ID的父OU，否则会留下空壳，
            # 下次父部门移动时与之重名冲突（name already in use）
            # DryRun 下父OU可能只是"将创建/将移动"，不据此判失败
            if (-not $DryRun -and -not (Test-OUExists -DN $parentPath)) {
                Write-Host "✗ 跳过 - $deptName : 父OU不存在 ($parentPath)" -ForegroundColor Red
                $failureCount++
                continue
            }
            
            # 先通过部门ID查找（可能已改名或移动）
            $byId = @(Get-ADOrganizationalUnit -Filter {Description -eq $expectedDesc} -SearchBase $BaseOU -Properties Description,DistinguishedName,Name -ErrorAction SilentlyContinue)
            if ($byId.Count -gt 1) {
                # 同一部门ID在多个 OU 上：无法确定哪个是本体，不做任何改动
                Write-Host "✗ 跳过 - $deptName : 部门ID '$expectedDesc' 同时出现在 $($byId.Count) 个 OU 上，需人工确认保留哪个" -ForegroundColor Red
                $byId | ForEach-Object { Write-Host "    - $($_.DistinguishedName)" -ForegroundColor Red }
                $failureCount++
                continue
            }
            $existingOUByID = if ($byId.Count -eq 1) { $byId[0] } else { $null }

            if ($existingOUByID) {
                # 找到具有相同部门ID的OU
                $currentDN = $existingOUByID.DistinguishedName
                $currentName = $existingOUByID.Name
                
                if (Test-SameDN $currentDN $expectedDN) {
                    # 位置和名称都正确
                    $skippedCount++
                } else {
                    # 需要重命名和/或移动
                    $currentParent = Get-ParentDN $currentDN
                    $currentRdn = ConvertTo-RdnValue $currentName
                    $needRename = ($currentName -ne $deptName)
                    $needMove = -not (Test-SameDN $currentParent $parentPath)

                    # 既移动又改名时要经过一个中间位置：先移动 → OU=旧名,新父；先改名 → OU=新名,旧父。
                    # 中间位置上的同名 OU 与本部门无关，不能合并，选一个空着的顺序；都被占用则放弃。
                    $order = if ($needMove -and $needRename) {
                        if (-not (Test-OUExists -DN "OU=$currentRdn,$parentPath")) { 'moveFirst' }
                        elseif (-not (Test-OUExists -DN "OU=$deptRdn,$currentParent")) { 'renameFirst' }
                        else { 'blocked' }
                    } elseif ($needMove) { 'moveOnly' } else { 'renameOnly' }

                    try {
                        if ($order -eq 'blocked') {
                            throw "移动并改名需经过的两个中间位置都已有同名OU: OU=$currentRdn,$parentPath / OU=$deptRdn,$currentParent"
                        }
                        # 最终位置（新名,新父）上的同名 OU：空壳删除，否则并入本部门 OU（预检不过则报错）
                        Clear-MoveTarget -TargetDN $expectedDN -SourceDN $currentDN -DryRun:$DryRun

                        if ($DryRun) {
                            if ($needRename -and $needMove) {
                                Write-Host "[DRY-RUN] 将重命名并移动 - $currentName -> $deptName, 移动到 $parentPath" -ForegroundColor Cyan
                                $renamedCount++; $movedCount++
                            } elseif ($needRename) {
                                Write-Host "[DRY-RUN] 将重命名 - $currentName -> $deptName" -ForegroundColor Cyan
                                $renamedCount++
                            } else {
                                Write-Host "[DRY-RUN] 将移动 - $deptName 到 $parentPath" -ForegroundColor Cyan
                                $movedCount++
                            }
                        } else {
                            if ($order -eq 'renameFirst') {
                                Rename-ADObject -Identity $currentDN -NewName $deptName -ErrorAction Stop
                                Write-Host "✓ 已重命名 - $currentName -> $deptName" -ForegroundColor Cyan
                                $renamedCount++
                                $currentDN = "OU=$deptRdn,$currentParent"
                            }
                            if ($needMove) {
                                Move-ObjectSafely -DN $currentDN -TargetPath $parentPath
                                Write-Host "✓ 已移动 - $deptName 到 $parentPath" -ForegroundColor Cyan
                                $movedCount++
                                $currentDN = "$(Get-RDN $currentDN),$parentPath"
                            }
                            if ($needRename -and $order -ne 'renameFirst') {
                                Rename-ADObject -Identity $currentDN -NewName $deptName -ErrorAction Stop
                                Write-Host "✓ 已重命名 - $currentName -> $deptName" -ForegroundColor Cyan
                                $renamedCount++
                            }
                        }
                    } catch {
                        # 失败，记录详细错误信息
                        Write-Host "✗ OU移动/重命名失败 - $deptName" -ForegroundColor Red
                        Write-Host "  错误: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host "  错误类型: $($_.Exception.GetType().FullName)" -ForegroundColor Red
                        Write-Host "  当前位置: $currentDN" -ForegroundColor Red
                        Write-Host "  目标位置: $expectedDN" -ForegroundColor Red
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
                        
                        Write-Host "  ⚠ 这是同一个OU（部门ID匹配），必须移动/重命名，不能创建新OU；其子部门本次将跳过" -ForegroundColor Red
                        Write-Host "  请手动检查权限或手动处理此OU" -ForegroundColor Red
                        $failureCount++
                        continue
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
                    
                    if ($existingOUByName -and -not (Test-SameDN $existingOUByName.DistinguishedName $expectedDN)) {
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
                                    Clear-MoveTarget -TargetDN $expectedDN -SourceDN $oldDN
                                    Move-ObjectSafely -DN $oldDN -TargetPath $parentPath
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
        finally {
            # finally 在 continue 时也会执行：本部门有任何失败，就标记，让子部门跳过
            if ($failureCount -gt $failBefore) { $failedDeptIds[$deptId] = $true }
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
