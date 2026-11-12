# 清理飞书中已不存在的部门 OU
#
# 判定依据是部门ID（OU 的 Description），不是名称——名称会重名，ID 不会。
# 须在部门同步（create_ou.ps1）和用户更新之后运行：此时有效部门都已就位，飞书员工都已移到正确 OU。
#
#   有效 OU：Description 是本次飞书部门ID 的 OU；或 DN 与某个飞书部门的期望位置（名称+层级）
#            完全一致的 OU——即使 Description 不一致也保留（如飞书删除重建了同名部门、
#            本次部门同步因上级失败未能修正其ID）——并顺带把其 Description 修正为飞书部门ID；
#            以及以上 OU 的所有上级 OU（上级没有ID也保留，否则会把下面的有效部门一起拆掉）
#   残留 OU：基础 OU 下（不含 DC_EXCLUDE_OU、DC_RESIGNED_OU 子树）其余所有 OU。
#            离职员工 OU 单独排除，不依赖它是否位于 DC_EXCLUDE_OU 之下——它可能为空，删掉后离职处理会失败
#
# 残留 OU 自底向上删除，只删"干净"的：没有链接 GPO、里面没有用户/计算机/组等任何非 OU 对象、
# 子 OU 都已删除。不满足的不删也不移动，列出原因交人工处理。
# 熔断：残留数超过现有部门 OU 的 MaxDeleteRatio（默认 20%）时整体跳过——多半是飞书数据不完整。
#       sync_to_ad.py 从环境变量 OU_CLEANUP_MAX_RATIO 读取并传入。
param(
    [switch]$DryRun,
    [double]$MaxDeleteRatio = 0.2,
    [string]$CsvPath = "$env:USERPROFILE\feishu_departments.csv"
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

# 加载 DN 转义/比较工具（与配置一同上传到域控，见 DnUtils.ps1）
$dnUtilsPath = "$env:USERPROFILE\DnUtils.ps1"
if (-not (Test-Path $dnUtilsPath)) {
    Write-Error "DN 工具不存在: $dnUtilsPath"
    exit 1
}
. $dnUtilsPath

$BaseOU = $DC_BASE_OU
$ExcludeOU = $DC_EXCLUDE_OU
$ResignedOU = $DC_RESIGNED_OU

function Get-Depth { param($DN) return (Split-DN $DN).Count }

$title = if ($DryRun) { "===== 清理飞书中已不存在的部门 OU [DRY-RUN] =====" } else { "===== 清理飞书中已不存在的部门 OU =====" }
Write-Host "`n$title" -ForegroundColor Cyan
Write-Host "基础 OU: $BaseOU"
if ($ExcludeOU) { Write-Host "排除 OU: $ExcludeOU" }
if ($ResignedOU -and -not (Test-UnderDN $ResignedOU $ExcludeOU)) { Write-Host "排除 OU（离职员工）: $ResignedOU" }

try {
    # ---- 飞书部门ID ----
    $feishuIds = @{}     # dept_id -> 部门行
    foreach ($row in @(Import-Csv -Path $CsvPath -Encoding UTF8 -ErrorAction Stop)) {
        if ($row.dept_id) { $feishuIds[[string]$row.dept_id] = $row }
    }
    if ($feishuIds.Count -eq 0) {
        Write-Host "✗ 飞书部门列表为空，跳过清理（防止误删全部 OU）" -ForegroundColor Red
        exit 1
    }

    # 每个飞书部门的期望 DN -> 部门ID（与 create_ou.ps1 的 Get-OUPath 规则一致：上级不在列表中时挂在基础 OU 下）
    $expectedDNs = @{}
    foreach ($id in @($feishuIds.Keys)) {
        $parts = @(); $cur = $id; $guard = 0
        while ($cur -and $cur -ne '0' -and $feishuIds.ContainsKey($cur) -and $guard -lt 64) {
            $parts += "OU=$(ConvertTo-RdnValue $feishuIds[$cur].dept_name)"
            $cur = [string]$feishuIds[$cur].parent_dept_id
            $guard++
        }
        $expectedDNs[(Get-NormalDN (($parts + $BaseOU) -join ','))] = $id
    }

    # ---- AD 中的部门 OU（不含基础 OU 本身、排除子树、离职员工 OU 子树）----
    $baseOUs = @(Get-ADOrganizationalUnit -Filter * -SearchBase $BaseOU -Properties Description, gPLink -ErrorAction Stop)
    $allOUs = @($baseOUs | Where-Object { -not (Test-SameDN $_.DistinguishedName $BaseOU) -and
        -not (Test-UnderDN $_.DistinguishedName $ExcludeOU) -and -not (Test-UnderDN $_.DistinguishedName $ResignedOU) })

    # ---- 有效 OU 及其所有上级 ----
    $keep = @{}
    foreach ($ou in $allOUs) {
        $idLive = $ou.Description -and $feishuIds.ContainsKey([string]$ou.Description)
        $posLive = $expectedDNs.ContainsKey((Get-NormalDN $ou.DistinguishedName))
        if ($idLive -or $posLive) {
            $dn = $ou.DistinguishedName
            while ($dn -and -not (Test-SameDN $dn $BaseOU) -and -not $keep.ContainsKey((Get-NormalDN $dn))) {
                $keep[(Get-NormalDN $dn)] = $true
                $dn = Get-ParentDN $dn
            }
        }
    }
    # ---- 位置与飞书部门一致、但 Description 不是该部门ID 的 OU：修正为飞书部门ID ----
    # 同一层级下 OU 不能重名，位置对上即可唯一确定部门。
    # 例外：该ID已在别的 OU 上（如该部门的正式 OU 搬迁失败仍在别处）——再写一份会让按ID查找到两个 OU，只报告不修改
    $idFixed = 0; $idSkipped = 0
    foreach ($ou in $allOUs) {
        $dnKey = Get-NormalDN $ou.DistinguishedName
        if (-not $expectedDNs.ContainsKey($dnKey)) { continue }
        $want = $expectedDNs[$dnKey]
        $old = [string]$ou.Description
        if ($old -eq $want) { continue }
        if ($old -and $feishuIds.ContainsKey($old)) {
            # 旧ID本身仍是飞书中有效的部门：这个 OU 是该部门的本体（如飞书改名后又新建了同名部门、
            # 而本次部门同步未能完成改名），改写会把它的人员/GUID/GPO 交给另一个部门
            Write-Host "⚠ 未修正部门ID: $($ou.DistinguishedName) (Description='$old')" -ForegroundColor Yellow
            Write-Host "    - 当前ID '$old' 仍是飞书部门「$($feishuIds[$old].dept_name)」的有效ID，此位置在飞书中属于 '$want'，需人工确认" -ForegroundColor Yellow
            $idSkipped++
            continue
        }
        $owner = @($baseOUs | Where-Object { [string]$_.Description -eq $want -and $_.ObjectGUID -ne $ou.ObjectGUID }) | Select-Object -First 1
        if ($owner) {
            Write-Host "⚠ 未修正部门ID: $($ou.DistinguishedName) (Description='$old')" -ForegroundColor Yellow
            Write-Host "    - 飞书部门ID '$want' 已在另一个 OU 上: $($owner.DistinguishedName)，需人工确认保留哪个" -ForegroundColor Yellow
            $idSkipped++
            continue
        }
        if ($DryRun) {
            Write-Host "[DRY-RUN] 将修正部门ID: $($ou.DistinguishedName) Description '$old' -> '$want'" -ForegroundColor Cyan
        } else {
            Set-ADOrganizationalUnit -Identity $ou.ObjectGUID -Description $want -ErrorAction Stop
            Write-Host "✓ 已修正部门ID: $($ou.DistinguishedName) Description '$old' -> '$want'" -ForegroundColor Green
        }
        $idFixed++
    }
    if ($idFixed -or $idSkipped) {
        $verb = if ($DryRun) { "将修正" } else { "已修正" }
        Write-Host "部门ID: $verb $idFixed 个，未修正 $idSkipped 个`n"
    }

    # 残留 OU，深层在前（先删子再删父）
    $stale = @($allOUs | Where-Object { -not $keep.ContainsKey((Get-NormalDN $_.DistinguishedName)) } |
        Sort-Object { Get-Depth $_.DistinguishedName } -Descending)

    Write-Host "部门 OU: $($allOUs.Count) 个，飞书部门: $($feishuIds.Count) 个，判定为残留: $($stale.Count) 个`n"
    if ($stale.Count -eq 0) {
        Write-Host "✓ 没有需要清理的 OU" -ForegroundColor Green
        exit 0
    }

    # ---- 熔断 ----
    $limit = [math]::Floor($allOUs.Count * $MaxDeleteRatio)
    $tripped = ($stale.Count -gt $limit) -and ($stale.Count -gt 5)
    if ($tripped) {
        Write-Host "⚠ 残留 OU 数 ($($stale.Count)) 超过现有部门 OU 的 $([int]($MaxDeleteRatio * 100))% ($limit)，疑似飞书数据不完整" -ForegroundColor Red
        if (-not $DryRun) {
            Write-Host "✗ 本次不做任何删除。核对下方清单无误后，设置环境变量 OU_CLEANUP_MAX_RATIO=1 重新运行一次同步即可" -ForegroundColor Red
            Write-Host "`n残留 OU 列表:"
            $stale | ForEach-Object { Write-Host "  - $($_.DistinguishedName)  (Description='$($_.Description)')" }
            exit 1
        }
        Write-Host "  [DRY-RUN] 继续预览；正式运行时将因熔断跳过`n" -ForegroundColor Yellow
    }

    # ---- 自底向上删除 ----
    $deleted = @{}       # 已删除（DryRun 下为"将删除"）的 DN
    $successCount = 0; $skippedCount = 0; $failedCount = 0

    foreach ($ou in $stale) {
        $dn = $ou.DistinguishedName
        $desc = if ($ou.Description) { " (Description='$($ou.Description)')" } else { "" }
        try {
            $children = @(Get-ADObject -SearchBase $dn -SearchScope OneLevel -Filter * -ErrorAction Stop)
            $objects = @($children | Where-Object { $_.ObjectClass -ne 'organizationalUnit' })
            $remainingChildOUs = @($children | Where-Object {
                $_.ObjectClass -eq 'organizationalUnit' -and -not $deleted.ContainsKey((Get-NormalDN $_.DistinguishedName)) })

            $reasons = @()
            if (([string]$ou.gPLink).Trim() -ne '') { $reasons += "链接了 GPO" }
            if ($objects.Count -gt 0) {
                $names = ($objects | Select-Object -First 5 | ForEach-Object { "[$($_.ObjectClass)] $($_.Name)" }) -join ', '
                if ($objects.Count -gt 5) { $names += " 等 $($objects.Count) 个" }
                $reasons += "包含对象: $names"
            }

            if ($reasons.Count -gt 0) {
                Write-Host "⚠ 跳过: $dn$desc" -ForegroundColor Yellow
                $reasons | ForEach-Object { Write-Host "    - $_" -ForegroundColor Yellow }
                $skippedCount++
                continue
            }
            if ($remainingChildOUs.Count -gt 0) {
                # 子 OU 的原因已在前面报告过（深层先处理），这里只记一行
                Write-Host "⚠ 跳过: $dn （子 OU 未能删除）" -ForegroundColor DarkYellow
                $skippedCount++
                continue
            }

            if ($DryRun) {
                Write-Host "[DRY-RUN] 将删除: $dn$desc" -ForegroundColor Cyan
            } else {
                Set-ADOrganizationalUnit -Identity $ou.ObjectGUID -ProtectedFromAccidentalDeletion $false -ErrorAction Stop
                # 不加 -Recursive：已确认为空，万一有残留对象则报错而不是连带删除
                Remove-ADOrganizationalUnit -Identity $ou.ObjectGUID -Confirm:$false -ErrorAction Stop
                Write-Host "✓ 已删除: $dn$desc" -ForegroundColor Green
            }
            $deleted[(Get-NormalDN $dn)] = $true
            $successCount++
        }
        catch {
            Write-Host "✗ 删除失败: $dn - $($_.Exception.Message)" -ForegroundColor Red
            $failedCount++
        }
    }

    Write-Host "`n===== 处理完成 =====" -ForegroundColor Cyan
    if ($DryRun) {
        Write-Host "将删除: $successCount 个 OU" -ForegroundColor Green
    } else {
        Write-Host "成功删除: $successCount 个 OU" -ForegroundColor Green
    }
    Write-Host "跳过: $skippedCount 个 OU（需人工处理）" -ForegroundColor Yellow
    Write-Host "失败: $failedCount 个 OU" -ForegroundColor Red
}
catch {
    Write-Host "错误: $_ (第 $($_.InvocationInfo.ScriptLineNumber) 行: $($_.InvocationInfo.Line.Trim()))" -ForegroundColor Red
    exit 1
}
