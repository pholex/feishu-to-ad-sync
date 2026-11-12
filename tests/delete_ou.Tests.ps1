# delete_ou.ps1 回归测试：按部门ID清理飞书中已不存在的 OU、修正位置一致 OU 的部门ID
# 运行：pwsh -NoProfile -File tests/delete_ou.Tests.ps1   （或 tests/run.ps1 跑全部）
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot/MockAD.ps1"

$ScriptPath = (Resolve-Path (Join-Path $PSScriptRoot '../powershell/delete_ou.ps1')).Path
$BaseOU = 'OU=员工,DC=t'
$Exclude = "OU=其它,$BaseOU"
$Resigned = "OU=已离职,$BaseOU"     # 离职员工 OU 故意不放在排除子树下，验证显式排除

# 被测脚本从 $env:USERPROFILE 读取 dc_config.ps1 和 feishu_departments.csv，指向临时目录
$work = Join-Path ([System.IO.Path]::GetTempPath()) ("feishu-ad-sync-tests-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $work | Out-Null
$savedProfile = $env:USERPROFILE
$env:USERPROFILE = $work
$cfg = "`$DC_BASE_OU = `"$BaseOU`"`n`$DC_EXCLUDE_OU = `"$Exclude`"`n`$DC_RESIGNED_OU = `"$Resigned`""
Set-Content -Path (Join-Path $work 'dc_config.ps1') -Encoding UTF8 -Value $cfg
# 脚本用 Windows 路径 "$env:USERPROFILE\dc_config.ps1" 交给 .NET ReadAllText：非 Windows 上反斜杠不是分隔符，
# 需另写一份字面文件名带反斜杠的副本（Windows 上就是同一个文件）
$cfgWinPath = "$work\dc_config.ps1"
[System.IO.File]::WriteAllText($cfgWinPath, $cfg)

function Write-Feishu($rows) {
    $rows | ForEach-Object { [pscustomobject]@{ dept_id = $_[0]; dept_name = $_[1]; parent_dept_id = $_[2] } } |
        Export-Csv -Path (Join-Path $work 'feishu_departments.csv') -NoTypeInformation -Encoding UTF8
}
$feishu = @(@('od-c', '业务中心', '0'), @('od-s', '销售部', 'od-c'), @('od-y', '真实部门Y', '0'))

function Reset-AD {
    Reset-Store
    New-Obj $BaseOU 'organizationalUnit'
    New-Obj "OU=业务中心,$BaseOU" 'organizationalUnit' 'od-c'                      # 有效（ID）
    New-Obj "OU=销售部,OU=业务中心,$BaseOU" 'organizationalUnit' 'od-s-old'       # ID 不一致但位置与飞书一致 → 保留并修正ID
    New-Obj "OU=手工分组,$BaseOU" 'organizationalUnit'                             # 无ID，但下面有有效部门 → 保留
    New-Obj "OU=真实部门Y,OU=手工分组,$BaseOU" 'organizationalUnit' 'od-y'
    New-Obj "OU=旧部门,$BaseOU" 'organizationalUnit' 'od-old'                       # 残留，空 → 删除（含子）
    New-Obj "OU=旧子组,OU=旧部门,$BaseOU" 'organizationalUnit'
    New-Obj "OU=旧父,$BaseOU" 'organizationalUnit' 'od-gone'                        # 残留，子 OU 有人 → 两级都跳过
    New-Obj "OU=旧部门2,OU=旧父,$BaseOU" 'organizationalUnit'
    New-Obj "CN=服务账号,OU=旧部门2,OU=旧父,$BaseOU" 'user'
    New-Obj "OU=带GPO,$BaseOU" 'organizationalUnit' -gplink '[LDAP://cn={X}]'      # 残留，链接 GPO → 跳过
    New-Obj $Exclude 'organizationalUnit'                                          # 排除子树 → 不动
    New-Obj "OU=离职员工,$Exclude" 'organizationalUnit'
    New-Obj $Resigned 'organizationalUnit'                                         # 离职员工 OU（空、不在排除子树下）→ 不动
    New-Obj "OU=2025,$Resigned" 'organizationalUnit'
}
function Run-Script { param([switch]$DryRun, $Ratio)
    $a = @{}; if ($DryRun) { $a.DryRun = $true }; if ($null -ne $Ratio) { $a.MaxDeleteRatio = $Ratio }
    $out = Invoke-Logged { & $ScriptPath @a }
    return ($out -join "`n")
}

try {
    Section 'DryRun 只预览'
    Write-Feishu $feishu; Reset-AD; $before = Snapshot
    $txt = Run-Script -DryRun -Ratio 1
    Check 'DryRun 不改动 AD' ((Snapshot) -eq $before)
    Check 'DryRun 预览了子 OU 和父 OU 的删除' (($txt -match '将删除: OU=旧子组') -and ($txt -match '将删除: OU=旧部门,'))
    Check 'DryRun 预览了部门ID修正' ($txt -match "将修正部门ID: OU=销售部.*'od-s-old' -> 'od-s'")

    Section '正式清理'
    Write-Feishu $feishu; Reset-AD
    $txt = Run-Script -Ratio 1
    Check '有效 OU（ID）保留' (Has "OU=业务中心,$BaseOU")
    Check 'ID 不一致但位置与飞书一致的 OU 保留' (Has "OU=销售部,OU=业务中心,$BaseOU")
    Check '无ID但下面有有效部门的上级保留' (Has "OU=手工分组,$BaseOU")
    Check '有效部门保留' (Has "OU=真实部门Y,OU=手工分组,$BaseOU")
    Check '空残留子 OU 删除' (-not (Has "OU=旧子组,OU=旧部门,$BaseOU"))
    Check '空残留 OU 删除（先删子再删父，且解除了保护）' (-not (Has "OU=旧部门,$BaseOU"))
    Check '有用户的残留 OU 不删' (Has "OU=旧部门2,OU=旧父,$BaseOU")
    Check '用户原地不动' (Has "CN=服务账号,OU=旧部门2,OU=旧父,$BaseOU")
    Check '子 OU 删不掉时上级也不删' (Has "OU=旧父,$BaseOU")
    Check '链接 GPO 的残留 OU 不删' (Has "OU=带GPO,$BaseOU")
    Check '排除子树不动' ((Has $Exclude) -and (Has "OU=离职员工,$Exclude"))
    Check '离职员工 OU（不在排除子树下、且为空）不动' ((Has $Resigned) -and (Has "OU=2025,$Resigned"))
    Check '跳过原因列出了用户' ($txt -match '包含对象: \[user\] 服务账号')
    Check '跳过原因列出了 GPO' ($txt -match '链接了 GPO')
    Check '上级只记一行"子 OU 未能删除"' ($txt -match 'OU=旧父,.*子 OU 未能删除')

    Section '位置与飞书一致的 OU 修正部门ID'
    Write-Feishu $feishu; Reset-AD
    $null = Run-Script -Ratio 1
    Check 'Description 修正为飞书部门ID' ((Find-Obj "OU=销售部,OU=业务中心,$BaseOU").Description -eq 'od-s')
    Write-Feishu $feishu; Reset-AD; (Find-Obj "OU=业务中心,$BaseOU").Description = $null
    $null = Run-Script -Ratio 1
    Check '原无ID的 OU 补上飞书部门ID' ((Find-Obj "OU=业务中心,$BaseOU").Description -eq 'od-c')
    # 该ID已在另一个 OU 上（如该部门的正式 OU 搬迁失败仍在别处）→ 不修正，报告
    Write-Feishu $feishu; Reset-AD
    New-Obj "OU=销售部,OU=手工分组,$BaseOU" 'organizationalUnit' 'od-s'
    $txt = Run-Script -Ratio 1
    Check 'ID 已被占用时不修正' ((Find-Obj "OU=销售部,OU=业务中心,$BaseOU").Description -eq 'od-s-old')
    Check '占用方不受影响' ((Find-Obj "OU=销售部,OU=手工分组,$BaseOU").Description -eq 'od-s')
    Check '报告了占用方位置' ($txt -match '已在另一个 OU 上: OU=销售部,OU=手工分组')
    Check '仍按位置保留（不删除）' (Has "OU=销售部,OU=业务中心,$BaseOU")

    # 旧ID本身仍是飞书有效部门（该部门的本体）→ 不改写，报告
    Write-Feishu ($feishu + , @('od-x', '部门X', '0')); Reset-AD
    (Find-Obj "OU=销售部,OU=业务中心,$BaseOU").Description = 'od-x'
    $txt = Run-Script -Ratio 1
    Check '旧ID仍有效时不修正' ((Find-Obj "OU=销售部,OU=业务中心,$BaseOU").Description -eq 'od-x')
    Check '报告了旧ID所属的飞书部门' ($txt -match '仍是飞书部门「部门X」的有效ID')

    Section '熔断（飞书数据不完整）'
    Write-Feishu @(, @('od-c', '业务中心', '0'))
    Reset-AD
    New-Obj "OU=残1,$BaseOU" 'organizationalUnit'; New-Obj "OU=残2,$BaseOU" 'organizationalUnit'; New-Obj "OU=残3,$BaseOU" 'organizationalUnit'
    $before = Snapshot
    $txt = Run-Script
    Check '默认 20% 熔断：正式运行不删除任何 OU' ((Snapshot) -eq $before)
    Check '熔断时列出残留清单' ($txt -match '残留 OU 列表')
    $txt = Run-Script -DryRun
    Check '熔断时 DryRun 仍给出预览' ($txt -match '继续预览')

    Section '飞书部门列表为空'
    Set-Content -Path (Join-Path $work 'feishu_departments.csv') -Value '"dept_id","dept_name","parent_dept_id"' -Encoding UTF8
    Reset-AD; $before = Snapshot
    $txt = Run-Script -Ratio 1
    Check '飞书列表为空时不做任何改动' (((Snapshot) -eq $before) -and ($txt -match '飞书部门列表为空'))

    Section '无残留'
    Write-Feishu $feishu; Reset-AD
    foreach ($dn in @("OU=旧子组,OU=旧部门,$BaseOU", "OU=旧部门,$BaseOU", "CN=服务账号,OU=旧部门2,OU=旧父,$BaseOU",
                      "OU=旧部门2,OU=旧父,$BaseOU", "OU=旧父,$BaseOU", "OU=带GPO,$BaseOU")) { $global:Store.Remove($dn.ToLower()) }
    $txt = Run-Script
    Check '无残留时提示无需清理' ($txt -match '没有需要清理的 OU')
}
finally {
    $env:USERPROFILE = $savedProfile
    Remove-Item -LiteralPath $cfgWinPath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}

Complete-Tests
