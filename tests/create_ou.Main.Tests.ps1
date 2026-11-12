# create_ou.ps1 主流程回归测试：整体运行脚本，覆盖部门改名/移动与目标位置同名 OU、部门ID重复
# 运行：pwsh -NoProfile -File tests/create_ou.Main.Tests.ps1   （或 tests/run.ps1 跑全部）
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot/MockAD.ps1"

$ScriptPath = (Resolve-Path (Join-Path $PSScriptRoot '../powershell/create_ou.ps1')).Path
$BaseOU = 'OU=员工,DC=t'
$hq = "OU=总部,$BaseOU"
$sh = "OU=共享中心,$BaseOU"
$old = "OU=客服,$hq"             # 部门 od-svc 在 AD 中的当前位置
$final = "OU=服务部,$sh"         # 飞书中：改名为"服务部"并移到"共享中心"下

function Reset-AD {
    Reset-Store
    New-Obj $BaseOU 'organizationalUnit'
    New-Obj $hq 'organizationalUnit' 'od-hq'
    New-Obj $sh 'organizationalUnit' 'od-sh'
    New-Obj $old 'organizationalUnit' 'od-svc'
    New-Obj "CN=员工甲,$old" 'user'
    New-Obj "OU=一组,$old" 'organizationalUnit' 'od-g1'
    New-Obj "CN=员工乙,OU=一组,$old" 'user'
}
# 飞书部门表：服务部(od-svc) 改名并移到共享中心下
$moveRename = @(@('od-hq', '总部', '0'), @('od-sh', '共享中心', '0'), @('od-svc', '服务部', 'od-sh'), @('od-g1', '一组', 'od-svc'))
# 只改名：服务部仍在总部下
$renameOnly = @(@('od-hq', '总部', '0'), @('od-sh', '共享中心', '0'), @('od-svc', '服务部', 'od-hq'), @('od-g1', '一组', 'od-svc'))

function Run-Script { param([switch]$DryRun)
    $a = @{}; if ($DryRun) { $a.DryRun = $true }
    return ((Invoke-Logged { & $ScriptPath @a }) -join "`n")
}
function Check-Arrived($label) {
    $o = Find-Obj $final
    Check "$label：部门到达最终位置且ID不变" ($o.Description -eq 'od-svc')
    Check "$label：原有用户随部门移动" (Has "CN=员工甲,$final")
    Check "$label：子部门随部门移动" ((Has "CN=员工乙,OU=一组,$final") -and ((Find-Obj "OU=一组,$final").Description -eq 'od-g1'))
    Check "$label：旧位置已无此部门" (-not (Has $old))
}

$work = New-TestProfile $BaseOU
try {
    Section '改名+移动：基本情况'
    Write-FeishuDepts $work $moveRename; Reset-AD
    $txt = Run-Script
    Check-Arrived '基本'
    Check '基本：子部门无需再处理' ($txt -notmatch '处理失败')

    Section '改名+移动：最终位置已有同名空壳'
    Write-FeishuDepts $work $moveRename; Reset-AD
    New-Obj $final 'organizationalUnit'
    $txt = Run-Script
    Check-Arrived '空壳'
    Check '空壳：日志记录删除空壳' ($txt -match '已删除目标位置的空壳OU.*OU=服务部,OU=共享中心')

    Section '改名+移动：最终位置已有同名 OU（含用户）→ 合并'
    Write-FeishuDepts $work $moveRename; Reset-AD
    New-Obj $final 'organizationalUnit'
    New-Obj "CN=新员工,$final" 'user'
    $txt = Run-Script
    Check-Arrived '合并'
    Check '合并：目标里的用户并入部门' (Has "CN=新员工,$final")
    Check '合并：只剩一个 od-svc' (@($global:Store.Values | Where-Object Description -eq 'od-svc').Count -eq 1)

    Section '改名+移动：先移动的中间位置（旧名,新父）被无关 OU 占用 → 改为先改名'
    Write-FeishuDepts $work $moveRename; Reset-AD
    New-Obj "OU=客服,$sh" 'organizationalUnit'
    New-Obj "CN=别人,OU=客服,$sh" 'user'
    $txt = Run-Script
    Check-Arrived '先改名'
    Check '先改名：中间位置上的无关 OU 未被合并' ((Has "OU=客服,$sh") -and (Has "CN=别人,OU=客服,$sh"))

    Section '改名+移动：两个中间位置都被占用 → 不动并报错'
    Write-FeishuDepts $work $moveRename; Reset-AD
    New-Obj "OU=客服,$sh" 'organizationalUnit'
    New-Obj "CN=别人,OU=客服,$sh" 'user'
    New-Obj "OU=服务部,$hq" 'organizationalUnit'
    New-Obj "CN=别人2,OU=服务部,$hq" 'user'
    $before = Snapshot
    $txt = Run-Script
    Check '都被占用：AD 不变' ((Snapshot) -eq $before)
    Check '都被占用：报错说明中间位置' ($txt -match '两个中间位置都已有同名OU')
    Check '都被占用：子部门跳过' ($txt -match '跳过 - 一组 : 上级部门处理失败')

    Section '只改名：新名字位置已有同名空壳'
    Write-FeishuDepts $work $renameOnly; Reset-AD
    New-Obj "OU=服务部,$hq" 'organizationalUnit'
    $txt = Run-Script
    $o = Find-Obj "OU=服务部,$hq"
    Check '只改名：改名成功且ID不变' ($o.Description -eq 'od-svc')
    Check '只改名：用户与子部门随之' ((Has "CN=员工甲,OU=服务部,$hq") -and (Has "OU=一组,OU=服务部,$hq"))
    Check '只改名：旧名已不存在' (-not (Has $old))

    Section 'DryRun：改名+移动遇到无法合并的同名 OU 时预览即报错'
    Write-FeishuDepts $work $moveRename; Reset-AD
    New-Obj $final 'organizationalUnit'
    New-Obj "CN=PC01,$final" 'computer'
    $before = Snapshot
    $txt = Run-Script -DryRun
    Check 'DryRun：AD 不变' ((Snapshot) -eq $before)
    Check 'DryRun：预览中给出无法合并的原因' ($txt -match '无法自动合并' -and $txt -match 'computer')

    Section '部门名含 DN 特殊字符：新建'
    $special = $moveRename + @(, @('od-rd', '研发,测试组', 'od-sh')) + @(, @('od-ab', 'A+B "组"', 'od-rd')) + @(, @('od-sp', ' 首尾空格 ', 'od-sh'))
    Write-FeishuDepts $work $special; Reset-AD
    $txt = Run-Script
    $rd = "OU=研发\,测试组,$sh"
    $ab = "OU=A\+B \`"组\`",$rd"
    Check '特殊字符：含逗号的部门 OU 已创建且ID正确' ((Has $rd) -and ((Find-Obj $rd).Description -eq 'od-rd'))
    Check '特殊字符：OU 名称为原始名称（未带转义符）' ((Find-Obj $rd).Name -ceq '研发,测试组')
    Check '特殊字符：子部门创建在正确的上级下' ((Has $ab) -and ((Find-Obj $ab).Description -eq 'od-ab'))
    Check '特殊字符：首尾空格的部门名' ((Has "OU=\ 首尾空格\ ,$sh") -and ((Find-Obj "OU=\ 首尾空格\ ,$sh").Name -ceq ' 首尾空格 '))
    Check '特殊字符：没有在错误位置多建 OU' (-not (Has "OU=研发,$sh") -and -not (Has "OU=A,$sh"))
    Check '特殊字符：无失败' ($txt -notmatch '处理失败|创建失败|跳过 -')

    Section '部门名含 DN 特殊字符：已就位时无需变更'
    $before = Snapshot
    $txt = Run-Script
    Check '特殊字符：再次运行 AD 不变' ((Snapshot) -eq $before)
    Check '特殊字符：再次运行不移动/不重命名/不新建' ($txt -notmatch '已移动|已重命名|成功创建|处理失败')

    Section '部门名含 DN 特殊字符：改名为含逗号的名称并移动'
    $renameSpecial = @(@('od-hq', '总部', '0'), @('od-sh', '共享中心', '0'), @('od-svc', '服务,支持', 'od-sh'), @('od-g1', '一组', 'od-svc'))
    Write-FeishuDepts $work $renameSpecial; Reset-AD
    $txt = Run-Script
    $fs = "OU=服务\,支持,$sh"
    Check '特殊字符改名：部门到达最终位置且ID不变' ((Has $fs) -and ((Find-Obj $fs).Description -eq 'od-svc'))
    Check '特殊字符改名：用户与子部门随之' ((Has "CN=员工甲,$fs") -and (Has "CN=员工乙,OU=一组,$fs"))
    Check '特殊字符改名：子部门无需再处理' ($txt -notmatch '处理失败')

    Section '部门ID重复'
    Write-FeishuDepts $work $moveRename; Reset-AD
    New-Obj "OU=客服副本,$hq" 'organizationalUnit' 'od-svc'
    $before = Snapshot
    $txt = Run-Script
    Check '重复ID：AD 不变' ((Snapshot) -eq $before)
    Check '重复ID：报错并列出两个 OU' (($txt -match "同时出现在 2 个 OU 上") -and ($txt -match 'OU=客服副本'))
    Check '重复ID：子部门跳过' ($txt -match '跳过 - 一组 : 上级部门处理失败')
}
finally {
    Remove-TestProfile $work
}

Complete-Tests
