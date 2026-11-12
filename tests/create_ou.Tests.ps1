# create_ou.ps1 回归测试：部门 OU 移动时与目标位置同名 OU 的合并（Clear-MoveTarget 及相关函数）
# 运行：pwsh -NoProfile -File tests/create_ou.Tests.ps1   （或 tests/run.ps1 跑全部）
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot/MockAD.ps1"

# 只加载 create_ou.ps1 中的函数定义，不执行脚本主体
$scriptPath = Join-Path $PSScriptRoot '../powershell/create_ou.ps1'
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path $scriptPath), [ref]$null, [ref]$null)
$ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false) |
    Where-Object { $_.Parent.Parent -eq $ast.EndBlock -or $_.Parent -eq $ast.EndBlock } |
    ForEach-Object { . ([scriptblock]::Create($_.Extent.Text)) }

$BaseOU = 'OU=员工,DC=t'
$src = "OU=服务部,OU=业务中心,$BaseOU"     # 源：真正的 服务部（带部门ID），在旧位置
$tgt = "OU=服务部,OU=共享中心,$BaseOU"     # 目标位置已有的同名 OU
$mid = "OU=共享中心,$BaseOU"

function Reset-Base {
    # 飞书部门表（create_ou.ps1 主流程加载 CSV 后填充的 $script:FeishuDepts）
    $script:FeishuDepts = @{}
    foreach ($r in @(@('od-biz', '业务中心', '0'), @('od-shared', '共享中心', '0'), @('od-svc', '服务部', 'od-shared'),
                    @('od-cs1', '客服一组', 'od-svc'), @('od-cs2', '客服二组', 'od-svc'), @('od-a', 'A组', 'od-cs1'),
                    @('od-cs3', '别的组', 'od-svc'))) {
        $script:FeishuDepts[$r[0]] = [pscustomobject]@{ dept_id = $r[0]; dept_name = $r[1]; parent_dept_id = $r[2] }
    }
    Reset-Store
    New-Obj $BaseOU 'organizationalUnit'
    New-Obj "OU=业务中心,$BaseOU" 'organizationalUnit' 'od-biz'
    New-Obj "OU=共享中心,$BaseOU" 'organizationalUnit' 'od-shared'
    New-Obj $src 'organizationalUnit' 'od-svc'
    New-Obj "OU=客服一组,$src" 'organizationalUnit' 'od-cs1'
    New-Obj "CN=老员工,OU=客服一组,$src" 'user'
    New-Obj "OU=客服二组,$src" 'organizationalUnit' 'od-cs2'
    # 目标：人工预建的同名结构
    New-Obj $tgt 'organizationalUnit'
    New-Obj "CN=新员工A,$tgt" 'user'
    New-Obj "OU=客服一组,$tgt" 'organizationalUnit' 'od-cs1'
    New-Obj "CN=新员工B,OU=客服一组,$tgt" 'user'
    New-Obj "OU=空壳组,$tgt" 'organizationalUnit'
    New-Obj "OU=临时组,$tgt" 'organizationalUnit'
    New-Obj "CN=新员工C,OU=临时组,$tgt" 'user'
}

# 预期拒绝合并：报错信息匹配 $pattern，且预检阶段未改动 AD
function Expect-Refuse($label, [scriptblock]$setup, $pattern) {
    Reset-Base; & $setup
    $before = Snapshot
    try {
        $null = Invoke-Logged { Clear-MoveTarget -TargetDN $tgt -SourceDN $src }
        Check "$label → 应拒绝合并" $false
    } catch {
        $msg = $_.Exception.Message
        Check "$label → 拒绝并列出原因" ($msg -match $pattern)
        if ($msg -notmatch $pattern) { "      实际: $msg" }
        Check "$label → 预检阶段未改动 AD" ((Snapshot) -eq $before)
    }
}

Section 'DryRun 不改动任何东西'
Reset-Base; $before = Snapshot
$log = Invoke-Logged { Clear-MoveTarget -TargetDN $tgt -SourceDN $src -DryRun }
Check 'DryRun 后 AD 不变' ((Snapshot) -eq $before)
Check 'DryRun 列出将移入的用户' (($log -join "`n") -match '将移入 \[user\] 新员工A')

Section '正常合并 + 移动（仿 create_ou 主流程）'
Reset-Base
$null = Invoke-Logged { Clear-MoveTarget -TargetDN $tgt -SourceDN $src; Move-ObjectSafely -DN $src -TargetPath $mid }
$crm = Find-Obj $tgt
Check '源 OU 本体保留（Description=od-svc）' ($crm.Description -eq 'od-svc')
Check '源 OU 移动后保护已恢复' ($crm.ProtectedFromAccidentalDeletion)
Check '目标里的用户并入 服务部' (Has "CN=新员工A,$tgt")
Check '同名子 OU 递归合并' (Has "CN=新员工B,OU=客服一组,$tgt")
Check '源原有用户仍在' (Has "CN=老员工,OU=客服一组,$tgt")
Check '同ID的 OU 只剩一个' (@($global:Store.Values | Where-Object Description -eq 'od-cs1').Count -eq 1)
Check '空壳子 OU 被删除' (-not (Has "OU=空壳组,$tgt"))
Check '源里没有的子 OU 整体移入（含用户）' (Has "CN=新员工C,OU=临时组,$tgt")
Check '旧位置已无 服务部' (-not (Has $src))
Check '用户移入后保护已恢复' ((Find-Obj "CN=新员工A,$tgt").ProtectedFromAccidentalDeletion)

Section '纯空壳目标直接删除'
Reset-Base
foreach ($k in @($global:Store.Keys)) { if ($k.EndsWith(",$($tgt.ToLower())")) { $global:Store.Remove($k) } }
$null = Invoke-Logged { Clear-MoveTarget -TargetDN $tgt -SourceDN $src }
Check '空壳目标被删除' (-not (Has $tgt))

Section '停止条件'
Expect-Refuse '目标 OU 链接了 GPO' { (Find-Obj $tgt).gPLink = '[LDAP://cn={X},cn=policies;0]' } 'GPO'
Expect-Refuse '子 OU 设置阻止继承' { (Find-Obj "OU=客服一组,$tgt").gPOptions = 1 } '阻止继承'
Expect-Refuse '有计算机对象' { New-Obj "CN=PC01,OU=临时组,$tgt" 'computer' } 'computer'
Expect-Refuse '用户同名冲突' { New-Obj "CN=新员工A,$src" 'user' } '同名对象'
Expect-Refuse '整体移入的 OU 部门ID（飞书中存在）在别处已有' { (Find-Obj "OU=临时组,$tgt").Description = 'od-cs2' } '部门ID重复'

Section '部门ID冲突按飞书判定'
Expect-Refuse '两个ID在飞书中都存在' { (Find-Obj "OU=客服一组,$tgt").Description = 'od-cs2' } '都存在'
Expect-Refuse '目标ID有效但飞书中属于别处' {
    (Find-Obj "OU=客服一组,$src").Description = 'od-old'
    (Find-Obj "OU=客服一组,$tgt").Description = 'od-cs3' } '位置不符'

# keepSource：目标ID已不在飞书 → 合并，保留源ID
Reset-Base; (Find-Obj "OU=客服一组,$tgt").Description = 'od-gone'
$null = Invoke-Logged { Clear-MoveTarget -TargetDN $tgt -SourceDN $src }
Check 'keepSource：合并成功，源 OU 仍为 od-cs1' ((Find-Obj "OU=客服一组,$src").Description -eq 'od-cs1')
Check 'keepSource：目标用户并入' (Has "CN=新员工B,OU=客服一组,$src")
Check 'keepSource：残留ID已不存在' (@($global:Store.Values | Where-Object Description -eq 'od-gone').Count -eq 0)

# adoptTarget：源ID已不在飞书、目标ID有效 → 合并，源 OU 改用目标ID；嵌套时子 OU 的上级ID按合并后的ID校验
Reset-Base
(Find-Obj "OU=客服一组,$src").Description = 'od-old'
New-Obj "OU=A组,OU=客服一组,$src" 'organizationalUnit' 'od-a-old'
New-Obj "OU=A组,OU=客服一组,$tgt" 'organizationalUnit' 'od-a'
New-Obj "CN=新员工D,OU=A组,OU=客服一组,$tgt" 'user'
$usGuid = (Find-Obj "OU=客服一组,$src").ObjectGUID
$probs = @(Get-MergeProblems -TargetDN $tgt -SourceDN $src)
Check 'adoptTarget：预检无问题（嵌套上级ID按合并后的ID校验）' ($probs.Count -eq 0)
if ($probs) { $probs | ForEach-Object { "      ! $_" } }
$log = Invoke-Logged { Clear-MoveTarget -TargetDN $tgt -SourceDN $src }
$us = Find-Obj "OU=客服一组,$src"
Check 'adoptTarget：源 OU Description 改为 od-cs1' ($us.Description -eq 'od-cs1')
Check 'adoptTarget：保留的是源 OU 对象（GUID 不变）' ($us.ObjectGUID -eq $usGuid)
Check 'adoptTarget：嵌套子 OU 改为 od-a' ((Find-Obj "OU=A组,OU=客服一组,$src").Description -eq 'od-a')
Check 'adoptTarget：嵌套用户并入' (Has "CN=新员工D,OU=A组,OU=客服一组,$src")
Check 'adoptTarget：每个ID只剩一个 OU' ((@($global:Store.Values | Where-Object Description -eq 'od-cs1').Count -eq 1) -and
                                     (@($global:Store.Values | Where-Object Description -eq 'od-a').Count -eq 1))
Check 'adoptTarget：日志记录了旧ID -> 新ID' (($log -join "`n") -match "'od-old' -> 'od-cs1'")

# 整体移入的 OU 带残留ID（飞书中不存在）且别处重复 → 允许
Reset-Base
(Find-Obj "OU=临时组,$tgt").Description = 'od-stale'
New-Obj "OU=旧组,OU=业务中心,$BaseOU" 'organizationalUnit' 'od-stale'
Check '残留ID重复不算问题' (@(Get-MergeProblems -TargetDN $tgt -SourceDN $src).Count -eq 0)

Section 'Get-RDN'
Check '普通 DN' ((Get-RDN 'OU=A,OU=B,DC=t') -eq 'OU=A')
Check '转义逗号' ((Get-RDN 'CN=Zhang\, San,OU=B,DC=t') -eq 'CN=Zhang\, San')

Complete-Tests
