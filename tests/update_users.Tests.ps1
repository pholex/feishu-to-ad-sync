# update_users.ps1 回归测试：按部门路径把用户移到对应 OU（部门名含 DN 特殊字符、新旧两种路径格式）
# 运行：pwsh -NoProfile -File tests/update_users.Tests.ps1   （或 tests/run.ps1 跑全部）
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot/MockAD.ps1"

$ScriptPath = (Resolve-Path (Join-Path $PSScriptRoot '../powershell/update_users.ps1')).Path
$BaseOU = 'OU=员工,DC=t'
$rd = "OU=研发\,测试组,$BaseOU"         # 部门"研发,测试组"
$sub = "OU=A\+B组,$rd"                 # 子部门"A+B组"
$bs = "OU=前端\\后端,$BaseOU"          # 部门"前端\后端"（名称含反斜杠）
$old = "OU=旧部门,$BaseOU"

function Reset-AD {
    Reset-Store
    New-Obj $BaseOU 'organizationalUnit'
    New-Obj $rd 'organizationalUnit' 'od-rd'
    New-Obj $sub 'organizationalUnit' 'od-ab'
    New-Obj $bs 'organizationalUnit' 'od-bs'
    New-Obj $old 'organizationalUnit'
    New-User "CN=张三,$old" 'san.zhang' @{ DisplayName = '张三' }
    New-User "CN=李四,$sub" 'si.li' @{ DisplayName = '李四' }
    New-User "CN=王五,$old" 'wu.wang' @{ DisplayName = '王五' }
    New-User "CN=赵六,$old" 'liu.zhao' @{ DisplayName = '赵六' }
}
function Write-Accounts($work, $rows) {
    $rows | Export-Csv -Path (Join-Path $work 'ad_check_accounts.csv') -NoTypeInformation -Encoding UTF8
}
function Row($sam, $name, $deptName, $deptPath) {
    [pscustomobject]@{ SamAccountName = $sam; DisplayName = $name; EmailAddress = ''; EmployeeID = ''; EmployeeNumber = ''
                       info = ''; DepartmentName = $deptName; DepartmentPath = $deptPath }
}

$work = New-TestProfile $BaseOU
try {
    Section '部门名含逗号/加号：移到正确 OU'
    Reset-AD
    Write-Accounts $work @(
        (Row 'san.zhang' '张三' '研发,测试组\A+B组' '["研发,测试组","A+B组"]'),
        (Row 'si.li' '李四' '研发,测试组\A+B组' '["研发,测试组","A+B组"]')
    )
    $txt = (Invoke-Logged { & $ScriptPath }) -join "`n"
    Check '用户移到含特殊字符的子部门 OU' (Has "CN=张三,$sub")
    Check '已在正确 OU 的用户不动、不计为更新' ((Has "CN=李四,$sub") -and ($txt -notmatch '李四'))
    Check '只更新了 1 个用户' ($txt -match '成功: 1 个用户')

    Section '部门名含反斜杠：DepartmentPath(JSON) 保留原名'
    Reset-AD
    Write-Accounts $work @((Row 'wu.wang' '王五' '前端\后端' '["前端\\后端"]'))
    $txt = (Invoke-Logged { & $ScriptPath }) -join "`n"
    Check '用户移到名称含反斜杠的部门 OU' (Has "CN=王五,$bs")
    Check '没有失败' ($txt -notmatch '更新失败')

    Section '兼容旧格式：只有 DepartmentName'
    Reset-AD
    [pscustomobject]@{ SamAccountName = 'liu.zhao'; DisplayName = '赵六'; EmailAddress = ''; EmployeeID = ''; EmployeeNumber = ''
                       info = ''; DepartmentName = '研发,测试组\A+B组' } |
        Export-Csv -Path (Join-Path $work 'ad_check_accounts.csv') -NoTypeInformation -Encoding UTF8
    $txt = (Invoke-Logged { & $ScriptPath }) -join "`n"
    Check '旧格式仍能移到正确 OU' (Has "CN=赵六,$sub")
}
finally {
    Remove-TestProfile $work
}

Complete-Tests
