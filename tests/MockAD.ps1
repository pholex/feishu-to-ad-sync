# 内存版 Active Directory：用同名函数顶替 AD cmdlet（函数优先于 cmdlet），供回归测试使用。
# 只模拟本项目脚本用到的参数与行为：
#   - 开启"防止意外删除"的对象：Move-ADObject / Remove-ADOrganizationalUnit 报 Access is denied
#   - 目标位置同名：Move-ADObject 报 name already in use
#   - 非递归删除非叶子 OU：报错
# 用法：在测试脚本中 . "$PSScriptRoot/MockAD.ps1"
#
# 注意：状态放在 $global: 而不是 $script:——被测脚本以 & 调用时，mock 函数里的 $script:
# 会解析到被测脚本自己的作用域。

$global:Store = @{}
$global:TestPass = 0
$global:TestFail = 0
$global:DnSplitter = [regex]::new('(?<!\\),')

function Import-Module { }
function Write-Host { param($Object, $ForegroundColor) Write-Output "    | $Object" }

# ---- 构造与查询 ----
function New-Obj($dn, $cls, $desc = $null, $gplink = $null, $prot = $true) {
    $name = ($global:DnSplitter.Split($dn, 2)[0] -split '=', 2)[1]
    $global:Store[$dn.ToLower()] = [pscustomobject]@{
        DistinguishedName = $dn; Name = $name; ObjectClass = $cls; Description = $desc
        gPLink = $gplink; gPOptions = 0; ObjectGUID = [guid]::NewGuid(); ProtectedFromAccidentalDeletion = $prot
    }
}
function Reset-Store {
    $global:Store = @{}
    New-Obj 'DC=t' 'domainDNS' -prot $false
}
function Find-Obj($id) {
    $o = if ($id -is [guid]) { $global:Store.Values | Where-Object { $_.ObjectGUID -eq $id } } else { $global:Store[([string]$id).ToLower()] }
    if (-not $o) { throw "Directory object not found: $id" }
    return $o
}
function Has($dn) { $global:Store.ContainsKey($dn.ToLower()) }
function Snapshot { ($global:Store.Keys | Sort-Object) -join '|' }

# ---- AD cmdlet mock ----
function Get-ADObject { param($Identity, $SearchBase, $SearchScope = 'Subtree', $Filter, $Properties, $ErrorAction)
    if ($Identity) { return (Find-Obj $Identity) }
    $null = Find-Obj $SearchBase
    $sb = $SearchBase.ToLower()
    $global:Store.Values | Where-Object {
        $dn = $_.DistinguishedName.ToLower()
        if ($SearchScope -eq 'OneLevel') { $global:DnSplitter.Split($dn, 2)[1] -eq $sb } else { $dn -eq $sb -or $dn.EndsWith(",$sb") }
    } | Sort-Object { $_.DistinguishedName.Length }
}
function Get-ADOrganizationalUnit { param($Identity, $SearchBase, $SearchScope = 'Subtree', $Filter, $Properties, $ErrorAction)
    if ($Identity) {
        $o = Find-Obj $Identity
        if ($o.ObjectClass -ne 'organizationalUnit') { throw "Directory object not found: $Identity" }
        return $o
    }
    $r = Get-ADObject -SearchBase $SearchBase -SearchScope $SearchScope | Where-Object ObjectClass -eq 'organizationalUnit'
    # 支持 -Filter { Description -eq $var }：从调用方作用域取变量值
    if ($Filter -is [scriptblock] -and "$Filter" -match 'Description -eq \$(\w+)') {
        $v = (Get-Variable $Matches[1] -Scope 1).Value
        $r = $r | Where-Object { $_.Description -eq $v }
    }
    $r
}
function Set-ADObject { param($Identity, $ProtectedFromAccidentalDeletion, $ErrorAction)
    (Find-Obj $Identity).ProtectedFromAccidentalDeletion = $ProtectedFromAccidentalDeletion }
function Set-ADOrganizationalUnit { param($Identity, $ProtectedFromAccidentalDeletion, $Description, $ErrorAction)
    $o = Find-Obj $Identity
    if ($PSBoundParameters.ContainsKey('ProtectedFromAccidentalDeletion')) { $o.ProtectedFromAccidentalDeletion = $ProtectedFromAccidentalDeletion }
    if ($PSBoundParameters.ContainsKey('Description')) { $o.Description = $Description } }
function Move-ADObject { param($Identity, $TargetPath, $ErrorAction)
    $o = Find-Obj $Identity
    if ($o.ProtectedFromAccidentalDeletion) { throw "Access is denied" }
    $null = Find-Obj $TargetPath
    $old = $o.DistinguishedName
    $new = "$($global:DnSplitter.Split($old, 2)[0]),$TargetPath"
    if ($global:Store.ContainsKey($new.ToLower())) { throw "An attempt was made to add an object to the directory with a name that is already in use" }
    foreach ($x in @($global:Store.Values)) {
        $d = $x.DistinguishedName
        if ($d -eq $old -or $d.ToLower().EndsWith(",$($old.ToLower())")) {
            $global:Store.Remove($d.ToLower())
            $x.DistinguishedName = $d.Substring(0, $d.Length - $old.Length) + $new
            $global:Store[$x.DistinguishedName.ToLower()] = $x
        }
    }
}
function Remove-ADOrganizationalUnit { param($Identity, [switch]$Recursive, $Confirm, $ErrorAction)
    $o = Find-Obj $Identity
    $sub = @(Get-ADObject -SearchBase $o.DistinguishedName)
    if (-not $Recursive -and $sub.Count -gt 1) { throw "The directory service can perform the requested operation only on a leaf object" }
    foreach ($x in $sub) { if ($x.ProtectedFromAccidentalDeletion) { throw "Access is denied (protected): $($x.DistinguishedName)" } }
    foreach ($x in $sub) { $global:Store.Remove($x.DistinguishedName.ToLower()) }
}

# ---- 断言与输出 ----
function Check($name, $cond) {
    if ($cond) { "  ✅ $name"; $global:TestPass++ } else { "  ❌ $name"; $global:TestFail++ }
}
function Section($title) { "=== $title" }

# 执行被测代码并返回其日志输出；设置环境变量 TEST_VERBOSE=1 时同时打印日志
function Invoke-Logged([scriptblock]$Script) {
    $o = @(& $Script)
    if ($env:TEST_VERBOSE) { $o | ForEach-Object { Microsoft.PowerShell.Utility\Write-Host $_ } }
    return $o
}

function Complete-Tests {
    "`n结果: $($global:TestPass) 通过, $($global:TestFail) 失败"
    if ($global:TestFail -gt 0) { exit 1 }
    exit 0
}
