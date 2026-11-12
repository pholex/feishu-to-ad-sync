# 内存版 Active Directory：用同名函数顶替 AD cmdlet（函数优先于 cmdlet），供回归测试使用。
# 只模拟本项目脚本用到的参数与行为：
#   - 开启"防止意外删除"的对象：Move-ADObject / Remove-ADOrganizationalUnit 报 Access is denied
#   - 目标位置同名：Move-ADObject 报 name already in use
#   - 非递归删除非叶子 OU：报错
#   - DN 按规范化形式匹配（不区分大小写与转义写法），Name 为去转义后的原始名称，
#     Rename-ADObject / New-ADOrganizationalUnit 自动转义名称——与真实 AD 一致
#   - 故障注入：$global:FailOn['move:<dn>'|'remove:<dn>'|'rename:<dn>'] = $true（dn 小写）时对应操作抛错
# 用法：在测试脚本中 . "$PSScriptRoot/MockAD.ps1"
#
# 注意：状态放在 $global: 而不是 $script:——被测脚本以 & 调用时，mock 函数里的 $script:
# 会解析到被测脚本自己的作用域。

. "$PSScriptRoot/../powershell/DnUtils.ps1"    # 规范化 DN（DnUtils 本身由 DnUtils.Tests.ps1 单独校验）

$global:Store = @{}       # 规范化 DN -> 对象
$global:FailOn = @{}
$global:TestPass = 0
$global:TestFail = 0

function Import-Module { }
function Write-Host { param($Object, $ForegroundColor) Write-Output "    | $Object" }

# ---- 构造与查询 ----
function Get-MockKey($dn) { Get-NormalDN ([string]$dn) }
function Get-RdnName($dn) { ConvertFrom-RdnValue ((Get-RDN $dn) -split '=', 2)[1] }
function New-Obj($dn, $cls, $desc = $null, $gplink = $null, $prot = $true) {
    $name = Get-RdnName $dn
    $global:Store[(Get-MockKey $dn)] = [pscustomobject]@{
        DistinguishedName = $dn; Name = $name; ObjectClass = $cls; Description = $desc
        gPLink = $gplink; gPOptions = 0; ObjectGUID = [guid]::NewGuid(); ProtectedFromAccidentalDeletion = $prot
    }
}
function Reset-Store {
    $global:Store = @{}
    $global:FailOn = @{}
    New-Obj 'DC=t' 'domainDNS' -prot $false
}
function Find-Obj($id) {
    $o = if ($id -is [guid]) { $global:Store.Values | Where-Object { $_.ObjectGUID -eq $id } } else { $global:Store[(Get-MockKey $id)] }
    if (-not $o) { throw "Directory object not found: $id" }
    return $o
}
function Has($dn) { $global:Store.ContainsKey((Get-MockKey $dn)) }
function Snapshot { ($global:Store.Keys | Sort-Object) -join '|' }

# ---- AD cmdlet mock ----
function Test-Fail($op, $dn) {
    if ($global:FailOn.ContainsKey("${op}:$(Get-MockKey $dn)")) { throw "Injected failure: $op $dn" }
}
# 修改对象 DN（连同子树）
function Set-ObjDN($o, $new) {
    $oldKey = Get-MockKey $o.DistinguishedName
    $oldDepth = (Split-DN $o.DistinguishedName).Count
    if ($global:Store.ContainsKey((Get-MockKey $new)) -and (Get-MockKey $new) -ne $oldKey) {
        throw "An attempt was made to add an object to the directory with a name that is already in use"
    }
    foreach ($k in @($global:Store.Keys)) {
        if ($k -eq $oldKey -or $k.EndsWith(",$oldKey")) {
            $x = $global:Store[$k]
            $parts = Split-DN $x.DistinguishedName
            $keepN = $parts.Count - $oldDepth          # 子对象在被移动对象之下的层数
            $x.DistinguishedName = (@($parts | Select-Object -First $keepN) + $new) -join ','
            $global:Store.Remove($k)
            $global:Store[(Get-MockKey $x.DistinguishedName)] = $x
        }
    }
    $o.Name = Get-RdnName $o.DistinguishedName
}

function Get-ADObject { param($Identity, $SearchBase, $SearchScope = 'Subtree', $Filter, $Properties, $ErrorAction)
    if ($Identity) { return (Find-Obj $Identity) }
    if ($SearchBase) {
        $null = Find-Obj $SearchBase
        $sb = Get-MockKey $SearchBase
        $r = foreach ($k in @($global:Store.Keys)) {
            $hit = if ($SearchScope -eq 'OneLevel') { (Get-ParentDN $k) -eq $sb } else { $k -eq $sb -or $k.EndsWith(",$sb") }
            if ($hit) { $global:Store[$k] }
        }
    } else { $r = $global:Store.Values }
    # 支持 -Filter { <属性> -eq $var }：从调用方作用域取变量值
    if ($Filter -is [scriptblock] -and "$Filter" -match '(\w+)\s+-eq\s+\$(\w+)') {
        $prop = $Matches[1]; $v = (Get-Variable $Matches[2] -Scope 1).Value
        $r = if ($prop -eq 'DistinguishedName') { $r | Where-Object { Test-SameDN $_.DistinguishedName $v } } else { $r | Where-Object { $_.$prop -eq $v } }
    }
    $r | Sort-Object { $_.DistinguishedName.Length }
}
function Get-ADOrganizationalUnit { param($Identity, $SearchBase, $SearchScope = 'Subtree', $Filter, $Properties, $ErrorAction)
    if ($Identity) {
        $o = Find-Obj $Identity
        if ($o.ObjectClass -ne 'organizationalUnit') { throw "Directory object not found: $Identity" }
        return $o
    }
    $r = Get-ADObject -SearchBase $SearchBase -SearchScope $SearchScope | Where-Object ObjectClass -eq 'organizationalUnit'
    # 支持 -Filter { <属性> -eq $var }：从调用方作用域取变量值
    if ($Filter -is [scriptblock] -and "$Filter" -match '(\w+)\s+-eq\s+\$(\w+)') {
        $prop = $Matches[1]; $v = (Get-Variable $Matches[2] -Scope 1).Value
        $r = $r | Where-Object { $_.$prop -eq $v }
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
    Test-Fail 'move' $o.DistinguishedName
    if ($o.ProtectedFromAccidentalDeletion) { throw "Access is denied" }
    $null = Find-Obj $TargetPath
    Set-ObjDN $o "$(Get-RDN $o.DistinguishedName),$TargetPath"
}
function Rename-ADObject { param($Identity, $NewName, $ErrorAction)
    $o = Find-Obj $Identity
    Test-Fail 'rename' $o.DistinguishedName
    $prefix = ((Get-RDN $o.DistinguishedName) -split '=', 2)[0]
    Set-ObjDN $o "$prefix=$(ConvertTo-RdnValue $NewName),$(Get-ParentDN $o.DistinguishedName)"
}
# 用户：按 SamAccountName 查找；New-User 创建带常用属性的用户对象
function New-User($dn, $sam, $props = @{}) {
    New-Obj $dn 'user' -prot $false
    $u = Find-Obj $dn
    $defaults = @{ SamAccountName = $sam; DisplayName = (Get-RdnName $dn); EmailAddress = ''; EmployeeID = ''; EmployeeNumber = ''; info = ''; Enabled = $true }
    foreach ($k in $defaults.Keys) { $u | Add-Member -NotePropertyName $k -NotePropertyValue $(if ($props.ContainsKey($k)) { $props[$k] } else { $defaults[$k] }) -Force }
}
function Get-ADUser { param($Identity, $Properties, $ErrorAction)
    $u = $global:Store.Values | Where-Object { $_.ObjectClass -eq 'user' -and $_.SamAccountName -eq $Identity } | Select-Object -First 1
    if (-not $u) { $u = Find-Obj $Identity }
    return $u
}
function Set-ADUser { param($Identity, $Enabled, $DisplayName, $EmailAddress, $EmployeeID, $EmployeeNumber, [hashtable]$Replace, $ErrorAction)
    $u = Find-Obj $Identity
    foreach ($k in @('Enabled', 'DisplayName', 'EmailAddress', 'EmployeeID', 'EmployeeNumber')) {
        if ($PSBoundParameters.ContainsKey($k)) { $u.$k = $PSBoundParameters[$k] }
    }
    if ($Replace) { foreach ($k in $Replace.Keys) { $u.$k = $Replace[$k] } }
}
function New-ADOrganizationalUnit { param($Name, $Path, $Description, $ErrorAction)
    $dn = "OU=$(ConvertTo-RdnValue $Name),$Path"
    $null = Find-Obj $Path
    if (Has $dn) { throw "An attempt was made to add an object to the directory with a name that is already in use" }
    New-Obj $dn 'organizationalUnit' $Description
}
function Remove-ADOrganizationalUnit { param($Identity, [switch]$Recursive, $Confirm, $ErrorAction)
    $o = Find-Obj $Identity
    Test-Fail 'remove' $o.DistinguishedName
    $sub = @(Get-ADObject -SearchBase $o.DistinguishedName)
    if (-not $Recursive -and $sub.Count -gt 1) { throw "The directory service can perform the requested operation only on a leaf object" }
    foreach ($x in $sub) { if ($x.ProtectedFromAccidentalDeletion) { throw "Access is denied (protected): $($x.DistinguishedName)" } }
    foreach ($x in $sub) { $global:Store.Remove((Get-MockKey $x.DistinguishedName)) }
}

# ---- 被测脚本运行环境 ----
# 被测脚本从 $env:USERPROFILE 读取 dc_config.ps1 和 CSV：建临时目录并指向它
function New-TestProfile($BaseOU, $ExcludeOU = '', $ResignedOU = '') {
    $work = Join-Path ([System.IO.Path]::GetTempPath()) ("feishu-ad-sync-tests-" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $work | Out-Null
    $global:SavedUserProfile = $env:USERPROFILE
    $env:USERPROFILE = $work
    $cfg = "`$DC_BASE_OU = `"$BaseOU`"`n`$DC_EXCLUDE_OU = `"$ExcludeOU`"`n`$DC_RESIGNED_OU = `"$ResignedOU`""
    Set-Content -Path (Join-Path $work 'dc_config.ps1') -Encoding UTF8 -Value $cfg
    # 脚本用 Windows 路径 "$env:USERPROFILE\dc_config.ps1" 交给 .NET ReadAllText：非 Windows 上反斜杠不是分隔符，
    # 需另写一份字面文件名带反斜杠的副本（Windows 上就是同一个文件）
    [System.IO.File]::WriteAllText("$work\dc_config.ps1", $cfg)
    # 被测脚本引用的 DN 工具（生产环境由 sync_to_ad.py 上传到 ~/DnUtils.ps1）
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot '../powershell/DnUtils.ps1') -Destination (Join-Path $work 'DnUtils.ps1')
    return $work
}
function Remove-TestProfile($work) {
    $env:USERPROFILE = $global:SavedUserProfile
    Remove-Item -LiteralPath "$work\dc_config.ps1" -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}
function Write-FeishuDepts($work, $rows) {
    $rows | ForEach-Object { [pscustomobject]@{ dept_id = $_[0]; dept_name = $_[1]; parent_dept_id = $_[2]; level = 1 } } |
        Export-Csv -Path (Join-Path $work 'feishu_departments.csv') -NoTypeInformation -Encoding UTF8
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
