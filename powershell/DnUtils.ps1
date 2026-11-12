# DN（Distinguished Name）转义、拆分与比较工具，各脚本共用。
# sync_to_ad.py 在上传配置时把本文件一并上传到域控 ~/DnUtils.ps1，各脚本以 . "$env:USERPROFILE\DnUtils.ps1" 引用。
#
# 为什么需要：飞书部门名可能含 DN 特殊字符（, + " \ < > ; = 以及首尾空格、开头的 #），
# 直接 "OU=$名称,..." 拼接会得到错误或指向别处的 DN；按 ',' 直接拆分 DN 也会把转义的逗号拆开。
# 比较 DN 时先规范化（去转义 → 按统一规则重新转义 → 小写），不依赖 AD 返回的具体转义写法。

# 名称 → RDN 值（转义特殊字符）。用于自己拼 DN；New-ADOrganizationalUnit -Name / Rename-ADObject -NewName 传原始名称即可，不要转义
function ConvertTo-RdnValue {
    param([string]$Value)
    if ($Value -eq ' ') { return '\ ' }
    $v = $Value -replace '([,+"\\<>;=])', '\$1'
    if ($Value.StartsWith(' ') -or $Value.StartsWith('#')) { $v = '\' + $v }
    if ($Value.Length -gt 1 -and $Value.EndsWith(' ')) { $v = $v.Substring(0, $v.Length - 1) + '\ ' }
    return $v
}

# RDN 值 → 原始名称（去转义，支持 \c 与 \hh 十六进制 UTF-8 字节）
function ConvertFrom-RdnValue {
    param([string]$Value)
    $out = New-Object System.Text.StringBuilder
    $bytes = New-Object 'System.Collections.Generic.List[byte]'
    $i = 0
    while ($i -lt $Value.Length) {
        $ch = $Value[$i]
        if ($ch -eq [char]'\' -and $i + 3 -le $Value.Length -and $Value.Substring($i + 1, 2) -match '^[0-9A-Fa-f]{2}$') {
            $bytes.Add([Convert]::ToByte($Value.Substring($i + 1, 2), 16)); $i += 3; continue
        }
        if ($bytes.Count -gt 0) { [void]$out.Append([System.Text.Encoding]::UTF8.GetString($bytes.ToArray())); $bytes.Clear() }
        if ($ch -eq [char]'\' -and $i + 1 -lt $Value.Length) { [void]$out.Append($Value[$i + 1]); $i += 2; continue }
        [void]$out.Append($ch); $i++
    }
    if ($bytes.Count -gt 0) { [void]$out.Append([System.Text.Encoding]::UTF8.GetString($bytes.ToArray())) }
    return $out.ToString()
}

# 按未转义的逗号拆分 DN，各段保持原样（含转义）
function Split-DN {
    param([string]$DN)
    $parts = New-Object 'System.Collections.Generic.List[string]'
    if (-not $DN) { return , $parts.ToArray() }
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $DN.Length; $i++) {
        $ch = $DN[$i]
        if ($ch -eq [char]'\' -and $i + 1 -lt $DN.Length) { [void]$sb.Append($ch).Append($DN[$i + 1]); $i++; continue }
        if ($ch -eq [char]',') { $parts.Add($sb.ToString()); [void]$sb.Clear(); continue }
        [void]$sb.Append($ch)
    }
    $parts.Add($sb.ToString())
    return , $parts.ToArray()
}

# DN 的第一段 RDN（如 "OU=研发\,测试组"）
function Get-RDN {
    param($DN)
    return (Split-DN $DN)[0]
}

# DN 的上级 DN；没有上级时返回空串
function Get-ParentDN {
    param($DN)
    $p = Split-DN $DN
    if ($p.Count -le 1) { return '' }
    return ($p[1..($p.Count - 1)] -join ',')
}

# 规范化 DN，用于比较或作哈希键
function Get-NormalDN {
    param([string]$DN)
    if (-not $DN) { return '' }
    $rdns = foreach ($rdn in (Split-DN $DN)) {
        $eq = $rdn.IndexOf('=')
        if ($eq -lt 0) { $rdn.Trim().ToLower(); continue }
        $type = $rdn.Substring(0, $eq).Trim().ToLower()
        $val = ConvertTo-RdnValue (ConvertFrom-RdnValue $rdn.Substring($eq + 1))
        "$type=$($val.ToLower())"
    }
    return ($rdns -join ',')
}

function Test-SameDN {
    param($A, $B)
    return (Get-NormalDN $A) -eq (Get-NormalDN $B)
}

# DN 是否等于 Root 或位于 Root 子树中
function Test-UnderDN {
    param($DN, $Root)
    if (-not $Root) { return $false }
    $d = Get-NormalDN $DN; $r = Get-NormalDN $Root
    return ($d -eq $r) -or $d.EndsWith(",$r")
}

# 由部门名路径（上级 → 下级）和基础 OU 拼出 OU 的 DN
function Join-OUPath {
    param([string[]]$Names, $BaseOU)
    $dn = $BaseOU
    foreach ($n in $Names) { if ($n -ne '') { $dn = "OU=$(ConvertTo-RdnValue $n),$dn" } }
    return $dn
}

# 用户 CSV 中的部门路径（上级 → 下级）：优先 DepartmentPath（JSON 数组，部门名可含任意字符），
# 兼容旧格式 DepartmentName（以 \ 分隔）。始终返回数组（单层部门也不会被展开成字符串）；
# 调用方直接赋值或传给 [string[]] 参数，不要再套 @()
function Get-UserDeptPath {
    param($Row)
    $json = [string]$Row.DepartmentPath
    if ($json) { return , @($json | ConvertFrom-Json) }
    $name = [string]$Row.DepartmentName
    if ($name) { return , @($name -split '\\') }
    return , @()
}
