# DnUtils.ps1 单元测试：DN 转义、拆分、规范化比较。期望值写死，不依赖被测函数自身
# 运行：pwsh -NoProfile -File tests/DnUtils.Tests.ps1   （或 tests/run.ps1 跑全部）
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot/MockAD.ps1"    # 只用其中的 Check / Section / Complete-Tests

Section 'ConvertTo-RdnValue：转义'
$cases = [ordered]@{
    '普通部门'    = '普通部门'
    '研发,测试组' = '研发\,测试组'
    'a+b'         = 'a\+b'
    'x"y'         = 'x\"y'
    'back\slash'  = 'back\\slash'
    '<lt>'        = '\<lt\>'
    'semi;colon'  = 'semi\;colon'
    'eq=val'      = 'eq\=val'
    ' lead'       = '\ lead'
    'trail '      = 'trail\ '
    '#hash'       = '\#hash'
    'mid#dle'     = 'mid#dle'
    ' '           = '\ '
}
foreach ($k in $cases.Keys) {
    $got = ConvertTo-RdnValue $k
    Check "转义 [$k] -> [$($cases[$k])]" ($got -ceq $cases[$k])
    Check "去转义还原 [$k]" ((ConvertFrom-RdnValue $got) -ceq $k)
}

Section 'ConvertFrom-RdnValue：十六进制转义'
Check '\2C -> 逗号' ((ConvertFrom-RdnValue 'a\2Cb') -ceq 'a,b')
Check 'UTF-8 多字节十六进制' ((ConvertFrom-RdnValue '\E7\A0\94\E5\8F\91') -ceq '研发')
Check '转义的反斜杠后跟十六进制字符不误判' ((ConvertFrom-RdnValue 'a\\2C') -ceq 'a\2C')

Section 'Split-DN / Get-RDN / Get-ParentDN'
$dn = 'OU=研发\,测试组,OU=a\\b,OU=员工,DC=t'
$parts = Split-DN $dn
Check '按未转义逗号拆分' (($parts.Count -eq 4) -and ($parts[0] -ceq 'OU=研发\,测试组') -and ($parts[1] -ceq 'OU=a\\b'))
Check 'Get-RDN' ((Get-RDN $dn) -ceq 'OU=研发\,测试组')
Check 'Get-ParentDN' ((Get-ParentDN $dn) -ceq 'OU=a\\b,OU=员工,DC=t')
Check '转义的反斜杠后的逗号是分隔符' ((Split-DN 'OU=a\\,OU=b').Count -eq 2)
Check '单段 DN 无上级' ((Get-ParentDN 'DC=t') -eq '')

Section 'Test-SameDN / Get-NormalDN'
Check '大小写不同视为相同' (Test-SameDN 'OU=Abc,DC=T' 'ou=abc,dc=t')
Check '十六进制与反斜杠转义视为相同' (Test-SameDN 'OU=研发\2C测试组,DC=t' 'OU=研发\,测试组,DC=t')
Check '= 转义与否视为相同' (Test-SameDN 'OU=a=b,DC=t' 'OU=a\=b,DC=t')
Check '逗号转义与否是不同 DN' (-not (Test-SameDN 'OU=a\,OU=b,DC=t' 'OU=a,OU=b,DC=t'))
Check '逗号后空格不影响' (Test-SameDN 'OU=a, DC=t' 'OU=a,DC=t')

Section 'Test-UnderDN'
Check '子树中' (Test-UnderDN 'OU=x,OU=其它,OU=员工,DC=t' 'OU=其它,OU=员工,DC=t')
Check '等于根' (Test-UnderDN 'ou=其它,ou=员工,dc=t' 'OU=其它,OU=员工,DC=t')
Check '名称前缀相同但不在子树中' (-not (Test-UnderDN 'OU=其它2,OU=员工,DC=t' 'OU=其它,OU=员工,DC=t'))
Check '根为空时为否' (-not (Test-UnderDN 'OU=x,DC=t' ''))

Section 'Join-OUPath / Get-UserDeptPath'
Check '逐段转义拼接（上级在前）' ((Join-OUPath @('父,部门', '子') 'OU=员工,DC=t') -ceq 'OU=子,OU=父\,部门,OU=员工,DC=t')
Check '空路径即基础 OU' ((Join-OUPath @() 'OU=员工,DC=t') -ceq 'OU=员工,DC=t')
$p = Get-UserDeptPath ([pscustomobject]@{ DepartmentPath = '["研发\\中心","A,组"]'; DepartmentName = '忽略' })
Check 'DepartmentPath(JSON) 优先，名称中可含 \ 和 ,' (($p.Count -eq 2) -and ($p[0] -ceq '研发\中心') -and ($p[1] -ceq 'A,组'))
$p = Get-UserDeptPath ([pscustomobject]@{ DepartmentPath = ''; DepartmentName = '物流部\头程组' })
Check '兼容旧格式 DepartmentName（\ 分隔）' (($p.Count -eq 2) -and ($p[1] -ceq '头程组'))
$p = Get-UserDeptPath ([pscustomobject]@{ DepartmentName = '单层' })
Check '没有 DepartmentPath 列时回退' (($p.Count -eq 1) -and ($p[0] -ceq '单层'))
$p = Get-UserDeptPath ([pscustomobject]@{ DepartmentPath = ''; DepartmentName = '' })
Check '无部门' ($p.Count -eq 0)
$p = Get-UserDeptPath ([pscustomobject]@{ DepartmentPath = '["单层"]' })
Check 'JSON 单层部门仍为数组' (($p.Count -eq 1) -and ($p[0] -ceq '单层'))

Complete-Tests
