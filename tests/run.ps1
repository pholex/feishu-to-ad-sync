# 运行全部回归测试（不依赖 AD 环境，AD cmdlet 由 tests/MockAD.ps1 模拟），任一失败则退出码为 1
#   pwsh -NoProfile -File tests/run.ps1
#   TEST_VERBOSE=1 pwsh -NoProfile -File tests/run.ps1    # 同时打印被测脚本的日志
$failed = @()
foreach ($t in Get-ChildItem -Path $PSScriptRoot -Filter '*.Tests.ps1' | Sort-Object Name) {
    Write-Host "`n##### $($t.Name)" -ForegroundColor Cyan
    # 每个测试文件独立进程运行，互不影响 mock 状态
    pwsh -NoProfile -File $t.FullName
    if ($LASTEXITCODE -ne 0) { $failed += $t.Name }
}
if ($failed) {
    Write-Host "`n✗ 失败: $($failed -join ', ')" -ForegroundColor Red
    exit 1
}
Write-Host "`n✓ 全部通过" -ForegroundColor Green
