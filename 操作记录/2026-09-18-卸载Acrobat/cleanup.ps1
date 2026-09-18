# Acrobat 卸载残留清理（需要管理员权限）
# 回滚办法：双击 备份\reg_*.reg 恢复注册表；Common Files\Adobe\Color 重新装 Adobe 产品时会重建

$log = 'C:\Users\HZJ\Desktop\Git\操作记录\2026-09-18-卸载Acrobat\备份\cleanup_log.txt'
"=== START $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ===" | Out-File $log -Encoding UTF8

"--- [3] C:\Program Files\Common Files\Adobe ---" | Out-File -Append $log -Encoding UTF8
$p = 'C:\Program Files\Common Files\Adobe'
if (Test-Path $p) {
    Remove-Item $p -Recurse -Force -ErrorAction Continue
}
"    still exists after delete: $(Test-Path $p)" | Out-File -Append $log -Encoding UTF8

"--- [6] registry keys ---" | Out-File -Append $log -Encoding UTF8
$targets = @(
    'HKLM:\SOFTWARE\Adobe\Adobe Acrobat',
    'HKLM:\SOFTWARE\WOW6432Node\Adobe\Adobe Acrobat',
    'HKLM:\SOFTWARE\Policies\Adobe\Adobe Acrobat'
)
foreach ($k in $targets) {
    if (Test-Path $k) {
        Remove-Item $k -Recurse -Force -ErrorAction Continue
        "    deleted  $k   (still exists: $(Test-Path $k))" | Out-File -Append $log -Encoding UTF8
    } else {
        "    absent   $k" | Out-File -Append $log -Encoding UTF8
    }
}

"--- parent keys (only removed if they end up empty) ---" | Out-File -Append $log -Encoding UTF8
foreach ($k in @('HKLM:\SOFTWARE\Policies\Adobe','HKLM:\SOFTWARE\Adobe','HKLM:\SOFTWARE\WOW6432Node\Adobe')) {
    if (Test-Path $k) {
        $n = @(Get-ChildItem $k -ErrorAction SilentlyContinue).Count
        if ($n -eq 0) {
            Remove-Item $k -Force -ErrorAction Continue
            "    empty parent REMOVED: $k" | Out-File -Append $log -Encoding UTF8
        } else {
            "    parent KEPT (has $n subkeys): $k" | Out-File -Append $log -Encoding UTF8
        }
    } else {
        "    parent absent: $k" | Out-File -Append $log -Encoding UTF8
    }
}

"=== END $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ===" | Out-File -Append $log -Encoding UTF8
