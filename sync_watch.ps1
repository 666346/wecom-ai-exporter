# 后台静默增量同步：每 5 分钟检查 watch_groups.txt 中的群，有变化才覆盖刷新
# 由 start_sync.bat 以隐藏窗口启动；不要手动双击运行（会一直占用窗口）
$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$Root     = $PSScriptRoot
$Py       = "$Root\.venv\Scripts\python.exe"
$Tools    = "$Root\wecom_tools.py"
$Log      = "$Root\runtime\sync.log"
$ErrFile  = "$Root\runtime\_sync_err.txt"
$Interval = 300
New-Item -ItemType Directory -Force -Path "$Root\runtime" | Out-Null

function Write-Log([string]$Line) {
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    try {
        Add-Content -Path $Log -Value "[$stamp] $Line" -Encoding UTF8
        if ((Get-Item $Log).Length -gt 1MB) {
            $keep = @(Get-Content $Log -Encoding UTF8 | Select-Object -Last 500)
            Set-Content -Path $Log -Value $keep -Encoding UTF8
        }
    } catch {}
}

Write-Log "后台同步启动，间隔 ${Interval}s，PID=$PID"
while ($true) {
    try {
        if (Test-Path $ErrFile) { Remove-Item $ErrFile -Force -ErrorAction SilentlyContinue }
        $raw = & $Py -X utf8 $Tools sync 2>$ErrFile
        $text = ($raw | Out-String).Trim()
        if ($text) {
            $obj = $text | ConvertFrom-Json
            if ($obj.error) {
                Write-Log "ERROR: $($obj.error)"
            } elseif ($obj.skipped) {
                Write-Log "SKIP $($obj.reason)"
            } else {
                $done = @(@($obj.results) | Where-Object { -not $_.skipped }).Count
                Write-Log "SYNC changed=[$($obj.changed -join ',')] exported=$done errors=$(@($obj.errors).Count)"
                foreach ($r in @($obj.results)) {
                    if ($r.messages) {
                        Write-Log ("  {0} -> {1} 条消息 / {2} 附件 / missing={3} / encrypted={4}" -f `
                            $r.name, $r.messages, $r.attachments_ok, @($r.missing).Count, @($r.encrypted).Count)
                    }
                }
                foreach ($e in @($obj.errors)) { if ($e) { Write-Log "  ERR $e" } }
            }
        } else {
            Write-Log 'ERROR: 工具无输出'
        }
        if (Test-Path $ErrFile) {
            foreach ($e in @(Get-Content $ErrFile -Encoding UTF8 -ErrorAction SilentlyContinue)) {
                if ($e.Trim()) { Write-Log "  stderr: $e" }
            }
        }
    } catch {
        Write-Log "EXCEPTION: $($_.Exception.Message)"
    }
    Start-Sleep -Seconds $Interval
}
