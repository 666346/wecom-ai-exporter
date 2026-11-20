# 按群名一键导出企业微信聊天记录 + 附件
# 用法: 双击 export_group.bat，或 powershell -ExecutionPolicy Bypass -File export_group.ps1 [-Refresh]
param([switch]$Refresh)
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$OutputEncoding = [System.Text.Encoding]::UTF8

$Root  = $PSScriptRoot
$Py    = "$Root\.venv\Scripts\python.exe"
$Tools = "$Root\wecom_tools.py"
$Cfg   = Get-Content "$Root\config.json" -Raw -Encoding UTF8 | ConvertFrom-Json
$ErrLog = "$Root\runtime\_export_err.txt"
New-Item -ItemType Directory -Force -Path "$Root\runtime" | Out-Null

function Invoke-Tools {
    param([string[]]$ToolArgs)
    if (Test-Path $ErrLog) { Remove-Item $ErrLog -Force -ErrorAction SilentlyContinue }
    $raw = & $Py -X utf8 $Tools @ToolArgs 2>$ErrLog
    if (Test-Path $ErrLog) {
        $lines = Get-Content $ErrLog -Encoding UTF8 -ErrorAction SilentlyContinue
        foreach ($line in $lines) { if ($line.Trim()) { Write-Host "    $line" -ForegroundColor DarkGray } }
    }
    $text = ($raw | Out-String).Trim()
    if (-not $text) { throw '工具无输出，请检查 config.json 与企业微信是否运行' }
    try { return ($text | ConvertFrom-Json) } catch { throw "工具输出不是有效 JSON: $text" }
}

function Write-Summary {
    param($S)
    Write-Host ''
    Write-Host '────────────── 导出完成 ──────────────' -ForegroundColor Green
    Write-Host ("群名        : {0}" -f $S.name) -ForegroundColor White
    Write-Host ("conversation-id : {0}" -f $S.conversation_id) -ForegroundColor Cyan
    Write-Host ("账号目录    : {0}" -f $S.account) -ForegroundColor DarkGray
    Write-Host ("消息条数    : {0} 条（时间正序，UTF-8）" -f $S.messages) -ForegroundColor White
    Write-Host ("成功打包附件: {0} 个" -f $S.attachments_ok) -ForegroundColor Green
    Write-Host ("输出目录    : {0}" -f $S.out_dir) -ForegroundColor Cyan
    $st = $S.media_status
    if ($st) {
        $parts = @()
        foreach ($p in $st.PSObject.Properties) { $parts += "$($p.Name)=$($p.Value)" }
        Write-Host ("清单状态    : {0}" -f ($parts -join '  ')) -ForegroundColor DarkGray
    }
    if ($S.missing -and $S.missing.Count -gt 0) {
        Write-Host ''
        Write-Host ("【missing】{0} 个附件本地无缓存，需在企业微信里点开下载：" -f $S.missing.Count) -ForegroundColor Yellow
        $i = 0
        foreach ($m in $S.missing) {
            $i++
            if ($i -gt 20) { Write-Host ("  … 其余 {0} 个见 {1}" -f ($S.missing.Count - 20), $S.manifest) -ForegroundColor DarkGray; break }
            Write-Host ("  {0}. [{1}] 消息local_id={2}  {3}" -f $i, $m.kind, $m.local_id, $m.reason) -ForegroundColor Yellow
        }
    }
    if ($S.encrypted -and $S.encrypted.Count -gt 0) {
        Write-Host ''
        Write-Host ("【encrypted】{0} 个加密/无法识别的图片：" -f $S.encrypted.Count) -ForegroundColor Magenta
        $i = 0
        foreach ($m in $S.encrypted) {
            $i++
            if ($i -gt 20) { break }
            Write-Host ("  {0}. 消息local_id={1}  {2}" -f $i, $m.local_id, $m.reason) -ForegroundColor Magenta
        }
    }
    if ($S.other -and $S.other.Count -gt 0) {
        Write-Host ''
        Write-Host ("【other】{0} 个歧义/不支持的媒体（见清单）" -f $S.other.Count) -ForegroundColor DarkGray
    }
    if (($S.missing.Count + $S.encrypted.Count + $S.other.Count) -eq 0) {
        Write-Host '清单检查    : 无 missing / encrypted，附件齐全 ✅' -ForegroundColor Green
    }
    Write-Host ''
    Write-Host '文件: chat.md（正文） / chat.media.json（逐消息清单） / attachments/（图片·文件·语音）' -ForegroundColor DarkGray
    if ($S.missing.Count -gt 0) {
        Write-Host '提示: 在企微里点开上述消息让附件落地缓存后，重新运行本脚本即可补齐。' -ForegroundColor Yellow
    }
    Write-Host '把整个群文件夹路径直接给 Codex / 其他 AI 工具即可。' -ForegroundColor DarkGray
}

Write-Host ''
Write-Host '════ 企业微信 · 按群名一键导出 ════' -ForegroundColor Cyan
Write-Host ("输出根目录: {0}" -f $Cfg.export_root) -ForegroundColor DarkGray
Write-Host ("已配置账号: {0} 个" -f $Cfg.accounts.Count) -ForegroundColor DarkGray
Write-Host ''
Write-Host '正在刷新会话列表（首次或缓存过期时需要解密，请稍候）…' -ForegroundColor DarkGray
$refreshArgs = @('refresh')
if ($Refresh.IsPresent) { $refreshArgs += '--force' }
$ref = Invoke-Tools -ToolArgs $refreshArgs
foreach ($r in @($ref.refreshed)) { Write-Host ("  账号 {0}: {1} 个会话" -f $r.account, $r.conversations) -ForegroundColor DarkGray }

while ($true) {
    Write-Host ''
    $kw = Read-Host '请输入群名关键词（支持模糊匹配；直接回车退出）'
    if ([string]::IsNullOrWhiteSpace($kw)) { break }
    $res = Invoke-Tools @('match', $kw)
    if ($res.error) { Write-Host ("匹配失败: {0}" -f $res.error) -ForegroundColor Red; continue }
    if ($res.count -eq 0) {
        Write-Host ("没有匹配到包含「{0}」的会话，换个关键词试试。" -f $kw) -ForegroundColor Yellow
        continue
    }
    $chosen = $null
    $candidates = @($res.candidates)
    if ($res.count -eq 1) {
        $chosen = $candidates[0]
        Write-Host ("唯一匹配: {0}  (ID={1})" -f $chosen.name, $chosen.id) -ForegroundColor Green
    } else {
        Write-Host ''
        Write-Host ("匹配到 {0} 个会话，请输入编号选择：" -f $res.count) -ForegroundColor Cyan
        $i = 0
        foreach ($c in $candidates) {
            $i++
            Write-Host ("  [{0,2}] {1}  | {2} 条 | 最后 {3} | 账号 {4} | ID={5}" -f `
                $i, $c.name, $c.messages, $c.last_time, $c.account, $c.id) -ForegroundColor White
        }
        $sel = Read-Host '编号'
        $n = 0
        if (-not [int]::TryParse($sel, [ref]$n) -or $n -lt 1 -or $n -gt $res.count) {
            Write-Host '编号无效，返回重新输入。' -ForegroundColor Yellow
            continue
        }
        $chosen = $candidates[$n - 1]
    }
    Write-Host ''
    Write-Host ("正在导出「{0}」…（含 --with-media，可能需要几十秒）" -f $chosen.name) -ForegroundColor Cyan
    try {
        $summary = Invoke-Tools @('export', '--cid', $chosen.id, '--account', $chosen.account, '--name', $chosen.name)
        if ($summary.error) { throw $summary.error }
        Write-Summary $summary
    } catch {
        Write-Host ("导出失败: {0}" -f $_.Exception.Message) -ForegroundColor Red
    }
    $again = Read-Host '继续导出其它群？直接回车退出，输入 y 继续'
    if ($again.Trim().ToLower() -ne 'y') { break }
}

Write-Host ''
Write-Host '完成。窗口由启动脚本保持，按任意键关闭。' -ForegroundColor DarkGray
exit 0
