' 开机自启用（可选）—— 默认【未启用】，不要直接双击测试
' 使用方法：Win+R 输入 shell:startup，把本文件复制进打开的「启动」文件夹即可。
' 它用隐藏窗口启动 sync_watch.ps1，不会弹任何窗口；停止请用 stop_sync.bat。
Set ws = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "sync_watch.ps1")
ws.Run "powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """", 0, False
