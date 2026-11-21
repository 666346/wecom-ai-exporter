' Silent launcher for WeCom AI Export Tool web UI.
' - Starts the local server with pythonw.exe if it is not already running.
' - Waits for the health endpoint to become ready.
' - Opens Edge/Chrome app window when available, otherwise the default browser.

Option Explicit

Dim APP_DIR, APP_URL, PYW, APP_PY
Dim shell, env, oldNoBrowser, i, q
Dim edge1, edge2, chrome1, chrome2, pf, pf86, localAppData
Dim fso, logFile

APP_URL = "http://127.0.0.1:8765"
q = Chr(34)

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
' 以脚本所在目录为项目根目录，开源版不写死任何本机绝对路径
APP_DIR = fso.GetParentFolderName(WScript.ScriptFullName)
PYW = APP_DIR & "\.venv\Scripts\pythonw.exe"
APP_PY = APP_DIR & "\web_app.py"
If Not fso.FolderExists(APP_DIR & "\runtime") Then fso.CreateFolder APP_DIR & "\runtime"
Set logFile = fso.OpenTextFile(APP_DIR & "\runtime\launch_web.log", 8, True)
LogLine "LAUNCH_BEGIN"

Function FileExists(path)
    Dim localFso
    Set localFso = CreateObject("Scripting.FileSystemObject")
    FileExists = localFso.FileExists(path)
End Function

Sub LogLine(text)
    On Error Resume Next
    logFile.WriteLine Now & " " & text
    On Error GoTo 0
End Sub

Function ServerReady()
    Dim req, errNo, errDesc
    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
    req.SetProxy 1
    req.SetTimeouts 500, 500, 1500, 1500
    On Error Resume Next
    req.Open "GET", APP_URL & "/api/health", False
    req.Send
    errNo = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    If errNo <> 0 Then
        LogLine "HEALTH_ERR=" & errNo & " " & errDesc
        ServerReady = False
        Exit Function
    End If
    ServerReady = (req.Status = 200)
End Function

If Not FileExists(PYW) Then
    LogLine "MISSING_PYW"
    WScript.Quit 3
End If
If Not FileExists(APP_PY) Then
    LogLine "MISSING_APP_PY"
    WScript.Quit 4
End If

If ServerReady() Then
    LogLine "SERVER_ALREADY_READY"
Else
    LogLine "STARTING_SERVER"
    Set env = shell.Environment("PROCESS")
    oldNoBrowser = env.Item("WECOM_NO_BROWSER")
    env.Item("WECOM_NO_BROWSER") = "1"
    shell.CurrentDirectory = APP_DIR
    shell.Run q & PYW & q & " " & q & APP_PY & q, 0, False
    env.Item("WECOM_NO_BROWSER") = oldNoBrowser
    For i = 1 To 40
        If ServerReady() Then Exit For
        WScript.Sleep 500
    Next
    If ServerReady() Then
        LogLine "SERVER_READY"
    Else
        LogLine "SERVER_TIMEOUT"
    End If
End If

pf86 = shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
pf = shell.ExpandEnvironmentStrings("%ProgramFiles%")
localAppData = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
edge1 = pf86 & "\Microsoft\Edge\Application\msedge.exe"
edge2 = pf & "\Microsoft\Edge\Application\msedge.exe"
chrome1 = pf & "\Google\Chrome\Application\chrome.exe"
chrome2 = localAppData & "\Google\Chrome\Application\chrome.exe"

If FileExists(edge1) Then
    LogLine "OPEN_EDGE_APP"
    shell.Run q & edge1 & q & " --app=" & q & APP_URL & q, 1, False
ElseIf FileExists(edge2) Then
    LogLine "OPEN_EDGE_APP_2"
    shell.Run q & edge2 & q & " --app=" & q & APP_URL & q, 1, False
ElseIf FileExists(chrome1) Then
    LogLine "OPEN_CHROME_APP"
    shell.Run q & chrome1 & q & " --app=" & q & APP_URL & q, 1, False
ElseIf FileExists(chrome2) Then
    LogLine "OPEN_CHROME_APP_2"
    shell.Run q & chrome2 & q & " --app=" & q & APP_URL & q, 1, False
Else
    LogLine "OPEN_DEFAULT_BROWSER"
    shell.Run "rundll32.exe url.dll,FileProtocolHandler " & APP_URL, 1, False
End If

LogLine "LAUNCH_END"
logFile.Close
