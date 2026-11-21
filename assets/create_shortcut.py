# -*- coding: utf-8 -*-
"""Create a desktop shortcut with the custom icon for the silent web launcher.

Windows only, requires pywin32. All paths are derived from this file's location
and from Windows shell/environment APIs, so no absolute machine path is stored
in the repository.
"""
from __future__ import annotations

import json
from pathlib import Path

import win32api
import win32com.client

ROOT = Path(__file__).resolve().parent.parent
VBS = ROOT / "launch_web.vbs"
ICON = ROOT / "assets" / "app.ico"
WERSCRIPT = Path(win32api.GetSystemDirectory()) / "wscript.exe"

if not VBS.is_file() or not ICON.is_file() or not WERSCRIPT.is_file():
    raise RuntimeError("launch_web.vbs / assets/app.ico / wscript.exe is missing")

shell = win32com.client.Dispatch("WScript.Shell")
# SpecialFolders resolves redirected / OneDrive desktops correctly.
desktop = Path(shell.SpecialFolders("Desktop"))
lnk = desktop / "企微AI导出工具.lnk"

shortcut = shell.CreateShortcut(str(lnk))
shortcut.TargetPath = str(WERSCRIPT)
shortcut.Arguments = f'"{VBS}"'
shortcut.WorkingDirectory = str(VBS.parent)
shortcut.IconLocation = f"{ICON},0"
shortcut.Description = "一键打开企业微信 AI 导出工具网页界面"
shortcut.Save()

check = shell.CreateShortcut(str(lnk))
print(json.dumps({
    "path": str(lnk),
    "exists": lnk.is_file(),
    "size": lnk.stat().st_size,
    "target": check.TargetPath,
    "arguments": check.Arguments,
    "working_directory": check.WorkingDirectory,
    "icon": check.IconLocation,
    "description": check.Description,
}, ensure_ascii=False, indent=2))
