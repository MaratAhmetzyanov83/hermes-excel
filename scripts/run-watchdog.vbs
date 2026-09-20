' Runs the bridge watchdog with NO window at all.
'
' Why this file exists: the watchdog is started once a minute by Task Scheduler and by the
' Startup shortcut. Launching powershell.exe (a console app) from there can flash a console
' window. wscript.exe lives in the GUI subsystem, so it creates no console, and Run(..., 0, ...)
' asks for a hidden window for the child process too.
'
' Optional argument is forwarded to the watchdog (the Startup shortcut passes -Force so the
' bridge is up immediately at logon instead of waiting for the next minute).
Option Explicit
Dim sh, fso, here, script, extra, cmd
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
script = fso.BuildPath(here, "bridge-watchdog.ps1")
extra = ""
If WScript.Arguments.Count > 0 Then extra = " " & WScript.Arguments(0)
cmd = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File """ & script & """" & extra
sh.Run cmd, 0, False
