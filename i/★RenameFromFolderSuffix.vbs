' Drag & drop folders onto this .vbs (runs PowerShell hidden)
Option Explicit
Dim sh, fso, scriptDir, ps1, i, a, args, cmd, ret
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ps1 = """" & scriptDir & "\RenameFromFolderSuffix.ps1" & """"

args = ""
For i = 0 To WScript.Arguments.Count - 1
  a = WScript.Arguments(i)
  a = Replace(a, """", """""")   ' quote-escape
  args = args & " """ & a & """"
Next

cmd = "powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File " & ps1 & args
ret = sh.Run(cmd, 0, True)  ' 0=hidden, True=wait until finished
WScript.Quit ret
