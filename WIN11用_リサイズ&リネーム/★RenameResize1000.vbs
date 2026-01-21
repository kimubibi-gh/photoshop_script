' Drag & drop folders onto this .vbs — runs PowerShell hidden (no black window)
Option Explicit
Dim sh, fso, ps1, args, i, a, cmd, psexe
Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' 明示的に Windows PowerShell 5.1 を指定（pwsh ではなく）
psexe = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"

ps1 = """" & fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "RenameResize1000.ps1") & """"

args = ""
For i = 0 To WScript.Arguments.Count - 1
  a = WScript.Arguments(i)
  a = Replace(a, """", """""")
  args = args & " """ & a & """"
Next

cmd = """" & psexe & """" & " -NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File " & ps1 & args

' 0=非表示, True=終了まで待機（処理中も画面は出ません）
sh.Run cmd, 0, True
