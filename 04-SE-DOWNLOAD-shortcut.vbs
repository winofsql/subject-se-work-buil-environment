' ********************************************************************************
' PC 環境の細々な設定
'
' dependemcy : C:\app2\Microsoft VS Code\bin に VisualStudioCode がインストールされている事
' dependemcy : G:\共有ドライブ\SE-WORK-DOWNLOAD が存在する
'
' act1 : PATH 環境変数の設定
' act2 : 以下のフォルダが無ければ作成
'	 c:\app
'	 c:\app2
'	 c:\app\workspace
'	 c:\TMP
'	 c:\TEMP
'	 c:\app2\保存
'	 c:\app\java23
'	 c:\app\cs23
' act3 : 以下のショートカットを作成
'	 G:\共有ドライブ\SE-WORK-DOWNLOAD
'	 C:\app\workspace
'	 C:\Users\" & username & "\Downloads
'	 mspaint.exe
'	 C:\Windows\System32\mstsc.exe ( リモート デスクトップ接続 )
'	 ディスク クリーンアップ
'	 7-Zip File Manager
'	 Microsoft Edge
'        Visual Studio Code
'	 Access
'	 Excel
' 	 Visual Studio 2022
'	 Google Chrome
'	 Google Drive
'	 WinMerge
'	 eclipse
'	 保存
'	 DB Browser for SQLite
'	 FileZilla Client
'	 Snipping Tool
'	 Firefox
' ********************************************************************************

Set Fso = CreateObject( "Scripting.FileSystemObject" )

strCurPath = WScript.ScriptFullName
Set obj = Fso.GetFile( strCurPath )
Set obj = obj.ParentFolder
strCurPath = obj.Path

set WshShell = WScript.CreateObject("WScript.Shell")

target1 = "c:\app"
target2 = "c:\app2"
target3 = "c:\app\workspace"
target4 = "c:\TMP"
target5 = "c:\TEMP"
target6 = "c:\app2\保存"
target7 = "c:\app\java23"
target8 = "c:\app\cs23"

' 環境変数
WshShell.Environment("User").Item("Path") = "C:\app2\Python\Scripts\;C:\app2\Python\;%USERPROFILE%\AppData\Local\Microsoft\WindowsApps;%USERPROFILE%\.dotnet\tools;C:\app2\Microsoft VS Code\bin;C:\java16\bin;C:\app2\git\bin;C:\app2\sqlite3;"

If Fso.FolderExists(target2 & "\DesktopOK_x64") Then

	on error resume next
	Call Fso.CopyFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\software\DesktopOK_x64\DesktopOK.ini", target2 & "\DesktopOK_x64\DesktopOK.ini", true )
	on error goto 0

End If


If not Fso.FolderExists(target1) Then

	Fso.CreateFolder(target1)

End If
If not Fso.FolderExists(target2) Then

	Fso.CreateFolder(target2)

End If
If not Fso.FolderExists(target3) Then

	Fso.CreateFolder(target3)

End If
If not Fso.FolderExists(target4) Then

	Fso.CreateFolder(target4)

End If
If not Fso.FolderExists(target5) Then

	Fso.CreateFolder(target5)

End If
If not Fso.FolderExists(target6) Then

	Fso.CreateFolder(target6)

End If
If not Fso.FolderExists(target7) Then

	Fso.CreateFolder(target7)

End If
If not Fso.FolderExists(target8) Then

	Fso.CreateFolder(target8)

End If

username = WshShell.ExpandEnvironmentStrings("%USERNAME%")

set oShellLink = WshShell.CreateShortcut("C:\Users\" & username & "\Desktop" & "\SE-WORK DOWNLOAD.lnk")
oShellLink.TargetPath = "G:\共有ドライブ\SE-WORK-DOWNLOAD"
'oShellLink.Arguments = ""
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "C:\WINDOWS\system32\imageres.dll,8"
'oShellLink.WorkingDirectory = strCurPath
oShellLink.Save

set oShellLink = WshShell.CreateShortcut("C:\Users\" & username & "\Desktop" & "\WORKSPACE.lnk")
oShellLink.TargetPath = "C:\app\workspace"
'oShellLink.Arguments = ""
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "C:\WINDOWS\system32\imageres.dll,157"
'oShellLink.WorkingDirectory = strCurPath
oShellLink.Save

set oShellLink = WshShell.CreateShortcut("C:\Users\" & username & "\Desktop" & "\DOWNLOAD.lnk")
oShellLink.TargetPath = "C:\Users\" & username & "\Downloads"
'oShellLink.Arguments = ""
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "C:\WINDOWS\system32\imageres.dll,166"
'oShellLink.WorkingDirectory = strCurPath
oShellLink.Save

set oShellLink = WshShell.CreateShortcut("C:\Users\" & username & "\Desktop" & "\ペイント.lnk")
oShellLink.TargetPath = "mspaint.exe"
'oShellLink.Arguments = ""
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "%SystemRoot%\System32\SHELL32.dll,127"
'oShellLink.WorkingDirectory = "C:\Program Files\WindowsApps\Microsoft.Paint_11.2311.30.0_x64__8wekyb3d8bbwe\PaintApp"
oShellLink.Save

set oShellLink = WshShell.CreateShortcut("C:\Users\" & username & "\Desktop" & "\リモート デスクトップ接続.lnk")
oShellLink.TargetPath = "C:\Windows\System32\mstsc.exe"
'oShellLink.Arguments = ""
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "C:\Windows\System32\mstsc.exe,0"
oShellLink.WorkingDirectory = "C:\Windows\System32"
oShellLink.Save

Call Fso.CopyFile( "G:\共有ドライブ\SE-WORK-DOWNLOAD\lnk\ディスク クリーンアップ.lnk", "C:\Users\" & username & "\Desktop" & "\ディスク クリーンアップ.lnk", true )
Call Fso.CopyFile( "G:\共有ドライブ\SE-WORK-DOWNLOAD\lnk\7-Zip File Manager.lnk", "C:\Users\" & username & "\Desktop" & "\7-Zip File Manager.lnk", true )
Call Fso.CopyFile( "G:\共有ドライブ\SE-WORK-DOWNLOAD\lnk\Microsoft Edge.lnk", "C:\Users\" & username & "\Desktop" & "\Microsoft Edge.lnk", true )
Call Fso.CopyFile( "G:\共有ドライブ\SE-WORK-DOWNLOAD\lnk\Visual Studio Code.lnk", "C:\Users\" & username & "\Desktop" & "\Visual Studio Code.lnk", true )
Call Fso.CopyFile( "G:\共有ドライブ\SE-WORK-DOWNLOAD\lnk\Access.lnk", "C:\Users\" & username & "\Desktop" & "\Access.lnk", true )
Call Fso.CopyFile( "G:\共有ドライブ\SE-WORK-DOWNLOAD\lnk\Excel.lnk", "C:\Users\" & username & "\Desktop" & "\Excel.lnk", true )
Call Fso.CopyFile( "G:\共有ドライブ\SE-WORK-DOWNLOAD\lnk\Visual Studio 2022.lnk", "C:\Users\" & username & "\Desktop" & "\Visual Studio 2022.lnk", true )
'Call Fso.CopyFile( "G:\共有ドライブ\SE-WORK-DOWNLOAD\lnk\Google Chrome.lnk", "C:\Users\" & username & "\Desktop" & "\Google Chrome.lnk", true )
Call Fso.CopyFile( "G:\共有ドライブ\SE-WORK-DOWNLOAD\lnk\Google Drive.lnk", "C:\Users\" & username & "\Desktop" & "\Google Drive.lnk", true )
'Call Fso.CopyFile( "G:\共有ドライブ\SE-WORK-DOWNLOAD\lnk\WinMerge.lnk", "C:\Users\" & username & "\Desktop" & "\WinMerge.lnk", true )
Call Fso.CopyFile( "G:\共有ドライブ\SE-WORK-DOWNLOAD\lnk\eclipse.lnk", "C:\Users\" & username & "\Desktop" & "\eclipse.lnk", true )
Call Fso.CopyFile( "G:\共有ドライブ\SE-WORK-DOWNLOAD\lnk\保存.lnk", "C:\Users\" & username & "\Desktop" & "\保存.lnk", true )
Call Fso.CopyFile( "G:\共有ドライブ\SE-WORK-DOWNLOAD\lnk\DB Browser for SQLite.lnk", "C:\Users\" & username & "\Desktop" & "\DB Browser for SQLite.lnk", true )
'Call Fso.CopyFile( "G:\共有ドライブ\SE-WORK-DOWNLOAD\lnk\FileZilla Client.lnk", "C:\Users\" & username & "\Desktop" & "\FileZilla Client.lnk", true )

set oShellLink = WshShell.CreateShortcut("C:\Users\" & username & "\Desktop" & "\ディスク クリーンアップ.lnk")
oShellLink.TargetPath = "C:\Windows\System32\cleanmgr.exe"
'oShellLink.Arguments = ""
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "C:\Windows\System32\cleanmgr.exe,0"
oShellLink.WorkingDirectory = "C:\Windows\System32"
oShellLink.Save

set oShellLink = WshShell.CreateShortcut("C:\Users\" & username & "\Desktop" & "\Snipping Tool.lnk")
oShellLink.TargetPath = "SnippingTool.exe"
'oShellLink.Arguments = ""
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "%SystemRoot%\System32\SHELL32.dll,318"
'oShellLink.WorkingDirectory = "C:\Windows\System32"
oShellLink.Save

If Fso.FileExists("C:\Users\" & username & "\Desktop" & "\WinMerge.lnk") Then
	Call Fso.DeleteFile("C:\Users\" & username & "\Desktop" & "\WinMerge.lnk", True)
End If
If Fso.FileExists("C:\Users\" & username & "\Desktop" & "\FileZilla Client.lnk") Then
	Call Fso.DeleteFile("C:\Users\" & username & "\Desktop" & "\FileZilla Client.lnk", True)
End If
If Fso.FileExists("C:\Users\" & username & "\Desktop" & "\Firefox.lnk") Then
	Call Fso.DeleteFile("C:\Users\" & username & "\Desktop" & "\Firefox.lnk", True)
End If

set oShellLink = WshShell.CreateShortcut("C:\Users\" & username & "\Desktop" & "\WinMerge.lnk")
oShellLink.TargetPath = "C:\Program Files\WinMerge\WinMergeU.exe"
'oShellLink.Arguments = ""
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "C:\Program Files\WinMerge\WinMergeU.exe,0"
oShellLink.WorkingDirectory = "C:\Program Files\WinMerge"
oShellLink.Save

set oShellLink = WshShell.CreateShortcut("C:\Users\" & username & "\Desktop" & "\FileZilla Client.lnk")
oShellLink.TargetPath = "C:\Program Files\FileZilla FTP Client\filezilla.exe"
'oShellLink.Arguments = ""
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "C:\Program Files\FileZilla FTP Client\filezilla.exe,0"
oShellLink.WorkingDirectory = "C:\Program Files\FileZilla FTP Client"
oShellLink.Save

set oShellLink = WshShell.CreateShortcut("C:\Users\" & username & "\Desktop" & "\Firefox.lnk")
oShellLink.TargetPath = "C:\Program Files\Mozilla Firefox\firefox.exe"
'oShellLink.Arguments = ""
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "C:\Program Files\Mozilla Firefox\firefox.exe,0"
oShellLink.WorkingDirectory = "C:\Program Files\Mozilla Firefox"
oShellLink.Save

Wscript.Echo "処理が終了しました"
