Set objShellApplication = CreateObject("Shell.Application")
If WScript.Arguments.Count = 0 Then
    objShellApplication.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " dummy", "", "runas", 1
	Wscript.Quit
End If

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

target1 = "c:\GIMPPortable"

If not objFSO.DriveExists("G:") Then
	MsgBox("G ドライブがありません。Google ドライブを実行してください")
	Wscript.Quit
End If

If not objFSO.FolderExists("G:\共有ドライブ\SE-WORK-DOWNLOAD") Then
	MsgBox("SE-WORK-DOWNLOAD フォルダが共有フォルダにありません。対象となるアカウントでログインして下さい")
	Wscript.Quit
End If

SevenZipPath = objShell.RegRead("HKLM\SOFTWARE\7-Zip\Path") & "7z.exe"

If not objFSO.FolderExists(target1) Then
	Wscript.Echo "GIMPPortable.zip をGドライブからコピーしています......"

	Set sourceFile = objFSO.GetFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\image\GIMPPortable.zip")
	sourceFile.Copy "c:\GIMPPortable.zip", True

	Wscript.Echo "GIMPPortable.zip をGドライブからコピーしました"

	Wscript.Echo "GIMPPortable.zip を解凍しています......"

	objShell.Run Chr(34) & SevenZipPath & Chr(34) & " x -o" & target1 & " " & "c:\GIMPPortable.zip", 0, True

	Wscript.Echo "GIMPPortable.zip を解凍しました"

	Wscript.Echo "GIMPPortable.zip を削除しています......"

	objFSO.DeleteFile "c:\GIMPPortable.zip"

	Wscript.Echo "GIMPPortable.zip を削除しました"

End If

work = objShell.ExpandEnvironmentStrings("C:\Users\%USERNAME%\Desktop")

set oShellLink = objShell.CreateShortcut(work & "\GIMPPortable.lnk")
oShellLink.TargetPath = target1 & "\GIMPPortable.exe"
'oShellLink.Arguments = ""
oShellLink.WindowStyle = 1
oShellLink.IconLocation = target1 & "\GIMPPortable.exe"
oShellLink.WorkingDirectory = target1
oShellLink.Save

MsgBox("GIMPPortable 設定を終了しました。")
