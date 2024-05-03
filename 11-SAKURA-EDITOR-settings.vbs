Set objShellApplication = CreateObject("Shell.Application")
If WScript.Arguments.Count = 0 Then
    objShellApplication.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " dummy", "", "runas", 1
	Wscript.Quit
End If

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

target1 = "c:\app2"
target2 = "c:\app2\sakura"

If not objFSO.FolderExists(target1) Then

	objFSO.CreateFolder(target1)

End If

If not objFSO.DriveExists("G:") Then
	MsgBox("G ドライブがありません。Google ドライブを実行してください")
	Wscript.Quit
End If

If not objFSO.FolderExists("G:\共有ドライブ\SE-WORK-DOWNLOAD") Then
	MsgBox("SE-WORK-DOWNLOAD フォルダが共有フォルダにありません。対象となるアカウントでログインして下さい")
	Wscript.Quit
End If

SevenZipPath = objShell.RegRead("HKLM\SOFTWARE\7-Zip\Path") & "7z.exe"

If not objFSO.FolderExists(target2) Then
	Set sourceFile = objFSO.GetFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\important\sakura.zip")
	sourceFile.Copy target1 & "\sakura.zip", True

	Wscript.Echo "sakura.zip をGドライブからコピーしました"

	objShell.Run Chr(34) & SevenZipPath & Chr(34) & " x -o" & target2 & " " & target1 & "\sakura.zip", 0, True

	Wscript.Echo "sakura.zip を解凍しました"

	objFSO.DeleteFile target1 & "\sakura.zip"

	Wscript.Echo "sakura.zip を削除しました"

End If

'objShell.Run Chr(34) & target2 & "\sakura.exe" & Chr(34), 0, True

work = objShell.ExpandEnvironmentStrings("C:\Users\%USERNAME%\AppData\Roaming\sakura")

If not objFSO.FolderExists(work) Then

	objFSO.CreateFolder(work)

End If

work = objShell.ExpandEnvironmentStrings("C:\Users\%USERNAME%\AppData\Roaming\sakura\sakura.ini")

Set sourceFile = objFSO.GetFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\important\sakura.ini")
sourceFile.Copy work, True

work = objShell.ExpandEnvironmentStrings("C:\Users\%USERNAME%\Desktop")

set oShellLink = objShell.CreateShortcut(work & "\sakura.lnk")
oShellLink.TargetPath = target2 & "\sakura.exe"
'oShellLink.Arguments = ""
oShellLink.WindowStyle = 1
oShellLink.IconLocation = target2 & "\sakura.exe"
oShellLink.WorkingDirectory = target2
oShellLink.Save

objShell.Run "reg import G:\共有ドライブ\SE-WORK-DOWNLOAD\_windows-software\sakura-de-hiraku.reg", 0, True

MsgBox("sakura エディタ 設定を終了しました。")
