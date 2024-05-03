Set objShellApplication = CreateObject("Shell.Application")
If WScript.Arguments.Count = 0 Then
	objShellApplication.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " dummy", "", "runas", 1
	Wscript.Quit
End If

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

target1 = "c:\app2"
target2 = "c:\app2\sqlite3"

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
	Set sourceFile = objFSO.GetFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\database\sqlite-tools-win-x64-3450100.zip")
	sourceFile.Copy target1 & "\sqlite3.zip", True

	Wscript.Echo "sqlite-tools-win-x64-3450100.zip をGドライブからコピーしました"

	objShell.Run Chr(34) & SevenZipPath & Chr(34) & " x -oc:\app2\sqlite3" & " " & target1 & "\sqlite3.zip", 0, True

	Wscript.Echo "sqlite-tools-win-x64-3450100.zip を解凍しました"

	objFSO.DeleteFile target1 & "\sqlite3.zip"

	Wscript.Echo "sqlite-tools-win-x64-3450100.zip を削除しました"

End If

Set sourceFile = objFSO.GetFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\database\DB.Browser.for.SQLite-3.12.2-win64.msi")
sourceFile.Copy target1 & "\DB.Browser.for.SQLite-3.12.2-win64.msi", True

objShell.Run target1 & "\DB.Browser.for.SQLite-3.12.2-win64.msi", 0, True

MsgBox("SQLITE3 設定を終了しました。")
