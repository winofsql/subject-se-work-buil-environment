Set objShellApplication = CreateObject("Shell.Application")
If WScript.Arguments.Count = 0 Then
	objShellApplication.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " dummy", "", "runas", 1
	Wscript.Quit
End If

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

target1 = "c:\app2"

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

Set sourceFile = objFSO.GetFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\database\sqliteodbc.exe")
sourceFile.Copy "c:\app2\sqliteodbc.exe", True

Set sourceFile = objFSO.GetFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\database\sqliteodbc_w64.exe")
sourceFile.Copy "c:\app2\sqliteodbc_w64.exe", True

objShell.Run "cmd /c c:\app2\sqliteodbc.exe", 0, True
objShell.Run "cmd /c c:\app2\sqliteodbc_w64.exe", 0, True

MsgBox("SQLITE3 の ODBC ドライバをインストールしました。")
