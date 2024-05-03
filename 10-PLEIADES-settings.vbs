Set objShellApplication = CreateObject("Shell.Application")
If WScript.Arguments.Count = 0 Then
    objShellApplication.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " dummy", "", "runas", 1
	Wscript.Quit
End If

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

target0 = "c:\pleiades"
target1 = "c:\app"
target2 = "c:\app\java23"
target3 = "c:\app\java23\subject-0511"

If not objFSO.FolderExists(target1) Then

	objFSO.CreateFolder(target1)

End If
If not objFSO.FolderExists(target2) Then

	objFSO.CreateFolder(target2)

End If
If not objFSO.FolderExists(target3) Then

	objFSO.CreateFolder(target3)

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

If not objFSO.FolderExists(target0) Then

	Wscript.Echo "pleiades.zip を copy しています。しばらくお待ちください"

	Set sourceFile = objFSO.GetFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\java\pleiades.zip")
	sourceFile.Copy target0 & ".zip", True

	Wscript.Echo "pleiades.zip の copy が終了しました"

	Wscript.Echo "pleiades.zip を 解凍しています。しばらくお待ちください"

	objShell.Run Chr(34) & SevenZipPath & Chr(34) & " x -o" & target0 & " " & target0 & ".zip", 0, True

	Wscript.Echo "pleiades.zip の 解凍が終了しました"

End If

MsgBox("pleiades 設定を終了しました。")
