Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

If not objFSO.DriveExists("G:") Then
	MsgBox("G ドライブがありません。Google ドライブを実行してください")
	Wscript.Quit
End If

If not objFSO.FolderExists("G:\共有ドライブ\SE-WORK-DOWNLOAD") Then
	MsgBox("SE-WORK-DOWNLOAD フォルダが共有フォルダにありません。対象となるアカウントでログインして下さい")
	Wscript.Quit
End If

objShell.Run "G:\共有ドライブ\SE-WORK-DOWNLOAD\network\FileZilla_3.63.2_win64-setup.exe", 0, True

MsgBox("FileZilla 設定を終了しました。")
