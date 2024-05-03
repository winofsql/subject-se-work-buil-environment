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

objShell.Run "G:\共有ドライブ\SE-WORK-DOWNLOAD\software\WinMerge-2.16.28-jp-3-x64-Setup.exe", 0, True

MsgBox("WinMerge 設定を終了しました。")
