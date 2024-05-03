' ********************************************************************************
' Google ドライブアプリのインストール
'
' check-1 : G ドライブがある事
' check-2 : G:\共有ドライブ\SE-WORK-DOWNLOAD がある事
'
' dependemcy : G:\共有ドライブ\SE-WORK-DOWNLOAD\important\GoogleDriveSetup.exe
' ********************************************************************************

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

objShell.Run "G:\共有ドライブ\SE-WORK-DOWNLOAD\important\GoogleDriveSetup.exe", 0, True

MsgBox("Google Drive File Stream インストールを終了しました。")
