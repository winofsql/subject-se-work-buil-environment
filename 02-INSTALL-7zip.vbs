' 管理者権限実行用 ( Shell.Application )
Set ShellApplication = CreateObject("Shell.Application")
If WScript.Arguments.Count = 0 Then
	ShellApplication.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " dummy", "", "runas", 1
	Wscript.Quit
End If

' 基本オブジェクト
Set WshShell = CreateObject("WScript.Shell")
' ファイル処理用オブジェクト
Set Fso = CreateObject("Scripting.FileSystemObject")

' 利用ソフトウェア用コピー先フォルダ
target2 = "c:\app2"

' コピー先フォルダが無い場合
If not Fso.FolderExists(target2) Then

	Fso.CreateFolder(target2)

End If

' G ドライブが無い場合
If not Fso.DriveExists("G:") Then
	' エラーダイアログの表示
	MsgBox("G ドライブがありません。Google ドライブを実行してください")
	' スクリプト終了
	Wscript.Quit
End If

' G ドライブにインストール用フォルダが無い場合
If not Fso.FolderExists("G:\共有ドライブ\SE-WORK-DOWNLOAD") Then
	' エラーダイアログの表示
	MsgBox("SE-WORK-DOWNLOAD フォルダが共有フォルダにありません。対象となるアカウントでログインして下さい")
	' スクリプト終了
	Wscript.Quit
End If

' 環境変数より、ログインユーザ名を取得
username = WshShell.ExpandEnvironmentStrings("%USERNAME%")
' デスクトップにショートカットをコピー
Call Fso.CopyFile( "G:\共有ドライブ\SE-WORK-DOWNLOAD\lnk\7-Zip File Manager.lnk", "C:\Users\" & username & "\Desktop" & "\7-Zip File Manager.lnk", true )
' インストール用 exe をコピー
Call Fso.CopyFile( "G:\共有ドライブ\SE-WORK-DOWNLOAD\important\7z2301-x64.exe",  target2 & "\7z2301-x64.exe", true )

' 外部コマンドの実行
' 引数省略 : コマンドプロンプトを開く, True : 同期処理
WshShell.Run "cmd /c c:\app2\7z2301-x64.exe", , True

' 終了ダイアログの表示
MsgBox("7-Zip インストールを終了しました。")
