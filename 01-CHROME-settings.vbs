' ********************************************************************************
' Google Charome ポリシー設定
' レジストリインポート
' 1) HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Google\Chrome
' 2) HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Google\Chrome\CookiesSessionOnlyForUrls
' 3) HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist
' check-1 : G ドライブがある事
' check-2 : G:\共有ドライブ\SE-WORK-DOWNLOAD がある事
' dependemcy : G:\共有ドライブ\SE-WORK-DOWNLOAD\_windows-basic\chrome-policy.reg
' ********************************************************************************

' ********************************************************************************
' 管理者権限実行用 ( Shell.Application )
' ********************************************************************************
Set ShellApplication = CreateObject("Shell.Application")
If WScript.Arguments.Count = 0 Then
	ShellApplication.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " dummy", "", "runas", 1
	Wscript.Quit
End If

' 基本オブジェクト
Set WshShell = CreateObject("WScript.Shell")
' ファイル処理用オブジェクト
Set Fso = CreateObject("Scripting.FileSystemObject")

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

' 外部コマンドの実行 ( chrome-policy.reg ) のインポート
' 0 : コマンドプロンプトを開かない, True : 同期処理
WshShell.Run "reg import G:\共有ドライブ\SE-WORK-DOWNLOAD\_windows-basic\chrome-policy.reg", 0, True

' 終了ダイアログの表示
MsgBox("Google Chrome の設定を終了しました。")
