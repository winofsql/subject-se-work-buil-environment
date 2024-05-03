' ********************************************************************************
' 既存 XAMPP 環境のダウンロードと設定
'
' check-1 : G ドライブがある事
' check-2 : G:\共有ドライブ\SE-WORK-DOWNLOAD がある事
'
' dependemcy1 : 7-zip がインストールされている事
' dependemcy2 : C:\xampp フォルダが無い事
'		既にある場合は、リネームすると新しいフォルダ( xampp-月日 )が作成されるので、それを C:\xampp に リネームする
'
' act1 : 以下のフォルダが無ければ作成
'	c:\app
'	c:\app\web23
'	c:\app\web23\index
' act2 : Apache 用 インデックス環境作成
' act3 : xampp のショートカットを作成 ( 管理者権限で実行 )
' act4 : TOMCAT 用の設定ファイルをダウンロード
' ********************************************************************************

' ********************************************************************************
' 管理者権限実行用 ( Shell.Application )
' ********************************************************************************
Set objShellApplication = CreateObject("Shell.Application")
If WScript.Arguments.Count = 0 Then
	objShellApplication.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " dummy", "", "runas", 1
	Wscript.Quit
End If

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

target1 = "c:\app"
target2 = "c:\app\web23"
target3 = "c:\app\web23\index"

If not objFSO.FolderExists(target1) Then

	objFSO.CreateFolder(target1)

End If
If not objFSO.FolderExists(target2) Then

	objFSO.CreateFolder(target2)

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

Wscript.Echo ".htaccess をダウンロードしています......"
Set sourceFile = objFSO.GetFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\settings\.htaccess")
sourceFile.Copy target2 & "\.htaccess", True
Wscript.Echo ".htaccess をダウンロードしました"

If not objFSO.FolderExists(target3) Then

	objShell.CurrentDirectory = "c:\app\web23"
	
	work = "C:\Users\%USERNAME%\AppData\Roaming\Code\User\script\download.vbs"
	work = objShell.ExpandEnvironmentStrings( work )
	
	Wscript.Echo "index.zip をダウンロードしています......"
	Set sourceFile = objFSO.GetFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\settings\index.zip")
	sourceFile.Copy "c:\app\web23\index.zip", True
	Wscript.Echo "index.zip をダウンロードしました"

	Wscript.Echo "index.zip を解凍しています......"
	objShell.Run Chr(34) & SevenZipPath & Chr(34) & " x -o" & target3 & " " & target2 & "\index.zip", 0, True
	Wscript.Echo "index.zip を解凍しました"
	
	Wscript.Echo "index.zip を削除しています......"
	objFSO.DeleteFile target2 & "\index.zip"
	Wscript.Echo "index.zip を削除しました"

else

	Wscript.Echo target3 & " は既に存在しています。内容を再度ダウンロードするには " & target3 & " フォルダを削除するかリネームしてください" & vbCrLf

End If

username = objShell.ExpandEnvironmentStrings("%USERNAME%")
Call objFSO.CopyFile( "G:\共有ドライブ\SE-WORK-DOWNLOAD\lnk\xampp.lnk", "C:\Users\" & username & "\Desktop" & "\xampp.lnk", true )

If not objFSO.FolderExists("c:\xampp") Then

	Wscript.Echo "xampp.zip をダウンロードしています......"
	Set sourceFile = objFSO.GetFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\database\xampp\xampp.zip")
	sourceFile.Copy "c:\xampp.zip", True
	Wscript.Echo "xampp.zip をダウンロードしました"

	suffixMonth = Right("0" & Month(Date), 2)
	suffixDay = Right("0" & Day(Date), 2)
	suffix = suffixMonth & suffixDay

	If not objFSO.FolderExists("c:\xampp-" & suffix) Then
		Wscript.Echo "xampp.zip を解凍しています......"
		objShell.Run Chr(34) & SevenZipPath & Chr(34) & " x -oc:\xampp-" & suffix & " " & "c:\xampp.zip", 0, True
		Wscript.Echo "xampp.zip を解凍しました"
	end if

	Wscript.Echo "Tomcat 用の server.xml をダウンロードしています......"
	Set sourceFile = objFSO.GetFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\database\xampp\server.xml")
	sourceFile.Copy "C:\xampp-" & suffix & "\tomcat\conf\server.xml", True
	Wscript.Echo "Tomcat 用の web.xml をダウンロードしています......"
	Set sourceFile = objFSO.GetFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\database\xampp\web.xml")
	sourceFile.Copy "C:\xampp-" & suffix & "\tomcat\conf\web.xml", True

	Wscript.Echo "Python 用の httpd.conf をダウンロードしています......"
	Set sourceFile = objFSO.GetFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\database\xampp\httpd.conf")
	sourceFile.Copy "C:\xampp-" & suffix & "\tomcat\conf\httpd.conf", True

	MsgBox("xampp 設定を終了しました。")
	Wscript.Quit

else

	Wscript.Echo "c:\xampp は既に存在しています。内容を再度ダウンロードするには c:\xampp フォルダを削除するかリネームしてください" & vbCrLf

End If

Wscript.Echo "Tomcat 用の server.xml をダウンロードしています......"
Set sourceFile = objFSO.GetFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\database\xampp\server.xml")
sourceFile.Copy "C:\xampp\tomcat\conf\server.xml", True

Wscript.Echo "Tomcat 用の web.xml をダウンロードしています......"
Set sourceFile = objFSO.GetFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\database\xampp\web.xml")
sourceFile.Copy "C:\xampp\tomcat\conf\web.xml", True

Wscript.Echo "Python 用の httpd.conf をダウンロードしています......"
Set sourceFile = objFSO.GetFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\database\xampp\httpd.conf")
sourceFile.Copy "C:\xampp\apache\conf\httpd.conf", True

MsgBox("xampp 設定を終了しました。")
