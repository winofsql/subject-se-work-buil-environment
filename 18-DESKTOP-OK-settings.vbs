Crun()

'Set objShellApplication = CreateObject("Shell.Application")
'If WScript.Arguments.Count = 0 Then
'    objShellApplication.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " dummy", "", "runas", 1
'	Wscript.Quit
'End If

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

target1 = "c:\app2"
target2 = "c:\app2\DesktopOK_x64"

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
	Wscript.Echo "DesktopOK_x64.zip をGドライブからコピーしています......"

	Set sourceFile = objFSO.GetFile("G:\共有ドライブ\SE-WORK-DOWNLOAD\software\DesktopOK_x64.zip")
	sourceFile.Copy target1 & "\DesktopOK_x64.zip", True

	Wscript.Echo "DesktopOK_x64.zip をGドライブからコピーしました"
	Wscript.Echo "DesktopOK_x64.zip を解凍しています......"

	objShell.Run Chr(34) & SevenZipPath & Chr(34) & " x -o" & target2 & " " & target1 & "\DesktopOK_x64.zip", 0, True

	Wscript.Echo "DesktopOK_x64.zip を解凍しました"
	Wscript.Echo "DesktopOK_x64.zip を削除しています......"

	objFSO.DeleteFile target1 & "\DesktopOK_x64.zip"

	Wscript.Echo "DesktopOK_x64.zip を削除しました"

End If

work = objShell.ExpandEnvironmentStrings("C:\Users\%USERNAME%\Desktop")

set oShellLink = objShell.CreateShortcut(work & "\DesktopOK.lnk")
oShellLink.TargetPath = target2 & "\DesktopOK_x64.exe"
'oShellLink.Arguments = ""
oShellLink.WindowStyle = 1
oShellLink.IconLocation = target2 & "\DesktopOK_x64.exe"
oShellLink.WorkingDirectory = target2
oShellLink.Save

MsgBox("DesktopOK_x64 設定を終了しました。")

' **********************************************************
' Cscript.exe で実行を強制
' Cscript.exe の実行終了後 pause で一時停止
' **********************************************************
Function Crun( )

	Dim str,WshShell

	' 実行中の WSH のフルパス
	str = WScript.FullName
	' 右から11文字取得
	str = Right( str, 11 )
	' 全て大文字に変更
	str = Ucase( str )
	' CSCRIPT.EXE でなければ処理を行う
	if str <> "CSCRIPT.EXE" then
		' 実行中の自分自身(スクリプト)のフルパスを取得
		str = WScript.ScriptFullName

		Set WshShell = CreateObject( "WScript.Shell" )

		' 実行中の自分自身(スクリプト)への引数を引き継ぐ為の文字列を作成
		strParam = " "
		For I = 0 to Wscript.Arguments.Count - 1
			if instr(Wscript.Arguments(I), " ") < 1 then
				strParam = strParam & Wscript.Arguments(I) & " "
			else
				strParam = strParam & Dd(Wscript.Arguments(I)) & " "
			end if
		Next
		' cscript.exe で実行しなおす為のコマンドラインを実行
		Call WshShell.Run( "cmd.exe /c cscript.exe " & Dd(str) & strParam, 1 )
		' 実行中の自分自身(スクリプト)を終了
		WScript.Quit
	end if

End Function
' **********************************************************
' 文字列を " で囲む関数
' **********************************************************
Function Dd( strValue )

	Dd = """" & strValue & """"

End function
