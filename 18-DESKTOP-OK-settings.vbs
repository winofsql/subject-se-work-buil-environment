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
	MsgBox("G �h���C�u������܂���BGoogle �h���C�u�����s���Ă�������")
	Wscript.Quit
End If

If not objFSO.FolderExists("G:\���L�h���C�u\SE-WORK-DOWNLOAD") Then
	MsgBox("SE-WORK-DOWNLOAD �t�H���_�����L�t�H���_�ɂ���܂���B�ΏۂƂȂ�A�J�E���g�Ń��O�C�����ĉ�����")
	Wscript.Quit
End If

SevenZipPath = objShell.RegRead("HKLM\SOFTWARE\7-Zip\Path") & "7z.exe"

If not objFSO.FolderExists(target2) Then
	Wscript.Echo "DesktopOK_x64.zip ��G�h���C�u����R�s�[���Ă��܂�......"

	Set sourceFile = objFSO.GetFile("G:\���L�h���C�u\SE-WORK-DOWNLOAD\software\DesktopOK_x64.zip")
	sourceFile.Copy target1 & "\DesktopOK_x64.zip", True

	Wscript.Echo "DesktopOK_x64.zip ��G�h���C�u����R�s�[���܂���"
	Wscript.Echo "DesktopOK_x64.zip ���𓀂��Ă��܂�......"

	objShell.Run Chr(34) & SevenZipPath & Chr(34) & " x -o" & target2 & " " & target1 & "\DesktopOK_x64.zip", 0, True

	Wscript.Echo "DesktopOK_x64.zip ���𓀂��܂���"
	Wscript.Echo "DesktopOK_x64.zip ���폜���Ă��܂�......"

	objFSO.DeleteFile target1 & "\DesktopOK_x64.zip"

	Wscript.Echo "DesktopOK_x64.zip ���폜���܂���"

End If

work = objShell.ExpandEnvironmentStrings("C:\Users\%USERNAME%\Desktop")

set oShellLink = objShell.CreateShortcut(work & "\DesktopOK.lnk")
oShellLink.TargetPath = target2 & "\DesktopOK_x64.exe"
'oShellLink.Arguments = ""
oShellLink.WindowStyle = 1
oShellLink.IconLocation = target2 & "\DesktopOK_x64.exe"
oShellLink.WorkingDirectory = target2
oShellLink.Save

MsgBox("DesktopOK_x64 �ݒ���I�����܂����B")

' **********************************************************
' Cscript.exe �Ŏ��s������
' Cscript.exe �̎��s�I���� pause �ňꎞ��~
' **********************************************************
Function Crun( )

	Dim str,WshShell

	' ���s���� WSH �̃t���p�X
	str = WScript.FullName
	' �E����11�����擾
	str = Right( str, 11 )
	' �S�đ啶���ɕύX
	str = Ucase( str )
	' CSCRIPT.EXE �łȂ���Ώ������s��
	if str <> "CSCRIPT.EXE" then
		' ���s���̎������g(�X�N���v�g)�̃t���p�X���擾
		str = WScript.ScriptFullName

		Set WshShell = CreateObject( "WScript.Shell" )

		' ���s���̎������g(�X�N���v�g)�ւ̈����������p���ׂ̕�������쐬
		strParam = " "
		For I = 0 to Wscript.Arguments.Count - 1
			if instr(Wscript.Arguments(I), " ") < 1 then
				strParam = strParam & Wscript.Arguments(I) & " "
			else
				strParam = strParam & Dd(Wscript.Arguments(I)) & " "
			end if
		Next
		' cscript.exe �Ŏ��s���Ȃ����ׂ̃R�}���h���C�������s
		Call WshShell.Run( "cmd.exe /c cscript.exe " & Dd(str) & strParam, 1 )
		' ���s���̎������g(�X�N���v�g)���I��
		WScript.Quit
	end if

End Function
' **********************************************************
' ������� " �ň͂ފ֐�
' **********************************************************
Function Dd( strValue )

	Dd = """" & strValue & """"

End function
