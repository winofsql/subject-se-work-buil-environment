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
	MsgBox("G �h���C�u������܂���BGoogle �h���C�u�����s���Ă�������")
	Wscript.Quit
End If

If not objFSO.FolderExists("G:\���L�h���C�u\SE-WORK-DOWNLOAD") Then
	MsgBox("SE-WORK-DOWNLOAD �t�H���_�����L�t�H���_�ɂ���܂���B�ΏۂƂȂ�A�J�E���g�Ń��O�C�����ĉ�����")
	Wscript.Quit
End If

Set sourceFile = objFSO.GetFile("G:\���L�h���C�u\SE-WORK-DOWNLOAD\database\mysql-connector-odbc-8.0.25-win32.msi")
sourceFile.Copy "c:\app2\mysql-connector-odbc-8.0.25-win32.msi", True

Set sourceFile = objFSO.GetFile("G:\���L�h���C�u\SE-WORK-DOWNLOAD\database\mysql-connector-odbc-8.0.25-winx64.msi")
sourceFile.Copy "c:\app2\mysql-connector-odbc-8.0.25-winx64.msi", True


objShell.Run "c:\app2\mysql-connector-odbc-8.0.25-win32.msi", 0, True
objShell.Run "c:\app2\mysql-connector-odbc-8.0.25-winx64.msi", 0, True

objShell.Run "reg import G:\���L�h���C�u\SE-WORK-DOWNLOAD\_windows-software\mysql-odbc-entry.reg", 0, True
objShell.Run "reg import G:\���L�h���C�u\SE-WORK-DOWNLOAD\_windows-software\mysql-odbc.reg", 0, True

MsgBox("MySQL �� ODBC �h���C�o���C���X�g�[�����܂����B")
