' �Ǘ��Ҍ������s�p ( Shell.Application )
Set ShellApplication = CreateObject("Shell.Application")
If WScript.Arguments.Count = 0 Then
	ShellApplication.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " dummy", "", "runas", 1
	Wscript.Quit
End If

' ��{�I�u�W�F�N�g
Set WshShell = CreateObject("WScript.Shell")
' �t�@�C�������p�I�u�W�F�N�g
Set Fso = CreateObject("Scripting.FileSystemObject")

' ���p�\�t�g�E�F�A�p�R�s�[��t�H���_
target2 = "c:\app2"

' �R�s�[��t�H���_�������ꍇ
If not Fso.FolderExists(target2) Then

	Fso.CreateFolder(target2)

End If

' G �h���C�u�������ꍇ
If not Fso.DriveExists("G:") Then
	' �G���[�_�C�A���O�̕\��
	MsgBox("G �h���C�u������܂���BGoogle �h���C�u�����s���Ă�������")
	' �X�N���v�g�I��
	Wscript.Quit
End If

' G �h���C�u�ɃC���X�g�[���p�t�H���_�������ꍇ
If not Fso.FolderExists("G:\���L�h���C�u\SE-WORK-DOWNLOAD") Then
	' �G���[�_�C�A���O�̕\��
	MsgBox("SE-WORK-DOWNLOAD �t�H���_�����L�t�H���_�ɂ���܂���B�ΏۂƂȂ�A�J�E���g�Ń��O�C�����ĉ�����")
	' �X�N���v�g�I��
	Wscript.Quit
End If

' ���ϐ����A���O�C�����[�U�����擾
username = WshShell.ExpandEnvironmentStrings("%USERNAME%")
' �f�X�N�g�b�v�ɃV���[�g�J�b�g���R�s�[
Call Fso.CopyFile( "G:\���L�h���C�u\SE-WORK-DOWNLOAD\lnk\7-Zip File Manager.lnk", "C:\Users\" & username & "\Desktop" & "\7-Zip File Manager.lnk", true )
' �C���X�g�[���p exe ���R�s�[
Call Fso.CopyFile( "G:\���L�h���C�u\SE-WORK-DOWNLOAD\important\7z2301-x64.exe",  target2 & "\7z2301-x64.exe", true )

' �O���R�}���h�̎��s
' �����ȗ� : �R�}���h�v�����v�g���J��, True : ��������
WshShell.Run "cmd /c c:\app2\7z2301-x64.exe", , True

' �I���_�C�A���O�̕\��
MsgBox("7-Zip �C���X�g�[�����I�����܂����B")
