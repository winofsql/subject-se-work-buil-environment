' Google Charome �|���V�[�ݒ�
' ���W�X�g���C���|�[�g
' 1) HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Google\Chrome
' 2) HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Google\Chrome\CookiesSessionOnlyForUrls
' 3) HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist

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

' �O���R�}���h�̎��s ( chrome-policy.reg ) �̃C���|�[�g
' 0 : �R�}���h�v�����v�g���J���Ȃ�, True : ��������
WshShell.Run "reg import G:\���L�h���C�u\SE-WORK-DOWNLOAD\_windows-basic\chrome-policy.reg", 0, True

' �I���_�C�A���O�̕\��
MsgBox("Google Chrome �̐ݒ���I�����܂����B")
