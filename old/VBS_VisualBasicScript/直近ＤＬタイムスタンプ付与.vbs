'�J�����g�p�X�ȊO����ł��A�t�@�C�����̕ύX���\�Ƃ����B�@20200410
Option Explicit
call henkan

Sub henkan()
Dim fso, flg, X, fl,folderpath,filepath
if msgbox("�݌v�����X�V���Ă��Ȃ��ꍇ�A���s���Ȃ��ł�������" & vbcrlf & "���s���Ă�낵���ł����H",vbyesno + vbinformation) <> vbyes then
exit sub
end if 

'�h���C�����[�U�[�����擾���Ak-kasuga��k-satou�ȊO�A������ʂ��Ȃ�
Set X = WScript.CreateObject("WScript.network")
If x.username <> "k-kasuga" and x.username <> "k-satou" Then
   MsgBox "k-kasuga�[���܂��́Ak-satou�[���̂ݎ��s�\�ł�",vbcritical
   Set X = Nothing
  Exit Sub
End If


Set fso = CreateObject("Scripting.FileSystemObject")

'�J�����g�t�H���_�̃p�X���擾���A�w�肷��ꍇ�͏��folderpath�擾���J��
'��΃p�X�Ŏw�肷��ꍇ�͉���folderpath�擾���J������
'folderpath = fso.getParentFolderName(WScript.ScriptFullName)�@
folderpath = "\\KPDSV1\Share\���蕔��G���A\00_�����V�X�e��\910_��ʁE���[�����i�["
filepath = "���߂c�k�F"

For Each fl In fso.GetFolder(folderpath).Files
   If InStr(1, fl.Path, filepath) > 0 Then
      fl.Name = "���߂c�k�F" & Year(Now) & "�N" & Month(Now) & "��" & Day(Now) & "��  " & Hour(Now) & "��" & Minute(Now) & "��" & second(Now) & "�b"
      flg = True
      Exit For
   End If
Next

'flg=false�̏ꍇ�A�ύX�������s���Ă��Ȃ��̂ŁA���l�[�������s���܂����ƕ\������Ă����ɉe���͖����B
If flg = True Then
   MsgBox "���l�[�����������܂���"
Else
   MsgBox "���l�[�������s���܂���",vbexclamation
End If

Set fso = Nothing
Set X = Nothing

End Sub