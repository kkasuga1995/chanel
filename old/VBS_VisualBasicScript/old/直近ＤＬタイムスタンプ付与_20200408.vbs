Option Explicit
call henkan

Sub henkan()
Dim fso, flg, X, fl,folderpath,filepath
if msgbox("�݌v�����X�V���Ă��Ȃ��ꍇ�A���s���Ȃ��ł�������" & vbcrlf & "���s���Ă�낵���ł����H",vbyesno + vbinformation) <> vbyes then
exit sub
end if 

Set X = WScript.CreateObject("WScript.network")
If x.username <> "k-kasuga" and x.username <> "k-satou"Then
   MsgBox "k-kasuga�[���̂ݎ��s�\�ł�",vbcritical
   Set X = Nothing
  Exit Sub
End If

Set fso = CreateObject("Scripting.FileSystemObject")
folderpath = fso.getParentFolderName(WScript.ScriptFullName)
filepath = "���߂c�k�F"

For Each fl In fso.GetFolder(folderpath).Files
   If InStr(1, fl.Path, filepath) > 0 Then
      fl.Name = "���߂c�k�F" & Year(Now) & "�N" & Month(Now) & "��" & Day(Now) & "��  " & Hour(Now) & "��" & Minute(Now) & "��" & second(Now) & "�b"
      flg = True
      Exit For
   End If
Next

If flg = True Then
   MsgBox "���l�[�����������܂���"
Else
   MsgBox "���l�[�������s���܂���",vbexclamation
End If

Set fso = Nothing
Set X = Nothing

End Sub