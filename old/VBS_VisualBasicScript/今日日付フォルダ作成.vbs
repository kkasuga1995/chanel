Option Explicit

Dim objFSO
Dim strFolder
Dim strMessage
Dim stryear 
Dim strmonth 
Dim strday 


Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
stryear = year(date)
strmonth = month(date)
strday = day(date)
if len(strmonth) <= 1 then
  strmonth = "0"&strmonth
end if
if len(strday) <= 1 then
  strday = "0" & strday
end if

'�쐬����t�H���_�� 


strFolder = stryear & strmonth & strday

if objFSO.FolderExists(strFolder) = True Then
    '�����̃t�H���_�����邩
    strMessage = strFolder + "�͊��ɑ��݂��Ă��܂��B"
else
    '�t�H���_�̍쐬
    objFSO.CreateFolder(strFolder)

    '�t�H���_�쐬�̏����ŁA�G���[���������Ă��Ȃ���
    if Err.Number = 0 Then
        strMessage = strFolder & "���쐬���܂����B"
    else
        strMessage = "�G���[�F" & Err.Description
    end if
end if
set objFSO = nothing
MsgBox(strMessage)