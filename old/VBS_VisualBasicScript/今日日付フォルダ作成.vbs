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

'作成するフォルダ名 


strFolder = stryear & strmonth & strday

if objFSO.FolderExists(strFolder) = True Then
    '同名のフォルダがあるか
    strMessage = strFolder + "は既に存在しています。"
else
    'フォルダの作成
    objFSO.CreateFolder(strFolder)

    'フォルダ作成の処理で、エラーが発生していないか
    if Err.Number = 0 Then
        strMessage = strFolder & "を作成しました。"
    else
        strMessage = "エラー：" & Err.Description
    end if
end if
set objFSO = nothing
MsgBox(strMessage)