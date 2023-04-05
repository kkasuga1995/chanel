Option Explicit

Dim objFSO
Dim strFolder
Dim strMessage
Dim stryear 
Dim strmonth 
Dim strday 
Dim macrofolderpath
Dim macrofilepath
Dim macroname
Dim Rmacroname
Dim tomovepath
macrofolderpath = "C:\Users\k-kasuga\AppData\Roaming\Microsoft\Excel\XLSTART\"
macrofilepath = "C:\Users\k-kasuga\AppData\Roaming\Microsoft\Excel\XLSTART\マクロまとめブック(20190820).xlsm"
macroname = "マクロまとめブック(20190820).xlsm"
tomovepath = "C:\Users\k-kasuga\AppData\Roaming\Microsoft\Excel\XLSTART\old\"


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

'ファイル存在チェック
if not objFSO.FolderExists(macroname)  Then
    strMessage = macroname & "は存在しません。"
else
    'ファイルのリネーム
    Rmacroname = macroname &"_" & year(date) & format(month(date),"mm") & format(day(date),"dd")
 
    'リネームしたファイルの移動
        '移動先に同名ファイルが存在するかチェック
　　  if objFSO.FolderExists(macroname)  Then
      	msgbox"移動先に同名ファイルが存在します"
        exit sub
      end if
        
　　  'リネームファイルを移動
      objFSO.MoveFile macrofolderpath & Rmacroname ,tomovepath & Rmacroname 

　'処理中で、エラーが発生していないか
    if Err.Number = 0 Then
        strMessage = strFolder & "を作成しました。"
    else
        strMessage = "エラー：" & Err.Description
    end if
end if
set objFSO = nothing
MsgBox(strMessage)