Option Explicit
call henkan

Sub henkan()
Dim fso, flg, X, fl,folderpath,filepath
if msgbox("設計書を更新していない場合、実行しないでください" & vbcrlf & "実行してよろしいですか？",vbyesno + vbinformation) <> vbyes then
exit sub
end if 

Set X = WScript.CreateObject("WScript.network")
If x.username <> "k-kasuga" and x.username <> "k-satou"Then
   MsgBox "k-kasuga端末のみ実行可能です",vbcritical
   Set X = Nothing
  Exit Sub
End If

Set fso = CreateObject("Scripting.FileSystemObject")
folderpath = fso.getParentFolderName(WScript.ScriptFullName)
filepath = "直近ＤＬ："

For Each fl In fso.GetFolder(folderpath).Files
   If InStr(1, fl.Path, filepath) > 0 Then
      fl.Name = "直近ＤＬ：" & Year(Now) & "年" & Month(Now) & "月" & Day(Now) & "日  " & Hour(Now) & "時" & Minute(Now) & "分" & second(Now) & "秒"
      flg = True
      Exit For
   End If
Next

If flg = True Then
   MsgBox "リネームが完了しました"
Else
   MsgBox "リネームが失敗しました",vbexclamation
End If

Set fso = Nothing
Set X = Nothing

End Sub