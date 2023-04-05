'カレントパス以外からでも、ファイル名の変更を可能とした。　20200410
Option Explicit
call henkan

Sub henkan()
Dim fso, flg, X, fl,folderpath,filepath
if msgbox("設計書を更新していない場合、実行しないでください" & vbcrlf & "実行してよろしいですか？",vbyesno + vbinformation) <> vbyes then
exit sub
end if 

'ドメインユーザー名を取得し、k-kasugaかk-satou以外、ここを通さない
Set X = WScript.CreateObject("WScript.network")
If x.username <> "k-kasuga" and x.username <> "k-satou" Then
   MsgBox "k-kasuga端末または、k-satou端末のみ実行可能です",vbcritical
   Set X = Nothing
  Exit Sub
End If


Set fso = CreateObject("Scripting.FileSystemObject")

'カレントフォルダのパスを取得し、指定する場合は上のfolderpath取得を開放
'絶対パスで指定する場合は下のfolderpath取得を開放する
'folderpath = fso.getParentFolderName(WScript.ScriptFullName)　
folderpath = "\\KPDSV1\Share\特定部門エリア\00_次期システム\910_画面・帳票資料格納"
filepath = "直近ＤＬ："

For Each fl In fso.GetFolder(folderpath).Files
   If InStr(1, fl.Path, filepath) > 0 Then
      fl.Name = "直近ＤＬ：" & Year(Now) & "年" & Month(Now) & "月" & Day(Now) & "日  " & Hour(Now) & "時" & Minute(Now) & "分" & second(Now) & "秒"
      flg = True
      Exit For
   End If
Next

'flg=falseの場合、変更処理を行っていないので、リネームが失敗しましたと表示されても他に影響は無し。
If flg = True Then
   MsgBox "リネームが完了しました"
Else
   MsgBox "リネームが失敗しました",vbexclamation
End If

Set fso = Nothing
Set X = Nothing

End Sub