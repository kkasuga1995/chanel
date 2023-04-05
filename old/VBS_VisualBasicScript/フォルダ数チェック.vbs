Option Explicit
call FileCountCheck

Sub FileCountCheck()
Dim fso, flg, X, fl,folderpath,filepath
dim i
dim folderO(10)
dim folderP(10)
dim folderCMAX(10)
dim folderC(10)


if msgbox("画面・帳票資料格納の各フォルダ内のファイル数をチェックします。" & vbcrlf & "実行してよろしいですか？",vbyesno + vbinformation) <> vbyes then
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
'folderpath = "C:\Users\k-kasuga\Desktop\910_画面・帳票資料格納"           
folderpath = "\\KPDSV1\Share\特定部門エリア\00_次期システム\910_画面・帳票資料格納"           


'チェックする各フォルダのパスを定義
folderP(1) =folderpath & "\00_その他・共通"
folderP(2) =folderpath & "\01_販売"
folderP(3) =folderpath & "\02_調達"
folderP(4) =folderpath & "\03_在庫"
folderP(5) =folderpath & "\04_生産"
folderP(6) =folderpath & "\05_原価"
folderP(7) =folderpath & "\06_債権"
folderP(8) =folderpath & "\07_予算"
folderP(9) =folderpath & "\08_日次・月次"
folderP(10) =folderpath & "\09_マスタ"         


'チェックする各フォルダの最大ファイル数を定義
folderCMAX(1) = 1
folderCMAX(2) = 30
folderCMAX(3) = 21
folderCMAX(4) = 34
folderCMAX(5) = 12
folderCMAX(6) = 5
folderCMAX(7) = 24
folderCMAX(8) = 14
folderCMAX(9) = 33
folderCMAX(10) = 49


'各フォルダのオブジェクトを取得
for i = 1 to 10
	set folderO(i) = fso.getfolder(folderP(i))
next 


'各フォルダ内のファイル数を取得
for i = 1 to 10
	folderC(i) = folderO(i).files.count
next   

'msgboxにファイル数の比較結果を表示
msgbox "【00_その他・共通】" & folderC(1) & " / " & folderCMAX(1)& vbcrlf & _
       "【01_販売】" & folderC(2) & " / " & folderCMAX(2)& vbcrlf & _
       "【02_調達】" & folderC(3) & " / " & folderCMAX(3)& vbcrlf & _
       "【03_在庫】" & folderC(4) & " / " & folderCMAX(4)& vbcrlf & _
       "【04_生産】" & folderC(5) & " / " & folderCMAX(5)& vbcrlf & _
       "【05_原価】" & folderC(6) & " / " & folderCMAX(6)& vbcrlf & _
       "【06_債権】" & folderC(7) & " / " & folderCMAX(7)& vbcrlf & _
       "【07_予算】" & folderC(8) & " / " & folderCMAX(8)& vbcrlf & _
       "【08_日次・月次】" & folderC(9) & " / " & folderCMAX(9)& vbcrlf & _
       "【09_マスタ】" & folderC(10) & " / " & folderCMAX(10)
        
'オブジェクトの廃棄
Set fso = Nothing

End Sub