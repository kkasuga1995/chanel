'------------------------------------------------------
' バッチ側で、このVBSファイルを指定する。
' バッチ側で、このVBSファイルを指定する際の引数の設定を、以下処理で定義する。
'		引数１：	マクロブックの絶対パス
'		引数２：	実行するマクロ名(プロシージャ名)。引数１のマクロブック内に存在するプロシージャ。
'------------------------------------------------------

'Excel操作を可能にする為のオブジェクトを作成
Dim obj
Dim WB
Set obj = WScript.CreateObject("Excel.Application")

'Excel処理を画面非表示
obj.Visible = false

'バッチファイルでこのファイルを実行する際の引数設定

'第一引数：指定したExcelマクロを開く
'但し、すでに開いている場合は処理を飛ばす

  On Error Resume Next
    obj.Workbooks.Open WScript.Arguments(0)
  On Error goto 0

'第二引数：指定したマクロを実行
obj.Application.Run WScript.Arguments(1)