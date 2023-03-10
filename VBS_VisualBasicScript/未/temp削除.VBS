'削除するフォルダのパス       
strFrom = "C:\Users\k-kasuga\AppData\Local\Temp*"
strfile = "C:\Users\k-kasuga\AppData\Local\Temp\*.*"

'ファイルシステムオブジェクト作成
Set objFS = CreateObject("Scripting.FileSystemObject")

'フォルダ削除
objFS.DeleteFolder strFrom

'ファイル削除
objFS.DeleteFile strfile