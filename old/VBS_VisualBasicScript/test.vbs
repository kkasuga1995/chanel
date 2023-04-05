option explicit
Dim fs
Dim fn
Dim renfile as string
Dim fname_before as string
dim fname_after as string

set fs = wscript.createobject("scripting.filesystemobject")
fname_before = "C:\Users\k-kasuga\AppData\Roaming\Microsoft\Excel\XLSTART\マクロまとめブック(20190820).xlsm"
'マクロブック存在チェック
if fso.fileexists("C:\Users\k-kasuga\AppData\Roaming\Microsoft\Excel\XLSTART\マクロまとめブック(20190820).xlsm")=false then
   msgbox"移動対象のファイルが存在しません",vbexclamation
   goto ext
end if

'リネーム後ファイル重複チェック
’リネーム後ファイル名取得

'現フォルダ



set fn = fs.getfile("C:\Users\k-kasuga\AppData\Roaming\Microsoft\Excel\XLSTART\マクロまとめブック(20190820).xlsm")

fn.name = "マクロまとめブック(20190820)_20210413.xlsm"

set fs = nothing
set fn = nothing