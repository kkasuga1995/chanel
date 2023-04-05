Option Explicit
'------------------------------------------------------
'	カレントフォルダ内 拡張子 = .DATのファイルを消す。
'		条件：拡張子 = .csv or .CSV
'		             ※ 発注伝票データ,発注明細データ,発注製品データ は消さない
'------------------------------------------------------

If MsgBox ("カレントフォルダ内 拡張子 = .DATのファイルを消す。", vbyesno + vbInformation, "Information") <>  vbyes then
    WScript.Quit
End If


Dim FSO
Dim objFolder
Dim objFile

Dim objShell
Dim curDir
Dim FileExt
Dim ad
Dim a



' CurrentDirectory get
		Set objShell = CreateObject( "WScript.Shell" )
		curDir = objShell.CurrentDirectory
		Set FSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = FSO.GetFolder(curdir)
		For Each objFile in objFolder.files
			FileExt = FSO.GetExtensionName(objFile.path)
				if (fileext = "DAT" or fileext = "dat")  _
						and objFile.name <> "発注伝票データ.csv" _
						and objFile.name <> "発注明細データ.csv" _
						and objFile.name <> "発注製品データ.csv" then
						'発注データは変換に含めない。
							FSO.DeleteFile objFile,true
				END IF
		NEXT 


    'Trush object
    Set FSO = Nothing
    Set objShell = Nothing
    Set objFolder = nothing
		set  objFile   = nothing    
