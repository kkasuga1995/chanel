Option Explicit
'------------------------------------------------------
'	カレントフォルダ内 .CSVファイルの拡張子を.DATに変換する
'------------------------------------------------------


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
				if (fileext = "CSV" or fileext = "csv")  _
						and objFile.name <> 発注伝票データ.csv _
						and objFile.name <> 発注明細データ.csv _
						and objFile.name <> 発注製品データ.csv then
						'発注データは変換に含めない。
					   		a = objFile.NAME
					   		a = replace(a,fileext,"DAT")	'.CSVを.DATに変換
								a = replace(a,"C_","")        'C_XXXXXXXX → XXXXXXXX  に変換
								a = replace(a,"SCENARIO.","") 'SCENARIO. を削除
								a = replace(a,"_20190401以前","") '_20190401以前 を削除								
								objFile.name = a
				END IF
		NEXT 


    'Trush object
    Set FSO = Nothing
    Set objShell = Nothing
    Set objFolder = nothing
		set  objFile   = nothing    
