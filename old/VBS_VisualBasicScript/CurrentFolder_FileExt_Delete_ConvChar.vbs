Option Explicit

Dim objFS
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

'		msgbox curDir

    Set objFS = CreateObject("Scripting.FileSystemObject")

    Set objFolder = objFS.GetFolder(curdir)

    For Each objFile in objFolder.files

			FileExt = objFS.GetExtensionName(objFile.path)
				if fileext = "cls" or _
					 fileext = "frm" or _
					 fileext = "frx" then
'					 	msgbox objFile.name
						objFS.DeleteFile objFile,true
				else
					' Convert Charset ANSI --> UTF8

					Set ad = CreateObject("ADODB.Stream")
					ad.Type = 2
					ad.Charset = "Shift-JIS"
					ad.Open
					ad.LoadFromFile objFile.fullname
					a = ad.ReadText(-1)
					ad.Close					
					Set ad = Nothing


					Set ad = CreateObject("ADODB.Stream")
					ad.Type = 2
					ad.Charset = "UTF-8"
					ad.Open
					ad.WriteText a, 0
					ad.SaveToFile objFile.fullname, 2
					ad.Close
					Set ad = Nothing

					
				end if

    Next



    'Trush object
    Set objFS = Nothing
    Set objShell = Nothing
    Set objFolder = nothing
		set  objFile   = nothing    
