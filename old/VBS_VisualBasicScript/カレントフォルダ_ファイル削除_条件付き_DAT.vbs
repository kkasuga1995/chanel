Option Explicit
'------------------------------------------------------
'	�J�����g�t�H���_�� �g���q = .DAT�̃t�@�C���������B
'		�����F�g���q = .csv or .CSV
'		             �� �����`�[�f�[�^,�������׃f�[�^,�������i�f�[�^ �͏����Ȃ�
'------------------------------------------------------

If MsgBox ("�J�����g�t�H���_�� �g���q = .DAT�̃t�@�C���������B", vbyesno + vbInformation, "Information") <>  vbyes then
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
						and objFile.name <> "�����`�[�f�[�^.csv" _
						and objFile.name <> "�������׃f�[�^.csv" _
						and objFile.name <> "�������i�f�[�^.csv" then
						'�����f�[�^�͕ϊ��Ɋ܂߂Ȃ��B
							FSO.DeleteFile objFile,true
				END IF
		NEXT 


    'Trush object
    Set FSO = Nothing
    Set objShell = Nothing
    Set objFolder = nothing
		set  objFile   = nothing    
