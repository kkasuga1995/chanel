Option Explicit
'------------------------------------------------------
'	�J�����g�t�H���_�� .CSV�t�@�C���̊g���q��.DAT�ɕϊ�����
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
						and objFile.name <> �����`�[�f�[�^.csv _
						and objFile.name <> �������׃f�[�^.csv _
						and objFile.name <> �������i�f�[�^.csv then
						'�����f�[�^�͕ϊ��Ɋ܂߂Ȃ��B
					   		a = objFile.NAME
					   		a = replace(a,fileext,"DAT")	'.CSV��.DAT�ɕϊ�
								a = replace(a,"C_","")        'C_XXXXXXXX �� XXXXXXXX  �ɕϊ�
								a = replace(a,"SCENARIO.","") 'SCENARIO. ���폜
								a = replace(a,"_20190401�ȑO","") '_20190401�ȑO ���폜								
								objFile.name = a
				END IF
		NEXT 


    'Trush object
    Set FSO = Nothing
    Set objShell = Nothing
    Set objFolder = nothing
		set  objFile   = nothing    
