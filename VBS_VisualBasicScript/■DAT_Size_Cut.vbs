'====================================================================================
'
'�� CSV�t�@�C���̐擪�s�ɓ���̕�����s��}�����鏈�� 
'
'�yVBS �@�\�ǉ��˗��z
'  �ړI    �F�ϊ��}�N�����s�ɂ�����A�ŏ��T�C�Y�̃f�[�^�ő��i�K�̌��؂�f�����s���ׁB
'  �d�l�T�v�FDAT�t�@�C���ɑ΂��āA���100���R�[�h�قǂɏk�������t�@�C�����Đ�������B
'            �w�b�_�͎c���B
'            �ŏI�s�̂��݂͎c���Ȃ��B
'            ���R�[�h�T�C�Y(���x���R�[�h)�́A�ϐ��Œ����ύX�\�B
'            ������̃t�@�C������ XXXXXXXX.DAT
'            �w�背�R�[�h��(���100���R�[�h)�ȉ��̃t�@�C���͏������珜�O�B
'            �o�b�N�A�b�v�͕s�v�B�㏑��
'  ���s���@�F�Ώۃt�@�C�������݂���t�H���_��.VBS�t�@�C����u���Ď��s����B
'            ���t�H���_�ւ͊������Ȃ��B
'
'  �Ώۂ̃t�@�C���F
'      �@�g���q=DAT OR dat
'      �A�{�̖���8����
'      �B�J�����g�f�B���N�g��(vbs�t�@�C�����u����Ă���t�H���_)���̃t�@�C��
'      �C���O����t�@�C����ϐ��Ŏw��\�Ƃ���B(BUHTANF2.DAT�͏��O)

'====================================================================================

'----------- �萔�̍쐬
Const ForReading = 1                                 '�ǎ�t���O
Const ForWriting = 2                                 '�����t���O
Const ForAppending = 8                               '�ǋL�t���O

Dim i, j
Dim objFso, objFolder, objFile
Dim strFileName
Dim strFileBaseName
Dim strCurrentPath
Dim strPath
Dim strMyName
Dim strExt


'===================================��������̂͂���������=================================
Const CutRecordCnt = 100                             '���N���R�[�h�܂Ńf�[�^��ۑ�����
Dim ExclusionFile(1)                                 '�ΏۊO�Ƃ���t�@�C�������v�f���Ƃ���iExclusionFile(1)�F2�̃t�@�C�������O)
ExclusionFile(0) = "BUHTANF2"                        '�ΏۊO�Ƃ���t�@�C�������w�肷��
ExclusionFile(1) = "RZANIJF1"
'==========================================================================================


'�t�@�C���V�X�e���I�u�W�F�N�g�쐬
Set objFso = CreateObject("Scripting.FileSystemobject")

If MsgBox("�J�����g�t�@�C���̃f�[�^�T�C�Y��ύX���܂����A��낵���ł����H", vbYesNo + vbInfomation, "�m�F") = vbNo Then
	WScript.Quit '�����I��
End If

'�J�����g�t�H���_�̃p�X
strCurrentPath = objFso.GetAbsolutePathName(".")  

'�J�����g�t�H���_�̃I�u�W�F�N�g���Z�b�g
Set objFolder = objFso.GetFolder(".\")

'���X�N���v�g�����擾����
strMyName = WScript.ScriptName

'�t�H���_���̃t�@�C�������擾
For Each objFile In objFolder.Files

    '�擾�����t�@�C���̃t���p�X��ێ�
	strPath = strCurrentPath & "\" & objFile.Name

	'�e�L�X�g�̍s�����m�F
	Set objRead = objFso.OpenTextFile(strPath , ForReading)    '�ǎ惂�[�h�Ńe�L�X�g���J��
	objRead.ReadAll                                            '�S���ǂނ��ƂōŏI�s�ֈړ�
	intLine = objRead.Line                                     '���݂̍s�����m�F
	objRead.Close                                              '�ǎ惂�[�h����

	strFileName = objFile.Name
	strFileBaseName = objFso.getBaseName(strFileName)

	'�g���q���擾
	strExt = UCase(objFso.GetExtensionName(strPath))

	If strMyName = strFileName Then
		'MsgBox "�����͏��O"
	ElseIf strExt <> "DAT" Then
		'MsgBox "�g���q��DAT�ł͂Ȃ�"
	ElseIf Len(strFileBaseName) <> 8 Then
		'MsgBox "�t�@�C������8�����ł͂Ȃ����ߏ��O"
	Else

		Dim vFlg
		vFlg = True
		
		For i = 0 to Ubound(ExclusionFile)
			'�擾�����t�@�C�����Ǎ��ΏۊO�������ꍇ�͏����𔲂���
			If strFileBaseName = ExclusionFile(i) then
				'MsgBox "�����𔲂���"
				vFlg = False
				Exit For
			End if
		Next

		if vFlg = True Then
			'MsgBox "�������J�n[" & strFileBaseName &  "]"

			Dim WritingText                                           '�����p�̕�����i�ȗ��j
			WritingText = ""

			Set objRead2 = objFso.OpenTextFile(strPath , ForReading)  '�ǎ惂�[�h�Ńe�L�X�g���J��

			row = 1                                                   '�s���̊m�F�p�̐��l
			Do Until objRead2.AtEndOfStream = True                    '�I���s�܂ŌJ��Ԃ�

				If row > CutRecordCnt Then                                '�ő僌�R�[�h�s�܂ŗ����珈���𔲂���
					Exit Do
				End If

				Dim ReadingText 
				ReadingText = objRead2.ReadLine
				
				if Len(ReadingText) = 1  Then '�󕶎��s�͏����ΏۊO�Ƃ���
					'msgbox "NULL"
				Else
					WritingText = WritingText & ReadingText & vbCrLf        '1�s�ǂݎ��A�����p�̕�����ɒǉ�
					row = row + 1                                         '�ǂݎ�����s����1���₷
				end IF
								
			Loop

			objRead2.Close                                            '�ǎ惂�[�h����

			Set objWriting = objFso.OpenTextFile(strPath , ForWriting)          '�������[�h�Ńe�L�X�g���J��
			objWriting.Write WritingText                             '�����p�̕�����l����C�ɏ�����
			objWriting.Close                                    '�������[�h����

		End if

	End if




Next 


'----------- �������b�Z�[�W
MsgBox "�ϊ�����!!"
