Option Explicit
'============================================================================
'�@�w�t�H���_�쐬.VBS�x���s
'�A�f�X�N�g�b�v��́w���t�t�H���_�x���̃t�@�C�����A�@�ō쐬�����t�H���_�Ɉړ�
'�B����
 
'	���Ώۃt�@�C���̏�����
'		�E�J�����g�t�H���_���̃t�@�C��
'   �E�g���q��DAT									 (�ȉ�	�����@)
'   �E�t�@�C����(�{��)��8���ł���B(�ȉ�	�����A)(��IRAKUKAF.DAT �~IRAKUKAF_2021.DAT)
'
' �����t�t�H�[�}�b�g��
'   �t�@�C����(�{��)�̖�����_YYYYMMDD_HHNNSS��t�^����B
' 	�y��z
'			�t�@�C����	�FIRAKUKAF.DAT
'   	���t        �F2021�N6��25�� 13:05:33
'   	����        �FIRAKUKAF_20210625_130533.DAT
'	�������̗��ꁄ
'  �P�D�J�����g�t�H���_���̕t�^�Ώۃt�@�C�������l�[������ 		(�ȉ�	�����@)
'  �Q�D���l�[�������t�@�C�����w�ߋ����x�t�H���_�ɐ؂��肷�� (�ȉ�	�����A)
'  �R�D�I���
'============================================================================
'�@�w�t�H���_�쐬.VBS�x���s
	'�N���p�̃I�u�W�F�N�g�𐶐�
	Dim objWsh
	Set objWsh = WScript.CreateObject("WScript.Shell")

	'���s
	objWsh.Run "\\KPDSV1\Share\�d�Z\800 �t����ƃt�H���_�[\��t�H���_\Sou_���t\�t�H���_�쐬.vbs",,True



''	���t�^���t���擾���遄
'    Dim StrYear							'�V�X�e�����t����N�𒊏o�����l(YYYY)���i�[����
'    Dim StrMonth            '�V�X�e�����t���猎�𒊏o�����l(MM)���i�[����
'    Dim StrDay              '�V�X�e�����t������𒊏o�����l(DD)���i�[����
'    Dim StrHour             '�V�X�e�����t���玞�𒊏o�����l(HH)���i�[����
'    Dim StrMin              '�V�X�e�����t���番�𒊏o�����l(NN)���i�[����
'    Dim StrSec             	'�V�X�e�����t����b�𒊏o�����l(SS)���i�[����
'    Dim AttachDate          '�t�@�C���ɕt�^������t��������i�[����
'
''	���V�X�e�����t����A�N�A���A���A���A���A�b�𒊏o���遄
'    StrYear = Left(Now, 4): StrMonth = Mid(Now, 6, 2): StrDay = Mid(Now, 9, 2)
'    StrHour = Mid(Now, 12, 2): StrMin = Mid(Now, 15, 2): StrSec = Mid(Now, 18, 2)
''	�����o�����N�A���A���A���A���A�b���������遄
'    AttachDate = StrYear & StrMonth & StrDay & "_" & StrHour & StrMin & StrSec
'
'' ���t�@�C������p�I�u�W�F�N�g��`�Ɗi�[��
'    Dim FSO									'FileSystemObject�̃C���X�^���X���i�[����
'    Dim FLtmp               '���[�v���Ńt�@�C���I�u�W�F�N�g��s�x�i�[����
'    Dim FolderObj						'�����Ώۂ̃t�H���_�I�u�W�F�N�g���i�[����
'    Dim CurrentPath       	'�����Ώۂ̃t�H���_�p�X���i�[����
'		Set FSO = WScript.CreateObject("Scripting.FileSystemObject")	'FileSystemObject�̃C���X�^���X�𐶐����AFSO�Ɋi�[
'   	Set FolderObj = FSO.GetFolder(CurrentPath)	'�J�����g�t�H���_�̃I�u�W�F�N�g���i�[		
'    CurrentPath = FSO.GetAbsolutePathName("./")	'�{���s�t�@�C�����u����Ă���t�H���_(�J�����g�t�H���_)�̃p�X���i�[
'
'' �����l�[���ƃt�@�C���ړ���
'		'�J�����g�t�H���_�������[�v���āA�����@�A�����A�ɍ��v����t�@�C���ɑ΂��āA�����@�Ə����A�����s����
'    For Each FLtmp In FolderObj.Files
'			If FSO.GetExtensionName(FLtmp.Path) = "DAT" Then		'�����@
'      	If Len(FSO.GetBaseName(FLtmp.Path)) = 8 Then				'�����A	
'					FLtmp.Name = FSO.GetBaseName(FLtmp.Path) & "_" & AttachDate & ".DAT"	'�����@
'          FSO.MoveFile FLtmp.Path, Currentpath & "\�ߋ���\"                     '�����A          
'        End If
'      End If
'    Next                                                                                   
'
''	���i�[�����I�u�W�F�N�g�̔j����	
'    Set FSO = Nothing
'    Set FLtmp = Nothing