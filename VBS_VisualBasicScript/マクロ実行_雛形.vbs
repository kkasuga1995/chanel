'------------------------------------------------------
' �o�b�`���ŁA����VBS�t�@�C�����w�肷��B
' �o�b�`���ŁA����VBS�t�@�C�����w�肷��ۂ̈����̐ݒ���A�ȉ������Œ�`����B
'		�����P�F	�}�N���u�b�N�̐�΃p�X
'		�����Q�F	���s����}�N����(�v���V�[�W����)�B�����P�̃}�N���u�b�N���ɑ��݂���v���V�[�W���B
'------------------------------------------------------

'Excel������\�ɂ���ׂ̃I�u�W�F�N�g���쐬
Dim obj
Dim WB
Set obj = WScript.CreateObject("Excel.Application")

'Excel��������ʔ�\��
obj.Visible = false

'�o�b�`�t�@�C���ł��̃t�@�C�������s����ۂ̈����ݒ�

'�������F�w�肵��Excel�}�N�����J��
'�A���A���łɊJ���Ă���ꍇ�͏������΂�

  On Error Resume Next
    obj.Workbooks.Open WScript.Arguments(0)
  On Error goto 0

'�������F�w�肵���}�N�������s
obj.Application.Run WScript.Arguments(1)