�Z�b�g�o�^(���[�g)
	���V�[�g(�Z�b�g�o�^_���[�g)��ǉ�                          NO ���[�g���@�ꏊ��
	���V�[�g(�Z�b�g�o�^_���[�g)                                0		1		2
	�@�@�Z�b�g���b�A�� | ���[�g���b�ꏊ��
	�@�@�e�X�g���[�g�@�@�b00001 | �Z�ۂ̒r_��_1_1 | �����h

	���t�H�[��
		�E�R���{�{�b�N�X(�Z�b�g�ꗗ)��ǉ� . �w�V�K�x�A�w�ǉ��x�A�w�Ɖ�x�A�w�폜�x�A�w�ҏW�x�{�^���ǉ�
������������������������������������������������
��                       �� �Z�b�g�ꗗCB ��  ����
��                       ������������������  ����
��                        �V�K�@�ǉ��@�Ɖ�   ����
��                        �폜               ����
��                        �Z�b�g�ҏW         ����
�� ����������������������������������������������
�� ��                                      ������
�� ��                                      ������
�� ��                                      ������
�� ��                                      ������
�� ��                                      ������
�� ����������������������������������������������
����������������������������������������������������

	�w�V�K�x:���X�g1���ȏ�I����ԁB

			 �Z�b�g�����̓��b�Z�[�W�{�b�N�X�\���B(���łɑ��݂���Z�b�g���͂͂���)
			 �Z�b�g�o�^

			 

	
	�w�ǉ��x:���X�g1���ȏ�I����ԁB
			 �Z�b�g�ꗗ�b�a�ɓo�^�ς݃Z�b�g���\������Ă�����
			 (���łɂ��̃Z�b�g�ɓo�^����Ă��郋�[�g�͂͂���)

			 ���b�Z�[�W�{�b�N�X(�ʒm�F��낵���ł����H yes no)
			 �Z�b�g�ꗗ�b�a�Ƀ��[�g�ǉ�

	�w�Ɖ�x�F�Z�b�g�ꗗ�b�a�ɓo�^�ς݃Z�b�g���\������Ă�����

			 ���X�g�ɁA�Z�b�g�ꗗ�b�a�Ŏw�肳�ꂽ�Z�b�g�̃��[�g��\��

	�w�폜�x�F�Z�b�g�ꗗ�b�a�ɓo�^�ς݃Z�b�g���\������Ă�����
	          ���X�g1���ȏ�I����ԁB

	         ���X�g�I�𒆂̃��[�g���A�Z�b�g�ꗗ�b�a�Ŏw�肳��Ă���Z�b�g����폜�B(���݃`�F�b�N)

	�w�Z�b�g�ҏW�x�F�Z�b�g�ҏW�t�H�[�����J���B�o�^���ꂽ�Z�b�g��ҏW����B
			�ҏW�\
				�E���[�g�̏���
				�E���[�g�̍폜(�ꗗ�\��)
            	�E���[�g�̒ǉ�(�ꗗ�\��)



-���ǉ���������-----------------------------------------------------------------------------------------------------------
Private Sub �Z�b�g�o�^�V�K()
'    �w�V�K�x:���X�g1���ȏ�I����ԁB
'
'             �Z�b�g�����̓��b�Z�[�W�{�b�N�X�\��� (���łɑ��݂���Z�b�g���͂͂���)
'             �Z�b�g�o�^
'�z���`
    Dim X() As String
    Dim XR As Long, XC As Long
    Dim XRc As Long
    Dim LvselectCnt As Long
    
    Dim LvItem As ListItem
    For Each LvItem In Lv1.ListItems
        If LvItem.Selected = True Then
            LvselectCnt = LvselectCnt + 1
        End If
    Next
    If LvselectCnt < 1 Then
        MsgBox "���X�g��1���ȏ�I�����Ă�������", vbInformation
        GoTo ext
    End If
    
    '�V�[�g�o�^���e�F�Z�b�g���b�A�� | ���[�g���b�ꏊ��
    XR = LvselectCnt
    XC = 4
    ReDim X(1 To XR, 1 To XC) As String
    
    
    '�Z�b�g�����̓_�C�A���O�\��
    Dim SetName As String
    On Error Resume Next
        SetName = Application.InputBox("�Z�b�g�������", Type:=2)
    On Error GoTo 0
    If SetName = "" Then
        GoTo ext
    End If
    
    '���ɓo�^�ς݂̃Z�b�g���͂͂���
    Dim WS As Worksheet
    Dim LastRow As Long
    Set WS = ThisWorkbook.Sheets("�Z�b�g�o�^_���[�g")
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).row
    Dim i As Long, j As Long, k As Long
    For i = 2 To LastRow
        If WS.Cells(i, 1).Value = SetName Then
            MsgBox "���ɓo�^�ς݂̃Z�b�g���ł��B", vbExclamation
            GoTo ext
        End If
    Next i
    
    
    '�V�[�g�o�^���e��I�𒆃��X�g �� �z��ɓ]�L
    For Each LvItem In Lv1.ListItems
        If LvItem.Selected = True Then
            XRc = XRc + 1
            X(XRc, 1) = SetName '�Z�b�g��
            X(XRc, 2) = CStr(Format(XRc, "00000")) '�A��
            X(XRc, 3) = LvItem.SubItems(1) '���[�g��
            X(XRc, 4) = LvItem.SubItems(2) '�ꏊ��
        End If
    Next
    
    WS.Cells(LastRow + 1, 1).Resize(UBound(X, 1), UBound(X, 2)) = X
    
    '�]�L��A�ŏI�s�̍Ď擾
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).row
    
    '�V�[�g��Ń\�[�g���ёւ�
    Call �Z�b�g�o�^�V�[�g_�\�[�g����


ext:
    If err.Number <> 0 Then
        MsgBox "�G���[���������܂���" & vbCrLf & "�ԍ��F" & err.Number & vbCrLf & "���e�F" & err.Description
    End If
    Set WS = Nothing
End Sub
Private Sub �Z�b�g�o�^_�ǉ�()
'    �w�ǉ��x:���X�g1���ȏ�I����ԁB
'             �Z�b�g�ꗗCB�ɓo�^�ς݃Z�b�g���\������Ă�����
'             (���łɂ��̃Z�b�g�ɓo�^����Ă��郋�[�g�͂͂���)
'
'             ���b�Z�[�W�{�b�N�X(�ʒm�F��낵���ł����H yes no)
'             �Z�b�g�ꗗCB�Ƀ��[�g�ǉ�


'�c�^�X�N
'���Z�b�g�ꗗ�b�a���t�H�[����ɒǉ�����


    '�Z�b�g�ꗗCB�ɓo�^�ς݃Z�b�g���\������Ă�����
    If �Z�b�g�ꗗCB.Value = "" Then
        MsgBox "�Z�b�g�ꗗ�����I���ł�", vbExclamation
        GoTo ext
    End If


    Dim X() As String
    Dim XR As Long, XC As Long
    Dim XRc As Long
    Dim LvselectCnt As Long
    Dim LvItem As ListItem
    
    '���X�g1���ȏ�I����ԁB
    For Each LvItem In Lv1.ListItems
        If LvItem.Selected = True Then
            LvselectCnt = LvselectCnt + 1
        End If
    Next
    If LvselectCnt < 1 Then
        MsgBox "���X�g��1���ȏ�I�����Ă�������", vbInformation
        GoTo ext
    End If
    
    
    ���������܂�
    
    '���łɂ��̃Z�b�g�ɓo�^����Ă��郋�[�g�͂͂���
    Dim WS As Worksheet
    Dim LastRow As Long
    Set WS = ThisWorkbook.Sheets("�Z�b�g�o�^_���[�g")
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).row
    Dim i As Long, j As Long, k As Long
     '�A�z�z��Ɏw�肵���Z�b�g���̃��[�g�ꗗ���i�[
    Dim ���[�gDic As Object
    Dim ���[�gDicKey As String
    Dim ���[�gDicItem As String
    Set ���[�gDic = CreateObject("Scripting.Dictionary")
    
    
    
    For i = 2 To LastRow
        If WS.Cells(i, 1).Value = SetName Then
            MsgBox "���ɓo�^�ς݂̃Z�b�g���ł��B", vbExclamation
            GoTo ext
        End If
    Next i
    


        
        
    
    
    '�V�[�g�o�^���e�F�Z�b�g���b�A�� | ���[�g���b�ꏊ��
    XR = LvselectCnt
    XC = 4
    ReDim X(1 To XR, 1 To XC) As String
    
    
    '�Z�b�g�����̓_�C�A���O�\��
    Dim SetName As String
    On Error Resume Next
        SetName = Application.InputBox("�Z�b�g�������", Type:=2)
    On Error GoTo 0
    If SetName = "" Then
        GoTo ext
    End If
    

    
    
    '�V�[�g�o�^���e��I�𒆃��X�g �� �z��ɓ]�L
    For Each LvItem In Lv1.ListItems
        If LvItem.Selected = True Then
            XRc = XRc + 1
            X(XRc, 1) = SetName '�Z�b�g��
            X(XRc, 2) = CStr(Format(XRc, "00000")) '�A��
            X(XRc, 3) = LvItem.SubItems(1) '���[�g��
            X(XRc, 4) = LvItem.SubItems(2) '�ꏊ��
        End If
    Next
    
    WS.Cells(LastRow + 1, 1).Resize(UBound(X, 1), UBound(X, 2)) = X
    
    '�]�L��A�ŏI�s�̍Ď擾
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).row
    
    '�V�[�g��Ń\�[�g���ёւ�
    Call �Z�b�g�o�^�V�[�g_�\�[�g����


ext:
    If err.Number <> 0 Then
        MsgBox "�G���[���������܂���" & vbCrLf & "�ԍ��F" & err.Number & vbCrLf & "���e�F" & err.Description
    End If
    Set WS = Nothing
End Sub

Private Sub �Z�b�g�o�^�V�[�g_�\�[�g����()
    Dim WS As Worksheet
    Dim LastRow As Long
    Set WS = ThisWorkbook.Sheets("�Z�b�g�o�^_���[�g")
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).row
    Dim i As Long, j As Long, k As Long
    With WS.Sort
        With .SortFields
            .Clear
            .Add Key:=WS.Range("A1"), Order:=xlAscending
            .Add Key:=WS.Range("B1"), Order:=xlAscending
        End With
        .SetRange WS.Range("A1:D" & LastRow)
        .Header = xlYes
        .Apply
    End With

ext:
    If err.Number <> 0 Then
        MsgBox "�G���[���������܂���" & vbCrLf & "�ԍ��F" & err.Number & vbCrLf & "���e�F" & err.Description
    End If
    Set WS = Nothing


End Sub