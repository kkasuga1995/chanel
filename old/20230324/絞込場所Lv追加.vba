
�i���ꏊLv�ǉ��E�E�E�ꏊ�𕡐��w�肵�či�荞�߂�悤�ɁA�ꏊCB�ł͂Ȃ��A�ꏊLv�ōi���������w�肷��B

���t�H�[����Lv�ǉ�

���L��ϐ��ɒǉ��@Dim Basyo_�ꏊ�i��Lv_Dic As Object

�����W���[���ǉ�
Private Sub �ꏊLvIni()
    Dim Basyo()
    Basyo = Array("�����h", "�h���X�p", "����", "�w��n��", "�w��n��", "��Ȗ{�y", "�C�_��", "�Z�C���C��", "�ߌ�", "�����{", "�X���[��(�X��)", "�X���[��(����)")

    Dim �ꏊLvCnt As Long
    With �ꏊLv
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .HideSelection = False
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .MultiSelect = True
        .Gridlines = True
        .CheckBoxes = True
'        .ColumnHeaders.Add 1, "NUM", "��", Width:=19
        .ColumnHeaders.Add 1, "PRA", "�ꏊ", Width:=80
        
        For i = 0 To UBound(Basyo, 1)
            �ꏊLvCnt = �ꏊLvCnt + 1
            .ListItems.Add Text:=Basyo(i)
'            .ListItems(�ꏊLvCnt).SubItems(1) =
        Next i
    End With
    
    Dim Basyo_�ꏊ�i��Lv_Dic_Key As String
    Dim Basyo_�ꏊ�i��Lv_Dic_Item As String
    Set Basyo_�ꏊ�i��Lv_Dic = CreateObject("Scripting.Dictionary")



End Sub

��UserForm_Initialize�@�ǋL

Call �ꏊLvIni


�����X�g�\���@�X�V
Private Sub ���X�g�\��()

    
    Dim �ꏊ�� As String
    Dim ���ސ��]�� As Double
    Dim ���s�]�� As Double
    Dim ���^�]�� As Double
    Dim �����]�� As Double
    Dim �o���l�]�� As Double
    
    '�i������(�ꏊ�j�̎擾
        �ꏊ�� = basyo_�i�荞��CB.Value
    
       
    If Gekitaihyouka_�i�荞��T = "" Then
        ���ސ��]�� = 0
    Else
        ���ސ��]�� = Gekitaihyouka_�i�荞��T
    End If
    
    If Seieihyouka_�i�荞��T = "" Then
        ���s�]�� = 0
    Else
    
        ���s�]�� = Seieihyouka_�i�荞��T
    End If
    
    If Kogatahyouka_�i�荞��T = "" Then
        ���^�]�� = 0
    Else
        ���^�]�� = Kogatahyouka_�i�荞��T
    End If
    
    If MOrahyouka_�i�荞��T = "" Then
        �����]�� = 0
    Else
        �����]�� = MOrahyouka_�i�荞��T
    End If
    
    If Keikenhyouka_�i�荞��T = "" Then
        �o���l�]�� = 0
    Else
        �o���l�]�� = Keikenhyouka_�i�荞��T
    End If
    
    
    
'    If �ꏊ�� = "" And ���ސ��]�� = 0 And ���s�]�� = 0 And ���^�]�� = 0 And �����]�� = 0 And �o���l�]�� = 0 Then
'        Exit Sub
'    End If
    
    
    '�ꏊ�i��Lv�������擾
    Dim �ꏊLvItem As ListItem
    Basyo_�ꏊ�i��Lv_Dic.RemoveAll
    Basyo_�ꏊ�i��Lv_Dic_Key = ""
    For Each �ꏊLvItem In �ꏊLv.ListItems
        If �ꏊLvItem.Checked = True Then
            Basyo_�ꏊ�i��Lv_Dic_Key = CStr(�ꏊLvItem.Text)
            If Not Basyo_�ꏊ�i��Lv_Dic.Exists(Basyo_�ꏊ�i��Lv_Dic_Key) Then
                Basyo_�ꏊ�i��Lv_Dic.Add Basyo_�ꏊ�i��Lv_Dic_Key, 1
            End If
        End If
    Next
    
    
    
    'Lv�\��
    Dim WS As Worksheet
    Dim LastRow As Long
    Dim LastCol As Long
    
    Set WS = ThisWorkbook.Sheets("�G��胋�[�g��������")
    
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).row
    LastCol = WS.Cells(4, Columns.Count).End(xlToLeft).Column
    Dim i As Long, j As Long, k As Long
    Dim Cnt As Long
    
    
    
    With Lv1
        .Sorted = False
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        .View = lvwReport
        .HideSelection = False
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .Gridlines = True
        .ColumnHeaders.Add 1, "NUM", "��", Width:=19
        For j = 1 To LastCol
            If j > 3 Then
                .ColumnHeaders.Add j + 1, "NUM" & j, WS.Cells(4, j).Value, Alignment:=lvwColumnRight   '2��ڈȍ~�́A�}�X�^�[�̕\�̃w�b�_�񐔂Ɉˑ�
            Else
                .ColumnHeaders.Add j + 1, "NUM" & j, WS.Cells(4, j).Value
            End If
        Next j
    End With
    
    
    Dim ���[�g������w As String
    With Lv1
        For i = 8 To LastRow
            ���[�g������w = Val(Right(WS.Cells(i, 1).Value, Len(WS.Cells(i, 1).Value) - InStrRev(WS.Cells(i, 1).Value, "_")))
            If ���[�g������_1CH.Value = True Then
                If ���[�g������w = "1" Then
'If �ꏊ�� <> "" Then
'If WS.Cells(i, 2).Value = �ꏊ�� Then
                    If Basyo_�ꏊ�i��Lv_Dic(WS.Cells(i, 2).Value) Then
                        If WS.Cells(i, 5).Value >= ���ސ��]�� And WS.Cells(i, 6).Value >= ���^�]�� And _
                            WS.Cells(i, 7).Value >= ���s�]�� And WS.Cells(i, 8).Value >= �����]�� And _
                            WS.Cells(i, 9).Value >= �o���l�]�� Then
                                Cnt = Cnt + 1
                                .ListItems.Add Text:=Cnt
                                For j = 1 To LastCol
                                    .ListItems(Cnt).SubItems(j) = WS.Cells(i, j).Value
                                Next j
                        End If
'End If
'                    Else
'                        If WS.Cells(i, 5).Value >= ���ސ��]�� And WS.Cells(i, 6).Value >= ���^�]�� And _
'                            WS.Cells(i, 7).Value >= ���s�]�� And WS.Cells(i, 8).Value >= �����]�� And _
'                            WS.Cells(i, 9).Value >= �o���l�]�� Then
'                                Cnt = Cnt + 1
'                                .ListItems.Add Text:=Cnt
'                                For j = 1 To LastCol
'                                    .ListItems(Cnt).SubItems(j) = WS.Cells(i, j).Value
'                                Next j
'                        End If
                    End If
                End If  '�����[�g������w = "1"
            Else    '�����[�g������_1CH.Value = True


'If �ꏊ�� <> "" Then
'If WS.Cells(i, 2).Value = �ꏊ�� Then
                If Basyo_�ꏊ�i��Lv_Dic(WS.Cells(i, 2).Value) Then
                    If WS.Cells(i, 5).Value >= ���ސ��]�� And WS.Cells(i, 6).Value >= ���^�]�� And _
                        WS.Cells(i, 7).Value >= ���s�]�� And WS.Cells(i, 8).Value >= �����]�� And _
                        WS.Cells(i, 9).Value >= �o���l�]�� Then
                            Cnt = Cnt + 1
                            .ListItems.Add Text:=Cnt
                            For j = 1 To LastCol
                                .ListItems(Cnt).SubItems(j) = WS.Cells(i, j).Value
                            Next j
                    End If

'End If
'                Else
'                    If WS.Cells(i, 5).Value >= ���ސ��]�� And WS.Cells(i, 6).Value >= ���^�]�� And _
'                        WS.Cells(i, 7).Value >= ���s�]�� And WS.Cells(i, 8).Value >= �����]�� And _
'                        WS.Cells(i, 9).Value >= �o���l�]�� Then
'                            Cnt = Cnt + 1
'                            .ListItems.Add Text:=Cnt
'                            For j = 1 To LastCol
'                                .ListItems(Cnt).SubItems(j) = WS.Cells(i, j).Value
'                            Next j
'                    End If
                End If

            End If  '�����[�g������_1CH.Value = True
        Next i
    End With


    
End Sub
