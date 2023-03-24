
絞込場所Lv追加・・・場所を複数指定して絞り込めるように、場所CBではなく、場所Lvで絞込条件を指定する。

◆フォームにLv追加

◆広域変数に追加　Dim Basyo_場所絞込Lv_Dic As Object

◆モジュール追加
Private Sub 場所LvIni()
    Dim Basyo()
    Basyo = Array("モンド", "ドラスパ", "璃月", "層岩地上", "層岩地下", "稲妻本土", "海祇島", "セイライ島", "鶴見", "淵下宮", "スメール(森林)", "スメール(砂漠)")

    Dim 場所LvCnt As Long
    With 場所Lv
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
        .ColumnHeaders.Add 1, "PRA", "場所", Width:=80
        
        For i = 0 To UBound(Basyo, 1)
            場所LvCnt = 場所LvCnt + 1
            .ListItems.Add Text:=Basyo(i)
'            .ListItems(場所LvCnt).SubItems(1) =
        Next i
    End With
    
    Dim Basyo_場所絞込Lv_Dic_Key As String
    Dim Basyo_場所絞込Lv_Dic_Item As String
    Set Basyo_場所絞込Lv_Dic = CreateObject("Scripting.Dictionary")



End Sub

◆UserForm_Initialize　追記

Call 場所LvIni


◆リスト表示　更新
Private Sub リスト表示()

    
    Dim 場所名 As String
    Dim 撃退数評価 As Double
    Dim 精鋭評価 As Double
    Dim 小型評価 As Double
    Dim モラ評価 As Double
    Dim 経験値評価 As Double
    
    '絞込条件(場所）の取得
        場所名 = basyo_絞り込みCB.Value
    
       
    If Gekitaihyouka_絞り込みT = "" Then
        撃退数評価 = 0
    Else
        撃退数評価 = Gekitaihyouka_絞り込みT
    End If
    
    If Seieihyouka_絞り込みT = "" Then
        精鋭評価 = 0
    Else
    
        精鋭評価 = Seieihyouka_絞り込みT
    End If
    
    If Kogatahyouka_絞り込みT = "" Then
        小型評価 = 0
    Else
        小型評価 = Kogatahyouka_絞り込みT
    End If
    
    If MOrahyouka_絞り込みT = "" Then
        モラ評価 = 0
    Else
        モラ評価 = MOrahyouka_絞り込みT
    End If
    
    If Keikenhyouka_絞り込みT = "" Then
        経験値評価 = 0
    Else
        経験値評価 = Keikenhyouka_絞り込みT
    End If
    
    
    
'    If 場所名 = "" And 撃退数評価 = 0 And 精鋭評価 = 0 And 小型評価 = 0 And モラ評価 = 0 And 経験値評価 = 0 Then
'        Exit Sub
'    End If
    
    
    '場所絞込Lv条件を取得
    Dim 場所LvItem As ListItem
    Basyo_場所絞込Lv_Dic.RemoveAll
    Basyo_場所絞込Lv_Dic_Key = ""
    For Each 場所LvItem In 場所Lv.ListItems
        If 場所LvItem.Checked = True Then
            Basyo_場所絞込Lv_Dic_Key = CStr(場所LvItem.Text)
            If Not Basyo_場所絞込Lv_Dic.Exists(Basyo_場所絞込Lv_Dic_Key) Then
                Basyo_場所絞込Lv_Dic.Add Basyo_場所絞込Lv_Dic_Key, 1
            End If
        End If
    Next
    
    
    
    'Lv表示
    Dim WS As Worksheet
    Dim LastRow As Long
    Dim LastCol As Long
    
    Set WS = ThisWorkbook.Sheets("敵狩りルート調査実績")
    
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
                .ColumnHeaders.Add j + 1, "NUM" & j, WS.Cells(4, j).Value, Alignment:=lvwColumnRight   '2列目以降は、マスターの表のヘッダ列数に依存
            Else
                .ColumnHeaders.Add j + 1, "NUM" & j, WS.Cells(4, j).Value
            End If
        Next j
    End With
    
    
    Dim ルート名末尾w As String
    With Lv1
        For i = 8 To LastRow
            ルート名末尾w = Val(Right(WS.Cells(i, 1).Value, Len(WS.Cells(i, 1).Value) - InStrRev(WS.Cells(i, 1).Value, "_")))
            If ルート名末尾_1CH.Value = True Then
                If ルート名末尾w = "1" Then
'If 場所名 <> "" Then
'If WS.Cells(i, 2).Value = 場所名 Then
                    If Basyo_場所絞込Lv_Dic(WS.Cells(i, 2).Value) Then
                        If WS.Cells(i, 5).Value >= 撃退数評価 And WS.Cells(i, 6).Value >= 小型評価 And _
                            WS.Cells(i, 7).Value >= 精鋭評価 And WS.Cells(i, 8).Value >= モラ評価 And _
                            WS.Cells(i, 9).Value >= 経験値評価 Then
                                Cnt = Cnt + 1
                                .ListItems.Add Text:=Cnt
                                For j = 1 To LastCol
                                    .ListItems(Cnt).SubItems(j) = WS.Cells(i, j).Value
                                Next j
                        End If
'End If
'                    Else
'                        If WS.Cells(i, 5).Value >= 撃退数評価 And WS.Cells(i, 6).Value >= 小型評価 And _
'                            WS.Cells(i, 7).Value >= 精鋭評価 And WS.Cells(i, 8).Value >= モラ評価 And _
'                            WS.Cells(i, 9).Value >= 経験値評価 Then
'                                Cnt = Cnt + 1
'                                .ListItems.Add Text:=Cnt
'                                For j = 1 To LastCol
'                                    .ListItems(Cnt).SubItems(j) = WS.Cells(i, j).Value
'                                Next j
'                        End If
                    End If
                End If  '＜ルート名末尾w = "1"
            Else    '＜ルート名末尾_1CH.Value = True


'If 場所名 <> "" Then
'If WS.Cells(i, 2).Value = 場所名 Then
                If Basyo_場所絞込Lv_Dic(WS.Cells(i, 2).Value) Then
                    If WS.Cells(i, 5).Value >= 撃退数評価 And WS.Cells(i, 6).Value >= 小型評価 And _
                        WS.Cells(i, 7).Value >= 精鋭評価 And WS.Cells(i, 8).Value >= モラ評価 And _
                        WS.Cells(i, 9).Value >= 経験値評価 Then
                            Cnt = Cnt + 1
                            .ListItems.Add Text:=Cnt
                            For j = 1 To LastCol
                                .ListItems(Cnt).SubItems(j) = WS.Cells(i, j).Value
                            Next j
                    End If

'End If
'                Else
'                    If WS.Cells(i, 5).Value >= 撃退数評価 And WS.Cells(i, 6).Value >= 小型評価 And _
'                        WS.Cells(i, 7).Value >= 精鋭評価 And WS.Cells(i, 8).Value >= モラ評価 And _
'                        WS.Cells(i, 9).Value >= 経験値評価 Then
'                            Cnt = Cnt + 1
'                            .ListItems.Add Text:=Cnt
'                            For j = 1 To LastCol
'                                .ListItems(Cnt).SubItems(j) = WS.Cells(i, j).Value
'                            Next j
'                    End If
                End If

            End If  '＜ルート名末尾_1CH.Value = True
        Next i
    End With


    
End Sub
