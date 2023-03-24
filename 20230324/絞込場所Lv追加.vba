
iêŠLv’Ç‰ÁEEEêŠ‚ğ•¡”w’è‚µ‚Äi‚è‚ß‚é‚æ‚¤‚ÉAêŠCB‚Å‚Í‚È‚­AêŠLv‚ÅiğŒ‚ğw’è‚·‚éB

ŸƒtƒH[ƒ€‚ÉLv’Ç‰Á

ŸLˆæ•Ï”‚É’Ç‰Á@Dim Basyo_êŠiLv_Dic As Object

Ÿƒ‚ƒWƒ…[ƒ‹’Ç‰Á
Private Sub êŠLvIni()
    Dim Basyo()
    Basyo = Array("ƒ‚ƒ“ƒh", "ƒhƒ‰ƒXƒp", "—Œ", "‘wŠâ’nã", "‘wŠâ’n‰º", "ˆîÈ–{“y", "ŠC‹_“‡", "ƒZƒCƒ‰ƒC“‡", "’ßŒ©", "•£‰º‹{", "ƒXƒ[ƒ‹(X—Ñ)", "ƒXƒ[ƒ‹(»”™)")

    Dim êŠLvCnt As Long
    With êŠLv
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .HideSelection = False
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .MultiSelect = True
        .Gridlines = True
        .CheckBoxes = True
'        .ColumnHeaders.Add 1, "NUM", "‡‚", Width:=19
        .ColumnHeaders.Add 1, "PRA", "êŠ", Width:=80
        
        For i = 0 To UBound(Basyo, 1)
            êŠLvCnt = êŠLvCnt + 1
            .ListItems.Add Text:=Basyo(i)
'            .ListItems(êŠLvCnt).SubItems(1) =
        Next i
    End With
    
    Dim Basyo_êŠiLv_Dic_Key As String
    Dim Basyo_êŠiLv_Dic_Item As String
    Set Basyo_êŠiLv_Dic = CreateObject("Scripting.Dictionary")



End Sub

ŸUserForm_Initialize@’Ç‹L

Call êŠLvIni


ŸƒŠƒXƒg•\¦@XV
Private Sub ƒŠƒXƒg•\¦()

    
    Dim êŠ–¼ As String
    Dim Œ‚‘Ş”•]‰¿ As Double
    Dim ¸‰s•]‰¿ As Double
    Dim ¬Œ^•]‰¿ As Double
    Dim ƒ‚ƒ‰•]‰¿ As Double
    Dim ŒoŒ±’l•]‰¿ As Double
    
    'iğŒ(êŠj‚Ìæ“¾
        êŠ–¼ = basyo_i‚è‚İCB.Value
    
       
    If Gekitaihyouka_i‚è‚İT = "" Then
        Œ‚‘Ş”•]‰¿ = 0
    Else
        Œ‚‘Ş”•]‰¿ = Gekitaihyouka_i‚è‚İT
    End If
    
    If Seieihyouka_i‚è‚İT = "" Then
        ¸‰s•]‰¿ = 0
    Else
    
        ¸‰s•]‰¿ = Seieihyouka_i‚è‚İT
    End If
    
    If Kogatahyouka_i‚è‚İT = "" Then
        ¬Œ^•]‰¿ = 0
    Else
        ¬Œ^•]‰¿ = Kogatahyouka_i‚è‚İT
    End If
    
    If MOrahyouka_i‚è‚İT = "" Then
        ƒ‚ƒ‰•]‰¿ = 0
    Else
        ƒ‚ƒ‰•]‰¿ = MOrahyouka_i‚è‚İT
    End If
    
    If Keikenhyouka_i‚è‚İT = "" Then
        ŒoŒ±’l•]‰¿ = 0
    Else
        ŒoŒ±’l•]‰¿ = Keikenhyouka_i‚è‚İT
    End If
    
    
    
'    If êŠ–¼ = "" And Œ‚‘Ş”•]‰¿ = 0 And ¸‰s•]‰¿ = 0 And ¬Œ^•]‰¿ = 0 And ƒ‚ƒ‰•]‰¿ = 0 And ŒoŒ±’l•]‰¿ = 0 Then
'        Exit Sub
'    End If
    
    
    'êŠiLvğŒ‚ğæ“¾
    Dim êŠLvItem As ListItem
    Basyo_êŠiLv_Dic.RemoveAll
    Basyo_êŠiLv_Dic_Key = ""
    For Each êŠLvItem In êŠLv.ListItems
        If êŠLvItem.Checked = True Then
            Basyo_êŠiLv_Dic_Key = CStr(êŠLvItem.Text)
            If Not Basyo_êŠiLv_Dic.Exists(Basyo_êŠiLv_Dic_Key) Then
                Basyo_êŠiLv_Dic.Add Basyo_êŠiLv_Dic_Key, 1
            End If
        End If
    Next
    
    
    
    'Lv•\¦
    Dim WS As Worksheet
    Dim LastRow As Long
    Dim LastCol As Long
    
    Set WS = ThisWorkbook.Sheets("“Gë‚èƒ‹[ƒg’²¸ÀÑ")
    
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
        .ColumnHeaders.Add 1, "NUM", "‡‚", Width:=19
        For j = 1 To LastCol
            If j > 3 Then
                .ColumnHeaders.Add j + 1, "NUM" & j, WS.Cells(4, j).Value, Alignment:=lvwColumnRight   '2—ñ–ÚˆÈ~‚ÍAƒ}ƒXƒ^[‚Ì•\‚Ìƒwƒbƒ_—ñ”‚ÉˆË‘¶
            Else
                .ColumnHeaders.Add j + 1, "NUM" & j, WS.Cells(4, j).Value
            End If
        Next j
    End With
    
    
    Dim ƒ‹[ƒg–¼––”öw As String
    With Lv1
        For i = 8 To LastRow
            ƒ‹[ƒg–¼––”öw = Val(Right(WS.Cells(i, 1).Value, Len(WS.Cells(i, 1).Value) - InStrRev(WS.Cells(i, 1).Value, "_")))
            If ƒ‹[ƒg–¼––”ö_1CH.Value = True Then
                If ƒ‹[ƒg–¼––”öw = "1" Then
'If êŠ–¼ <> "" Then
'If WS.Cells(i, 2).Value = êŠ–¼ Then
                    If Basyo_êŠiLv_Dic(WS.Cells(i, 2).Value) Then
                        If WS.Cells(i, 5).Value >= Œ‚‘Ş”•]‰¿ And WS.Cells(i, 6).Value >= ¬Œ^•]‰¿ And _
                            WS.Cells(i, 7).Value >= ¸‰s•]‰¿ And WS.Cells(i, 8).Value >= ƒ‚ƒ‰•]‰¿ And _
                            WS.Cells(i, 9).Value >= ŒoŒ±’l•]‰¿ Then
                                Cnt = Cnt + 1
                                .ListItems.Add Text:=Cnt
                                For j = 1 To LastCol
                                    .ListItems(Cnt).SubItems(j) = WS.Cells(i, j).Value
                                Next j
                        End If
'End If
'                    Else
'                        If WS.Cells(i, 5).Value >= Œ‚‘Ş”•]‰¿ And WS.Cells(i, 6).Value >= ¬Œ^•]‰¿ And _
'                            WS.Cells(i, 7).Value >= ¸‰s•]‰¿ And WS.Cells(i, 8).Value >= ƒ‚ƒ‰•]‰¿ And _
'                            WS.Cells(i, 9).Value >= ŒoŒ±’l•]‰¿ Then
'                                Cnt = Cnt + 1
'                                .ListItems.Add Text:=Cnt
'                                For j = 1 To LastCol
'                                    .ListItems(Cnt).SubItems(j) = WS.Cells(i, j).Value
'                                Next j
'                        End If
                    End If
                End If  'ƒƒ‹[ƒg–¼––”öw = "1"
            Else    'ƒƒ‹[ƒg–¼––”ö_1CH.Value = True


'If êŠ–¼ <> "" Then
'If WS.Cells(i, 2).Value = êŠ–¼ Then
                If Basyo_êŠiLv_Dic(WS.Cells(i, 2).Value) Then
                    If WS.Cells(i, 5).Value >= Œ‚‘Ş”•]‰¿ And WS.Cells(i, 6).Value >= ¬Œ^•]‰¿ And _
                        WS.Cells(i, 7).Value >= ¸‰s•]‰¿ And WS.Cells(i, 8).Value >= ƒ‚ƒ‰•]‰¿ And _
                        WS.Cells(i, 9).Value >= ŒoŒ±’l•]‰¿ Then
                            Cnt = Cnt + 1
                            .ListItems.Add Text:=Cnt
                            For j = 1 To LastCol
                                .ListItems(Cnt).SubItems(j) = WS.Cells(i, j).Value
                            Next j
                    End If

'End If
'                Else
'                    If WS.Cells(i, 5).Value >= Œ‚‘Ş”•]‰¿ And WS.Cells(i, 6).Value >= ¬Œ^•]‰¿ And _
'                        WS.Cells(i, 7).Value >= ¸‰s•]‰¿ And WS.Cells(i, 8).Value >= ƒ‚ƒ‰•]‰¿ And _
'                        WS.Cells(i, 9).Value >= ŒoŒ±’l•]‰¿ Then
'                            Cnt = Cnt + 1
'                            .ListItems.Add Text:=Cnt
'                            For j = 1 To LastCol
'                                .ListItems(Cnt).SubItems(j) = WS.Cells(i, j).Value
'                            Next j
'                    End If
                End If

            End If  'ƒƒ‹[ƒg–¼––”ö_1CH.Value = True
        Next i
    End With


    
End Sub
