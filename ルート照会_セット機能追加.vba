セット登録(ルート)
	▼シート(セット登録_ルート)を追加                          NO ルート名　場所名
	▼シート(セット登録_ルート)                                0		1		2
	　　セット名｜連番 | ルート名｜場所名
	　　テストルート①　｜00001 | 六課の池_下_1_1 | モンド

	▼フォーム
		・コンボボックス(セット一覧)を追加 . 『新規』、『追加』、『照会』、『削除』、『編集』ボタン追加
┌────────────────────┬┬┐
│                       │ セット一覧CB ▼  ││
│                       └───────┘  ││
│                        新規　追加　照会   ││
│                        削除               ││
│                        セット編集         ││
│ ┌───────────────────┐││
│ │                                      │││
│ │                                      │││
│ │                                      │││
│ │                                      │││
│ │                                      │││
│ └───────────────────┘├┤
│─────────────────────┴┴──

	『新規』:リスト1件以上選択状態。

			 セット名入力メッセージボックス表示。(すでに存在するセット名ははじく)
			 セット登録

			 

	
	『追加』:リスト1件以上選択状態。
			 セット一覧ＣＢに登録済みセットが表示されている状態
			 (すでにそのセットに登録されているルートははじく)

			 メッセージボックス(通知：よろしいですか？ yes no)
			 セット一覧ＣＢにルート追加

	『照会』：セット一覧ＣＢに登録済みセットが表示されている状態

			 リストに、セット一覧ＣＢで指定されたセットのルートを表示

	『削除』：セット一覧ＣＢに登録済みセットが表示されている状態
	          リスト1件以上選択状態。

	         リスト選択中のルートを、セット一覧ＣＢで指定されているセットから削除。(存在チェック)

	『セット編集』：セット編集フォームを開く。登録されたセットを編集する。
			編集可能
				・ルートの順番
				・ルートの削除(一覧表示)
            	・ルートの追加(一覧表示)



-▼追加した処理-----------------------------------------------------------------------------------------------------------
Private Sub セット登録新規()
'    『新規』:リスト1件以上選択状態。
'
'             セット名入力メッセージボックス表示｡ (すでに存在するセット名ははじく)
'             セット登録
'配列定義
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
        MsgBox "リストを1件以上選択してください", vbInformation
        GoTo ext
    End If
    
    'シート登録内容：セット名｜連番 | ルート名｜場所名
    XR = LvselectCnt
    XC = 4
    ReDim X(1 To XR, 1 To XC) As String
    
    
    'セット名入力ダイアログ表示
    Dim SetName As String
    On Error Resume Next
        SetName = Application.InputBox("セット名を入力", Type:=2)
    On Error GoTo 0
    If SetName = "" Then
        GoTo ext
    End If
    
    '既に登録済みのセット名ははじく
    Dim WS As Worksheet
    Dim LastRow As Long
    Set WS = ThisWorkbook.Sheets("セット登録_ルート")
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).row
    Dim i As Long, j As Long, k As Long
    For i = 2 To LastRow
        If WS.Cells(i, 1).Value = SetName Then
            MsgBox "既に登録済みのセット名です。", vbExclamation
            GoTo ext
        End If
    Next i
    
    
    'シート登録内容を選択中リスト → 配列に転記
    For Each LvItem In Lv1.ListItems
        If LvItem.Selected = True Then
            XRc = XRc + 1
            X(XRc, 1) = SetName 'セット名
            X(XRc, 2) = CStr(Format(XRc, "00000")) '連番
            X(XRc, 3) = LvItem.SubItems(1) 'ルート名
            X(XRc, 4) = LvItem.SubItems(2) '場所名
        End If
    Next
    
    WS.Cells(LastRow + 1, 1).Resize(UBound(X, 1), UBound(X, 2)) = X
    
    '転記後、最終行の再取得
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).row
    
    'シート上でソート並び替え
    Call セット登録シート_ソート処理


ext:
    If err.Number <> 0 Then
        MsgBox "エラーが発生しました" & vbCrLf & "番号：" & err.Number & vbCrLf & "内容：" & err.Description
    End If
    Set WS = Nothing
End Sub
Private Sub セット登録_追加()
'    『追加』:リスト1件以上選択状態。
'             セット一覧CBに登録済みセットが表示されている状態
'             (すでにそのセットに登録されているルートははじく)
'
'             メッセージボックス(通知：よろしいですか？ yes no)
'             セット一覧CBにルート追加


'残タスク
'★セット一覧ＣＢをフォーム上に追加する


    'セット一覧CBに登録済みセットが表示されている状態
    If セット一覧CB.Value = "" Then
        MsgBox "セット一覧が未選択です", vbExclamation
        GoTo ext
    End If


    Dim X() As String
    Dim XR As Long, XC As Long
    Dim XRc As Long
    Dim LvselectCnt As Long
    Dim LvItem As ListItem
    
    'リスト1件以上選択状態。
    For Each LvItem In Lv1.ListItems
        If LvItem.Selected = True Then
            LvselectCnt = LvselectCnt + 1
        End If
    Next
    If LvselectCnt < 1 Then
        MsgBox "リストを1件以上選択してください", vbInformation
        GoTo ext
    End If
    
    
    ★★ここまで
    
    'すでにそのセットに登録されているルートははじく
    Dim WS As Worksheet
    Dim LastRow As Long
    Set WS = ThisWorkbook.Sheets("セット登録_ルート")
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).row
    Dim i As Long, j As Long, k As Long
     '連想配列に指定したセット名のルート一覧を格納
    Dim ルートDic As Object
    Dim ルートDicKey As String
    Dim ルートDicItem As String
    Set ルートDic = CreateObject("Scripting.Dictionary")
    
    
    
    For i = 2 To LastRow
        If WS.Cells(i, 1).Value = SetName Then
            MsgBox "既に登録済みのセット名です。", vbExclamation
            GoTo ext
        End If
    Next i
    


        
        
    
    
    'シート登録内容：セット名｜連番 | ルート名｜場所名
    XR = LvselectCnt
    XC = 4
    ReDim X(1 To XR, 1 To XC) As String
    
    
    'セット名入力ダイアログ表示
    Dim SetName As String
    On Error Resume Next
        SetName = Application.InputBox("セット名を入力", Type:=2)
    On Error GoTo 0
    If SetName = "" Then
        GoTo ext
    End If
    

    
    
    'シート登録内容を選択中リスト → 配列に転記
    For Each LvItem In Lv1.ListItems
        If LvItem.Selected = True Then
            XRc = XRc + 1
            X(XRc, 1) = SetName 'セット名
            X(XRc, 2) = CStr(Format(XRc, "00000")) '連番
            X(XRc, 3) = LvItem.SubItems(1) 'ルート名
            X(XRc, 4) = LvItem.SubItems(2) '場所名
        End If
    Next
    
    WS.Cells(LastRow + 1, 1).Resize(UBound(X, 1), UBound(X, 2)) = X
    
    '転記後、最終行の再取得
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).row
    
    'シート上でソート並び替え
    Call セット登録シート_ソート処理


ext:
    If err.Number <> 0 Then
        MsgBox "エラーが発生しました" & vbCrLf & "番号：" & err.Number & vbCrLf & "内容：" & err.Description
    End If
    Set WS = Nothing
End Sub

Private Sub セット登録シート_ソート処理()
    Dim WS As Worksheet
    Dim LastRow As Long
    Set WS = ThisWorkbook.Sheets("セット登録_ルート")
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
        MsgBox "エラーが発生しました" & vbCrLf & "番号：" & err.Number & vbCrLf & "内容：" & err.Description
    End If
    Set WS = Nothing


End Sub