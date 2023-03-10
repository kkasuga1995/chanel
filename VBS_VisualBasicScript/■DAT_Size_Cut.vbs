'====================================================================================
'
'■ CSVファイルの先頭行に特定の文字列行を挿入する処理 
'
'【VBS 機能追加依頼】
'  目的    ：変換マクロ実行にあたり、最小サイズのデータで第一段階の検証を素早く行う為。
'  仕様概要：DATファイルに対して、上位100レコードほどに縮小したファイルを再生成する。
'            ヘッダは残す。
'            最終行のごみは残さない。
'            レコードサイズ(上位xレコード)は、変数で逐次変更可能。
'            生成後のファイル名は XXXXXXXX.DAT
'            指定レコード数(上位100レコード)以下のファイルは処理から除外。
'            バックアップは不要。上書き
'  実行方法：対象ファイルが存在するフォルダに.VBSファイルを置いて実行する。
'            他フォルダへは干渉させない。
'
'  対象のファイル：
'      ①拡張子=DAT OR dat
'      ②本体名が8文字
'      ③カレントディレクトリ(vbsファイルが置かれているフォルダ)内のファイル
'      ④除外するファイルを変数で指定可能とする。(BUHTANF2.DATは除外)

'====================================================================================

'----------- 定数の作成
Const ForReading = 1                                 '読取フラグ
Const ForWriting = 2                                 '書込フラグ
Const ForAppending = 8                               '追記フラグ

Dim i, j
Dim objFso, objFolder, objFile
Dim strFileName
Dim strFileBaseName
Dim strCurrentPath
Dim strPath
Dim strMyName
Dim strExt


'===================================★いじるのはここだけ★=================================
Const CutRecordCnt = 100                             '上位Nレコードまでデータを保存する
Dim ExclusionFile(1)                                 '対象外とするファイル数＝要素数とする（ExclusionFile(1)：2個のファイルを除外)
ExclusionFile(0) = "BUHTANF2"                        '対象外とするファイル名を指定する
ExclusionFile(1) = "RZANIJF1"
'==========================================================================================


'ファイルシステムオブジェクト作成
Set objFso = CreateObject("Scripting.FileSystemobject")

If MsgBox("カレントファイルのデータサイズを変更しますが、よろしいですか？", vbYesNo + vbInfomation, "確認") = vbNo Then
	WScript.Quit '処理終了
End If

'カレントフォルダのパス
strCurrentPath = objFso.GetAbsolutePathName(".")  

'カレントフォルダのオブジェクトをセット
Set objFolder = objFso.GetFolder(".\")

'自スクリプト名を取得する
strMyName = WScript.ScriptName

'フォルダ内のファイル名を取得
For Each objFile In objFolder.Files

    '取得したファイルのフルパスを保持
	strPath = strCurrentPath & "\" & objFile.Name

	'テキストの行数を確認
	Set objRead = objFso.OpenTextFile(strPath , ForReading)    '読取モードでテキストを開く
	objRead.ReadAll                                            '全部読むことで最終行へ移動
	intLine = objRead.Line                                     '現在の行数を確認
	objRead.Close                                              '読取モード閉じる

	strFileName = objFile.Name
	strFileBaseName = objFso.getBaseName(strFileName)

	'拡張子を取得
	strExt = UCase(objFso.GetExtensionName(strPath))

	If strMyName = strFileName Then
		'MsgBox "自分は除外"
	ElseIf strExt <> "DAT" Then
		'MsgBox "拡張子がDATではない"
	ElseIf Len(strFileBaseName) <> 8 Then
		'MsgBox "ファイル名が8文字ではないため除外"
	Else

		Dim vFlg
		vFlg = True
		
		For i = 0 to Ubound(ExclusionFile)
			'取得したファイルが読込対象外だった場合は処理を抜ける
			If strFileBaseName = ExclusionFile(i) then
				'MsgBox "処理を抜ける"
				vFlg = False
				Exit For
			End if
		Next

		if vFlg = True Then
			'MsgBox "処理を開始[" & strFileBaseName &  "]"

			Dim WritingText                                           '書込用の文字列（省略可）
			WritingText = ""

			Set objRead2 = objFso.OpenTextFile(strPath , ForReading)  '読取モードでテキストを開く

			row = 1                                                   '行数の確認用の数値
			Do Until objRead2.AtEndOfStream = True                    '終了行まで繰り返し

				If row > CutRecordCnt Then                                '最大レコード行まで来たら処理を抜ける
					Exit Do
				End If

				Dim ReadingText 
				ReadingText = objRead2.ReadLine
				
				if Len(ReadingText) = 1  Then '空文字行は処理対象外とする
					'msgbox "NULL"
				Else
					WritingText = WritingText & ReadingText & vbCrLf        '1行読み取り、書込用の文字列に追加
					row = row + 1                                         '読み取った行数を1増やす
				end IF
								
			Loop

			objRead2.Close                                            '読取モード閉じる

			Set objWriting = objFso.OpenTextFile(strPath , ForWriting)          '書込モードでテキストを開く
			objWriting.Write WritingText                             '書込用の文字列値を一気に書込み
			objWriting.Close                                    '書込モード閉じる

		End if

	End if




Next 


'----------- 完了メッセージ
MsgBox "変換完了!!"
