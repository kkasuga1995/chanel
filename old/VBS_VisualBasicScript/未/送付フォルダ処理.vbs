Option Explicit
'============================================================================
'①『フォルダ作成.VBS』実行
'②デスクトップ上の『送付フォルダ』内のファイルを、①で作成したフォルダに移動
'③完了
 
'	＜対象ファイルの条件＞
'		・カレントフォルダ内のファイル
'   ・拡張子がDAT									 (以下	条件①)
'   ・ファイル名(本体)が8桁である。(以下	条件②)(○IRAKUKAF.DAT ×IRAKUKAF_2021.DAT)
'
' ＜日付フォーマット＞
'   ファイル名(本体)の末尾に_YYYYMMDD_HHNNSSを付与する。
' 	【例】
'			ファイル名	：IRAKUKAF.DAT
'   	日付        ：2021年6月25日 13:05:33
'   	結果        ：IRAKUKAF_20210625_130533.DAT
'	＜処理の流れ＞
'  １．カレントフォルダ内の付与対象ファイルをリネームする 		(以下	処理①)
'  ２．リネームしたファイルを『過去分』フォルダに切り取りする (以下	処理②)
'  ３．終わり
'============================================================================
'①『フォルダ作成.VBS』実行
	'起動用のオブジェクトを生成
	Dim objWsh
	Set objWsh = WScript.CreateObject("WScript.Shell")

	'実行
	objWsh.Run "\\KPDSV1\Share\電算\800 春日作業フォルダー\空フォルダ\Sou_送付\フォルダ作成.vbs",,True



''	＜付与日付を取得する＞
'    Dim StrYear							'システム日付から年を抽出した値(YYYY)を格納する
'    Dim StrMonth            'システム日付から月を抽出した値(MM)を格納する
'    Dim StrDay              'システム日付から日を抽出した値(DD)を格納する
'    Dim StrHour             'システム日付から時を抽出した値(HH)を格納する
'    Dim StrMin              'システム日付から分を抽出した値(NN)を格納する
'    Dim StrSec             	'システム日付から秒を抽出した値(SS)を格納する
'    Dim AttachDate          'ファイルに付与する日付文字列を格納する
'
''	＜システム日付から、年、月、日、時、分、秒を抽出する＞
'    StrYear = Left(Now, 4): StrMonth = Mid(Now, 6, 2): StrDay = Mid(Now, 9, 2)
'    StrHour = Mid(Now, 12, 2): StrMin = Mid(Now, 15, 2): StrSec = Mid(Now, 18, 2)
''	＜抽出した年、月、日、時、分、秒を結合する＞
'    AttachDate = StrYear & StrMonth & StrDay & "_" & StrHour & StrMin & StrSec
'
'' ＜ファイル操作用オブジェクト定義と格納＞
'    Dim FSO									'FileSystemObjectのインスタンスを格納する
'    Dim FLtmp               'ループ内でファイルオブジェクトを都度格納する
'    Dim FolderObj						'処理対象のフォルダオブジェクトを格納する
'    Dim CurrentPath       	'処理対象のフォルダパスを格納する
'		Set FSO = WScript.CreateObject("Scripting.FileSystemObject")	'FileSystemObjectのインスタンスを生成し、FSOに格納
'   	Set FolderObj = FSO.GetFolder(CurrentPath)	'カレントフォルダのオブジェクトを格納		
'    CurrentPath = FSO.GetAbsolutePathName("./")	'本実行ファイルが置かれているフォルダ(カレントフォルダ)のパスを格納
'
'' ＜リネームとファイル移動＞
'		'カレントフォルダ内をループして、条件①、条件②に合致するファイルに対して、処理①と処理②を実行する
'    For Each FLtmp In FolderObj.Files
'			If FSO.GetExtensionName(FLtmp.Path) = "DAT" Then		'条件①
'      	If Len(FSO.GetBaseName(FLtmp.Path)) = 8 Then				'条件②	
'					FLtmp.Name = FSO.GetBaseName(FLtmp.Path) & "_" & AttachDate & ".DAT"	'処理①
'          FSO.MoveFile FLtmp.Path, Currentpath & "\過去分\"                     '処理②          
'        End If
'      End If
'    Next                                                                                   
'
''	＜格納したオブジェクトの破棄＞	
'    Set FSO = Nothing
'    Set FLtmp = Nothing