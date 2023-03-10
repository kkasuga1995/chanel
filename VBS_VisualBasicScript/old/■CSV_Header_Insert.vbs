'====================================================================================
'
'■ CSVファイルの先頭行に特定の文字列行を挿入する処理 
'
'元ファイル        変換先ファイル(挿入後)
'[ZSETJIF2.DAT] → [ZSETJIF2_20220202_142420.DAT]（ヘッダなし）
'                   └ [ZSETJIF2.DAT]（ヘッダ有り）
'
'====================================================================================

'----------- 定数の作成
Const ForReading = 1                                 '読取フラグ
Const ForWriting = 2                                 '書込フラグ
Const ForAppending = 8                               '追記フラグ
Const InsertLine = 1                                 '文字列の挿入行

Dim i, j

Dim InsertFilePath                                   '処理するテキストファイルのパス

'ファイルシステムオブジェクト作成
Set objFso = CreateObject("Scripting.FileSystemobject")


If MsgBox("CSVファイルにヘッダー行を追加しますが、よろしいですか？", vbYesNo + vbInfomation, "確認") = vbNo Then
	WScript.Quit '処理終了
End If


'ルートフォルダのパス
InsertFilePath = objFso.GetAbsolutePathName(".")

'対象ファイルの追加・変更する場合、以下を適宜訂正する。＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'Dim InsertFileName(1) と Dim InsertText(1) の要素数( (1)の部分 )を対象ファイル分変更する。(2ファイルであれば[1])
'InsertFileName(N) に、追加ファイル名を定義する。(InsertFileName(N)= "XXXXXXXXX.DAT")
Dim InsertFileName(1)                                '処理するテキストファイルの名前 ：対象ファイル分、要素数を加算する
Dim InsertText(1)                                    '挿入する文字列                 ：対象ファイル分、要素数を加算する

InsertFileName(0) = "ZSETJIF2.DAT"
InsertText(0) = """作業日"",""伝票番号"",""工程"",""行番"",""発注番号"",""材料区分"",""梱包番号"",""分割"",""投入枚数"",""原板使用禁止"",""原板保管場所"",""支給区分"",""板仕様"",""切断場所"",""元得意先コード(1)"",""元製品コード(1)"",""元製品区分(1)"",""元部品(1)"",""元取数(1)"",""元部品枚数(1)"",""元得意先コード(2)"",""元製品コード(2)"",""元製品区分(2)"",""元部品(2)"",""元取数(2)"",""元部品枚数(2)"",""元得意先コード(3)"",""元製品コード(3)"",""元製品区分(3)"",""元部品(3)"",""元取数(3)"",""元部品枚数(3)"",""元得意先コード(4)"",""元製品コード(4)"",""元製品区分(4)"",""元部品(4)"",""元取数(4)"",""元部品枚数(4)"",""元得意先コード(5)"",""元製品コード(5)"",""元製品区分(5)"",""元部品(5)"",""元取数(5)"",""元部品枚数(5)"",""元得意先コード(6)"",""元製品コード(6)"",""元製品区分(6)"",""元部品(6)"",""元取数(6)"",""元部品枚数(6)"",""得意先コード(1)"",""製品コード(1)"",""製品区分(1)"",""部品(1)"",""部品枚数(1)"",""取数(1)"",""投入数(1)"",""切断不良(1)"",""材料不良(1)"",""印刷不良(1)"",""テスト(1)"",""完成数(1)"",""得意先コード(2)"",""製品コード(2)"",""製品区分(2)"",""部品(2)"",""部品枚数(2)"",""取数(2)"",""投入数(2)"",""切断不良(2)"",""材料不良(2)"",""印刷不良(2)"",""テスト(2)"",""完成数(2)"",""得意先コード(3)"",""製品コード(3)"",""製品区分(3)"",""部品(3)"",""部品枚数(3)"",""取数(3)"",""投入数(3)"",""切断不良(3)"",""材料不良(3)"",""印刷不良(3)"",""テスト(3)"",""完成数(3)"",""得意先コード(4)"",""製品コード(4)"",""製品区分(4)"",""部品(4)"",""部品枚数(4)"",""取数(4)"",""投入数(4)"",""切断不良(4)"",""材料不良(4)"",""印刷不良(4)"",""テスト(4)"",""完成数(4)"",""得意先コード(5)"",""製品コード(5)"",""製品区分(5)"",""部品(5)"",""部品枚数(5)"",""取数(5)"",""投入数(5)"",""切断不良(5)"",""材料不良(5)"",""印刷不良(5)"",""テスト(5)"",""完成数(5)"",""得意先コード(6)"",""製品コード(6)"",""製品区分(6)"",""部品(6)"",""部品枚数(6)"",""取数(6)"",""投入数(6)"",""切断不良(6)"",""材料不良(6)"",""印刷不良(6)"",""テスト(6)"",""完成数(6)"",""保管場所(1)"",""分割２(1)"",""使用禁止(1)"",""行２(1)"",""部品２(1)"",""取数２(1)"",""数量２(1)"",""行２(2)"",""部品２(2)"",""取数２(2)"",""数量２(2)"",""行２(3)"",""部品２(3)"",""取数２(3)"",""数量２(3)"",""行２(4)"",""部品２(4)"",""取数２(4)"",""数量２(4)"",""行２(5)"",""部品２(5)"",""取数２(5)"",""数量２(5)"",""行２(6)"",""部品２(6)"",""取数２(6)"",""数量２(6)"",""保管場所(2)"",""分割２(2)"",""使用禁止(2)"",""行２(7)"",""部品２(7)"",""取数２(7)"",""数量２(7)"",""行２(8)"",""部品２(8)"",""取数２(8)"",""数量２(8)"",""行２(9)"",""部品２(9)"",""取数２(9)"",""数量２(9)"",""行２(10)"",""部品２(10)"",""取数２(10)"",""数量２(10)"",""行２(11)"",""部品２(11)"",""取数２(11)"",""数量２(11)"",""行２(12)"",""部品２(12)"",""取数２(12)"",""数量２(12)"",""保管場所(3)"",""分割２(3)"",""使用禁止(3)"",""行２(13)"",""部品２(13)"",""取数２(13)"",""数量２(13)"",""行２(14)"",""部品２(14)"",""取数２(14)"",""数量２(14)"",""行２(15)"",""部品２(15)"",""取数２(15)"",""数量２(15)"",""行２(16)"",""部品２(16)"",""取数２(16)"",""数量２(16)"",""行２(17)"",""部品２(17)"",""取数２(17)"",""数量２(17)"",""行２(18)"",""部品２(18)"",""取数２(18)"",""数量２(18)"",""保管場所(4)"",""分割２(4)"",""使用禁止(4)"",""行２(19)"",""部品２(19)"",""取数２(19)"",""数量２(19)"",""行２(20)"",""部品２(20)"",""取数２(20)"",""数量２(20)"",""行２(21)"",""部品２(21)"",""取数２(21)"",""数量２(21)"",""行２(22)"",""部品２(22)"",""取数２(22)"",""数量２(22)"",""行２(23)"",""部品２(23)"",""取数２(23)"",""数量２(23)"",""行２(24)"",""部品２(24)"",""取数２(24)"",""数量２(24)"",""保管場所(5)"","" 分割２(5)"",""使用禁止(5)"",""行２(25)"",""部品２(25)"",""取数２(25)"",""数量２(25)"",""行２(26)"",""部品２(26)"",""取数２(26)"",""数量２(26)"",""行２(27)"",""部品２(27)"",""取数２(27)"",""数量２(27)"",""行２(28)"",""部品２(28)"",""取数２(28)"",""数量２(28)"",""行２(29)"",""部品２(29)"",""取数２(29)"",""数量２(29)"",""行２(30)"",""部品２(30)"",""取数２(30)"",""数量２(30)"",""摘要１"",""摘要２"",""処理者"",""処理日"",""処理時刻"",""端末ＩＤ"""

InsertFileName(1) = "RZANIJF1.DAT"
InsertText(1) = """作業日"",""伝票番号"",""伝区"",""行番"",""製造ロットＮＯ"",""加工場所"",""二次加工"",""入庫場所"",""入庫得意先コード"",""入庫製品コード"",""製品完成数"",""部品得意先コード(1)"",""部品製品コード(1)"",""部品ロットＮＯ(1)"",""部品投入数(1)"",""部品印刷不良(1)"",""印刷不良判断(1)"",""部品加工不良(1)"",""加工内訳キズ(1)"",""加工判断キズ(1)"",""加工内訳あたり(1)"",""加工判断あたり(1)"",""部品作業不良(1)"",""作業内訳キズ(1)"",""作業判断キズ(1)"",""作業内訳あたり(1)"",""作業判断あたり(1)"",""部品差替え(1)"",""部品得意先コード(2)"",""部品製品コード(2)"",""部品ロットＮＯ(2)"",""部品投入数(2)"",""部品印刷不良(2)"",""印刷不良判断(2)"",""部品加工不良(2)"",""加工内訳キズ(2)"",""加工判断キズ(2)"",""加工内訳あたり(2)"",""加工判断あたり(2)"",""部品作業不良(2)"",""作業内訳キズ(2)"",""作業判断キズ(2)"",""作業内訳あたり(2)"",""作業判断あたり(2)"",""部品差替え(2)"",""部品得意先コード(3)"",""部品製品コード(3)"",""部品ロットＮＯ(3)"",""部品投入数(3)"",""部品印刷不良(3)"",""印刷不良判断(3)"",""部品加工不良(3)"",""加工内訳キズ(3)"",""加工判断キズ(3)"",""加工内訳あたり(3)"",""加工判断あたり(3)"",""部品作業不良(3)"",""作業内訳キズ(3)"",""作業判断キズ(3)"",""作業内訳あたり(3)"",""作業判断あたり(3)"",""部品差替え(3)"",""部品得意先コード(4)"",""部品製品コード(4)"",""部品ロットＮＯ(4)"",""部品投入数(4)"",""部品印刷不良(4)"",""印刷不良判断(4)"",""部品加工不良(4)"",""加工内訳キズ(4)"",""加工判断キズ(4)"",""加工内訳あたり(4)"",""加工判断あたり(4)"",""部品作業不良(4)"",""作業内訳キズ(4)"",""作業判断キズ(4)"",""作業内訳あたり(4)"",""作業判断あたり(4)"",""部品差替え(4)"",""部品得意先コード(5)"",""部品製品コード(5)"",""部品ロットＮＯ(5)"",""部品投入数(5)"",""部品印刷不良(5)"",""印刷不良判断(5)"",""部品加工不良(5)"",""加工内訳キズ(5)"",""加工判断キズ(5)"",""加工内訳あたり(5)"",""加工判断あたり(5)"",""部品作業不良(5)"",""作業内訳キズ(5)"",""作業判断キズ(5)"",""作業内訳あたり(5)"",""作業判断あたり(5)"",""部品差替え(5)"",""部品得意先コード(6)"",""部品製品コード(6)"",""部品ロットＮＯ(6)"",""部品投入数(6)"",""部品印刷不良(6)"",""印刷不良判断(6)"",""部品加工不良(6)"",""加工内訳キズ(6)"",""加工判断キズ(6)"",""加工内訳あたり(6)"",""加工判断あたり(6)"",""部品作業不良(6)"",""作業内訳キズ(6)"",""作業判断キズ(6)"",""作業内訳あたり(6)"",""作業判断あたり(6)"",""部品差替え(6)"",""部品得意先コード(7)"",""部品製品コード(7)"",""部品ロットＮＯ(7)"",""部品投入数(7)"",""部品印刷不良(7)"",""印刷不良判断(7)"",""部品加工不良(7)"",""加工内訳キズ(7)"",""加工判断キズ(7)"",""加工内訳あたり(7)"",""加工判断あたり(7)"",""部品作業不良(7)"",""作業内訳キズ(7)"",""作業判断キズ(7)"",""作業内訳あたり(7)"",""作業判断あたり(7)"",""部品差替え(7)"",""部品得意先コード(8)"",""部品製品コード(8)"",""部品ロットＮＯ(8)"",""部品投入数(8)"",""部品印刷不良(8)"",""印刷不良判断(8)"",""部品加工不良(8)"",""加工内訳キズ(8)"",""加工判断キズ(8)"",""加工内訳あたり(8)"",""加工判断あたり(8)"",""部品作業不良(8)"",""作業内訳キズ(8)"",""作業判断キズ(8)"",""作業内訳あたり(8)"",""作業判断あたり(8)"",""部品差替え(8)"",""部品得意先コード(9)"",""部品製品コード(9)"",""部品ロットＮＯ(9)"",""部品投入数(9)"",""部品印刷不良(9)"",""印刷不良判断(9)"",""部品加工不良(9)"",""加工内訳キズ(9)"",""加工判断キズ(9)"",""加工内訳あたり(9)"",""加工判断あたり(9)"",""部品作業不良(9)"",""作業内訳キズ(9)"",""作業判断キズ(9)"",""作業内訳あたり(9)"",""作業判断あたり(9)"",""部品差替え(9)"",""部品得意先コード(10)"",""部品製品コード(10)"",""部品ロットＮＯ(10)"",""部品投入数(10)"",""部品印刷不良(10)"",""印刷不良判断(10)"",""部品加工不良(10)"",""加工内訳キズ(10)"",""加工判断キズ(10)"",""加工内訳あたり(10)"",""加工判断あたり(10)"",""部品作業不良(10)"",""作業内訳キズ(10)"",""作業判断キズ(10)"",""作業内訳あたり(10)"",""作業判断あたり(10)"",""部品差替え(10)"",""部品得意先コード(11)"",""部品製品コード(11)"",""部品ロットＮＯ(11)"",""部品投入数(11)"",""部品印刷不良(11)"",""印刷不良判断(11)"",""部品加工不良(11)"",""加工内訳キズ(11)"",""加工判断キズ(11)"",""加工内訳あたり(11)"",""加工判断あたり(11)"",""部品作業不良(11)"",""作業内訳キズ(11)"",""作業判断キズ(11)"",""作業内訳あたり(11)"",""作業判断あたり(11)"",""部品差替え(11)"",""部品得意先コード(12)"",""部品製品コード(12)"",""部品ロットＮＯ(12)"",""部品投入数(12)"",""部品印刷不良(12)"",""印刷不良判断(12)"",""部品加工不良(12)"",""加工内訳キズ(12)"",""加工判断キズ(12)"",""加工内訳あたり(12)"",""加工判断あたり(12)"",""部品作業不良(12)"",""作業内訳キズ(12)"",""作業判断キズ(12)"",""作業内訳あたり(12)"",""作業判断あたり(12)"",""部品差替え(12)"",""開始時間(1)"",""終了時間(1)"",""正社員男性(1)"",""正社員女性(1)"",""パート(1)"",""開始時間(2)"",""終了時間(2)"",""正社員男性(2)"",""正社員女性(2)"",""パート(2)"",""開始時間(3)"",""終了時間(3)"",""正社員男性(3)"",""正社員女性(3)"",""パート(3)"",""開始時間(4)"",""終了時間(4)"",""正社員男性(4)"",""正社員女性(4)"",""パート(4)"",""総工数"",""製造工賃"",""工程(1)"",""担当者(1)"",""担当者名(1)"",""工程(2)"",""担当者(2)"",""担当者名(2)"",""工程(3)"",""担当者(3)"",""担当者名(3)"",""工程(4)"",""担当者(4)"",""担当者名(4)"",""工程(5)"",""担当者(5)"",""担当者名(5)"",""工程(6)"",""担当者(6)"",""担当者名(6)"",""工程(7)"",""担当者(7)"",""担当者名(7)"",""工程(8)"",""担当者(8)"",""担当者名(8)"",""工程(9)"",""担当者(9)"",""担当者名(9)"",""工程(10)"",""担当者(10)"",""担当者名(10)"",""工程(11)"",""担当者(11)"",""担当者名(11)"",""工程(12)"",""担当者(12)"",""担当者名(12)"",""製缶不良コード(1)"",""製缶不良名(1)"",""製缶責任部門(1)"",""製缶胴不良１(1)"",""製缶蓋不良１(1)"",""製缶請求判断１(1)"",""製缶胴不良２(1)"",""製缶蓋不良２(1)"",""製缶請求判断２(1)"",""製缶不良コード(2)"",""製缶不良名(2)"",""製缶責任部門(2)"",""製缶胴不良１(2)"",""製缶蓋不良１(2)"",""製缶請求判断１(2)"",""製缶胴不良２(2)"",""製缶蓋不良２(2)"",""製缶請求判断２(2)"",""製缶不良コード(3)"",""製缶不良名(3)"",""製缶責任部門(3)"",""製缶胴不良１(3)"",""製缶蓋不良１(3)"",""製缶請求判断１(3)"",""製缶胴不良２(3)"",""製缶蓋不良２(3)"",""製缶請求判断２(3)"",""製缶不良コード(4)"",""製缶不良名(4)"",""製缶責任部門(4)"",""製缶胴不良１(4)"",""製缶蓋不良１(4)"",""製缶請求判断１(4)"",""製缶胴不良２(4)"",""製缶蓋不良２(4)"",""製缶請求判断２(4)"",""製缶不良コード(5)"",""製缶不良名(5)"",""製缶責任部門(5)"",""製缶胴不良１(5)"",""製缶蓋不良１(5)"",""製缶請求判断１(5)"",""製缶胴不良２(5)"",""製缶蓋不良２(5)"",""製缶請求判断２(5)"",""製缶不良コード(6)"",""製缶不良名(6)"",""製缶責任部門(6)"",""製缶胴不良１(6)"",""製缶蓋不良１(6)"",""製缶請求判断１(6)"",""製缶胴不良２(6)"",""製缶蓋不良２(6)"",""製缶請求判断２(6)"",""製缶不良コード(7)"",""製缶不良名(7)"",""製缶責任部門(7)"",""製缶胴不良１(7)"",""製缶蓋不良１(7)"",""製缶請求判断１(7)"",""製缶胴不良２(7)"",""製缶蓋不良２(7)"",""製缶請求判断２(7)"",""製缶不良コード(8)"",""製缶不良名(8)"",""製缶責任部門(8)"",""製缶胴不良１(8)"",""製缶蓋不良１(8)"",""製缶請求判断１(8)"",""製缶胴不良２(8)"",""製缶蓋不良２(8)"",""製缶請求判断２(8)"",""作業不良コード(1)"",""作業不良名(1)"",""作業責任部門(1)"",""作業胴不良１(1)"",""作業蓋不良１(1)"",""作業請求判断１(1)"",""作業胴不良２(1)"",""作業蓋不良２(1)"",""作業請求判断２(1)"",""作業不良コード(2)"",""作業不良名(2)"",""作業責任部門(2)"",""作業胴不良１(2)"",""作業蓋不良１(2)"",""作業請求判断１(2)"",""作業胴不良２(2)"",""作業蓋不良２(2)"",""作業請求判断２(2)"",""作業不良コード(3)"",""作業不良名(3)"",""作業責任部門(3)"",""作業胴不良１(3)"",""作業蓋不良１(3)"",""作業請求判断１(3)"",""作業胴不良２(3)"",""作業蓋不良２(3)"",""作業請求判断２(3)"",""作業不良コード(4)"",""作業不良名(4)"",""作業責任部門(4)"",""作業胴不良１(4)"",""作業蓋不良１(4)"",""作業請求判断１(4)"",""作業胴不良２(4)"",""作業蓋不良２(4)"",""作業請求判断２(4)"",""作業不良コード(5)"",""作業不良名(5)"",""作業責任部門(5)"",""作業胴不良１(5)"",""作業蓋不良１(5)"",""作業請求判断１(5)"",""作業胴不良２(5)"",""作業蓋不良２(5)"",""作業請求判断２(5)"",""作業不良コード(6)"",""作業不良名(6)"",""作業責任部門(6)"",""作業胴不良１(6)"",""作業蓋不良１(6)"",""作業請求判断１(6)"",""作業胴不良２(6)"",""作業蓋不良２(6)"",""作業請求判断２(6)"",""胴のみ１"",""蓋のみ１"",""胴のみ２"",""蓋のみ２"",""請求保留Ｆ"",""請求処理日"",""請求済みＦ"",""入力者"",""入力日"",""端末ＩＤ"""

'例 【REZAIRFL.DATを追加する場合】
'Dim InsertFileName(2)    (1) → (2)★
'Dim InsertText(2)        (1) → (2)★
'
'InsertFileName(0) = "ZSETJIF2.DAT"
'InsertText(0) = """作業日"",""伝票番号"",""工程"",""行番"",""発注番号"",""材料区分"",""梱包番号"",""分割"",""投入枚数"",""原板使用禁止"",""原板保管場所"",""支給区分"",""板仕様"",""切断場所"",""元得意先コード(1)"",""元製品コード(1)"",""元製品区分(1)"",""元部品(1)"",""元取数(1)"",""元部品枚数(1)"",""元得意先コード(2)"",""元製品コード(2)"",""元製品区分(2)"",""元部品(2)"",""元取数(2)"",""元部品枚数(2)"",""元得意先コード(3)"",""元製品コード(3)"",""元製品区分(3)"",""元部品(3)"",""元取数(3)"",""元部品枚数(3)"",""元得意先コード(4)"",""元製品コード(4)"",""元製品区分(4)"",""元部品(4)"",""元取数(4)"",""元部品枚数(4)"",""元得意先コード(5)"",""元製品コード(5)"",""元製品区分(5)"",""元部品(5)"",""元取数(5)"",""元部品枚数(5)"",""元得意先コード(6)"",""元製品コード(6)"",""元製品区分(6)"",""元部品(6)"",""元取数(6)"",""元部品枚数(6)"",""得意先コード(1)"",""製品コード(1)"",""製品区分(1)"",""部品(1)"",""部品枚数(1)"",""取数(1)"",""投入数(1)"",""切断不良(1)"",""材料不良(1)"",""印刷不良(1)"",""テスト(1)"",""完成数(1)"",""得意先コード(2)"",""製品コード(2)"",""製品区分(2)"",""部品(2)"",""部品枚数(2)"",""取数(2)"",""投入数(2)"",""切断不良(2)"",""材料不良(2)"",""印刷不良(2)"",""テスト(2)"",""完成数(2)"",""得意先コード(3)"",""製品コード(3)"",""製品区分(3)"",""部品(3)"",""部品枚数(3)"",""取数(3)"",""投入数(3)"",""切断不良(3)"",""材料不良(3)"",""印刷不良(3)"",""テスト(3)"",""完成数(3)"",""得意先コード(4)"",""製品コード(4)"",""製品区分(4)"",""部品(4)"",""部品枚数(4)"",""取数(4)"",""投入数(4)"",""切断不良(4)"",""材料不良(4)"",""印刷不良(4)"",""テスト(4)"",""完成数(4)"",""得意先コード(5)"",""製品コード(5)"",""製品区分(5)"",""部品(5)"",""部品枚数(5)"",""取数(5)"",""投入数(5)"",""切断不良(5)"",""材料不良(5)"",""印刷不良(5)"",""テスト(5)"",""完成数(5)"",""得意先コード(6)"",""製品コード(6)"",""製品区分(6)"",""部品(6)"",""部品枚数(6)"",""取数(6)"",""投入数(6)"",""切断不良(6)"",""材料不良(6)"",""印刷不良(6)"",""テスト(6)"",""完成数(6)"",""保管場所(1)"",""分割２(1)"",""使用禁止(1)"",""行２(1)"",""部品２(1)"",""取数２(1)"",""数量２(1)"",""行２(2)"",""部品２(2)"",""取数２(2)"",""数量２(2)"",""行２(3)"",""部品２(3)"",""取数２(3)"",""数量２(3)"",""行２(4)"",""部品２(4)"",""取数２(4)"",""数量２(4)"",""行２(5)"",""部品２(5)"",""取数２(5)"",""数量２(5)"",""行２(6)"",""部品２(6)"",""取数２(6)"",""数量２(6)"",""保管場所(2)"",""分割２(2)"",""使用禁止(2)"",""行２(7)"",""部品２(7)"",""取数２(7)"",""数量２(7)"",""行２(8)"",""部品２(8)"",""取数２(8)"",""数量２(8)"",""行２(9)"",""部品２(9)"",""取数２(9)"",""数量２(9)"",""行２(10)"",""部品２(10)"",""取数２(10)"",""数量２(10)"",""行２(11)"",""部品２(11)"",""取数２(11)"",""数量２(11)"",""行２(12)"",""部品２(12)"",""取数２(12)"",""数量２(12)"",""保管場所(3)"",""分割２(3)"",""使用禁止(3)"",""行２(13)"",""部品２(13)"",""取数２(13)"",""数量２(13)"",""行２(14)"",""部品２(14)"",""取数２(14)"",""数量２(14)"",""行２(15)"",""部品２(15)"",""取数２(15)"",""数量２(15)"",""行２(16)"",""部品２(16)"",""取数２(16)"",""数量２(16)"",""行２(17)"",""部品２(17)"",""取数２(17)"",""数量２(17)"",""行２(18)"",""部品２(18)"",""取数２(18)"",""数量２(18)"",""保管場所(4)"",""分割２(4)"",""使用禁止(4)"",""行２(19)"",""部品２(19)"",""取数２(19)"",""数量２(19)"",""行２(20)"",""部品２(20)"",""取数２(20)"",""数量２(20)"",""行２(21)"",""部品２(21)"",""取数２(21)"",""数量２(21)"",""行２(22)"",""部品２(22)"",""取数２(22)"",""数量２(22)"",""行２(23)"",""部品２(23)"",""取数２(23)"",""数量２(23)"",""行２(24)"",""部品２(24)"",""取数２(24)"",""数量２(24)"",""保管場所(5)"","" 分割２(5)"",""使用禁止(5)"",""行２(25)"",""部品２(25)"",""取数２(25)"",""数量２(25)"",""行２(26)"",""部品２(26)"",""取数２(26)"",""数量２(26)"",""行２(27)"",""部品２(27)"",""取数２(27)"",""数量２(27)"",""行２(28)"",""部品２(28)"",""取数２(28)"",""数量２(28)"",""行２(29)"",""部品２(29)"",""取数２(29)"",""数量２(29)"",""行２(30)"",""部品２(30)"",""取数２(30)"",""数量２(30)"",""摘要１"",""摘要２"",""処理者"",""処理日"",""処理時刻"",""端末ＩＤ"""
'
'InsertFileName(1) = "RZANIJF1.DAT"
'InsertText(1) = """作業日"",""伝票番号"",""伝区"",""行番"",""製造ロットＮＯ"",""加工場所"",""二次加工"",""入庫場所"",""入庫得意先コード"",""入庫製品コード"",""製品完成数"",""部品得意先コード(1)"",""部品製品コード(1)"",""部品ロットＮＯ(1)"",""部品投入数(1)"",""部品印刷不良(1)"",""印刷不良判断(1)"",""部品加工不良(1)"",""加工内訳キズ(1)"",""加工判断キズ(1)"",""加工内訳あたり(1)"",""加工判断あたり(1)"",""部品作業不良(1)"",""作業内訳キズ(1)"",""作業判断キズ(1)"",""作業内訳あたり(1)"",""作業判断あたり(1)"",""部品差替え(1)"",""部品得意先コード(2)"",""部品製品コード(2)"",""部品ロットＮＯ(2)"",""部品投入数(2)"",""部品印刷不良(2)"",""印刷不良判断(2)"",""部品加工不良(2)"",""加工内訳キズ(2)"",""加工判断キズ(2)"",""加工内訳あたり(2)"",""加工判断あたり(2)"",""部品作業不良(2)"",""作業内訳キズ(2)"",""作業判断キズ(2)"",""作業内訳あたり(2)"",""作業判断あたり(2)"",""部品差替え(2)"",""部品得意先コード(3)"",""部品製品コード(3)"",""部品ロットＮＯ(3)"",""部品投入数(3)"",""部品印刷不良(3)"",""印刷不良判断(3)"",""部品加工不良(3)"",""加工内訳キズ(3)"",""加工判断キズ(3)"",""加工内訳あたり(3)"",""加工判断あたり(3)"",""部品作業不良(3)"",""作業内訳キズ(3)"",""作業判断キズ(3)"",""作業内訳あたり(3)"",""作業判断あたり(3)"",""部品差替え(3)"",""部品得意先コード(4)"",""部品製品コード(4)"",""部品ロットＮＯ(4)"",""部品投入数(4)"",""部品印刷不良(4)"",""印刷不良判断(4)"",""部品加工不良(4)"",""加工内訳キズ(4)"",""加工判断キズ(4)"",""加工内訳あたり(4)"",""加工判断あたり(4)"",""部品作業不良(4)"",""作業内訳キズ(4)"",""作業判断キズ(4)"",""作業内訳あたり(4)"",""作業判断あたり(4)"",""部品差替え(4)"",""部品得意先コード(5)"",""部品製品コード(5)"",""部品ロットＮＯ(5)"",""部品投入数(5)"",""部品印刷不良(5)"",""印刷不良判断(5)"",""部品加工不良(5)"",""加工内訳キズ(5)"",""加工判断キズ(5)"",""加工内訳あたり(5)"",""加工判断あたり(5)"",""部品作業不良(5)"",""作業内訳キズ(5)"",""作業判断キズ(5)"",""作業内訳あたり(5)"",""作業判断あたり(5)"",""部品差替え(5)"",""部品得意先コード(6)"",""部品製品コード(6)"",""部品ロットＮＯ(6)"",""部品投入数(6)"",""部品印刷不良(6)"",""印刷不良判断(6)"",""部品加工不良(6)"",""加工内訳キズ(6)"",""加工判断キズ(6)"",""加工内訳あたり(6)"",""加工判断あたり(6)"",""部品作業不良(6)"",""作業内訳キズ(6)"",""作業判断キズ(6)"",""作業内訳あたり(6)"",""作業判断あたり(6)"",""部品差替え(6)"",""部品得意先コード(7)"",""部品製品コード(7)"",""部品ロットＮＯ(7)"",""部品投入数(7)"",""部品印刷不良(7)"",""印刷不良判断(7)"",""部品加工不良(7)"",""加工内訳キズ(7)"",""加工判断キズ(7)"",""加工内訳あたり(7)"",""加工判断あたり(7)"",""部品作業不良(7)"",""作業内訳キズ(7)"",""作業判断キズ(7)"",""作業内訳あたり(7)"",""作業判断あたり(7)"",""部品差替え(7)"",""部品得意先コード(8)"",""部品製品コード(8)"",""部品ロットＮＯ(8)"",""部品投入数(8)"",""部品印刷不良(8)"",""印刷不良判断(8)"",""部品加工不良(8)"",""加工内訳キズ(8)"",""加工判断キズ(8)"",""加工内訳あたり(8)"",""加工判断あたり(8)"",""部品作業不良(8)"",""作業内訳キズ(8)"",""作業判断キズ(8)"",""作業内訳あたり(8)"",""作業判断あたり(8)"",""部品差替え(8)"",""部品得意先コード(9)"",""部品製品コード(9)"",""部品ロットＮＯ(9)"",""部品投入数(9)"",""部品印刷不良(9)"",""印刷不良判断(9)"",""部品加工不良(9)"",""加工内訳キズ(9)"",""加工判断キズ(9)"",""加工内訳あたり(9)"",""加工判断あたり(9)"",""部品作業不良(9)"",""作業内訳キズ(9)"",""作業判断キズ(9)"",""作業内訳あたり(9)"",""作業判断あたり(9)"",""部品差替え(9)"",""部品得意先コード(10)"",""部品製品コード(10)"",""部品ロットＮＯ(10)"",""部品投入数(10)"",""部品印刷不良(10)"",""印刷不良判断(10)"",""部品加工不良(10)"",""加工内訳キズ(10)"",""加工判断キズ(10)"",""加工内訳あたり(10)"",""加工判断あたり(10)"",""部品作業不良(10)"",""作業内訳キズ(10)"",""作業判断キズ(10)"",""作業内訳あたり(10)"",""作業判断あたり(10)"",""部品差替え(10)"",""部品得意先コード(11)"",""部品製品コード(11)"",""部品ロットＮＯ(11)"",""部品投入数(11)"",""部品印刷不良(11)"",""印刷不良判断(11)"",""部品加工不良(11)"",""加工内訳キズ(11)"",""加工判断キズ(11)"",""加工内訳あたり(11)"",""加工判断あたり(11)"",""部品作業不良(11)"",""作業内訳キズ(11)"",""作業判断キズ(11)"",""作業内訳あたり(11)"",""作業判断あたり(11)"",""部品差替え(11)"",""部品得意先コード(12)"",""部品製品コード(12)"",""部品ロットＮＯ(12)"",""部品投入数(12)"",""部品印刷不良(12)"",""印刷不良判断(12)"",""部品加工不良(12)"",""加工内訳キズ(12)"",""加工判断キズ(12)"",""加工内訳あたり(12)"",""加工判断あたり(12)"",""部品作業不良(12)"",""作業内訳キズ(12)"",""作業判断キズ(12)"",""作業内訳あたり(12)"",""作業判断あたり(12)"",""部品差替え(12)"",""開始時間(1)"",""終了時間(1)"",""正社員男性(1)"",""正社員女性(1)"",""パート(1)"",""開始時間(2)"",""終了時間(2)"",""正社員男性(2)"",""正社員女性(2)"",""パート(2)"",""開始時間(3)"",""終了時間(3)"",""正社員男性(3)"",""正社員女性(3)"",""パート(3)"",""開始時間(4)"",""終了時間(4)"",""正社員男性(4)"",""正社員女性(4)"",""パート(4)"",""総工数"",""製造工賃"",""工程(1)"",""担当者(1)"",""担当者名(1)"",""工程(2)"",""担当者(2)"",""担当者名(2)"",""工程(3)"",""担当者(3)"",""担当者名(3)"",""工程(4)"",""担当者(4)"",""担当者名(4)"",""工程(5)"",""担当者(5)"",""担当者名(5)"",""工程(6)"",""担当者(6)"",""担当者名(6)"",""工程(7)"",""担当者(7)"",""担当者名(7)"",""工程(8)"",""担当者(8)"",""担当者名(8)"",""工程(9)"",""担当者(9)"",""担当者名(9)"",""工程(10)"",""担当者(10)"",""担当者名(10)"",""工程(11)"",""担当者(11)"",""担当者名(11)"",""工程(12)"",""担当者(12)"",""担当者名(12)"",""製缶不良コード(1)"",""製缶不良名(1)"",""製缶責任部門(1)"",""製缶胴不良１(1)"",""製缶蓋不良１(1)"",""製缶請求判断１(1)"",""製缶胴不良２(1)"",""製缶蓋不良２(1)"",""製缶請求判断２(1)"",""製缶不良コード(2)"",""製缶不良名(2)"",""製缶責任部門(2)"",""製缶胴不良１(2)"",""製缶蓋不良１(2)"",""製缶請求判断１(2)"",""製缶胴不良２(2)"",""製缶蓋不良２(2)"",""製缶請求判断２(2)"",""製缶不良コード(3)"",""製缶不良名(3)"",""製缶責任部門(3)"",""製缶胴不良１(3)"",""製缶蓋不良１(3)"",""製缶請求判断１(3)"",""製缶胴不良２(3)"",""製缶蓋不良２(3)"",""製缶請求判断２(3)"",""製缶不良コード(4)"",""製缶不良名(4)"",""製缶責任部門(4)"",""製缶胴不良１(4)"",""製缶蓋不良１(4)"",""製缶請求判断１(4)"",""製缶胴不良２(4)"",""製缶蓋不良２(4)"",""製缶請求判断２(4)"",""製缶不良コード(5)"",""製缶不良名(5)"",""製缶責任部門(5)"",""製缶胴不良１(5)"",""製缶蓋不良１(5)"",""製缶請求判断１(5)"",""製缶胴不良２(5)"",""製缶蓋不良２(5)"",""製缶請求判断２(5)"",""製缶不良コード(6)"",""製缶不良名(6)"",""製缶責任部門(6)"",""製缶胴不良１(6)"",""製缶蓋不良１(6)"",""製缶請求判断１(6)"",""製缶胴不良２(6)"",""製缶蓋不良２(6)"",""製缶請求判断２(6)"",""製缶不良コード(7)"",""製缶不良名(7)"",""製缶責任部門(7)"",""製缶胴不良１(7)"",""製缶蓋不良１(7)"",""製缶請求判断１(7)"",""製缶胴不良２(7)"",""製缶蓋不良２(7)"",""製缶請求判断２(7)"",""製缶不良コード(8)"",""製缶不良名(8)"",""製缶責任部門(8)"",""製缶胴不良１(8)"",""製缶蓋不良１(8)"",""製缶請求判断１(8)"",""製缶胴不良２(8)"",""製缶蓋不良２(8)"",""製缶請求判断２(8)"",""作業不良コード(1)"",""作業不良名(1)"",""作業責任部門(1)"",""作業胴不良１(1)"",""作業蓋不良１(1)"",""作業請求判断１(1)"",""作業胴不良２(1)"",""作業蓋不良２(1)"",""作業請求判断２(1)"",""作業不良コード(2)"",""作業不良名(2)"",""作業責任部門(2)"",""作業胴不良１(2)"",""作業蓋不良１(2)"",""作業請求判断１(2)"",""作業胴不良２(2)"",""作業蓋不良２(2)"",""作業請求判断２(2)"",""作業不良コード(3)"",""作業不良名(3)"",""作業責任部門(3)"",""作業胴不良１(3)"",""作業蓋不良１(3)"",""作業請求判断１(3)"",""作業胴不良２(3)"",""作業蓋不良２(3)"",""作業請求判断２(3)"",""作業不良コード(4)"",""作業不良名(4)"",""作業責任部門(4)"",""作業胴不良１(4)"",""作業蓋不良１(4)"",""作業請求判断１(4)"",""作業胴不良２(4)"",""作業蓋不良２(4)"",""作業請求判断２(4)"",""作業不良コード(5)"",""作業不良名(5)"",""作業責任部門(5)"",""作業胴不良１(5)"",""作業蓋不良１(5)"",""作業請求判断１(5)"",""作業胴不良２(5)"",""作業蓋不良２(5)"",""作業請求判断２(5)"",""作業不良コード(6)"",""作業不良名(6)"",""作業責任部門(6)"",""作業胴不良１(6)"",""作業蓋不良１(6)"",""作業請求判断１(6)"",""作業胴不良２(6)"",""作業蓋不良２(6)"",""作業請求判断２(6)"",""胴のみ１"",""蓋のみ１"",""胴のみ２"",""蓋のみ２"",""請求保留Ｆ"",""請求処理日"",""請求済みＦ"",""入力者"",""入力日"",""端末ＩＤ"""
'
'★ 以下を追加
'InsertFileName(2) = "REZAIRFL.DAT"
'InsertText(2) = """発注番号"",""材料区分"",""発注回数"",""受入回数"",""受注番号"",""材質"",""調質"",""板厚"",""表面"",""板縦サイズ"",""縦ロール"",""板横サイズ"",""横ロール"",""材料発注先Ｃ"",""印刷発注先Ｃ"",""得意先コード(1)"",""製品コード(1)"",""部品コード(1)"",""取数(1)"",""部品板枚数(1)"",""得意先コード(2)"",""製品コード(2)"",""部品コード(2)"",""取数(2)"",""部品板枚数(2)"",""得意先コード(3)"",""製品コード(3)"",""部品コード(3)"",""取数(3)"",""部品板枚数(3)"",""得意先コード(4)"",""製品コード(4)"",""部品コード(4)"",""取数(4)"",""部品板枚数(4)"",""得意先コード(5)"",""製品コード(5)"",""部品コード(5)"",""取数(5)"",""部品板枚数(5)"",""得意先コード(6)"",""製品コード(6)"",""部品コード(6)"",""取数(6)"",""部品板枚数(6)"",""得意先コード(7)"",""製品コード(7)"",""部品コード(7)"",""取数(7)"",""部品板枚数(7)"",""得意先コード(8)"",""製品コード(8)"",""部品コード(8)"",""取数(8)"",""部品板枚数(8)"",""得意先コード(9)"",""製品コード(9)"",""部品コード(9)"",""取数(9)"",""部品板枚数(9)"",""得意先コード(10)"",""製品コード(10)"",""部品コード(10)"",""取数(10)"",""部品板枚数(10)"",""今回発注数"",""発注日"",""今回受入数"",""単価"",""金額"",""納入日"",""受入取消Ｆ"",""入力者"",""集計処理済Ｆ"",""原価処理済Ｆ"",""処理日"",""処理時刻"",""端末ＩＤ"""
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
For i = 0 to Ubound(InsertText) 

	strPath = InsertFilePath & "\" & InsertFileName(i)

	'ファイルが存在することを確認
	If objFso.FileExists(strPath) = False Then
		MsgBox "ファイル[" & InsertFileName(i) & "]がみつかりません"
		WScript.Quit '処理終了
	End If

	'yyyy/MM/dd HH:mm:ss を yyyyMMdd_HHmmss に変換
	Dim dtNow
	dtNow = Now
	strDate = FormatDateTime(dtNow, vbShortDate) & " " & Right("0" & FormatDateTime(dtNow, vbLongTime),8)
	strFormattedDate = Replace(Replace(Replace(strDate, "/", ""), ":", ""), " ", "_")

	Dim strBackupPath 
	strBackupPath = InsertFilePath & "\" & Replace(InsertFileName(i), ".DAT", "") & "_" & strFormattedDate & ".DAT"

	'実行前にファイルをシステム日付の名前にしてバックアップする
	Call objFso.CopyFile(strPath , strBackupPath)

	'テキストの行数を確認
	Set objRead = objFso.OpenTextFile(strPath , ForReading)    '読取モードでテキストを開く
	objRead.ReadAll                                            '全部読むことで最終行へ移動
	intLine = objRead.Line                                     '現在の行数を確認
	objRead.Close                                              '読取モード閉じる

	'テキストの挿入
	If intLine <= InsertLine Then                              '挿入行がテキストの行数より大きいか、同じの場合

		Set objAppending = objFso.OpenTextFile(strPath , ForAppending)      '追記モードでテキストを開く
		objAppending.WriteLine InsertText(i)               '挿入行の追記
		objAppending.Close                                 '追記モード閉じる

	Else                                                   '挿入行がテキストの行数より小さい場合

		Dim WritingText                                    '書込用の文字列（省略可）
		Set objRead2 = objFso.OpenTextFile(strPath , ForReading)  '読取モードでテキストを開く

		row = 1                                                     '行数の確認用の数値
		Do Until objRead2.AtEndOfStream = True                    '終了行まで繰り返し
			If row = InsertLine Then                                '挿入行が来たら、文字列を挿入
				WritingText = WritingText & InsertText(i) & vbCrLf  'vbCrLfは改行コード
			End If

			WritingText = WritingText & objRead2.ReadLine & vbCrLf        '1行読み取り、書込用の文字列に追加
			row = row + 1                                       '読み取った行数を1増やす
		Loop

		objRead2.Close                                    '読取モード閉じる

		Set objWriting = objFso.OpenTextFile(strPath , ForWriting)          '書込モードでテキストを開く
		objWriting.Write WritingText                             '書込用の文字列値を一気に書込み
		objWriting.Close                                    '書込モード閉じる

	End If

Next 


'----------- 完了メッセージ
MsgBox "変換完了!!"
