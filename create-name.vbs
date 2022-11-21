' ***********************************************************
' 処理開始
' create table [社員マスタ] (
' 	[社員コード] VARCHAR(4)
' 	,[氏名] VARCHAR(50)
' 	,[フリガナ] VARCHAR(50)
' 	,[所属] VARCHAR(4)
' 	,[性別] INT
' 	,[作成日] DATETIME
' 	,[更新日] DATETIME
' 	,[給与] INT
' 	,[手当] INT
' 	,[管理者] VARCHAR(4)
' 	,[生年月日] DATETIME
' 	,primary key([社員コード])
' )
' ***********************************************************

nMax = 50

strName1 = "山川森鈴木高田本多村吉岡松丸杉浦中尾安原野内"
strName1k = "ヤマ,カワ,モリ,スズ,キ,タカ,タ,モト,タ,ムラ,ヨシ,オカ,マツ,マル,スギ,ウラ,ナカ,オ,ヤス,ハラ,ノ,ウチ"
strName2 = "和元雅正由克友浩春冬洋輝"
strName2k = "カズ,モト,マサ,マサ,ヨシ,カツ,トモ,ヒロ,ハル,フユ,ヒロ,テル"

strName3 = "男也一行樹之"
strName3k = "オ,ヤ,カズ,ユキ,キ,ユキ"
strName4 = "子代美恵"
strName4k = "コ,ヨ,ミ,エ"

strNo = ""
Query = ""

For i = 1 to nMax


	Query = Query & vbCrLf & "insert into [社員マスタ] values("

	strNo = Fzero( i , 4 )

	Query = Query & Ss(strNo)

	' 姓1文字目
	nTarget = SameRandom( 1, Len(strName1) )
	strName = Mid( strName1, nTarget, 1 )
	aData = Split(strName1k,",")
	strKana = aData(nTarget-1)
	' 1文字目と2文字目が一致したら除外
	nTarget2 = nTarget
	Do while( nTarget = nTarget2 )
		nTarget2 = SameRandom( 1, Len(strName1) )
	Loop
	' 姓2文字目
	strName = strName & Mid( strName1, nTarget2, 1 ) & " "
	strKana = strKana & aData(nTarget2-1) & " "
	' 名1文字目
	nTarget = SameRandom( 1, Len(strName2) )
	strName = strName & Mid( strName2, nTarget, 1 )
	aData = Split(strName2k,",")
	strKana = strKana & aData(nTarget-1)
	' 性別
	nTarget = SameRandom( 0, 1 )
	nS = nTarget
	' 性別によって名2文字目を決定
	if nTarget = 0 then
		nTarget = SameRandom( 1, Len(strName3) )
		strName = strName & Mid( strName3, nTarget, 1 )
		aData = Split(strName3k,",")
		strKana = strKana & aData(nTarget-1)
	else
		nTarget = SameRandom( 1, Len(strName4) )
		strName = strName & Mid( strName4, nTarget, 1 )
		aData = Split(strName4k,",")
		strKana = strKana & aData(nTarget-1)
	end if

	Query = Query & "," & Ss( strName )
	Query = Query & "," & Ss( strKana )
	nTarget = SameRandom( 1, 3 )
	Query = Query & "," & Ss( Fzero( nTarget, 4 ) )
	Query = Query & "," & nS
	strWork = Date() - SameRandom( 0, 100 )
	Query = Query & "," & Ss( strWork )
	strWork = Date() + SameRandom( 0, 100 )
	Query = Query & "," & Ss( strWork )
	Query = Query & "," & SameRandom( 14, 30 ) * 10000
	if i mod 5 = 1 then
		Query = Query & "," & SameRandom( 5, 10 ) * 1000
	else
		Query = Query & ",NULL"
	end if
	if i <= 5 then
		Query = Query & ",NULL"
	else
		Query = Query & "," & Ss( Fzero( SameRandom( 1, 5 ) , 4 ) )
	end if
	Query = Query & ",'2000/01/01');"

	Wscript.Echo strNo & " " & strName & "(" & strKana & ")"

Next

Wscript.Echo Query

Wscript.Echo "update 社員マスタ set 生年月日 = 生年月日 + 社員コード;"

' ************************************************
' 指定範囲の整数の乱数を取得
' ************************************************
Function Random( nMin, nMax )

	Randomize
	Random = nMin + Int(Rnd * (nMax - nMin + 1))

End function

Function SameRandom( nMin, nMax )

	SameRandom = nMin + Int(Rnd * (nMax - nMin + 1))

End function


' **********************************************************
' ダブルクォートで囲む
' **********************************************************
Function Dd( strValue )

	Dd = """" & strValue & """"

End function

' **********************************************************
' 前ゼロ編集
' **********************************************************
Function Fzero( nData, nLen )

	Dim str

	str = String( nLen, "0" )
	str = str & nData
	Fzero = Right( str, nLen )

End Function

' **********************************************************
' シングルクォート挟み込み
' **********************************************************
Function Ss( str )

	Ss = "'" & str & "'"

End Function

