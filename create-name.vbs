' ***********************************************************
' �����J�n
' create table [�Ј��}�X�^] (
' 	[�Ј��R�[�h] VARCHAR(4)
' 	,[����] VARCHAR(50)
' 	,[�t���K�i] VARCHAR(50)
' 	,[����] VARCHAR(4)
' 	,[����] INT
' 	,[�쐬��] DATETIME
' 	,[�X�V��] DATETIME
' 	,[���^] INT
' 	,[�蓖] INT
' 	,[�Ǘ���] VARCHAR(4)
' 	,[���N����] DATETIME
' 	,primary key([�Ј��R�[�h])
' )
' ***********************************************************

nMax = 50

strName1 = "�R��X��؍��c�{�����g�����ې��Y�����������"
strName1k = "���},�J��,����,�X�Y,�L,�^�J,�^,���g,�^,����,���V,�I�J,�}�c,�}��,�X�M,�E��,�i�J,�I,���X,�n��,�m,�E�`"
strName2 = "�a���됳�R���F�_�t�~�m�P"
strName2k = "�J�Y,���g,�}�T,�}�T,���V,�J�c,�g��,�q��,�n��,�t��,�q��,�e��"

strName3 = "�j���s���V"
strName3k = "�I,��,�J�Y,���L,�L,���L"
strName4 = "�q����b"
strName4k = "�R,��,�~,�G"

strNo = ""
Query = ""

For i = 1 to nMax


	Query = Query & vbCrLf & "insert into [�Ј��}�X�^] values("

	strNo = Fzero( i , 4 )

	Query = Query & Ss(strNo)

	' ��1������
	nTarget = SameRandom( 1, Len(strName1) )
	strName = Mid( strName1, nTarget, 1 )
	aData = Split(strName1k,",")
	strKana = aData(nTarget-1)
	' 1�����ڂ�2�����ڂ���v�����珜�O
	nTarget2 = nTarget
	Do while( nTarget = nTarget2 )
		nTarget2 = SameRandom( 1, Len(strName1) )
	Loop
	' ��2������
	strName = strName & Mid( strName1, nTarget2, 1 ) & " "
	strKana = strKana & aData(nTarget2-1) & " "
	' ��1������
	nTarget = SameRandom( 1, Len(strName2) )
	strName = strName & Mid( strName2, nTarget, 1 )
	aData = Split(strName2k,",")
	strKana = strKana & aData(nTarget-1)
	' ����
	nTarget = SameRandom( 0, 1 )
	nS = nTarget
	' ���ʂɂ���Ė�2�����ڂ�����
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

Wscript.Echo "update �Ј��}�X�^ set ���N���� = ���N���� + �Ј��R�[�h;"

' ************************************************
' �w��͈͂̐����̗������擾
' ************************************************
Function Random( nMin, nMax )

	Randomize
	Random = nMin + Int(Rnd * (nMax - nMin + 1))

End function

Function SameRandom( nMin, nMax )

	SameRandom = nMin + Int(Rnd * (nMax - nMin + 1))

End function


' **********************************************************
' �_�u���N�H�[�g�ň͂�
' **********************************************************
Function Dd( strValue )

	Dd = """" & strValue & """"

End function

' **********************************************************
' �O�[���ҏW
' **********************************************************
Function Fzero( nData, nLen )

	Dim str

	str = String( nLen, "0" )
	str = str & nData
	Fzero = Right( str, nLen )

End Function

' **********************************************************
' �V���O���N�H�[�g���ݍ���
' **********************************************************
Function Ss( str )

	Ss = "'" & str & "'"

End Function

