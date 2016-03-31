'------------------------------------------------
' databuffer
' Example:
'    Set MyBuffer = CreateBuff
'
'    Wscript.Echo MyBuffer.Length
'
'    MyBuffer.SetData("001")
'
'    Wscript.Echo MyBuffer.Length
'
'    MyBuffer.SetData("日本語表示")
'    MyBuffer.SetData("XYZ")
'
'    Wscript.Echo MyBuffer.Length
'
'    ' 内部配列の直接参照
'    Wscript.Echo MyBuffer.Buff(1)
'    MyBuffer.Buff(1) = "日本語"
'
'    ' 改行コードで連結した内部配列
'    Wscript.Echo MyBuffer.GetData(vbCrLf)
'
'    ' カンマで連結した内部配列
'    Wscript.Echo MyBuffer.GetData(",")
Class buffCon

	Public Buff()

	' ************************************************
	' コンストラクタ
	' ************************************************
	Public Default Function InitSetting()

		Redim Buff(0)

	end function

	' ************************************************
	' メソッド ( データセット )
	' ************************************************
	function Length()

		if IsEmpty( Buff(0) ) then
			Length = 0
		else
			Length =  Ubound(Buff)+1
		end if

	end function

	' ************************************************
	' メソッド ( データセット )
	' ************************************************
	function SetData( strRow )

		if IsEmpty( Buff(0) ) then
			Buff(0) = strRow
		else
			ReDim Preserve Buff(Ubound(Buff)+1)
			Buff(Ubound(Buff)) = strRow
		end if

	end function
	 
	' ************************************************
	' メソッド ( データ取得 )
	' ************************************************
	function GetData( strDelim )

		GetData = Join( Buff, strDelim )

	end function

End Class

' **********************************************************
' バッファ作成
' **********************************************************
Function CreateBuff( )

	ExecuteString = "Dim gblCreateBuff : "
	ExecuteString = ExecuteString & "Set gblCreateBuff = new buffCon"
	ExecuteGlobal ExecuteString
	Call gblCreateBuff

	Set CreateBuff = gblCreateBuff

End Function
