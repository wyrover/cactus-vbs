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
'    MyBuffer.SetData("�ձ��Z��ʾ")
'    MyBuffer.SetData("XYZ")
'
'    Wscript.Echo MyBuffer.Length
'
'    ' �ڲ����Ф�ֱ�Ӳ���
'    Wscript.Echo MyBuffer.Buff(1)
'    MyBuffer.Buff(1) = "�ձ��Z"
'
'    ' ���Х��`�ɤ��B�Y�����ڲ�����
'    Wscript.Echo MyBuffer.GetData(vbCrLf)
'
'    ' ����ޤ��B�Y�����ڲ�����
'    Wscript.Echo MyBuffer.GetData(",")
Class buffCon

	Public Buff()

	' ************************************************
	' ���󥹥ȥ饯��
	' ************************************************
	Public Default Function InitSetting()

		Redim Buff(0)

	end function

	' ************************************************
	' �᥽�å� ( �ǩ`�����å� )
	' ************************************************
	function Length()

		if IsEmpty( Buff(0) ) then
			Length = 0
		else
			Length =  Ubound(Buff)+1
		end if

	end function

	' ************************************************
	' �᥽�å� ( �ǩ`�����å� )
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
	' �᥽�å� ( �ǩ`��ȡ�� )
	' ************************************************
	function GetData( strDelim )

		GetData = Join( Buff, strDelim )

	end function

End Class

' **********************************************************
' �Хåե�����
' **********************************************************
Function CreateBuff( )

	ExecuteString = "Dim gblCreateBuff : "
	ExecuteString = ExecuteString & "Set gblCreateBuff = new buffCon"
	ExecuteGlobal ExecuteString
	Call gblCreateBuff

	Set CreateBuff = gblCreateBuff

End Function
