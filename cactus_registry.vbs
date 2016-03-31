'------------------------------------------------
' É¾³ý×¢²á±í¼ü
Sub RegDelete(fullkey)
    Set objShell = CreateObject(COM_SHELL)
    objShell.RegDelete fullkey
End Sub

'------------------------------------------------
' É¾³ý×¢²á±í¼ü
Sub RegDeleteKey(rootkey, key)
    Dim computer   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv")
         
    oReg.DeleteKey rootkey, key
End Sub 

'------------------------------------------------
' É¾³ý×¢²á±í¼üÖµ
Sub RegDeleteValue(rootkey, key, name)
    Dim computer   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv") 
    oReg.DeleteValue rootkey, key, name
End Sub


'------------------------------------------------
' Ð´×¢²á±íMultiStringÖµ
Sub RegWriteMultiStringValue(rootkey, key, name, ByRef values)
    Dim computer   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv")
         
    oReg.SetMultiStringValue rootkey, key, name, values
End Sub

'------------------------------------------------
' ¶Á×¢²á±íMultiStringÖµ
Function RegReadMultiString(rootkey, key, name)
    Dim computer, arrValues   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv") 

    oReg.GetMultiStringValue rootkey, key, name, arrValues
    RegReadMultiString = arrValues
End Function

'------------------------------------------------
' Ð´×¢²á±íStringÖµ
Sub RegWriteStringValue(rootkey, key, name, value)
    Dim computer   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv") 

    oReg.SetStringValue rootkey, key, name, value
End Sub

'------------------------------------------------
' ¶Á×¢²á±íStringÖµ
Function RegReadStringValue(rootkey, key, name)
    Dim computer, value   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv") 

    oReg.GetStringValue rootkey, key, name, value
    RegReadStringValue = value
End Function

'------------------------------------------------
' Ð´×¢²á±íDWORDÖµ
Sub RegWriteDWORDValue(rootkey, key, name, value)
    Dim computer   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv") 

    oReg.SetDWORDValue rootkey, key, name, value
End Sub

'------------------------------------------------
' ¶Á×¢²á±íDWORDÖµ
Function RegReadDWORDValue(rootkey, key, name)
    Dim computer, value   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv") 

    oReg.GetDWORDValue rootkey, key, name, value
    RegReadDWORDValue = value
End Function


'------------------------------------------------
' Ã¶¾Ù×¢²á±í¼ü
Function RegEnumKeys(rootkey, key)
    Dim computer, arrSubKeys   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv") 
    oReg.EnumKey rootkey, key, arrSubKeys 
    RegEnumKeys = arrSubKeys
End Function

'------------------------------------------------
' Ã¶¾Ù×¢²á±íÖµ
Function RegEnumValues(rootkey, key, ByRef arrValueTypes)
    Dim computer, arrValueNames   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv")
    oReg.EnumValues rootkey, key, arrValueNames, arrValueTypes
    RegEnumValues = arrValueNames
End Function

'------------------------------------------------
' ×¢²á±íÀà
Class Registry
	Private FRootKey
	Private FReg
	Private FKey	

	Private Sub Class_Initialize
		Dim computer
		computer = "."	
        Set FReg = 	GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
				computer & "\root\default:StdRegProv")		
    End Sub

	Private Sub Class_Terminate
		Set FReg = Nothing	
    End Sub

	Public Default Function Init(ARootKey, AKey)		
		FRootKey = ARootKey		
		FKey = AKey
		Set Init = Me
    End Function

	Public Function ExistsKey(AKey)			
		If FReg.EnumKey(FRootKey, AKey, "", "") = 0 Then
			ExistsKey = True
		Else
			ExistsKey = False
		End If		
	End Function

		
	Public Function CreateKey(AKey)			
		FReg.Createkey FRootKey, AKey		
		FKey = AKey
	End Function

	Public Function DeleteKey(AKey)
		FReg.DeleteKey FRootKey, AKey		
	End Function

	Public Function OpenKey(AKey)
		FKey = AKey
	End Function

	Public Function EnumKey(AKey)		
		Dim arrSubKeys
		FReg.EnumKey FRootKey, AKey, arrSubKeys
		EnumKey = arrSubKeys
	End Function


	Public Sub SetStringValue(AValueName, AValue)					
		FReg.SetStringValue FRootKey, FKey, AValueName, AValue
	End Sub

	Public Sub SetDWORDValue(AValueName, AValue)					
		FReg.SetDWORDValue FRootKey, FKey, AValueName, AValue
	End Sub

	Public Sub SetExpandedStringValue(AValueName, AValue)					
		FReg.SetExpandedStringValue FRootKey, FKey, AValueName, AValue
	End Sub
	 
	Public Function DeleteValue(AValueName)					
		FReg.DeleteValue FRootKey, FKey, AValueName
	End Function	

	Public Function GetDWORDValue(AValueName)		
		Dim dwValue
		FReg.GetDWORDValue FRootKey, FKey, AValueName, dwValue
		GetDWORDValue = dwValue
	End Function

	Public Function GetStringValue(AValueName)		
		Dim strValue
		FReg.GetStringValue FRootKey, FKey, AValueName, dwValue
		GetStringValue = strValue
	End Function

	Public Function GetMultiStringValue(AValueName)		
		Dim arrValues
		FReg.GetMultiStringValue FRootKey, FKey, AValueName, arrValues
		GetMultiStringValue = arrValues
	End Function

	Public Function GetExpandedStringValue(AValueName)		
		Dim strValue
		FReg.GetExpandedStringValue FRootKey, FKey, AValueName, strValue
		GetExpandedStringValue = strValue
	End Function

	Public Function GetBinaryValue(AValueName)		
		Dim strValue
		FReg.GetBinaryValue FRootKey, FKey, AValueName, strValue
		GetBinaryValue = strValue
	End Function

	Public Sub EnumValues(AKey, ByRef arrValueNames, ByRef arrTypes)		
		FReg.EnumValues FRootKey, AKey, arrValueNames, arrTypes		
	End Sub

End Class