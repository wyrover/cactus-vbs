'------------------------------------------------
' ObjectList

Class ObjectList
    Public List
    
    Sub Class_Initialize()
        Set List = CreateObject(COM_DICT)
    End Sub
    
    Sub Class_Terminate()
        Set List = Nothing
    End Sub
    
    Function Append(Anything) 
        List.Add CStr(List.Count + 1), Anything 
        Set Append = Anything
    End Function
    
    Function Item(id) 
        If List.Exists(CStr(id)) Then
            Set Item = List(CStr(id))
        Else
            Set Item = Nothing
        End If
    End Function
End Class