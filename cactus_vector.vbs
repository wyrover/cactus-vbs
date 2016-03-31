'------------------------------------------------
' Vector Class
Class Vector
    Private myStack
    Private myCount

    Private Sub Class_Initialize()
        Redim myStack(8)
        myCount = -1
    End Sub

    Private Sub Class_Terminate()
    End Sub

    Public Property Let Dimension(pDim)
        Redim myStack(pDim)
    End Property

    Public Property Get Count()
        Count = myCount + 1
    End Property

    Public Sub Push(pElem)
        myCount = myCount + 1
        If (UBound(myStack) < myCount) Then
            Redim Preserve myStack(UBound(myStack) * 2)
        End If
        Call SetElementAt(myCount, pElem)
    End Sub

    Public Function Pop()
        If IsObject(myStack(myCount)) Then
            Set Pop = myStack(myCount)
        Else
            Pop = myStack(myCount)
        End If
        myCount = myCount - 1
    End Function

    Public Function Top()
        If IsObject(myStack(myCount)) Then
            Set Top = myStack(myCount)
        Else
            Top = myStack(myCount)
        End If
    End Function

    Public Function ElementAt(pIndex)
        If IsObject(myStack(pIndex)) Then
            Set ElementAt = myStack(pIndex)
        Else
            ElementAt = myStack(pIndex)
        End If
    End Function

    Public Sub SetElementAt(pIndex, pValue)
        If IsObject(pValue) Then
            Set myStack(pIndex) = pValue
        Else
            myStack(pIndex) = pValue
        End If
    End Sub

    Public Sub RemoveElementAt(pIndex)
        Do While pIndex < myCount
            Call SetElementAt(pIndex, ElementAt(pIndex + 1))
            pIndex = pIndex + 1
        Loop
        myCount = myCount - 1
    End Sub

    Public Function IsEmpty()
        IsEmpty = (myCount < 0)
    End Function
End Class
