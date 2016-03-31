Class PEReader
    Private m_stm
    Private m_strFileName
    Private m_szBuf
    Private m_nHeaderSize
    Private m_nMajorLinkerVer
    Private m_nMinorLinkerVer
    
    Public Property Get PEHeaderSize
        PEHeaderSize = m_nHeaderSize
    End Property
    
    Public Property Get MajorLinkerVer
        MajorLinkerVer = m_nMajorLinkerVer
    End Property
    
    Public Property Get MinorLinkerVer
        MinorLinkerVer = m_nMinorLinkerVer
    End Property
    
    Public Property Get LinkerVer
        LinkerVer = CStr(MajorLinkerVer) & "." & CStr(MinorLinkerVer)
    End Property
    
    Public Property Get FileName
        FileName = m_strFileName
    End Property
    
    Private Sub Class_Initialize
        Set m_stm = CreateObject("ADODB.Stream")
        m_stm.Type = 1 ' adTypeBinary
        m_stm.Open
    End Sub
    
    Private Sub Class_Terminate
        Set m_stm = Nothing
    End Sub
    
    Public Function Load( strFileName )
        On Error Resume Next
        m_strFileName = ExpandEnv( strFileName )
        m_stm.LoadFromFile m_strFileName
        If Err.Number <> 0 Then
            Load = False
            Exit Function
        End If
        m_stm.Position = 1
        m_szBuf = m_stm.Read( 512 ) ' 多めに読み込み
        
        m_nHeaderSize     = GetPEHeaderSize( m_szBuf )
        m_nMajorLinkerVer = GetMajorLinkerVer( m_szBuf )
        m_nMinorLinkerVer = GetMinorLinkerVer( m_szBuf )
        Load = True
    End Function
    
    Public Function ExpandEnv( strFileName )
        Dim strResult
        Dim shell
        Set shell = CreateObject("WScript.Shell")
        strResult = shell.ExpandEnvironmentStrings( strFileName )
        Set shell = Nothing
        ExpandEnv = strResult
    End Function
    
    Public Sub Show()
        WScript.Echo "FileName        = [" & Me.FileName & "]"
        WScript.Echo "PEHeaderSize    = [" & Me.PEHeaderSize & "]"
        WScript.Echo "MajorLinkerVer  = [" & Me.MajorLinkerVer & "]"
        WScript.Echo "MinorLinkerVer  = [" & Me.MinorLinkerVer & "]"
        WScript.Echo "LinkerVer       = [" & Me.LinkerVer & "]"
    End Sub
    
    Private Function GetPEHeaderSize( szBuf )
        Dim nResult
        Dim nPosition
        Dim nSize
        
        nPosition = 60 ' IMAGE_DOS_HEADER の e_lfanew の位置
        nSize = 4
        nResult = GetFieldValueFromBinary( szBuf, nPosition, nSize )
        
        GetPEHeaderSize = nResult
    End Function
    
    Private Function GetMajorLinkerVer( szBuf )
        Dim nResult
        Dim nPosition
        Dim nSize
        
        nPosition = GetPEHeaderSize( szBuf )
        nPosition = nPosition + 4 + 20 + 2     ' IMAGE_OPTIONAL_HEADER の MajorLinkerVersion の位置
        nSize = 1
        nResult = GetFieldValueFromBinary( szBuf, nPosition, nSize )
        
        GetMajorLinkerVer = nResult
    End Function
    
    Private Function GetMinorLinkerVer( szBuf )
        Dim nResult
        Dim nPosition
        Dim nSize
        
        nPosition = GetPEHeaderSize( szBuf )
        nPosition = nPosition + 4 + 20 + 2 + 1 ' IMAGE_OPTIONAL_HEADER の MinorLinkerVersion の位置
        nSize = 1
        nResult = GetFieldValueFromBinary( szBuf, nPosition, nSize )
        
        GetMinorLinkerVer = nResult
    End Function
    
    Private Function GetFieldValueFromBinary( szBuf, nPosition, nSize )
        Dim nResult
        
        Dim szField
        szField = MidB( szBuf, nPosition, nSize )
        nResult = ConvertBinaryToNumber( szField, nSize )
        
        GetFieldValueFromBinary = nResult
    End Function
    
    Private Function ConvertBinaryToNumber( szBuf, nSize )
        Dim nResult
        nResult = 0
        Dim ch
        Dim i
        For i = 1 To nSize
            ch = AscB( MidB( szBuf, i, 1 ) )
            nResult = nResult + ch * 256 ^ (i-1) ' リトルエンディアンを想定
        Next
        ConvertBinaryToNumber = nResult
    End Function
End Class