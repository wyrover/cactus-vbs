' COMMENT: This vbscript class lets you manipulate byte arrays.

' How do you manipulate a byte array with vbscript?? For example, how do you edit an
' image file? Say you need to change some color values.
' The only way I have found is to read chunks of bytes using the MidB function:

' 1) In your application: Convert the byte array to a string using MidB
' 2) Convert this string to a string of hex values (see: OctetToHexStr)
' 3) Add these strings to a temp text file (see: AddBytes)
' 4) Open that file and convert the hex values back into a byte array using ADO stream.
'     This is where the magic is done. (see: ConvertHexStringToByteArray)
' 5) Save the byte array to binary file(see: SaveToFile)

' I take no credit for this smart conversion technic :-)
' See http://blogs.brnets.com/michael/archive/2005/03/09/387.aspx
'============================================================
' ROUTINES:

' - Public Property Get BytesTotal()

' - Private Sub Class_Initialize()
' - Private Sub Class_Terminate()
' - Public Sub AddBytes(bytes)
' - Public Function ReturnBytes()
' - Private Function OctetToHexStr(arrbytOctet)
' - Private Sub ConvertHexStringToByteArray(ByVal strHexString, ByRef pByteArray)
' - Private Function ReadFile()
' - Private Sub DeleteTempFile()
'============================================================
Class cByteArray    
    Private m_lngByteSize                       '// Size of byte array
    Private m_sTmpPath                          '// Temp file holding the hexed string(s)

    '// MODULE PROPERTIES
    Public Property Get BytesTotal()
        BytesTotal = m_lngByteSize
    End Property

    '------------------------------------------------------------------------------------------------------------
    ' Comment: Initialize our module variables
    '------------------------------------------------------------------------------------------------------------
    Private Sub Class_Initialize()
        On Error Resume Next

        m_lngByteSize = 0
        m_sTmpPath = Server.MapPath("temp.txt")

    End Sub

    '--------------------------------------------------------------------------------------------------------
    ' Comment: Clean up
    '--------------------------------------------------------------------------------------------------------
    Private Sub Class_Terminate()
        On Error Resume Next

        Call DeleteTempFile

    End Sub

    '------------------------------------------------------------------------------------------------------------
    ' Comment: Append string chuncks to a temp file.
    '------------------------------------------------------------------------------------------------------------
    Public Sub AddBytes(bytes)
        On Error Resume Next

        Dim oFSO, oFile, sHexString

        If LenB(bytes) = 0 Then Exit Sub

        '// Convert the string chunks to hexed characters and add them to a temp text file.
        sHexString = OctetToHexStr(bytes)

        '// Set the new number of bytes
        m_lngByteSize = m_lngByteSize + LenB(bytes)

        If IsEmpty(oFSO) Then Set oFSO = CreateObject(COM_FSO)
        Set oFile = oFSO.OpenTextFile(m_sTmpPath, ForAppending, True)
        oFile.Write sHexString
        oFile.Close
        
        Set oFile = Nothing
        Set oFSO = Nothing

    End Sub

    '------------------------------------------------------------------------------------------------------------
    ' Comment: Return all bytes.
    '------------------------------------------------------------------------------------------------------------
    Public Function ReturnBytes()
        On Error Resume Next

        Dim sHex, arrBytes

        '// Load the hexed characters from the temp text file ..
        sHex = ReadFile
        '// .. and convert them back into a byte array. We pass an empty variant byref
        '// that will be set as a byte array.
        Call ConvertHexStringToByteArray(sHex, arrBytes)

        ReturnBytes = arrBytes

    End Function

    '------------------------------------------------------------------------------------------------------------
    ' Comment: http://blogs.brnets.com/michael/archive/2005/03/09/387.aspx
    '------------------------------------------------------------------------------------------------------------
    Private Function OctetToHexStr(arrbytOctet)
        On Error Resume Next
        ' Function to convert OctetString (byte array) to Hex string.
        ' Code from Richard Mueller, a MS MVP in Scripting and ADSI

        Dim k

        OctetToHexStr = ""

        For k = 1 To LenB(arrbytOctet)
            OctetToHexStr = OctetToHexStr & Right("0" & Hex(AscB(MidB(arrbytOctet, k, 1))), 2)
        Next

    End Function

    '------------------------------------------------------------------------------------------------------------
    ' Comment: http://blogs.brnets.com/michael/archive/2005/03/09/387.aspx
    '------------------------------------------------------------------------------------------------------------
    Private Sub ConvertHexStringToByteArray(ByVal strHexString, _
                                            ByRef pByteArray)
        On Error Resume Next

        Dim fso, stream, temp, ts, n
        ' This is an elegant way to convert a hex string to a Byte
        ' array. Typename(pByteArray) will return Byte(). pByteArray
        ' should be a null variant upon entry. strHexString should be
        ' an ASCII string containing nothing but hex characters, e.g.,
        ' FD70C1BC2206240B828F7AE31FEB55BE
        ' Code from Michael Harris, a MS MVP in Scripting

        Set fso = CreateObject(COM_FSO)
        Set stream = CreateObject(COM_ADOSTREAM)

        temp = Server.MapPath(fso.gettempname)
        Set ts = fso.createtextfile(temp)

        For n = 1 To (Len(strHexString) - 1) Step 2
            ts.Write Chr("&h" & Mid(strHexString, n, 2))
        Next

        ts.Close

        stream.Type = 1
        stream.Open
        stream.LoadFromFile temp
        pByteArray = stream.Read

        stream.Close
        fso.DeleteFile temp

        Set stream = Nothing
        Set fso = Nothing

    End Sub

    '------------------------------------------------------------------------------------------------------------
    ' Comment: Read in a text file.
    '------------------------------------------------------------------------------------------------------------
    Private Function ReadFile()
        On Error Resume Next

        Dim oFSO
        Dim oFile
     
        If IsEmpty(oFSO) Then Set oFSO = Server.CreateObject(COM_FSO)
        If Not oFSO.FileExists(m_sTmpPath) Then Exit Function
        Set oFile = oFSO.OpenTextFile(m_sTmpPath, ForReading, False)
        
        ReadFile = oFile.ReadAll

        Set oFile = Nothing
        Set oFSO = Nothing

    End Function

    '------------------------------------------------------------------------------------------------------------
    ' Comment: Delete temporary file.
    '------------------------------------------------------------------------------------------------------------
    Private Sub DeleteTempFile()
        On Error Resume Next        
        Dim oFSO
        If IsEmpty(oFSO) Then Set oFSO = Server.CreateObject(COM_FSO)
        If oFSO.FileExists(m_sTmpPath) Then oFSO.DeleteFile (m_sTmpPath)
        Set oFSO = Nothing
    End Sub

End Class