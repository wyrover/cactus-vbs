'------------------------------------------------
' 脚本目录
Function ScriptPath
    ScriptPath = left(Wscript.ScriptFullName,len(Wscript.ScriptFullName)-len(Wscript.ScriptName))
    'ScriptPath = Replace(WScript.ScriptFullName, "\" & WScript.ScriptName, "")
End Function

'------------------------------------------------
' 文件夹是否存在
Function FolderExists(dir)
    Dim FSO 
    Set FSO = CreateObject(COM_FSO) 
    FolderExists = FSO.FolderExists(dir)
    Set FSO = Nothing 
End Function

'------------------------------------------------
' 文件是否存在
Function FileExists(filename)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    FileExists = FSO.FileExists(filename)
End Function

'------------------------------------------------
' 目录是否存在
Function DirExists(dirname)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    DirExists = FSO.FolderExists(dirname)
End Function

'------------------------------------------------
' 移动文件夹
' sourcedir = "C:\Scripts"
' destdir = "D:\Archive"
Sub MoveFolder(sourcedir, destdir)    
    Dim objShell, objFolder
    Set objShell = CreateObject(COM_SHELLAPP)
    Set objFolder = objShell.NameSpace(destdir) 
    objFolder.MoveHere sourcedir, FOF_CREATEPROGRESSDLG
End Sub

'------------------------------------------------
' 删除文件
' 删除.txt文件，"C:\FSO\*.txt"
Function DeleteFiles(filename)
    
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    
    If FSO.FileExists(filename) Then
        FSO.DeleteFile filename, True
        DeleteFiles = True
    Else
        DeleteFiles = False
    End If
    
    Set FSO = Nothing
    
End Function 

'------------------------------------------------
' 删除特定文件
' @delfilesname         文件列表"test1.txt|test2.txt"
' @dirname            文件目录
Sub DelFiles(delfilesname, dirname) 
    Dim FSO, files, fullpath, I
    If Right(dirname, 1) <> "\" Then dirname = dirname & "\"
    If delfilesname <> "" And Not IsNull(delfilesname) Then
        Set FSO = CreateObject(COM_FSO)
        files = Split(delfilesname & "|", "|")
        For I = 0 to Ubound(files) - 1
            fullpath = dirname + files(I)
            If FSO.FileExists(fullpath) Then FSO.DeleteFile(fullpath)
        Next
    End If
End Sub

'------------------------------------------------
' 删除特定文件
' @dir          文件目录
' @days         当前日期减去多少天
Sub DeleteFilesByDate(dir, days)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    Call DeleteSubFolders(FSO.GetFolder(dir), days, FSO)
End Sub

Sub DeleteSubFolders(folder, days, fso)
    Dim subfolder, files
    For Each subfolder in folder.SubFolders
        Set files = subfolder.Files
        If files.Count <> 0 Then
            For Each file in Files
                If file.DateLastModified < (Now - days) Then
                    fso.DeleteFile(subfolder.Path & "\" & file.Name)    
                End If
            Next
        End If
        Call DeleteSubFolders(subfolder, days, fso)
    Next
End Sub

'------------------------------------------------
' 文件重命名
Function ReFilename(filename, name)
    Dim FSO, file
    Set FSO = CreateObject(COM_FSO)
    Set file = FSO.GetFile(filename)
    file.Name = name
    Set FSO = Nothing
End Function 

'------------------------------------------------
' 文件夹重命名
Function ReDir(source, dest)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    FSO.MoveFolder source, dest
    Set FSO = Nothing
End Function

'------------------------------------------------
' 获取文件路径
Function GetFilePath(filename)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    GetFilePath = DisposePath(FSO.GetParentFolderName(filename))
End Function 


'------------------------------------------------
' 获取文件绝对路径
Function GetAbsolutePathName(filename)
    Dim FSO, file
    Set FSO = CreateObject(COM_FSO)
    Set file = FSO.GetFile(filename)
    GetAbsolutePathName = FSO.GetAbsolutePathName(file)
End Function

'------------------------------------------------
' 获取文件名
Function GetFileName(filename)
    Dim FSO, file
    Set FSO = CreateObject(COM_FSO)
    Set file = FSO.GetFile(filename)
    GetFileName = FSO.GetFileName(file)
End Function

'------------------------------------------------
' 获取基本文件名
Function GetBaseName(filename)
    Dim FSO, file
    Set FSO = CreateObject(COM_FSO)
    Set file = FSO.GetFile(filename)
    GetBaseName = FSO.GetBaseName(file)
End Function

'------------------------------------------------
' 获取文件扩展名
Function GetExtensionName(filename)
    Dim FSO, file
    Set FSO = CreateObject(COM_FSO)
    Set file = FSO.GetFile(filename)
    GetExtensionName = FSO.GetExtensionName(file)
End Function

'------------------------------------------------
' 获取文件扩展名
Function GetAnExtension(filename)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    GetAnExtension = FSO.GetExtensionName(filename)
End Function

'------------------------------------------------
' 获取工作目录
Function GetCurrentDirectory() 
    Dim objShell
    Set objShell = CreateObject(COM_SHELL)
    GetCurrentDirectory = objShell.CurrentDirectory 
End Function 



'------------------------------------------------
' 枚举当前脚本目录的子目录
Sub EnumCurrentDirectory(ByRef arrFolders)
    Dim objShell, objFSO, objFolder, currentDirectory, folder
    Set objShell = CreateObject(COM_SHELL)
    currentDirectory = objShell.CurrentDirectory
    Set objFSO = CreateObject(COM_FSO)
    Set objFolder = objFSO.GetFolder(currentDirectory)
    Set arrFolders = objFolder.SubFolders	
    Set objShell = Nothing
    Set objFSO = Nothing
End Sub

'------------------------------------------------
' 重命名文件夹
Sub RenameFolders(folder1, folder2)
    On Error Resume Next
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    FSO.MoveFolder (folder1),(folder2)
    If Err.Number <> 0 Then
        WScript.Echo "sorry you have a file open in that directory"
        WScript.Echo Err.Description
        WScript.Echo Err.Number
        Err.Clear 
    End If
End Sub

'------------------------------------------------
' 重命名文件
Sub RenameFile(sourcefile, destfile)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    FSO.MoveFile sourcefile, destfile
End Sub

'------------------------------------------------
' 创建级联目录
' Example:
'   ForceCreateFolder("C:\d\e\f\g\h")
Sub ForceCreateFolder(dir)
    On Error Resume Next
    Dim FSO, dirpath
    Set FSO = CreateObject(COM_FSO)
    dirpath = FSO.GetAbsolutePathName(dir)
    If (Not FSO.folderExists(FSO.GetParentFolderName(dirpath))) then    
        Call ForceCreateFolder(fso.GetParentFolderName(dirpath))
    End If
    
    FSO.CreateFolder(dirpath)
End Sub

'------------------------------------------------
' 删除目录
' Example:
'   ForceDeleteFolder("C:\d")
Sub ForceDeleteFolder(dir)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    dir = FSO.GetAbsolutePathName(dir)
    If (FSO.FolderExists(dir)) Then
        FSO.DeleteFolder(dir)
    End If
End Sub

'------------------------------------------------
' 拷贝文件
Sub CopyFile(SourceFile, DestinationFile)
    
    Set FSO = CreateObject(COM_FSO)
    
    'Check to see if the file already exists in the destination folder
    Dim wasReadOnly
    wasReadOnly = False
    If FSO.FileExists(DestinationFile) Then
        'Check to see if the file is read-only
        If fso.GetFile(DestinationFile).Attributes And 1 Then 
            'The file exists and is read-only.
            WScript.Echo "Removing the read-only attribute"
            'Remove the read-only attribute
            FSO.GetFile(DestinationFile).Attributes = FSO.GetFile(DestinationFile).Attributes - 1
            wasReadOnly = True
        End If
        
        WScript.Echo "Deleting the file"
        FSO.DeleteFile DestinationFile, True
    End If
    
    'Copy the file
    WScript.Echo "Copying " & SourceFile & " to " & DestinationFile
    FSO.CopyFile SourceFile, DestinationFile, True
    
    If wasReadOnly Then
        'Reapply the read-only attribute
        FSO.GetFile(DestinationFile).Attributes = FSO.GetFile(DestinationFile).Attributes + 1
    End If
    
    Set FSO = Nothing
    
End Sub

'------------------------------------------------
' 桌面文件夹
Function DesktopDir
    Dim objShell
    Set objShell = CreateObject(COM_SHELL)
    DesktopDir = objShell.SpecialFolders("desktop")
    Set objShell = Nothing
End Function


'------------------------------------------------
' 路径末尾添加\
Function DisposePath(sPath)
    On Error Resume Next
    
    If Right(sPath, 1) = "\" Then
        DisposePath = sPath
    Else
        DisposePath = sPath & "\"
    End If
    
    DisposePath = Trim(DisposePath)
End Function 

'------------------------------------------------
' 替换文件内容
Function ReplaceFileContent(filepath, pattern, text, is_utf8)
    Set objFSO = CreateObject(COM_FSO)
    Set objFile = objFSO.GetFile(filepath)
    Dim objStream
    
    If objFile.Size > 0 Then
        
        If is_utf8 = 1 Then			
            Set objStream = CreateObject(COM_ADOSTREAM)
            objStream.Open
            objStream.Type = adTypeText
            objStream.Position = 0
            objStream.Charset = CdoUTF_8
            objStream.LoadFromFile filepath
            strContents = objstream.ReadText
            objStream.Close
            Set objStream = Nothing
        Else
            Set objReadFile = objFSO.OpenTextFile(filepath, 1)
            strContents = objReadFile.ReadAll
            objReadFile.Close
        End If
    End If
    
    Dim re
    Set re = new RegExp
    re.IgnoreCase = False
    re.Global = True
    re.MultiLine = True
    re.Pattern = pattern
    strContents = re.replace(strContents, text)
    
    're.Pattern="^Public\s+Const\s+APP_VERSION.*""$"
    'strContents = re.replace(strContents,"Public Const APP_VERSION = ""Version: " & appversion & """")
    
    Set re = Nothing
    
    If is_utf8 = 1 Then
        Set objStream = CreateObject(COM_ADOSTREAM)
        objStream.Open
        objStream.Type = adTypeText
        objStream.Position = 0
        objStream.Charset = CdoUTF_8
        objStream.WriteText = strContents
        objStream.SaveToFile filepath, adSaveCreateOverWrite
        objStream.Close
        Set objStream = Nothing
    Else
        Set objWriteFile = objFSO.OpenTextFile(filepath, 2, False)
        objWriteFile.Write(strContents)
        objWriteFile.Close
    End If
End Function 

'------------------------------------------------
' 获取桌面路径
Function GetDesktopPath()
    Set objShell = CreateObject(COM_SHELLAPP)
    Set objFolder = objShell.Namespace(DESKTOP)
    Set objFolderItem = objFolder.Self
    GetDesktopPath = objFolderItem.Path
End Function

'------------------------------------------------
' 获取应用程序数据路径
Function GetApplicationDataPath()
    Dim SHELL, folder, folder_item
    Set SHELL = CreateObject(COM_SHELLAPP)
    Set folder = SHELL.Namespace(LOCAL_APPLICATION_DATA)
    Set folder_item = folder.Self
    GetApplicationDataPath = folder_item.Path
End Function 


'------------------------------------------------
' 获取临时文件夹路径
Function GetTempPath()
    Set objShell = CreateObject(COM_SHELLAPP)
    Set objFolder = objShell.Namespace(TEMPORARY_INTERNET_FILES)
    Set objFolderItem = objFolder.Self
    GetTempPath = objFolderItem.Path
End Function 

'------------------------------------------------
' 创建临时文件
Function CreateTempFile(dir)
    Dim FSO, tempname, fullname, file
    Set FSO = CreateObject(COM_FSO)
    tempname = FSO.GetTempName
    fullname = FSO.BuildPath(dir, tempname)
    Set file = FSO.CreateTextFile(fullname)
    file.Close
End Function


'------------------------------------------------
' 选择文件
Function SelectFile
    Dim objDialog
    Set objDialog = CreateObject(COM_COMMONDIALOG)
    objDialog.Filter = "Windows Media 音频(*.wma;*.wav)|*.wma;*.wav|MP3(*.mp3)|*.mp3|All Files(*.*)|*.*"
    objDialog.InitialDir = ScriptPath
    intResult = objDialog.ShowOpen
    If intResult = 0 Then
        SelectFile = ""
    Else
        SelectFile = objDialog.FileName
    End If
End Function


'------------------------------------------------
' 读文本文件
Function ReadFile(ByVal filename)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    
    If InStr(filename, ":\") = 0 And Left(filename, 2) <> "\\" Then 
        filename = FSO.GetSpecialFolder(0) & "\" & filename
    End If
    
    On Error Resume Next
    ReadFile = FSO.OpenTextFile(filename).ReadAll
End Function

'------------------------------------------------
' 写文本文件
Function WriteFile(ByVal filename, ByVal Contents)
    
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    
    If InStr(filename, ":\") = 0 And Left(filename, 2) <> "\\" Then 
        filename = FSO.GetSpecialFolder(0) & "\" & filename
    End If
    
    Dim OutStream
    Set OutStream = FSO.OpenTextFile(filename, 2, True)
    OutStream.Write Contents
End Function

'------------------------------------------------
' 读文本文件到数组
Function ReadFile2Array(ByVal filename)
    Dim arrFileLines(), FSO, file, I
    I = 0    
    Set FSO = CreateObject(COM_FSO)
    Set file = FSO.OpenTextFile(filename, ForReading)    
    Do Until file.AtEndOfStream
        Redim Preserve arrFileLines(i)
        arrFileLines(i) = file.ReadLine
        I = I + 1
    Loop    
    file.Close
    Set FSO = Nothing        
    ReadFile2Array = arrFileLines
End Function


'------------------------------------------------
' 读二进制文件
Function ReadBinary(FileName)    
    Dim adodbStream, xmldom, node
    Set xmldom = CreateObject(COM_XMLDOM)
    Set node = xmldom.CreateElement("binary")
    node.DataType = "bin.hex"
    Set adodbStream = CreateObject(COM_ADOSTREAM)
    adodbStream.Type = adTypeBinary
    adodbStream.Open
    adodbStream.LoadFromFile FileName
    node.NodeTypedValue = adodbStream.Read
    adodbStream.Close
    Set adodbStream = Nothing
    ReadBinary = node.Text
    Set node = Nothing
    Set xmldom = Nothing
End Function

'------------------------------------------------
' 写二进制文件
Function WriteBinary(FileName, Buf)    
    Dim adodbStream, xmldom, node
    Set xmldom = CreateObject(COM_XMLDOM)
    Set node = xmldom.CreateElement("binary")
    node.DataType = "bin.hex"
    node.Text = Buf
    Set adodbStream = CreateObject(COM_ADOSTREAM)
    adodbStream.Type = adTypeBinary
    adodbStream.Open
    adodbStream.write node.NodeTypedValue
    adodbStream.saveToFile FileName, adSaveCreateOverWrite
    adodbStream.Close
    Set adodbStream = Nothing
    Set node = Nothing
    Set xmldom = Nothing
End Function 

'------------------------------------------------
' 读二进制文件到数组
Function ReadBinary2(FileName)
    Dim Buf(), I
    With CreateObject("ADODB.Stream")
        .Mode = 3
        .Type = 1
        .Open
        .LoadFromFile FileName
        ReDim Buf(.Size - 1)
        For I = 0 To .Size - 1
            Buf(I) = AscB(.Read(1))
        Next
        .Close
    End With
    ReadBinary = Buf
End Function

'------------------------------------------------
' 写二进制文件
Sub WriteBinary2(FileName, Buf)
    Dim I, aBuf, Size, bStream
    Size = UBound(Buf): ReDim aBuf(Size \ 2)
    For I = 0 To Size - 1 Step 2
        aBuf(I \ 2) = ChrW(Buf(I + 1) * 256 + Buf(I))
    Next
    If I = Size Then aBuf(I \ 2) = ChrW(Buf(I))
    aBuf = Join(aBuf, "")
    Set bStream = CreateObject("ADODB.Stream")
    bStream.Type = 1: bStream.Open
    With CreateObject("ADODB.Stream")
        .Type = 2 : .Open: .WriteText aBuf
        .Position = 2: .CopyTo bStream: .Close
    End With
    bStream.SaveToFile FileName, 2: bStream.Close
    Set bStream = Nothing
End Sub



'------------------------------------------------
' 读文本文件
Function ReadTextFile(ByVal filename)
    Dim adodbStream, retval, charset
    charset = dected_file_charset(filename)
    If charset = "ANSI" Then
        charset = "GB2312"
    ElseIf charset = "UTF-8" Then
        charset = "UTF-8" 
    ElseIf charset = "UTF8_NOBOM" Then
        charset = "UTF-8" 
    ElseIf charset = "UTF-16LE" Then
        charset = "unicode" 
    ElseIf charset = "UTF-16BE" Then
        charset = "unicode"      
    End If

    If charset <> "None" Then
        Set adodbStream = CreateObject(COM_ADOSTREAM)
        adodbStream.Type = adTypeText 
        adodbStream.mode = adModeReadWrite 
        adodbStream.charset = charset
        adodbStream.Open
        adodbStream.loadfromfile filename
        retval = adodbStream.readtext
        adodbStream.Close
        Set adodbStream = Nothing
        ReadTextFile = retval
    End If
End Function 


'------------------------------------------------
' 写文本文件
Function WriteTextFile(ByVal filename, byval Str, charset) 
    On Error Resume Next
    If charset = "" Then
        If FileExists(filename) Then
            charset = dected_file_charset(filename)           
        End If
        
        If charset = "UTF8_NOBOM" Or charset = "UTF-8" Then   
            charset = "UTF-8"            
        ElseIf charset = "UTF-16L" Or charset = "UTF-16BE" Then
            charset = "unicode"      
        ElseIf charset = "ANSI" Or charset = "" Then
            charset = "GB2312"
        End If
    End If

    If UCASE(charset) = "UTF-8" Then
        WriteTextFile = WriteUTF8WithoutBOM(filename, Str)
        Exit Function
    End If

    Dim adodbStream
    Set adodbStream = CreateObject(COM_ADOSTREAM)
    adodbStream.Type = adTypeText '以本模式读取
    adodbStream.mode = adModeReadWrite
    adodbStream.charset = charset
    adodbStream.Open
    adodbStream.WriteText str
    adodbStream.SaveToFile filename, 2 
    adodbStream.flush
    adodbStream.Close
    Set adodbStream = nothing

    If Err = 0 Then
        WriteTextFile = True
    Else
        WriteTextFile = False
    End If
End Function 

'------------------------------------------------
' 读文本文件
Function pvReadFile(sFile)    
    Dim sPrefix
    
    With CreateObject(COM_FSO)
        sPrefix = .OpenTextFile(sFile, ForReading, False, False).Read(3)
    End With
    If Left(sPrefix, 3) <> Chr(&HEF) & Chr(&HBB) & Chr(&HBF) Then
        With CreateObject(COM_FSO)
            pvReadFile = .OpenTextFile(sFile, ForReading, False, Left(sPrefix, 2) = Chr(&HFF) & Chr(&HFE)).ReadAll()
        End With
    Else
        With CreateObject(COM_ADOSTREAM)
            .Open
            If Left(sPrefix, 2) = Chr(&HFF) & Chr(&HFE) Then
                .Charset = "Unicode"
            ElseIf Left(sPrefix, 3) = Chr(&HEF) & Chr(&HBB) & Chr(&HBF) Then
                .Charset = "UTF-8"
            Else
                .Charset = "_autodetect"
            End If
            .LoadFromFile sFile
            pvReadFile = .ReadText
        End With
    End If
End Function

'------------------------------------------------
' 写文本文件
Function pvWriteFile(sFile, sText, lType)    
    With CreateObject(COM_ADOSTREAM)
        .Open
        If lType = 2 Then
            .Charset = "Unicode"
        ElseIf lType = 3 Then
            .Charset = "UTF-8"
        Else
            .Charset = "_autodetect"
        End If
        .WriteText sText
        .SaveToFile sFile, adSaveCreateOverWrite
    End With
End Function

'------------------------------------------------
' 写utf-8无BOM文件
Function WriteUTF8WithoutBOM(ByVal filename, ByRef content)
    On Error Resume Next
    dim stm:set stm = CreateObject(COM_ADOSTREAM)   
    stm.Type = 2 '以文本模式读取   
    stm.mode = 3   
    stm.charset = "utf-8"  
    stm.open   
    stm.Writetext(content)   
    stm.Position = 3   
    dim newStream : Set newStream = CreateObject(COM_ADOSTREAM)   
    With newStream   
        .Mode = 3   
        .Type = 1   
        .Open()   
    End With  
    stm.CopyTo(newStream)   
    newStream.SaveToFile filename, 2   
    stm.flush   
    stm.Close   
    Set stm = Nothing  
    Set newStream = Nothing  
    If Err = 0 Then
        WriteUTF8WithoutBOM = True
    Else
        WriteUTF8WithoutBOM = False
    End If
End Function

'------------------------------------------------
' 使用IE保存UTF8无BOM的文件
Function RemoveUTF8BOM_IE(filename, content)
    Dim ie : Set ie = CreateObject(COM_IE) 
    ie.Navigate filename
    Do While ie.Busy Or ie.ReadyState<>4
        WScript.Sleep 100
    Loop
    If ie.Document.Charset="utf-8" Then 
        ie.ExecWB OLECMDID_SAVE,OLECMDEXECOPT_DODEFAULT
    End If
    ie.Quit
End Function


'------------------------------------------------
' 字符串转字节数组
Function Str2Bytes(str, charset)
    Dim adodbStream, strRet 
    Set adodbStream = CreateObject(COM_ADOSTREAM)     
    adodbStream.Type = adTypeText              
    adodbStream.Charset = charset    
    adodbStream.Open                     
    adodbStream.WriteText str                  
    adodbStream.Position = 0         
    adodbStream.Type = adTypeBinary        
    vout = adodbStream.Read(adodbStream.Size)   
    adodbStream.Close                
    Set adodbStream = nothing 
    Str2Bytes = vout 
End Function

'------------------------------------------------
' 字节数组转字符串
Function BytesToBstr(str, charset)
    If LenB(str) = 0 Then  
        BytesToBstr = "" 
        Exit Function 
    End If 
    
    Dim adodbStream 
    Set adodbStream = CreateObject(COM_ADOSTREAM) 
    adodbStream.Type = adTypeBinary 
    adodbStream.Mode = adModeReadWrite 
    adodbStream.Open 
    adodbStream.Write str 
    adodbStream.Position = 0 
    adodbStream.Type = adTypeText 
    adodbStream.Charset = charset 
    BytesToBstr = adodbStream.ReadText 
    adodbStream.Close 
    Set adodbStream = nothing 
End Function

'------------------------------------------------
' String2Bytes
Function String2Bytes(str)
    Dim k,char,code,bytes
    For k=1 To Len(str)
        char=Mid(str,k,1)
        code=Asc(char)
        If code<0 Then code=code+256*256
        If code<256 Then
            bytes=bytes & ChrB(code)
        Else
            bytes=bytes & ChrB(code\256) & ChrB(code Mod 256)
        End If
    Next
    String2Bytes=bytes
End Function


'------------------------------------------------
' toUTF8
'Function toUTF8 (szInput) 
'    Dim wch, uch, szRet 
'    Dim x 
'    Dim nAsc, nAsc2, nAsc3 
'    'If the input parameter is empty, then exit the function 
'    If szInput = "" Then 
'        toUTF8 = szInput 
'        Exit Function 
'    End If 
'    'Start conversion 
'    For x = 1 To Len (szInput) 
'        'Mid function split GB encoded text 
'        wch = Mid (szInput, x, 1) 
'        'To use AscW function returns a GB encoded text Unicode character code 
'        'Note: The the asc function returns the ANSI character code, note the difference 
'        NASC = AscW (WCH) 
'        If nAsc <0 Then nAsc = nAsc + 65536 
'        
'        If (nAsc And & HFF80) = 0 Then 
'            szRet = szRet & wch 
'        Else 
'            If (nAsc And & HF000) = 0 Then 
'                uch = "%" & Hex (((nAsc / 2 ^ 6)) Or & HC0) & Hex (nAsc And & H3F Or & H80) 
'                szRet = szRet & uch 
'            Else 
'                'GB encoded text Unicode character code in 0800 - FFFF three-byte template 
'                uch = "%" & Hex ((nAsc / 2 ^ 12) Or & HE0) & "%" & _ 
'                Hex ((nAsc / 2 ^ 6) And & H3F Or & H80) & "%" & _ 
'                Hex (nAsc And & H3F Or & H80) 
'                szRet = szRet & uch 
'            End If 
'        End If 
'    Next 
'    
'    toUTF8 = szRet 
'End Function

'------------------------------------------------
' chinese2unicode
' GB transfer unicode --- GB encoded text is converted to Unicode encoded text
Function chinese2unicode(Str) 
    Dim i 
    Dim Str_one 
    Dim Str_unicode 
    If (IsNull (Str)) Then 
        Exit Function 
    End If 
    For i = 1 To Len (STR) 
        Str_one = Mid (Str, i, 1) 
        Str_unicode = Str_unicode & Chr (38) 
        Str_unicode = Str_unicode & Chr (35) 
        Str_unicode = Str_unicode & Chr (120) 
        Str_unicode = Str_unicode & Hex (AscW (Str_one)) 
        Str_unicode = Str_unicode & Chr (59) 
    Next 
    chinese2unicode = Str_unicode 
End Function

'------------------------------------------------
' URLDecode
Function URLDecode (enStr) 
    Dim deStr 
    Dim c, i, v 
    deStr = "" 
    For i = 1 To Len (enStr) 
        c = Mid (enStr, i, 1) 
        If c = "%" Then 
            v = Eval ("& h" + Mid (enStr, i +1,2)) 
            If v <128 Then 
                deStr = deStr & Chr (v) 
                i = i +1 
            Else 
                If isvalidhex (Mid (enstr, i, 3)) Then 
                    If isvalidhex (Mid (enstr, i +3,3)) Then 
                        v = Eval ("& h" + Mid (enStr, i +1,2) + Mid (enStr, i +4,2)) 
                        deStr = deStr & Chr (v) 
                        i = i +5 
                    Else 
                        v = Eval ("& h" + Mid (enStr, i +1,2) + CStr (Hex (Asc (Mid (enStr, i +3,1))))) 
                        deStr = deStr & Chr (v) 
                        i = i +3 
                    End If 
                Else 
                    destr = destr & c 
                End If 
            End If 
        Else 
            If c = "+" Then 
                deStr = deStr & "" 
            Else 
                deStr = deStr & c 
            End If 
        End If 
    Next 
    URLDecode = deStr 
End Function

'------------------------------------------------
' isvalidhex
' To determine whether a valid hexadecimal code 
'Function isvalidhex (str) 
'    Dim c 
'    isvalidhex = True 
'    str = UCase (str) 
'    If Len (str) <> 3 Then isvalidhex = False: Exit Function 
'    If Left (str, 1) <> "%" Then isvalidhex = False: Exit Function 
'    c = Mid (str, 2,1) 
'    If Not (((c> = "0") And (c <= "9")) Or ((c> = "A") And (c <= "Z"))) Then isvalidhex = False: Exit Function 
'    c = Mid (str, 3,1) 
'    If Not (((c> = "0") And (c <= "9")) Or ((c> = "A") And (c <= "Z"))) Then isvalidhex = False: Exit Function 
'End Function 


'------------------------------------------------
' UTF2GB
'Function UTF2GB (UTFStr)
'    
'    For Dig = 1 To Len (UTFStr) 
'        'If the UTF8 encoded text% at the beginning of the conversion 
'        If Mid (UTFStr, Dig, 1) = "%" Then 
'            'UTF8 encoded text more than eight converted into Chinese characters 
'            If Len (UTFStr)> = Dig +8 Then 
'                GBStr = GBStr & ConvChinese (Mid (UTFStr, Dig, 9)) 
'                Dig = Dig +8 
'            Else 
'                GBStr = GBStr & Mid (UTFStr, Dig, 1) 
'            End If 
'        Else 
'            GBStr = GBStr & Mid (UTFStr, Dig, 1) 
'        End If 
'    Next 
'    UTF2GB = GBStr 
'End Function

'------------------------------------------------
' ConvChinese
' UTF8 encoded text will be converted to Chinese characters 
Function ConvChinese (x) 
    A = Split (Mid (x, 2), "%") 
    i = 0 
    j = 0 
    For i = 0 To UBound (A) 
        A (i) = c16to2 (A (i)) 
    Next 
    For i = 0 To UBound (A) -1 
        DigS = InStr (A (i), "0") 
        Unicode = "" 
        For j = 1 To DigS-1 
            If j = 1 Then 
                A (i) = Right (A (i), Len (A (i))-DIGS) 
                Unicode = Unicode & A (i) 
            Else 
                i = i +1 
                A (i) = Right (A (i), Len (A (i)) -2) 
                Unicode = Unicode & A (i) 
            End If 
        Next
        
        If Len (c2to16 (Unicode)) = 4 Then 
            ConvChinese = ConvChinese & chrw (Int ("& H" & c2to16 (Unicode))) 
        Else 
            ConvChinese = ConvChinese & Chr (Int ("& H" & c2to16 (Unicode))) 
        End If 
    Next 
End Function

'------------------------------------------------
' c2to16
' Binary code into hex code 
Function c2to16 (x) 
    i = 1 
    For i = 1 To Len (x) Step 4 
        c2to16 = c2to16 & Hex (c2to10 (Mid (x, i, 4))) 
    Next 
End Function

'------------------------------------------------
' c2to10
' Binary code converted to decimal code 
Function c2to10 (x) 
    c2to10 = 0 
    If x = "0" Then Exit Function 
    i = 0 
    For i = 0 To Len (x) -1 
        If Mid (x, Len (x)-i, 1) = "1" Then c2to10 = c2to10 +2 ^ (i) 
    Next 
End Function

'------------------------------------------------
' c16to2
' Hexadecimal code is converted to binary code 
'Function c16to2 (x) 
'    i = 0 
'    For i = 1 To Len (Trim (x)) 
'        tempstr = c10to2 (CInt (Int ("& h" & Mid (x, i, 1)))) 
'        Do While Len (tempstr) <4 
'            tempstr = "0" & ??tempstr 
'        Loop 
'        c16to2 = c16to2 & tempstr 
'    Next 
'End Function

'------------------------------------------------
' c10to2
' Decimal code is converted into a binary code 
'Function c10to2 (x) 
'    mysign = Sgn (x) 
'    x = Abs (x) 
'    DIGS = 1 
'    Do 
'        If x <2 ^ DigS Then 
'            Exit Do 
'        Else 
'            DigS = DigS +1 
'        End If 
'    Loop 
'    tempnum = x
'    
'    i = 0 
'    For i = DigS To 1 Step-1 
'        If tempnum> = 2 ^ (i-1) Then 
'            tempnum = tempnum-2 ^ (i-1) 
'            c10to2 = c10to2 & "1" 
'        Else 
'            c10to2 = c10to2 & "0" 
'        End If 
'    Next 
'    If mysign = -1 Then c10to2 = "-" & c10to2 
'End Function


'------------------------------------------------
' 获取正则匹配内容
Function GetMatchText(filename, pattern)
    Dim text, re, matches, tmpstr
    text = ReadTextFile(filename, "gb2312")
    
    Set re = new RegExp
    re.IgnoreCase = False
    re.Global = True
    re.MultiLine = True
    re.Pattern = pattern
    
    Set matches = re.Execute(text)
    If matches.Count > 0 Then
        For Each m In matches
            If m.SubMatches.Count > 0 Then
                GetMatchText = m.SubMatches(0)
            End If
        Next
    End If
End Function 


'------------------------------------------------
' 获取文件行数
Function GetFileLines(filename)
    Dim FSO, file
    Set FSO = CreateObject(COM_FSO)
    Set file = FSO.OpenTextFile(filename, ForReading)
    ' Skip lines one by one
    Do While file.AtEndOfStream <> True
        file.SkipLine
    Loop
    
    GetFileLines = file.Line
    
    Set FSO = Nothing
End Function



'------------------------------------------------
' 遍历文件夹
'	Function testfile(filename)
'		WScript.Echo filename
'	End Function
'
'	Call EachFiles("D:\tools\7-Zip", "\.txt", "testfile")
Sub EachFiles(dir, pattern, method)
    Dim FSO, re
    Set FSO = CreateObject(COM_FSO)
    Set root = FSO.GetFolder(dir)
    Set re = new RegExp
    re.Pattern    = pattern
    re.IgnoreCase = True
    
    Call EachSubFolder(root, re, method)
    
    Set FSO = Nothing
    Set re = Nothing
End Sub

Sub EachSubFolder(root, re, method)
    Dim subfolder, file, script
    
    For Each file In root.Files
        If re.Test(file.Name) Then
            script = "Call " & method & "(""" & file.Path & """)"
            ExecuteGlobal script
        End If
    Next
    
    For Each subfolder In root.SubFolders
        Call EachSubFolder(subfolder, re, method)    
    Next
End Sub

'------------------------------------------------
' 根据原文件名，自动以日期YYYY-MM-DD-RANDOM格式生成新文件名
Function GetfileExt(byval filename)
    Dim fileExt_a
    fileExt_a = Split(filename,".")
    GetfileExt = Lcase(fileExt_a(Ubound(fileExt_a)))
End Function

'------------------------------------------------
' 根据原文件名，自动以日期YYYY-MM-DD-RANDOM格式生成新文件名
Function GenerateRandomFileName(ByVal filename)
    Randomize
    ranNum = Int(90000 * Rnd) + 10000
    If Month(Now) < 10 Then c_month = "0" & Month(Now) Else c_month = Month(Now)
    If Day(Now) < 10 Then c_day = "0" & Day(Now) Else c_day = Day(Now)
    If Hour(Now) < 10 Then c_hour = "0" & Hour(Now) Else c_hour = Hour(Now)
    If Minute(Now) < 10 Then c_minute = "0" & Minute(Now) Else c_minute = Minute(Now)
    If Second(Now) < 10 Then c_second = "0" & Second(Now) Else c_second = Minute(Now)
    fileExt_a = Split(filename, ".")
    FileExt = LCase(fileExt_a(UBound(fileExt_a)))
    GenerateRandomFileName = Year(Now) & c_month & c_day & c_hour & c_minute & c_second & "_" & ranNum & "." & FileExt
End Function


'------------------------------------------------
' 建立目录的程序，如果有多级目录，则一级一级的创建
Function CreateDir(ByVal LocalPath) 
    On Error Resume Next
    Dim FSO
    LocalPath = Replace(LocalPath, "\", "/")
    Set FSO = CreateObject(COM_FSO)
    patharr = Split(LocalPath, "/")
    path_level = UBound(patharr)
    For I = 0 To path_level
        If I = 0 Then pathtmp = patharr(0) & "/" Else pathtmp = pathtmp & patharr(I) & "/"
        cpath = Left(pathtmp, Len(pathtmp) - 1)
        If Not FSO.FolderExists(cpath) Then FSO.CreateFolder cpath
    Next
    Set FSO = Nothing
    If Err.Number <> 0 Then
        CreateDir = False
        Err.Clear
    Else
        CreateDir = True
    End If
End Function

'------------------------------------------------
' 移除utf8 bom
Sub RemoveBOM(INfile)
    Dim OUTfile
    Dim fileName
    
    Dim dirPos
    Dim pathOUT
    Dim OUTName
    
    Dim oADOST_R     'As Object
    Dim readPos      'As Long ' or Currency or Double
    
    Dim oADOST_W     'As Object
    
    Dim dathead01
    Dim dathead02
    Dim dathead03
    
    dirPos   = InStrRev(INfile,"\")
    pathOUT  = Left(INfile,dirPos)
    fileName = Mid(INfile , dirPos + 1 , 999)
    
    OUTName  = "_NoBom_" & fileName
    OUTfile  = pathOUT & OUTName
    
    Set oADOST_R = CreateObject("ADODB.Stream")
    
    oADOST_R.Type = 1   '1=adTypeBinary 2=adTypeText
    oADOST_R.Open
    oADOST_R.LoadFromFile INfile
    readPos = 0
    oADOST_R.Position = readPos '読込開始位置
    
    dathead01  = UCase(Right("0" & Hex(AscB(oADOST_R.Read(1) )) , 2)) 
    readPos = readPos + 1
    dathead02  = UCase(Right("0" & Hex(AscB(oADOST_R.Read(1) )) , 2))  
    readPos = readPos + 1
    dathead03  = UCase(Right("0" & Hex(AscB(oADOST_R.Read(1) )) , 2)) 
    
    If  ( dathead01 = "FF" And dathead02 = "FE" ) Or _
        ( dathead01 = "FE" And dathead02 = "FF" ) Then
        readPos = 2  'utf-16
    ElseIf ( dathead01 = "EF"   And _
        dathead02 = "BB"   And _
        dathead03 = "BF" ) Then
        readPos = 3  'utf-8
    Else
        readPos = 0  'BOM無し
    End If
    oADOST_R.Position = readPos '読込開始位置
    
    '書込Object設定
    Set oADOST_W = CreateObject("ADODB.Stream")
    oADOST_W.Mode = 3
    oADOST_W.Type = 1 '1=adTypeBinary 2=adTypeText
    'oADOST_W.Charset = "utf-8"
    'oADOST_W.Charset = "iso-8859-1" 'キャラクタセット＝Latin-1
    oADOST_W.Open
    
    oADOST_R.CopyTo(oADOST_W)   
    
    '既にファイルが存在する場合　1=実行時エラー、2=上書保存
    oADOST_W.SaveToFile OUTfile, 2
    
    oADOST_R.Close
    oADOST_W.Close
    Set oADOST_R = Nothing
    Set oADOST_W = Nothing
End Sub

'------------------------------------------------
' 搜索文件
Function SearchFileByExt(ext)
    Dim computer, objFile, colFiles
    computer = "."
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" _
        & computer & "\root\cimv2")
    Set colFiles = objWMIService. _
        ExecQuery("Select * from CIM_DataFile where Extension = '" & ext & "'")
    For Each objFile in colFiles
        Wscript.Echo objFile.Name
    Next
End Function

'------------------------------------------------
' 检测文件编码
' https://en.wikipedia.org/wiki/Byte_order_mark
Function dected_file_charset(ByVal filename)
    Dim ado_stream, readpos, datahead, charset, dict, byte_array, length, i
    Set dict = CreateObject(COM_DICT)
    readpos = 0
    Set ado_stream = CreateObject(COM_ADOSTREAM)
    ado_stream.Type = adTypeBinary   
    ado_stream.Open
    ado_stream.LoadFromFile filename    
    ado_stream.Position = readpos    

    If ado_stream.Size > 2 Then
        datahead = ado_stream.Read(4)

        head4bytes = Hex(AscB(MidB(datahead, 1, 1))) _
			& Hex(AscB(MidB(datahead, 2, 1))) _
			& Hex(AscB(MidB(datahead, 3, 1))) _
            & Hex(AscB(MidB(datahead, 4, 1)))

        head3bytes = Hex(AscB(MidB(datahead, 1, 1))) _
			& Hex(AscB(MidB(datahead, 2, 1))) _
			& Hex(AscB(MidB(datahead, 3, 1)))

        head2bytes = Hex(AscB(MidB(datahead, 1, 1))) _
			& Hex(AscB(MidB(datahead, 2, 1))) 
        
        If head4bytes = "FFFE0000" Then
            charset = "UTF-32LE"
        ElseIf head4bytes = "0000FEFF" Then
            charset = "UTF-32BE"
        ElseIf head4bytes = "2B2F7638" Or head4bytes = "2B2F7639" Or head4bytes = "2B2F762B" Or head4bytes = "2B2F762F" Then
            charset = "UTF-7"
        ElseIf head3bytes = "EFBBBF" Then
            charset = "UTF-8"
        ElseIf head3bytes = "F7644C" Then
            charset = "UTF-1"
        ElseIf head2bytes = "FFFE" Then
            charset = "UTF-16LE"
        ElseIf head2bytes = "FEFF" Then
            charset = "UTF-16BE"
        Else
            charset = "None"
        End If
    End If

    If charset <> "None" Then
        dected_file_charset = charset
        Exit Function
    End If


    ado_stream.Position = 0
    Do Until ado_stream.EOS
        dict.Add dict.Count, ado_stream.Read(1)
    Loop

    byte_array = dict.Items
    length = UBound(byte_array)
    For i = 0 To length
        byte_array(i) = AscB(byte_array(i))
    Next

    charset = CheckUTF8(byte_array)  
    If charset <> "None" Then
        dected_file_charset = charset
        Exit Function
    End If

    ' ANSI or None (binary) then
    If Not DoesContainNulls(byte_array) Then
        charset = "ANSI"
    Else
        charset = "None"
    End If

    dected_file_charset = charset

End Function

Function CheckUTF8(ByRef byte_array)
    ' UTF8 Valid sequences
	' 0xxxxxxx  ASCII
	' 110xxxxx 10xxxxxx  2-byte
	' 1110xxxx 10xxxxxx 10xxxxxx  3-byte
	' 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx  4-byte
	'
	' Width in UTF8
	' Decimal		Width
	' 0-127		    1 byte
	' 194-223		2 bytes
	' 224-239		3 bytes
	' 240-244		4 bytes
	'
	' Subsequent chars are in the range 128-191

    Dim pos, length, ch, more_chars, only_saw_ascii_range 
    only_saw_ascii_range = TRUE
    pos = 0
    length = UBound(byte_array)
    Do While pos < length 
        
        ch = byte_array(pos)
        pos = pos + 1

        If ch = 0 Then
            CheckUTF8 = "None"            
            Exit Function
        ElseIf ch <= 127 Then		
			' 1 byte
			more_chars = 0	
		ElseIf ch >= 194 AND ch <= 223 Then		
			' 2 Byte
			more_chars = 1	
		ElseIf ch >= 224 AND ch <= 239 Then
			' 3 Byte
			more_chars = 2		
		ElseIf ch >= 240 AND ch <= 244 Then		
			' 4 Byte
			more_chars = 3		
		Else		
			CheckUTF8 = "None"						' Not utf8
            Exit Function
		End If

        ' Check secondary chars are in range if we are expecting any
		Do While more_chars AND pos < length		
			only_saw_ascii_range = FALSE		' Seen non-ascii chars now

			ch = byte_array(pos)
            pos = pos + 1
			If ch < 128 Or ch > 191 Then
				CheckUTF8 = "None"					' Not utf8                
                Exit Function
            End If

			more_chars = more_chars - 1
		Loop

    Loop

    ' If we get to here then only valid UTF-8 sequences have been processed

	' If we only saw chars in the range 0-127 then we can't assume UTF8 (the caller will need to decide)
	If only_saw_ascii_range Then
		CheckUTF8 =  "ASCII"
	Else        
		CheckUTF8 =  "UTF8_NOBOM"
    End If
End Function


Function CheckUTF16NewlineChars(ByRef byte_array)

    Dim pos, length, ch1, ch2, le_control_chars, be_control_chars 
    
    pos = 0
    length = UBound(byte_array)

	If length < 2 Then
		CheckUTF16NewlineChars = "None"
        Exit Function
    End If

	' Reduce size by 1 so we don't need to worry about bounds checking for pairs of bytes
	'size--;

	le_control_chars = 0
	be_control_chars = 0	

	pos = 0
	Do While pos < length
		ch1 = byte_array(pos)
        pos = pos + 1
		ch2 = byte_array(pos)
        pos = pos + 1

		If ch1 = 0 Then		
			If (ch2 = &H0A) Or (ch2 = &H0D) Then
                be_control_chars = be_control_chars + 1
            End If
		
		ElseIf ch2 = 0 Then		
			If (ch1 = &H0A) Or (ch1 = &H0D) Then
                le_control_chars = le_control_chars + 1				
            End If
		End If

		' If we are getting both LE and BE control chars then this file is not utf16
		If le_control_chars And be_control_chars Then
			CheckUTF16NewlineChars = "None"
            Exit Function
        End If
	Loop

	If le_control_chars Then
		CheckUTF16NewlineChars = "UTF16_LE_NOBOM"
	ElseIf be_control_chars Then
		CheckUTF16NewlineChars = "UTF16_BE_NOBOM"
	Else
		CheckUTF16NewlineChars = "None"
    End If
End Function

'------------------------------------------------
' DoesContainNulls
' Checks if a buffer contains any nulls. Used to check for binary vs text data.
Function DoesContainNulls(ByRef byte_array)
    Dim pos, length
    pos = 0
    length = UBound(byte_array)	
	Do While pos < length	
		If byte_array(pos) = 0 Then
			DoesContainNulls = True
            Exit Function
        End If

        pos = pos + 1
	Loop

	DoesContainNulls = False
End Function




Function UnicodeToUTF8(ByRef pstrUnicode)
    ' Written 2007 by Alexander Klink for the OpenXPKI Project
    ' (c) 2007 by the OpenXPKI Project, released under the Apache License v2.0
    ' converts a unicode string to UTF8 (well, sort of)
    ' reference: http://en.wikipedia.org/wiki/UTF8
    Dim i, result

    result = ""
    For i = 1 To Len(pStrUnicode)
        CurrentChar = Mid(PstrUnicode, i, 1)
        CodePoint = AscW(CurrentChar)

        If (CodePoint < 0) Then
            ' AscW is broken. Badly. It can only return an integer,
            ' which is 32767 at most. So everything up to 65535 is
            ' AscW() + 65536. That Unicode chars exist beyond 65535
            ' is apparently unknown to Microsoft ...
            CodePoint = CodePoint + 65536
        End If

        MaskSixBits   = 2^6 - 1 ' the lower 6 bits are 1
        MaskFourBits  = 2^4 - 1 ' the lower 4 bits are 1
        MaskThreeBits = 2^3 - 1 ' the lower 3 bits are 1
        MaskTwoBits   = 2^2 - 1 ' the lower 3 bits are 1

        'MsgBox CurrentChar & " : " & CodePoint
        If (CodePoint >= 0) And (CodePoint < 128) Then
            ' for codepoints < 128, just add one byte with the
            ' value of the codepoint (this is the ASCII subset)
            Zs = CodePoint
            result = result & ChrB(Zs)
        End If
        ' this is common for all of the following
        Zs = CodePoint And MaskSixBits
        If (CodePoint >= 128) And (CodePoint < 2048) Then
            ' for naming, see the Wikipedia article referenced above
            Ys = RightShift(CodePoint, 6)
            FirstByte  = LeftShift(6, 5) Xor Ys ' 110yyyy 
            SecondByte = LeftShift(2, 6) Xor Zs ' 10zzzzz
            'MsgBox "Case 1: " & FirstByte & ", " & SecondByte
            result = result & ChrB(FirstByte) & ChrB(SecondByte)
        End If
        If (CodePoint >= 2048) And (CodePoint < 65536) Then
            Ys = RightShift(CodePoint, 6) And MaskSixBits
            Xs = RightShift(CodePoint, 12) And MaskFourBits
            FirstByte  = LeftShift(14, 4) Xor Xs ' 1110xxxx
            SecondByte = LeftShift(2, 6) Xor Ys  ' 10yyyyyy
            ThirdByte  = LeftShift(2, 6) Xor Zs  ' 10zzzzzz
            'MsgBox "Case 2: " & FirstByte & ", " & SecondByte & ", " & ThirdByte
            result = result & ChrB(FirstByte) & ChrB(SecondByte) & ChrB(ThirdByte)
        End If
    Next
    UnicodeToUTF8 = result
End Function


'------------------------------------------------
' 目录索引
Class IndexDir
    Private dir_path_
    Private Sub Class_Initialize        
    End Sub

	Private Sub Class_Terminate		
    End Sub

	Public Default Function Init(dir_path)	
        If dir_path = "" Then
            Dim fso
            set fso = createobject(COM_FSO)
            dir_path_ = fso.GetAbsolutePathName(".")
            set fso = nothing
        Else
            dir_path_ = dir_path
        End If
		Set Init = Me
    End Function   
    
    Public Function Index()        
        GetWorkingFolder dir_path_, 0, 1, "|"
    End Function

    ' called recursively to get a folder to work in
    Private function GetWorkingFolder(foldspec, foldcount, firsttime, spacer)

        dim fso
        Set fso = CreateObject(COM_FSO)

        dim fold
        set fold = fso.GetFolder(foldspec)
        
        dim foldcol
        set foldcol = fold.SubFolders
        
        if firsttime = 1 then
            wscript.echo fold.name
            spacer = ""
            foldcount = foldcol.count
            firsttime = 0
        end if
        
        dim remaincount
        remaincount = foldcol.count
        
        dim sf
        for each sf in foldcol
                    
            spacer = spacer + space(3) + "|"
            
            wscript.echo spacer + "-- " + sf.name 
            
            ' If you wanted to show the number of bytes, use this line instead of above
            'wscript.echo spacer + "-- " + sf.name + " (uses " + cstr(FormatNumber(sf.size)) + " bytes)"

            if remaincount = 1 then
                spacer = left(spacer, len(spacer) - 1)
                spacer = spacer + " "
            end if
            
            '
            ' if you want to do something more useful, put that function call, or just
            ' insert the code, here.
            '
            
            remaincount = GetWorkingFolder (foldspec +"\"+sf.name, remaincount, firsttime, spacer)
        
        next 
            
        if len(spacer) > 3 then
            spacer = left(spacer, len(spacer) - 4)
        end if
        
        set fso = nothing
        
        GetWorkingFolder = foldcount - 1

    end function



End Class 