'------------------------------------------------
' NewZip
Sub NewZip(filename) 
    'WScript.Echo "Newing up a zip file (" & pathToZipFile & ") "
    
    Dim FSO, file
    Set FSO = CreateObject(COM_FSO)   
    Set file = FSO.CreateTextFile(filename)
    
    file.Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0) 
    file.Close
    Set FSO = Nothing
    Set FSO = Nothing 
    WScript.Sleep 500 
End Sub

'------------------------------------------------
' CreateZip         空目录无法压缩
' Example:
'   CreateZip "results.zip", "results"
Sub CreateZip(filename, dir) 
    'WScript.Echo "Creating zip  (" & pathToZipFile & ") from (" & dirToZip & ")"
    
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    
    filename = FSO.GetAbsolutePathName(filename)
    dir = FSO.GetAbsolutePathName(dir)
    
    If FSO.FileExists(filename) Then
        'WScript.Echo "That zip file already exists - deleting it."
        FSO.DeleteFile filename
    End If
    
    If Not FSO.FolderExists(dir) Then
        'WScript.Echo "The directory to zip does not exist."
        Exit Sub
    End If
    
    NewZip filename
    
    Dim SHELLAPP, zip, d
    Set SHELLAPP = CreateObject(COM_SHELLAPP)   
    Set zip = SHELLAPP.NameSpace(filename) 
    
    'WScript.Echo "opening dir  (" & dir & ")" 
    
    Set d = SHELLAPP.NameSpace(dir)
    
    ' Look at http://msdn.microsoft.com/en-us/library/bb787866(VS.85).aspx
    ' for more information about the CopyHere function.
    zip.CopyHere d.items, 4
    
    Do Until d.Items.Count <= zip.Items.Count
        Wscript.Sleep(200)
    Loop
    
End Sub

'------------------------------------------------
' ExtractFilesFromZip
' Example:
'   ExtractFilesFromZip "results.zip", "."
Sub ExtractFilesFromZip(filename, dir)
    
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    
    filename = fso.GetAbsolutePathName(filename)
    dir = fso.GetAbsolutePathName(dir)
    
    If (Not fso.FileExists(filename)) Then
        WScript.Echo "Zip file does not exist: " & filename
        Exit Sub
    End If
    
    If Not fso.FolderExists(dir) Then
        WScript.Echo "Directory does not exist: " & dir
        Exit Sub
    End If
    
    Dim SHELLAPP, zip, d
    set SHELLAPP = CreateObject("Shell.Application")   
    Set zip = SHELLAPP.NameSpace(filename)  
    Set d = SHELLAPP.NameSpace(dir)
    
    ' Look at http://msdn.microsoft.com/en-us/library/bb787866(VS.85).aspx
    ' for more information about the CopyHere function.
    d.CopyHere zip.items, 4
    
    Do Until zip.Items.Count <= d.Items.Count
        Wscript.Sleep(200)
    Loop
    
End Sub

'------------------------------------------------
' ZipBy7Zip
' @archive_file_name        压缩文件名
' @filelist                 文件列表
' Example:
'   Call ZipBy7Zip("results_01.zip", "111.txt 222.txt") 文件列表
'   Call ZipBy7Zip("files.zip", """c:\program files\text files\*.txt""") 文件列表
'   Call ZipBy7Zip("resutls_02.zip", "dadfasd")     文件夹
Function ZipBy7Zip(archive_file_name, filelist)
    Dim FSO, SHELL, sWorkingDirectory
    Set FSO = CreateObject(COM_FSO)
    Set SHELL = CreateObject(COM_SHELL)   
    
    sWorkingDirectory = FSO.GetParentFolderName(Wscript.ScriptFullName) 
    
    '-------Ensure we can find 7za.exe------
    If FSO.FileExists(sWorkingDirectory & "\" & "7z.exe") Then
        s7zLocation = ""
    ElseIf FSO.FileExists("D:\tools\7-Zip\7z.exe") Then
        s7zLocation = "D:\tools\7-Zip\"
    Else
        ZipBy7Zip = "Error: Couldn't find 7za.exe"
        Exit Function
    End If
    '--------------------------------------
    
    SHELL.Run """" & s7zLocation & "7z.exe"" a -tzip -y """ & archive_file_name & """ " _
    & filelist, 0, True   
    
    If FSO.FileExists(archive_file_name) Then
        ZipBy7Zip = 1
    Else
        ZipBy7Zip = "Error: Archive Creation Failed."
    End If
End Function

'------------------------------------------------
' UnZipBy7Zip
' @archive_file_name        压缩文件名
' @dir                      解压目录
' Example:
'   Call UnZipBy7Zip("results_01.zip", "C:\ddddd\dddd\ddd")
Function UnZipBy7Zip(archive_file_name, dir)  
    Dim FSO, SHELL, sWorkingDirectory
    Set FSO = CreateObject(COM_FSO)
    Set SHELL = CreateObject(COM_SHELL)   
    
    sWorkingDirectory = FSO.GetParentFolderName(Wscript.ScriptFullName) 
    '--------------------------------------
    
    '-------Ensure we can find 7za.exe------
    If FSO.FileExists(sWorkingDirectory & "\" & "7z.exe") Then
        s7zLocation = ""
    ElseIf FSO.FileExists("D:\tools\7-Zip\7z.exe") Then
        s7zLocation = "D:\tools\7-Zip\"
    Else
        UnZipBy7Zip = "Error: Couldn't find 7za.exe"
        Exit Function
    End If
    '--------------------------------------
    
    '-Ensure we can find archive to uncompress-
    If Not FSO.FileExists(archive_file_name) Then
        UnZipBy7Zip = "Error: File Not Found."
        Exit Function
    End If
    '--------------------------------------
    
    SHELL.Run """" & s7zLocation & "7z.exe"" e -y -o""" & dir & """ """ & _
    archive_file_name & """", 0, True
    UnZipBy7Zip = 1
End Function