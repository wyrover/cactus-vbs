'------------------------------------------------
' ²¥·ÅMP3
Function PlayMp3(FileName)
    Dim objWWP, objShell 
    Set objShell = CreateObject(COM_SHELL)
    Set objWMP = CreateObject(COM_WMP)
    objWMP.url = FileName
    Do Until objWMP.playState = 1
        objShell.Sleep 100
    Loop
    Set objShell = Nothing
    Set objWMP = Nothing
End Function