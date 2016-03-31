'------------------------------------------------
' 打印字符串
Sub Echo(message)
    WScript.Echo message
End Sub

'------------------------------------------------
' 打印字符串，带换行符
Sub Println(message)
    Dim stdout
    Set stdout = WScript.StdOut
    stdout.WriteLine message
End Sub

'------------------------------------------------
' CScriptRun
Sub CScriptRun 
    Dim Args
    Dim Arg
    If LCase(Right(WScript.FullName,11)) = "wscript.exe" Then
        Args = Array("cmd.exe /k CScript.exe", """" & WScript.ScriptFullName & """" )
            For Each Arg In WScript.Arguments
            ReDim Preserve Args(UBound(Args)+1)
            Args(UBound(Args)) = """" & Arg & """"
        Next
        WScript.Quit CreateObject("WScript.Shell").Run(Join(Args), 1, True)
    End If
End Sub


'------------------------------------------------
' 不需要实时输出，执行，返回errorCode
Function Run(Cmd)
    Dim objShell, errorCode
    Set objShell = CreateObject(COM_SHELL)
    errorCode = objShell.Run(Cmd, 0, True)
    Run = errorCode
    Set objShell = Nothing
End Function

'------------------------------------------------
' 执行，实时输出
Sub Exec(Cmd)
    Dim objShell, objExec, comspec
    Set objShell = CreateObject(COM_SHELL)	
    comspec = objShell.ExpandEnvironmentStrings("%comspec%")
    Set objExec = objShell.Exec(comspec & " /c ipconfig")
    Do
        WScript.StdOut.WriteLine(objExec.StdOut.ReadLine())
    Loop While Not objExec.Stdout.atEndOfStream
    WScript.StdOut.WriteLine(objExec.StdOut.ReadAll)
    Set objShell = Nothing
End Sub


'------------------------------------------------
' 执行PHP脚本
Function ExecPHP(phpFile)
    Dim objShell, objExec, php, arrStr
    Set objShell = CreateObject(COM_SHELL)
    php = config("PHP")
    Set objExec = objShell.Exec(php & " " & phpFile)
    ExecPHP = objExec.StdOut.ReadAll
End Function

'------------------------------------------------
' 执行Jar文件
Function ExecJar(jarFile)
    Dim objShell, objExec, java, arrStr
    Set objShell = CreateObject(COM_SHELL)
    java = config("JAVA")
    Set objExec = objShell.Exec(java & " -jar " & jarFile)
    ExecJar = objExec.StdOut.ReadAll
End Function

'------------------------------------------------
' 执行iconv.exe
Function ExecIConv(source_charset, dest_charset, source_file, dest_file)
    Dim objShell, cmd, iconv
    Set objShell = CreateObject(COM_SHELL)
    iconv = DisposePath(lib.path) & "iconv.exe"
    cmd = "%comspec% /c """ & iconv & """ -f " & source_charset & " -t " & dest_charset & " " & source_file & " > " & dest_file
    Echo cmd
    Dim iRet : iRet = objShell.Run(cmd, 0, True)
    ExecIConv = iRet
End Function

'------------------------------------------------
' 暂停
Sub Pause(message)
    WScript.Echo message
    z = WScript.StdIn.Read(1)
End Sub

'------------------------------------------------
' IIf条件表达式
Function IIf(condition, resTrue, resFalse)
    If condition Then
        IIf = resTrue
    Else
        IIf = resFalse
    End if
End Function

'------------------------------------------------
' COM组件是否安装
Function IsObjectInstalled(classString)
	On Error Resume Next

	IsObjectInstalled = False
	Err = 0
	Dim objTest
	objTest = CreateObject(classString)
	If Err = 0 Or Err = -2147352567 Then
		IsObjectInstalled = True
	End If	
	Set objTest = Nothing
	Err = 0	
End Function


'------------------------------------------------
' 进制转换函数
' Example:
'    WScript.Echo base_convert("A37334", 16, 2)
'    WScript.Echo base_convert("http://demon.tw", 16, 10)
Function base_convert(number, frombase, tobase)
    'Author: Demon
    'Date: 2011/12/17
    'Website: http://demon.tw

    Dim digits, num, ptr, i, n, c
    digits = "0123456789abcdefghijklmnopqrstuvwxyz"
    
    If frombase < 2 Or frombase > 36 Then
        Err.Raise vbObjectError + 7575,,"Invalid from base"
    End If
    
    If tobase < 2 Or tobase > 36 Then
        Err.Raise vbObjectError + 7575,,"Invalid to base"
    End If
    
    number = CStr(number) : n = Len(number)
    
    For i = 1 To n
        c = Mid(number, i, 1)
        
        If c >= "0" And c <= "9" Then
            c = c - "0"
        ElseIf c >= "A" And c <= "Z" Then
            c = Asc(c) - Asc("A") + 10
        ElseIf c >= "a" And c <= "z" Then
            c = Asc(c) - Asc("a") + 10
        Else
            c = frombase
        End If
        
        If c < frombase Then
            num = num * frombase + c
        End If
    Next
        
    Do
        ptr = ptr & Mid(digits, (num Mod tobase + 1), 1)
        num = num \ tobase
    Loop While num
    
    base_convert = StrReverse(ptr)
End Function

'------------------------------------------------
' hex2dec
' Hex String to Decimal max lenght is long int
Function hex2dec(strHex)
	Dim lngResult
	Dim intIndex
	Dim strDigit
	Dim intDigit
	Dim intValue
	lngResult = 0
	For intIndex = Len(strHex) To 1 Step -1
		strDigit = Mid(strHex, intIndex, 1)
		intDigit = InStr("0123456789ABCDEF", UCase(strDigit))-1
		If intDigit >= 0 Then
			intValue = intDigit * (16 ^ (Len(strHex)-intIndex))
			lngResult = CLng(lngResult + intValue)
		Else
			lngResult = 0
			intIndex = 0 ' stop the loop
		End If
	Next
	hex2dec = lngResult
End Function

'------------------------------------------------
' HTA Sleep
Function HTA_Sleep(n)
    Dim SHELL
    Set SHELL = CreateObject(COM_SHELL)
    Call SHELL.Run("%comspec% /c ping -n " + n + " 127.0.0.1 > nul", 0, 1)
    Set SHELL = Nothing
End Function

'------------------------------------------------
' 枚举System环境变量
Sub EnumSystemEnvironment(ByRef arrEnvironment)
    Dim objShell, objEnv
    Set objShell = CreateObject(COM_SHELL)
    Set arrEnvironment = objShell.Environment("SYSTEM")
End Sub


'------------------------------------------------
' 获取屏幕分辨率
Sub GetScreenWidthHeight(ByRef width, ByRef height)
    Dim objHTML, objScreen
    Set objHTML = CreateObject(COM_HTML)
    Set objScreen = objHTML.parentwindow.screen
    width = objScreen.width
    height = objScreen.height
    Set objHTML = Nothing
End Sub

'------------------------------------------------
' 显示桌面
Sub ShowDesktop
    Dim objShell
    Set objShell = CreateObject(COM_SHELLAPP)
    objShell.ToggleDesktop
    Set objShell = Nothing
End Sub


'------------------------------------------------
' 重启计算机
Sub ShutDown
    Dim Result, SHELL
    Set SHELL = CreateObject(COM_SHELL)
    Result = MsgBox("你确定要重起计算机吗?",vbokcancel+vbexclamation,"注意！") 
    If Result = vbOk Then
        SHELL.Run("Shutdown.exe -r -t 0")
    End If
End Sub

'------------------------------------------------
' 是否64位操作系统
Function IsX64()
    Dim objWMI, colItems, objItem, computer
    IsX64 = False
    computer = "."
    Set objWMI = CreateObject("winmgmts:{impersonationLevel=impersonate}!\\"&computer&"\root\cimv2")
    Set colItems = objWMI.ExecQuery("Select * from Win32_ComputerSystem",,48)
    For Each objItem in colItems		
        If InStr(objItem.SystemType, "64") <> 0 Then
            IsX64 = True
            Exit For
        End If
    Next
    Set objWMI = Nothing
End Function


