'------------------------------------------------
' 杀死进程
' @name         进程名
' example:
'   Call KillProcess("rar.exe")
Sub KillProcess(name)
    Dim computer, WMI, processlist, process
    computer = "."
    Set WMI = GetObject("winmgmts:\" & computer & "\root\cimv2")
    Set processlist = WMI.ExecQuery("Select * from Win32_Process Where Name = '" & name & "'")
    For Each process in processlist
        process.Terminate()
    Next
End Sub


' Basic process class to help managed local or remote processes.
' This class should be expanded to do other functions in the future
Class std_process
	' Must be nothing on creation
	Private Sub Class_Initialize()
		Set objWMIService = Nothing
	End Sub
	
	' Connect to the reg provider for this registy object
	Public Function ConnectProvider( sComputerName )
		ConnectProvider = False
		On Error Resume Next
		Err.Clear
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & sComputerName & "\root\cimv2")
		If Err.Number = 0 Then			
			ConnectProvider = True
		End If
	End Function
	
	' Returns tr if a processes of strProc exist.
	' ie: strProc is the process name like calc.exe
	Function ProcessExists( strProc )
		ProcessExists = PROCESS_WMI_FAIL
		Dim colProcessList	
		If Not objWMIService Is Nothing Then
			Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process WHERE Name = '"& strProc &"'")
			If colProcessList.count > 0 Then
				ProcessExists = PROCESS_EXISTS
			Else
				ProcessExists = PROCESS_NOT_EXIST
			End If
		End If
		' Setting to Nothing is not needed, but we are doing it just to be sure
		Set colProcessList = Nothing
	End Function
	
	' Returns true if a processes were killed where process strProc is the name of that process(es).
	' ie: strProc is the process name without the .exe "calc" is an example
	Function ProcessKill( strProc )
		ProcessKill = PROCESS_WMI_FAIL
		Dim colProcessList, objProcess
		If Not objWMIService Is Nothing Then
			Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process WHERE Name = '"& strProc &"'")
			If colProcessList.count > 0 Then
				For Each objProcess In colProcessList
					ProcessKill = objProcess.Terminate()
				Next
			Else
				ProcessKill = PROCESS_NOT_EXIST
			End If 
		End If
		' Setting to Nothing is not needed, but we are doing it just to be sure
		Set colProcessList = Nothing
	End Function
	
	' Creates a process on the remote host and retuns the pid on success
	' The function returns PROCESS_SUCCESS on success. 
	Function CreateProcess( strExe , strWorkingDir , pid )
        CreateProcess = PROCESS_WMI_FAIL
		Dim objProcess
        If Not objWMIService Is Nothing Then
            Set objProcess = objWMIService.Get("Win32_Process")
            CreateProcess = objProcess.Create(strExe,strWorkingDir,Null,pid)
        End If
    End Function

	' Lists all the processes running where process strProc is the name of that process
	' The function returns PROCESS_SUCCESS on success. If successful arrProc is a list 
	' of processes running on strHost.
	Function ProcessList( byRef arrProc )
		ProcessList = PROCESS_WMI_FAIL
		Dim colProcessList, objProcess, i
		arrProc = Array
		If Not objWMIService Is Nothing Then
			Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process")
			If colProcessList.count > 0 Then
				For Each objProcess In colProcessList
					i = UBound(arrProc) + 1		
					ReDim Preserve arrProc(i)
					arrProc(i) = objProcess.Name		
				Next
				ProcessList = PROCESS_SUCCESS
			Else
				ProcessList = PROCESS_NOT_EXIST
			End If
		End If
		' Setting to Nothing is not needed, but we are doing it just to be sure
		Set colProcessList = Nothing
	End Function
	
	' Function waits for a process strProcName to exit for nMaxWaitTimeSeconds
	' If the function succeeds the return value is true, else false process didn't exit
	' The nMaxWaitTimSeconds should be 5-10 minutes just to take into account slow processes
	Function WaitOnProcessByName( strProcName , nNotExistThreashold , nMaxWaitTimeSeconds )
		WaitOnProcessByName = False
		Dim notExistCount : notExistCount = 0	
		Dim maxCount : maxCount = 0 
		Dim rVal 
		While notExistCount < nNotExistThreashold And maxCount < nMaxWaitTimeSeconds
			' This could return WMI_FAIL
			rVal = ProcessExists( strProcName )
			If rVal = PROCESS_NOT_EXIST Then 
				notExistCount = notExistCount + 1
			Else
				notExistCount = 0
			End If
			maxCount = maxCount + 1
			' Sleep 1 seconds keep looking for setup exec
			WScript.Sleep 1000
		Wend
		' TODO: Perhaps find a better way
		' The process we were watching ended if they are equal
		If notExistCount = nNotExistThreashold Then 
			WaitOnProcessByName = True
		End If		
	End Function
	
	' Members
	Private objWMIService
End Class
