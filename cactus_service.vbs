' WMI, Win32_Serviceのヘルパを提供



' サービスの状態を返す
'   サービス名で検索し、取得できなければ表示名で検索し、最初に見つかったサービスの状態を返す
'   停止：Stopped
'   開始：Running
'   サービスが見つからない：(Empty)
Function GetServiceState(strServiceName)
    Dim Services, Service
    Set Services = GetObject("winmgmts:").ExecQuery("Select * from Win32_Service Where Name='" & strServiceName & "'")
    If Services.Count = 0 Then
        Set Services = GetObject("winmgmts:").ExecQuery("Select * from Win32_Service Where DisplayName='" & StrServiceName & "'")
    End If
    
    For Each Service In Services
        GetServiceState = Service.State
        Exit For
    Next
End Function


' 指定されたサービスを開始し、結果を返す
'   サービス名で検索し、取得できなければ表示名で検索し、最初に見つかったサービスを開始する
Function StartService(strServiceName)
    Dim Services, Service
    Set Services = GetObject("winmgmts:").ExecQuery("Select * from Win32_Service Where Name='" & strServiceName & "'")
    If Services.Count = 0 Then
        Set Services = GetObject("winmgmts:").ExecQuery("Select * from Win32_Service Where DisplayName='" & StrServiceName & "'")
    End If
    
    StartService = False
    Dim RetVal
    For Each Service In Services
        RetVal = Service.StartService()
        If RetVal = 0 Then StartService = True  ' Started
        If RetVal = 10 Then StartService = True ' Running
        Exit For
    Next
End Function


' 指定されたサービスを停止し、結果を返す
'   サービス名で検索し、取得できなければ表示名で検索し、見つかった最初のサービスを停止する
Function StopService(strServiceName)  
    Dim Services, Service
    Set Services = GetObject("winmgmts:").ExecQuery("Select * from Win32_Service Where Name='" & strServiceName & "'")
    If Services.Count = 0 Then
        Set Services = GetObject("winmgmts:").ExecQuery("Select * from Win32_Service Where DisplayName='" & StrServiceName & "'")
    End If
    
    StopService = False
    Dim RetVal
    For Each Service In Services
        RetVal = Service.StopService()
        If RetVal = 0 Then StopService = True  ' Stopped
        If RetVal = 5 Then StopService = True  ' Already Stopped
        Exit For
    Next
End Function

' WMI::StartService, StopServiceの戻値は下記参照。
'	http://www.wmifun.net/library/win32_service.html
'	3 - 実行中のほかのサービスが依存しているので停止できません。


'====================
'Test

'MsgBox GetServiceState("Telephony")    ' 表示名
'MsgBox GetServiceState("TapiSrv")    ' サービス名
'MsgBox GetServiceState("Alerter")
'StopService("Security Center")
'StopService("Security Center")


' Service helper class allows creation, deletion, and manipulation of local or remote services.		
Class std_service
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
	
	' Issue a service stop where strService is the service name
	' Return error codes can be SERVICE_WMI_FAIL, SERVICE_NOT_EXIST
	' In all cases success returns SERVICE_SUCCESS
	Public Function ServiceStop( strService , bQueryByDisplayName )
		ServiceStop = SERVICE_WMI_FAIL
		Dim objService : Set objService = Nothing
		If Not objWMIService Is Nothing Then
			If bQueryByDisplayName Then
				On Error Resume Next
				Dim colServices : Set colServices = objWMIService.ExecQuery("Select * from Win32_Service WHERE DisplayName='"& strService &"'")
				For Each objService In colServices
					Exit For
				Next
			Else
				On Error Resume Next
				Set objService = objWMIService.Get("Win32_Service.Name='" & strService & "'")
			End If 
			If objService Is Nothing Then 
				ServiceStop = SERVICE_NOT_EXIST
			Else
				If objService.State <> SERVICE_STAT_STOPPED Then
					Dim rVal
					rVal = objService.StopService()
					If rVal = SERVICE_SUCCESS Or rVal = SERVICE_STATE_REQUESTED Then
						ServiceStop = SERVICE_SUCCESS
					Else
						ServiceStop = rVal
					End If
				Else
					ServiceStop = SERVICE_SUCCESS
				End If 
			End If
		End If
		' Setting to Nothing is not needed, but we are doing it just to be sure
		Set objService = Nothing
	End Function
	
	' Issue a service start where strService is the service name
	' Return error codes can be SERVICE_WMI_FAIL, SERVICE_NOT_EXIST
	' In all cases success returns SERVICE_SUCCESS
	Public Function ServiceStart( strService , bQueryByDisplayName )
		ServiceStart = SERVICE_WMI_FAIL
		Dim objService : Set objService = Nothing
		If Not objWMIService Is Nothing Then
			If bQueryByDisplayName Then 
				On Error Resume Next
				Dim colServices : Set colServices = objWMIService.ExecQuery("Select * from Win32_Service WHERE DisplayName='"& strService &"'")
				For Each objService In colServices
					Exit For
				Next
			Else
				On Error Resume Next
				Set objService = objWMIService.Get("Win32_Service.Name='" & strService & "'")
			End If 
			If objService Is Nothing Then 
				ServiceStart = SERVICE_NOT_EXIST
			Else
				If objService.State <> SERVICE_STAT_RUNNING Then
					Dim rVal
					rVal = objService.StartService()
					If rVal = SERVICE_SUCCESS Or rVal = SERVICE_STATE_REQUESTED Then
						ServiceStart = SERVICE_SUCCESS
					Else
						ServiceStart = rVal
					End If
				Else
					ServiceStart = SERVICE_SUCCESS
				End If 
			End If
		End If
		' Setting to Nothing is not needed, but we are doing it just to be sure
		Set objService = Nothing
	End Function
	
	
	' Query the service status where strService is the service name
	' and strStatus is the text returned status string
	' Return error codes can be SERVICE_WMI_FAIL, SERVICE_NOT_EXIST 
	' In all cases success returns SERVICE_SUCCESS
	Public Function ServiceStatus( strService , byRef strStatus , bQueryByDisplayName )
		ServiceStatus = SERVICE_WMI_FAIL
		strStatus = "SERVICE_WMI_FAIL"
		Dim objService : Set objService = Nothing
		If Not objWMIService Is Nothing Then
			If bQueryByDisplayName Then
				On Error Resume Next
				Dim colServices : Set colServices = objWMIService.ExecQuery("Select * from Win32_Service WHERE DisplayName='"& strService &"'")
				For Each objService In colServices
					Exit For
				Next
			Else
				On Error Resume Next
				Set objService = objWMIService.Get("Win32_Service.Name='" & strService & "'")
			End If
			If objService Is Nothing Then 
				ServiceStatus = SERVICE_NOT_EXIST
				strStatus = "SERVICE_NOT_EXIST"
			Else
				strStatus = objService.State
				ServiceStatus = SERVICE_SUCCESS
			End If
		End If
		' Setting to Nothing is not needed, but we are doing it just to be sure
		Set objService = Nothing
	End Function
	
	' Wait for a service state where strService is the service name
	' strState should be one the constants below:
	' SERVICE_STAT_STOPPED
	' SERVICE_STAT_RUNNING
	' SERVICE_STAT_PAUSED
	' nTimeoutSec should be the number of seconds to wait till deeming the service
	' state stuck or not changeable.
	' The the function returns false if the service state we reached and false
	' if the service state was not reached.
	Public Function ServiceStatusWait( strService , bQueryByDisplayName , strState , nTimeoutSec )
		ServiceStatusWait = False
		Dim Timeout : Timeout = 0
		Dim state , rval
		If nTimeoutSec < 1 Then 
			rval = ServiceStatus( strService , state , bQueryByDisplayName )
			If rval = SERVICE_SUCCESS And UCase(state) = UCase(strState) Then
				ServiceStatusWait = True
			Else
				ServiceStatusWait = False
			End If
		Else
			For Timeout = 0 To nTimeoutSec - 1
				rval = ServiceStatus( strService , state , bQueryByDisplayName )
				If rval = SERVICE_SUCCESS And UCase(state) = UCase(strState) Then 
					ServiceStatusWait = True
					Exit For
				End If 
				If rval = SERVICE_NOT_EXIST Then
					ServiceStatusWait = False
					Exit For
				End If
				WScript.Sleep(1000)
			Next
		End If 
	End Function
	
	' Issue a service deletion where strService is the service name
	' Return error codes can be SERVICE_WMI_FAIL, SERVICE_NOT_EXIST 
	' In all cases success returns SERVICE_SUCCESS
	Public Function ServiceDelete( strService , bQueryByDisplayName )
		ServiceDelete = SERVICE_WMI_FAIL
		Dim objService : Set objService = Nothing
		If Not objWMIService Is Nothing Then
			If bQueryByDisplayName Then
				On Error Resume Next
				Dim colServices : Set colServices = objWMIService.ExecQuery("Select * from Win32_Service WHERE DisplayName='"& strService &"'")
				For Each objService In colServices
					Exit For
				Next
			Else
				On Error Resume Next
				Set objService = objWMIService.Get("Win32_Service.Name='" & strService & "'")
			End If
			If objService Is Nothing Then 
				ServiceDelete = SERVICE_NOT_EXIST
			Else
				ServiceDelete = objService.Delete()
			End If
		End If 
		' Setting to Nothing is not needed, but we are doing it just to be sure
		Set objService = Nothing
	End Function
	
	' Create a service where strServiceName is the service name
	' See the section "TEMPLATE CONSTANTS" at the top of the script for "Service Types" and "Error Control" for
	' values to be used with nServiceType and nErrorControl
	' Return error codes can be SERVICE_WMI_FAIL, SERVICE_NOT_EXIST 
	' In all cases success returns SERVICE_SUCCESS
	Public Function ServiceCreate( strServiceName , strDescription , strSvcExe , nServiceType , nErrorControl , strStartMode , bDesktopInteract , strStartName , strStartPass , strLoadOrderGroup )
		ServiceCreate = SERVICE_WMI_FAIL
		Dim objService
		If Not objWMIService Is Nothing Then
			On Error Resume Next
			Set objService = objWMIService.Get("Win32_BaseService")
			If objService Is Nothing Then 
				ServiceCreate = SERVICE_NOT_EXIST ' Not sure why whis would happen
			Else
				ServiceCreate = objService.Create( strServiceName , strDescription , strSvcExe , nServiceType , nErrorControl , strStartMode , bDesktopInteract, strStartName , strStartPass )
			End If
		End If
		' Setting to Nothing is not needed, but we are doing it just to be sure
		Set objService = Nothing
	End Function
	
	
	' Query service start mode where strService is the service name
	' strMode is returned byRef if the function completes with success 
	' Return error codes can be SERVICE_WMI_FAIL, SERVICE_NOT_EXIST
	' In all cases success returns SERVICE_SUCCESS
	' See the "TEMPLATE CONSTANTS" section for a list of Service Start Modes:
	' SERVICE_START_BOOT
	' SERVICE_START_SYSTEM
	' SERVICE_START_AUTO
	' SERVICE_START_MANUAL
	' SERVICE_START_DISABLE
	Public Function ServiceStartMode( strService, byRef strMode , bQueryByDisplayName )
		ServiceStartMode = SERVICE_WMI_FAIL
		strMode = "SERVICE_WMI_FAIL"
		Dim objService : objService = Nothing
		If Not objWMIService Is Nothing Then
			If bQueryByDisplayName Then
				On Error Resume Next
				Dim colServices : Set colServices = objWMIService.ExecQuery("Select * from Win32_Service WHERE DisplayName='"& strService &"'")
				For Each objService In colServices
					Exit For
				Next
			Else
				On Error Resume Next
				Set objService = objWMIService.Get("Win32_Service.Name='" & strService & "'")
			End If
			If objService Is Nothing Then 
				ServiceStartMode = SERVICE_NOT_EXIST
				strMode = "SERVICE_NOT_EXIST"
			Else
				strMode = objService.StartMode
				ServiceStartMode = SERVICE_SUCCESS
			End If
		End If
		' Setting to Nothing is not needed, but we are doing it just to be sure
		Set objService = Nothing
	End Function
	
	
	' Modify service start mode where strService is the service name
	' strMode is the start mode which will be applied to the service if it exists
	' Return error codes can be SERVICE_WMI_FAIL, SERVICE_NOT_EXIST
	' In all cases success returns SERVICE_SUCCESS
	' See the "TEMPLATE CONSTANTS" section for a list of Service Start Modes:
	' SERVICE_START_BOOT
	' SERVICE_START_SYSTEM
	' SERVICE_START_AUTO
	' SERVICE_START_MANUAL
	' SERVICE_START_DISABLE
	Public Function ChangeStartMode( strService , strMode , bQueryByDisplayName )
		ChangeStartMode = SERVICE_WMI_FAIL
		Dim objService : objService = Nothing
		If Not objWMIService Is Nothing Then
			If bQueryByDisplayName Then
				On Error Resume Next
				Dim colServices : Set colServices = objWMIService.ExecQuery("Select * from Win32_Service WHERE DisplayName='"& strService &"'")
				For Each objService In colServices
					Exit For
				Next
			Else
				On Error Resume Next
				Set objService = objWMIService.Get("Win32_Service.Name='" & strService & "'")
			End If
			If objService Is Nothing Then 
				ChangeStartMode = SERVICE_NOT_EXIST
			Else
				ChangeStartMode = objService.ChangeStartMode(strMode)
			End If
		End If
		' Setting to Nothing is not needed, but we are doing it just to be sure
		Set objService = Nothing
	End Function
	
	' Members
	Private objWMIService
End Class