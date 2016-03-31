'---------------------------------------------------------
' This sample shows how to set the various task objects
' for creating a new scheduled task in Task Scheduler 2.0
'
' An alternative would be to to build the XML and create the t
' task with the .RegisterTask method 
' - See http://msdn.microsoft.com/en-us/library/windows/desktop/aa382575%28v=vs.85%29.aspx
'---------------------------------------------------------

' -------------------------------------------------------------------------------
' Define the Available Enumerations that can be used
' Since VBScript does not offer enumeration, we have to set these all as constants
' For all enumeration documentation see - http://msdn.microsoft.com/en-us/library/windows/desktop/aa383602%28v=vs.85%29.aspx

 Option Explicit

'  TASK_ACTION_TYPE
Const TASK_ACTION_EXEC = 0
Const TASK_ACTION_COM_HANDLER = 5
Const TASK_ACTION_SEND_EMAIL = 6
Const TASK_ACTION_SHOW_MESSAGE = 7

' TASK_COMPATIBILITY
Const TASK_COMPATIBILITY_AT = 0
Const TASK_COMPATIBILITY_V1 = 1
Const TASK_COMPATIBILITY_V2 = 2

' TASK_CREATION
Const TASK_VALIDATE_ONLY = &H01&
Const TASK_CREATE = &H02&
Const TASK_UPDATE = &H04&
Const TASK_CREATE_OR_UPDATE = &H06&
Const TASK_DISABLE = &H08&
Const TASK_DONT_ADD_PRINCIPAL_ACE = &H10&
Const TASK_IGNORE_REGISTRATION_TRIGGERS = &H20&

' TASK_INSTANCES_POLICY
Const TASK_INSTANCES_PARALLEL = 0
Const TASK_INSTANCES_QUEUE = 1
Const TASK_INSTANCES_IGNORE_NEW = 2
Const TASK_INSTANCES_STOP_EXISTING = 3

' TASK_LOGON_TYPE
Const TASK_LOGON_NONE = 0
Const TASK_LOGON_PASSWORD = 1
Const TASK_LOGON_S4U = 2
Const TASK_LOGON_INTERACTIVE_TOKEN = 3
Const TASK_LOGON_GROUP = 4
Const TASK_LOGON_SERVICE_ACCOUNT = 5
Const TASK_LOGON_INTERACTIVE_TOKEN_OR_PASSWORD = 6

' TASK_RUNLEVEL_TYPE
Const TASK_RUNLEVEL_LUA = 0
Const TASK_RUNLEVEL_HIGHEST = 1

' TASK_TRIGGER_TYPE2
Const TASK_TRIGGER_EVENT = 0
Const TASK_TRIGGER_TIME  = 1
Const TASK_TRIGGER_DAILY = 2
Const TASK_TRIGGER_WEEKLY = 3
Const TASK_TRIGGER_MONTHLY = 4
Const TASK_TRIGGER_MONTHLYDOW = 5
Const TASK_TRIGGER_IDLE = 6
Const TASK_TRIGGER_REGISTRATION = 7
Const TASK_TRIGGER_BOOT = 8
Const TASK_TRIGGER_LOGON = 9
Const TASK_TRIGGER_SESSION_STATE_CHANGE = 11
' -------------------------------------------------------------------------------

Dim objTaskService, objRootFolder, objTaskFolder, objNewTaskDefinition
Dim objTaskTrigger, objTaskAction, objTaskTriggers, blnFoundTask
Dim objTaskFolders

' Create the TaskService object and connect
Set objTaskService = CreateObject("Schedule.Service")
call objTaskService.Connect()

' Get the Root Folder where we will place this task
Set objTaskFolder = objTaskService.GetFolder("\")
' Or create a folder and use it.  I would first check if it exists.
' If it does exist CreateFolder will generate an error.  
' You have to loop through the folders to find the one you want
Set objRootFolder = objTaskService.GetFolder("\")

' Get all the sub folders and see if the one one want exists
Set objTaskFolders = objRootFolder.GetFolders(0)
    
For Each objTaskFolder In objTaskFolders
	If objTaskFolder.Path = "\MyNewTaskFolder" Then
		blnFoundTask = True
		Exit For
	End If
Next
    
If Not blnFoundTask Then Set objTaskFolder = objRootFolder.CreateFolder("\MyNewTaskFolder")
 
' -------------------------------------------------------------------------------
' Start Creation of the Task Definition
' -------------------------------------------------------------------------------
' The flags parameter is 0 because it is not used and reserved for future use
' http://msdn.microsoft.com/en-us/library/windows/desktop/aa383470%28v=vs.85%29.aspx
Set objNewTaskDefinition = objTaskService.NewTask(0) 

With objNewTaskDefinition
	' Text that is associated with the task. This data is ignored by the Task Scheduler 
	' service, but is used by third-parties who wish to extend the task format.
	.Data = "This is my sample task via script" 

	' -------------------------------------------------------------------------------
	' Set the values for the registration information - General Tab Top Section
	' -------------------------------------------------------------------------------
	'http://msdn.microsoft.com/en-us/library/windows/desktop/aa382100%28v=vs.85%29.aspx
	With .RegistrationInfo
		.Author = "Name or Process Creating Task"
		' or
		' .Author = objTaskService.ConnectedDomain  & "\" & objTaskService.ConnectedUser 
		.Date = ConvertTime(now())
		.Description = "Description of What this task does"
		.Documentation  = "My Document" ' See - http://msdn.microsoft.com/en-us/library/windows/desktop/aa382104%28v=vs.85%29.aspx
		' .SecurityDescriptor ' http://msdn.microsoft.com/en-us/library/windows/desktop/aa379567%28v=vs.85%29.aspx
		.Source = "VB Script"
		.URI = "http://mysite.com"	' Not Shown in GUI
		.Version = "1.0" ' Self defined version of this scheduled task. Not Shown in GUI
	End With 'objRegistrationInfo

	' -------------------------------------------------------------------------------
	' Set the values for the General Tab Security Section
	' -------------------------------------------------------------------------------
	' http://msdn.microsoft.com/en-us/library/windows/desktop/aa382071%28v=vs.85%29.aspx
	With .Principal
		.Id = "My ID" 	' Not shown in GUI
		.DisplayName = "Principal Description" ' Not Shown in GUI
		.UserId = "Domain\myuser"	' This script must be run with elevated privileges if this is not the current user.
		' or 
		.UserId = objTaskService.ConnectedDomain  & "\" & objTaskService.ConnectedUser 
		'.GroupId = "" ' The identifier of the user group that is associated with this principal. Do not set this property if a user identifier is specified in the UserId property.
		.LogonType = TASK_LOGON_INTERACTIVE_TOKEN	' TASK_LOGON_TYPE
		.RunLevel = TASK_RUNLEVEL_LUA ' TASK_RUNLEVEL_TYPE - If you use Highest Privilege, the script will need to run elevated.
	End With 'objPrincipal

	' -------------------------------------------------------------------------------
	' Set Triggers Tab - Examples of the different Types of Triggers
	' -------------------------------------------------------------------------------
	' http://msdn.microsoft.com/en-us/library/windows/desktop/aa383868%28v=vs.85%29.aspx
	Set objTaskTriggers = .Triggers
	' *** Event Trigger ***  ---  http://msdn.microsoft.com/en-us/library/windows/desktop/aa446882%28v=vs.85%29.aspx
	Set objTaskTrigger = objTaskTriggers.Create(TASK_TRIGGER_EVENT)
	With objTaskTrigger
		.Enabled = True
		.Id = "EventTriggerID1"
		' Time Format  YYYY-MM-DDTHH:MM:SS or use ConvertTime Format 
		'.StartBoundary = "2013-07-01T08:08:00"
		'.EndBoundary = "2013-07-01T08:08:00"
		.StartBoundary = ConvertTime(DateAdd("h", 1, now()))
		.EndBoundary = ConvertTime(DateAdd("h", 3, now()))
		' Stop Task if it runs longer than . . 
		' Format For Days = P#D where # is the number of days
		' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
		.ExecutionTimeLimit = "PT1M"
		.Delay = "PT5M"
		
		' Need to build Event Subscription XML.  This xml can 
		' be very complicated and I could not find documentation
		' I recommend you set up a task with a custom event trigger
		' On the custom trigger screen click on the XML tab and copy
		' that xml document here to create your trigger.  Change all the 
		' qoute marks (") to single tick marks (') in the copied xml
		.Subscription  = "<QueryList>" _ 
				& "<Query Id='0' Path='MyTest'>" _
				& "<Select Path='Application'>*[System[Provider[@Name='FedExAdminService'] and (Level=1  or Level=2 or Level=3 or Level=4 or Level=0 or Level=5)]]</Select>" _
				& "</Query>" _
				& "</QueryList>"

		With .Repetition
			' Format For Days = P#D where # is the number of days
			' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
			.Duration = "P1D"
			.Interval = "PT1H"
			.StopAtDurationEnd = True
		End With 'objTaskRepitition
	End With

	' *** Time Trigger ***  ---  http://msdn.microsoft.com/en-us/library/windows/desktop/aa383622%28v=vs.85%29.aspx
	Set objTaskTrigger = objTaskTriggers.Create(TASK_TRIGGER_TIME)
	With objTaskTrigger
		.Enabled = True
		.Id = "TimeTriggerID1"
		' Time Format  YYYY-MM-DDTHH:MM:SS or use ConvertTime Format 
		'.StartBoundary = "2013-07-01T08:08:00"
		'.EndBoundary = "2013-07-01T08:08:00"
		.StartBoundary = ConvertTime(DateAdd("h", 1, now()))
		.EndBoundary = ConvertTime(DateAdd("h", 3, now()))
		' Stop Task if it runs longer than . . 
		' Format For Days = P#D where # is the number of days
		' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
		.ExecutionTimeLimit = "PT1M"

		With .Repetition
			' Format For Days = P#D where # is the number of days
			' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
			.Duration = "P1D"
			.Interval = "PT1H"
			.StopAtDurationEnd = True
		End With 'objTaskRepitition
	End With

	' *** Daily Trigger ***  ---  http://msdn.microsoft.com/en-us/library/windows/desktop/aa446858%28v=vs.85%29.aspx
	Set objTaskTrigger = objTaskTriggers.Create(TASK_TRIGGER_DAILY)
	With objTaskTrigger
		.Enabled = True
		.Id = "DailyTriggerID1"
		' Time Format  YYYY-MM-DDTHH:MM:SS or use ConvertTime Format 
		'.StartBoundary = "2013-07-01T08:08:00"
		'.EndBoundary = "2013-07-01T08:08:00"
		.StartBoundary = ConvertTime(DateAdd("h", 1, now()))
		.EndBoundary = ConvertTime(DateAdd("h", 3, now()))
		' Stop Task if it runs longer than . . 
		' Format For Days = P#D where # is the number of days
		' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
		.ExecutionTimeLimit = "PT1M"
		.DaysInterval = 3	' Recur every x number of days

		With .Repetition
			' Format For Days = P#D where # is the number of days
			' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
			.Duration = "P1D"
			.Interval = "PT1H"
			.StopAtDurationEnd = True
		End With 'objTaskRepitition
	End With
	
	' *** Weekly Trigger ***  ---  http://msdn.microsoft.com/en-us/library/windows/desktop/aa384019%28v=vs.85%29.aspx
	Set objTaskTrigger = objTaskTriggers.Create(TASK_TRIGGER_WEEKLY)
	With objTaskTrigger
		.Enabled = True
		.Id = "WeeklyTriggerID1"
		' Time Format  YYYY-MM-DDTHH:MM:SS or use ConvertTime Format 
		'.StartBoundary = "2013-07-01T08:08:00"
		'.EndBoundary = "2013-07-01T08:08:00"
		.StartBoundary = ConvertTime(DateAdd("h", 1, now()))
		.EndBoundary = ConvertTime(DateAdd("h", 3, now()))
		' Stop Task if it runs longer than . . 
		' Format For Days = P#D where # is the number of days
		' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
		.ExecutionTimeLimit = "PT1M"
		' Days of Week is a Bit Flag.  http://msdn.microsoft.com/en-us/library/windows/desktop/aa384024%28v=vs.85%29.aspx
		' Options:
		' 1 = Sunday
		' 2 = Monday
		' 4 = Tuesday
		' 8 = Wednesday
		' 16 = Thursday
		' 32 = Friday
		' 64 = Saturday
		.DaysOfWeek = 28 ' 28 = Tuesday, Wednesday & Thursday
		.WeeksInterval = 2 ' Recur every x number of weeks
	
		With .Repetition
			' Format For Days = P#D where # is the number of days
			' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
			.Duration = "P1D"
			.Interval = "PT1H"
			.StopAtDurationEnd = True
		End With 'objTaskRepitition	
	End With

	' *** Monthly Trigger ***  ---  http://msdn.microsoft.com/en-us/library/windows/desktop/aa382062%28v=vs.85%29.aspx
	Set objTaskTrigger = objTaskTriggers.Create(TASK_TRIGGER_MONTHLY)
	With objTaskTrigger
		.Enabled = True
		.Id = "MonthlyTriggerID1"
		' Time Format  YYYY-MM-DDTHH:MM:SS or use ConvertTime Format 
		'.StartBoundary = "2013-07-01T08:08:00"
		'.EndBoundary = "2013-07-01T08:08:00"
		.StartBoundary = ConvertTime(DateAdd("h", 1, now()))
		.EndBoundary = ConvertTime(DateAdd("h", 3, now()))
		' Stop Task if it runs longer than . . 
		' Format For Days = P#D where # is the number of days
		' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
		.ExecutionTimeLimit = "PT1M"
		' Days of Month is a bit Flag -- http://msdn.microsoft.com/en-us/library/windows/desktop/aa382063%28v=vs.85%29.aspx
		' Day 1 =  &H01&, 2 =  &H02&, 3 =  &H04&, . . . . 31 =  &H40000000&
		.DaysOfMonth = &h01& + &h04& + &h40& + &H40000000&  ' 1st day, 3rd day, 7th day and 31st Day
		.RunOnLastDayOfMonth = True ' Flag for if this should run on the last day of the scheduled months  
		' Months of the year is a bit flag -- http://msdn.microsoft.com/en-us/library/windows/desktop/aa382064%28v=vs.85%29.aspx
		' January = 1, February = 2, March = 4, April = 8, May = 16, June = 32, July = 64, August = 128
		' September = 256, October = 512, November = 1024, December = 2048
		.MonthsOfYear = 1045 ' January, March, May, November
		' Randomly delay the start of the task
		' Format For Days = P#D where # is the number of days
		' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
		.RandomDelay = "PT4H"	' Randomly Delay for 4 Hours
	
		With .Repetition
			' Format For Days = P#D where # is the number of days
			' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
			.Duration = "P1D"
			.Interval = "PT1H"
			.StopAtDurationEnd = True
		End With 'objTaskRepitition	
	End With

	' *** Monthly Day of Week Trigger ***  ---  
	Set objTaskTrigger = objTaskTriggers.Create(TASK_TRIGGER_MONTHLYDOW)
	With objTaskTrigger
		.Enabled = True
		.Id = "Monthly DayOfMonth TriggerID1"
		' Time Format  YYYY-MM-DDTHH:MM:SS or use ConvertTime Format 
		'.StartBoundary = "2013-07-01T08:08:00"
		'.EndBoundary = "2013-07-01T08:08:00"
		.StartBoundary = ConvertTime(DateAdd("h", 1, now()))
		.EndBoundary = ConvertTime(DateAdd("h", 3, now()))
		' Stop Task if it runs longer than . . 
		' Format For Days = P#D where # is the number of days
		' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
		.ExecutionTimeLimit = "PT1M"
		' Days of Week is a Bit Flag.  http://msdn.microsoft.com/en-us/library/windows/desktop/aa384024%28v=vs.85%29.aspx
		' Options:
		' 1 = Sunday
		' 2 = Monday
		' 4 = Tuesday
		' 8 = Wednesday
		' 16 = Thursday
		' 32 = Friday
		' 64 = Saturday
		.DaysOfWeek = 28 ' 28 = Tuesday, Wednesday & Thursday
		' Weeks of The Month is a bit flag.   http://msdn.microsoft.com/en-us/library/windows/desktop/aa382061%28v=vs.85%29.aspx
		' First Week = 1, Second Week = 2, Third Week = 4, Fourth Week = 8 
		.WeeksOfMonth = 9 ' Run on 1st and 4th weeks of month
		.RunOnLastWeekOfMonth  = True 
		' Months of the year is a bit flag -- http://msdn.microsoft.com/en-us/library/windows/desktop/aa382064%28v=vs.85%29.aspx
		' January = 1, February = 2, March = 4, April = 8, May = 16, June = 32, July = 64, August = 128
		' September = 256, October = 512, November = 1024, December = 2048
		.MonthsOfYear = 1045 ' January, March, May, November
		' Randomly delay the start of the task
		' Format For Days = P#D where # is the number of days
		' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
		.RandomDelay = "PT4H"	' Randomly Delay for 4 Hours
	
		With .Repetition
			' Format For Days = P#D where # is the number of days
			' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
			.Duration = "P1D"
			.Interval = "PT1H"
			.StopAtDurationEnd = True
		End With 'objTaskRepitition	
	End With

	' *** On Idle Trigger ***  ---  http://msdn.microsoft.com/en-us/library/windows/desktop/aa380690%28v=vs.85%29.aspx
	' Make sure the IDLE Properties are also set for this trigger to work.
	Set objTaskTrigger = objTaskTriggers.Create(TASK_TRIGGER_IDLE)
	With objTaskTrigger
		.Enabled = True
		.Id = "OnIdlleTriggerID1"
		' Time Format  YYYY-MM-DDTHH:MM:SS or use ConvertTime Format 
		'.StartBoundary = "2013-07-01T08:08:00"
		'.EndBoundary = "2013-07-01T08:08:00"
		.StartBoundary = ConvertTime(DateAdd("h", 1, now()))
		.EndBoundary = ConvertTime(DateAdd("h", 3, now()))
		' Stop Task if it runs longer than . . 
		' Format For Days = P#D where # is the number of days
		' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
		.ExecutionTimeLimit = "PT1M"
	
		With .Repetition
			' Format For Days = P#D where # is the number of days
			' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
			.Duration = "P1D"
			.Interval = "PT1H"
			.StopAtDurationEnd = True
		End With 'objTaskRepitition	
	End With

	' *** On Task Creation/Modification ***  ---  http://msdn.microsoft.com/en-us/library/windows/desktop/aa382110%28v=vs.85%29.aspx
	Set objTaskTrigger = objTaskTriggers.Create(TASK_TRIGGER_REGISTRATION)
	With objTaskTrigger
		.Enabled = True
		.Id = "TaskCreateTriggerID1"
		' Time Format  YYYY-MM-DDTHH:MM:SS or use ConvertTime Format 
		'.StartBoundary = "2013-07-01T08:08:00"
		'.EndBoundary = "2013-07-01T08:08:00"
		.StartBoundary = ConvertTime(DateAdd("h", 1, now()))
		.EndBoundary = ConvertTime(DateAdd("h", 3, now()))
		' Stop Task if it runs longer than . . 
		' Format For Days = P#D where # is the number of days
		' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
		.ExecutionTimeLimit = "PT1M"
		.Delay = "PT45M"
	
		With .Repetition
			' Format For Days = P#D where # is the number of days
			' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
			.Duration = "P1D"
			.Interval = "PT1H"
			.StopAtDurationEnd = True
		End With 'objTaskRepitition	
	End With

	' *** AT Start Up (Boot) Trigger  ***  ---  http://msdn.microsoft.com/en-us/library/windows/desktop/aa446815%28v=vs.85%29.aspx
	Set objTaskTrigger = objTaskTriggers.Create(TASK_TRIGGER_BOOT)
	With objTaskTrigger
		.Enabled = True
		.Id = "BootTriggerID1"
		' Time Format  YYYY-MM-DDTHH:MM:SS or use ConvertTime Format 
		'.StartBoundary = "2013-07-01T08:08:00"
		'.EndBoundary = "2013-07-01T08:08:00"
		.StartBoundary = ConvertTime(DateAdd("h", 1, now()))
		.EndBoundary = ConvertTime(DateAdd("h", 3, now()))
		' Stop Task if it runs longer than . . 
		' Format For Days = P#D where # is the number of days
		' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
		.ExecutionTimeLimit = "PT1M"
		.Delay = "PT45M"
	
		With .Repetition
			' Format For Days = P#D where # is the number of days
			' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
			.Duration = "P1D"
			.Interval = "PT1H"
			.StopAtDurationEnd = True
		End With 'objTaskRepitition	
	End With

	' *** At Log on Trigger ***  ---  http://msdn.microsoft.com/en-us/library/windows/desktop/aa381908%28v=vs.85%29.aspx
	Set objTaskTrigger = objTaskTriggers.Create(TASK_TRIGGER_LOGON)
	With objTaskTrigger
		.Enabled = True
		.Id = "LogonTriggerID1"
		' Time Format  YYYY-MM-DDTHH:MM:SS or use ConvertTime Format 
		'.StartBoundary = "2013-07-01T08:08:00"
		'.EndBoundary = "2013-07-01T08:08:00"
		.StartBoundary = ConvertTime(DateAdd("h", 1, now()))
		.EndBoundary = ConvertTime(DateAdd("h", 3, now()))
		' Stop Task if it runs longer than . . 
		' Format For Days = P#D where # is the number of days
		' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
		.ExecutionTimeLimit = "PT1M"
		.Delay = "PT45M"
		' For User ID - comment out to leave for any user.  
		' This script must be run with elevated privileges if this is not the current user or if not used!
		.UserId = "Domain\myuser"	
		' or 
		.UserId = objTaskService.ConnectedDomain  & "\" & objTaskService.ConnectedUser 
	
		With .Repetition
			' Format For Days = P#D where # is the number of days
			' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
			.Duration = "P1D"
			.Interval = "PT1H"
			.StopAtDurationEnd = True
		End With 'objTaskRepitition	
	End With

	' *** Session State Change Trigger  ***  ---  http://msdn.microsoft.com/en-us/library/windows/desktop/aa382142%28v=vs.85%29.aspx
	Set objTaskTrigger = objTaskTriggers.Create(TASK_TRIGGER_SESSION_STATE_CHANGE)
	With objTaskTrigger
		.Enabled = True
		.Id = "Session State Changed TriggerID1"
		' Time Format  YYYY-MM-DDTHH:MM:SS or use ConvertTime Format 
		'.StartBoundary = "2013-07-01T08:08:00"
		'.EndBoundary = "2013-07-01T08:08:00"
		.StartBoundary = ConvertTime(DateAdd("h", 1, now()))
		.EndBoundary = ConvertTime(DateAdd("h", 3, now()))
		' Stop Task if it runs longer than . . 
		' Format For Days = P#D where # is the number of days
		' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
		.ExecutionTimeLimit = "PT1M"
		.Delay = "PT45M"
		' For User ID - comment out to leave for any user.  
		' This script must be run with elevated privileges if this is not the current user or if not used!
		.UserId = "Domain\myuser"	
		' or 
		.UserId = objTaskService.ConnectedDomain  & "\" & objTaskService.ConnectedUser 
		' State Change Event - http://msdn.microsoft.com/en-us/library/windows/desktop/aa382144%28v=vs.85%29.aspx
		' User Session Connect to Local Computer = 1
		' User Session Disconnect from Local Computer = 2
		' User Session Connect to Remote Computer = 3
		' User Session Disconnect from Remote Computer = 4
		' On Workstation Lock = 7
		' On Workstation Unlock = 8
		.StateChange = 7
	
		With .Repetition
			' Format For Days = P#D where # is the number of days
			' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
			.Duration = "PT1H"
			.Interval = "PT1M"
			.StopAtDurationEnd = False
		End With 'objTaskRepitition	
	End With

	' -------------------------------------------------------------------------------
	' Set Value for Actions Tab - Examples of the different types of actions
	' -------------------------------------------------------------------------------
	' *** Execute / Command Line Action  ***  ---  http://msdn.microsoft.com/en-us/library/windows/desktop/aa446890%28v=vs.85%29.aspx
	Set objTaskAction = .Actions.Create(TASK_ACTION_EXEC)
	With objTaskAction
		.Id = "ExecuteAction Sample"
		' File Path and Name to run or command line to execute
		.Path = "C:\Windows\System32\notepad.exe"
		.Arguments = WScript.ScriptFullName  
		.WorkingDirectory = "C:\Windows\System32"
	End With

	' *** Email Message Action  ***  ---  http://msdn.microsoft.com/en-us/library/windows/desktop/aa446868%28v=vs.85%29.aspx
	Set objTaskAction = .Actions.Create(TASK_ACTION_SEND_EMAIL)
	With objTaskAction
		.Id = "Email Message Sample"
		.From = "sender.email@abc.com"
		.ReplyTo = "replyto.email@abc.com"
		.To = "recipient.email@abc.com"
		.Cc = "ReceiveCopy.email@abc.com"
		.Bcc = "ReceiveBlindCopy.email@abc.com"
		.subject = "Test Email from Task Scheduler"
		.Body = "All the Text that you want to have in the email."
		.Server = "SMTP_Server _Name"
		Dim objAttachments(1)	' Array of Attachments - Zero Based
		objAttachments(0) = WScript.ScriptFullName
		objAttachments(1) = "C:\Windows\Win.ini"
		.Attachments = objAttachments
		
		' Create a custom Header Field and value for your email.
		' I could not get this to work.  It stores in the XML properly
		' but task scheduler cannot open it and I had to manually delete from 
		' the tasks folder in System32
		'Dim objHeaderPair
		'objHeaderPair = .HeaderFields.Create
		'objHeaderPair.Name = "TestHeaderName"
		'objHeaderPair.Value = "TestHeaderValue"
		'.HeaderFields = objHeaderPair
	End With

	' *** Show a Message Box Action  ***  ---  http://msdn.microsoft.com/en-us/library/windows/desktop/aa382149%28v=vs.85%29.aspx
	Set objTaskAction = .Actions.Create(TASK_ACTION_SHOW_MESSAGE)
	With objTaskAction
		.Id = "Show a Message Sample"
		.Title = "Title for Message Box"
		.MessageBody = "Text in the Message Box"
	End With

	' -------------------------------------------------------------------------------
	' Set Values for Conditions and Settings Tabs
	' -------------------------------------------------------------------------------
	'http://msdn.microsoft.com/en-us/library/windows/desktop/aa383480%28v=vs.85%29.aspx
	With .Settings
		.Enabled = True 	' Must be set to two or task will have a status of disabled
		' Compatibility.  A Value of 0 or 1 will greatly restrict what objects can be used.
		' it is recommended that you user a 2 or 3 with this script
		' This value is not required and may be omitted 
		' 0 = Compatible with AT Command
		' 1 = Compatible with Task Scheduler 1.0
		' 2 = Compatible with Task Scheduler 2.0 (Windows Vista / Windows 2008)
		' 3 = Compatible with Task Scheduler 2.0 (Windows 7 / Windows 2008 R2) - this is not listed in the documentation
		.Compatibility = 2
		' Optional to Set Priority Level.  Can be omitted (recommended)
		' 0 = High / 10 = Low.  Setting not visible in GUI
		.Priority = 5
	
		' -------------------------------------------------------------------------------
		' General Tab
		' -------------------------------------------------------------------------------
		.Hidden = False ' If you mark this as hidden then you must have View >> Show Hidden Tasks enabled to see it 

		' -------------------------------------------------------------------------------
		' Conditions Tab
		' -------------------------------------------------------------------------------
		' Idle Section
		' Format For Days = P#D where # is the number of days
		' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
		.RunOnlyIfIdle = True		' Run this task only if the computer has been idle for selected period of time
		'http://msdn.microsoft.com/en-us/library/windows/desktop/aa380669%28v=vs.85%29.aspx
		With .IdleSettings 
			.IdleDuration = "PT2M"  ' Start the task only if computer is idle for selected time
			.StopOnIdleEnd = True	' Stop task if computer ceases to be idle
			.RestartOnIdle = True 	' Restart if the idle state resumes
			.WaitTimeout= "PT2H"	' Time to wait for idle
		End With 'objIdleSettings
		
		' Power Section
		.DisallowStartIfOnBatteries = True ' Start the task only if the computer is on AC Power
		.StopIfGoingOnBatteries = True	' Stop if the computer switches to battery power
		.WakeToRun = False	' Wake the computer to run this task
				
		' Network Section
		.RunOnlyIfNetworkAvailable = False ' Run only if the selected network is available. Must set network settings
		'http://msdn.microsoft.com/en-us/library/windows/desktop/aa382067%28v=vs.85%29.aspx
		'With .NetworkSettings 
		'	.Id = "{}" ' SSID For Network
		'	.Name = "MyNetwork" ' Network Name
		'End With 'objTaskNetworkSettings
		
		' -------------------------------------------------------------------------------
		' Settings Tab
		' -------------------------------------------------------------------------------
		.AllowDemandStart = True ' Allow the Task to Be Run on Demand
		.StartWhenAvailable = True ' Run Task as soon as possible after a scheduled start is missed.
		' Format For Days = P#D where # is the number of days
		' Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
		.RestartInterval = "PT10M"	' If the task fails attempt to restart every x period of Time (Seconds Not Valid for this Object
		.RestartCount = 2	' Attempt to restart x number of times.  Must be set if .RestartInterval is set
		.ExecutionTimeLimit = "PT1H"	' Stop the task if it runs longer than chosen time
		.AllowHardTerminate = False ' If the task does not end when requested, force it to stop
		.DeleteExpiredTaskAfter = "P30D"	' Must have at least one trigger with an expiration date to use this field.
		' Tell the task how to function if the task is initiated again while it is already running
		' 0 = Run a second instance now (Parallel)
		' 1 = Put the new instance in line behind the current running instance (Add To Queue)
		' 2 = Ignore the new request"
		.MultipleInstances = 2

	End With 'objTaskSettings

	' Alternatively you could create the task by creating the full XML document and assigning it to .xml
	'.xml = "<YourProperlyFormatedXMLString/>"
End With ' objNewTaskDefinition
			
' http://msdn.microsoft.com/en-us/library/windows/desktop/aa382577%28v=vs.85%29.aspx

' Register The Task
' 	Path = Name of the scheduled task
'   Definiition = The Task Definition object set above
'	Flags = Task Creation Constants (Bit Flags)
'	userId = The user credentials that are used to register the task. 
'		If present, these credentials take priority over the credentials 
'		specified in the task definition object pointed to by the definition parameter.
'	password = The password for the userId that is used to register the task. When 
'		the TASK_LOGON_SERVICE_ACCOUNT logon type is used, the password must be an 
'		empty VARIANT value such as VT_NULL or VT_EMPTY.
'	logonType = Task Logon Type Constant
'	ssdl = The security descriptor that is associated with the registered task. 
Call objTaskFolder.RegisterTaskDefinition( _
    "SampleTask From VBscript", objNewTaskDefinition, TASK_CREATE_OR_UPDATE, , , _
	TASK_LOGON_INTERACTIVE_TOKEN)

WScript.Echo "Task submitted."

wscript.quit

Function ConvertTime(DateTimeValue)
	' Convert a DateTime value to the format needed by 
	' task scheduler
	' YYYY-MM-DDTHH:MM:SS 
	Dim strTime
	
	strTime = year(DateTimeValue) & "-"
	strTime = strTime & Right("0" & Month(DateTimeValue), 2) & "-"
	strTime = strTime & Right("0" & Day(DateTimeValue), 2) & "T"
	strTime = strTime & Right("0" & Hour(DateTimeValue), 2) & ":"
	strTime = strTime & Right("0" & Minute(DateTimeValue), 2) & ":"
	strTime = strTime & Right("0" & Day(DateTimeValue), 2)
	
	ConvertTime = strTime
End Function