Function SendMail(LogFile , PngFile)
    Set oMessage=WScript.CreateObject("CDO.Message")
    Set oConf=WScript.CreateObject("CDO.Configuration")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set pso = CreateObject("Scripting.FileSystemObject")

    'Set server,port and other information about CDO.Configuration Object.(IIS SMTP)
    oConf.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
    oConf.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")="smtp.gmail.com"
    oConf.Fields("http://schemas.microsoft.com/cdo/configuration/serverport")=25
    oConf.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")=1
    oConf.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername")="redmine@cocos2d-x.org"
    oConf.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword")="cocos2d-x.org"
    oConf.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl")=1
    oConf.Fields.Update()

    'Set subject,attachments and other information about CDO.Message Object.
    oMessage.Configuration = oConf
    oMessage.To = "739657621@qq.com"
    oMessage.From = "redmine@cocos2d-x.org"
    oMessage.Subject = "QTRunner Notification"

    file = fso.GetAbsolutePathName(LogFile)
    Set fso = Nothing
    oMessage.AddAttachment( file )
    picture = pso.GetAbsolutePathName(PngFile)
    Set pso = Nothing
    oMessage.AddAttachment( picture )

    TextBody = "QTRunner Finish! See attachment for logs."    
    oMessage.TextBody = TextBody
    oMessage.Send()
End Function



' Mail class allows direct email via smtp
Class std_mail
	Private Sub Class_Initialize()
		Set objMessage = CreateObject("CDO.Message") 
		' Since for SMTP mail only
		objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
		timeout = 60
		port = 25
		enablessl = False	
		server = "localhost"
		setAuthTypeAnonymous
	End Sub
	Public Property Let subject( strSubject )
	objMessage.Subject = strSubject
	End Property
	Public Property Let from( strFrom )
	objMessage.From = strFrom
	End Property
	Public Property Let recipient( strTo )
	objMessage.To = strTo
	End Property
	Public Property Let cc( strCc )
	objMessage.Cc = strCc
	End Property
	Public Property Let bcc( strBcc )
	objMessage.Bcc = strBcc
	End Property
	Public Property Let body( strBody )
	objMessage.TextBody = strBody
	End Property
	Public Property Let server( strServer )
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strServer
	End Property 
	Public Sub setAuthTypeAnonymous
		objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0 'No Authentication
	End Sub
	Public Sub setAuthTypeBasic
		objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'Basic Authentication
	End Sub
	Public Sub setAuthTypeNTLM
		objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 2 'NTLM
	End Sub
	Public Property Let user( strUser )
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = strUser
	End Property
	Public Property Let password( strPassword )
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strPassword
	End Property
	Public Property Let port( nPort )
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = nPort
	End Property
	Public Property Let enablessl( bUseSSL )
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = bUseSSL
	End Property
	Public Property Let timeout( nTimeout )
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = nTimeout
	End Property
	Public Property Let htmlbody( strHTML )
	objMessage.HTMLBody = strHTML
	End Property
	Public Property Let HtmlFileorUrlBody( strHtmlFileorUrl )
	Call objMessage.CreateMHTMLBody( strHtmlFileorUrl )
	End Property
	Public Property Let attachment( strAttachment )
	Call objMessage.AddAttachment( strAttachment )
	End Property
	Sub Send
		objMessage.Configuration.Fields.Update
		objMessage.Send
	End Sub 
	Private objMessage
End Class
