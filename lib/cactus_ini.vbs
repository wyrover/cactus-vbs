'------------------------------------------------
' WriteINIString
' example:
'   WriteINIString "Mail", "MAPI", "1", "win.ini"
'   wscript.echo GetINIString("Mail", "MAPI", "-", "win.ini")
Sub WriteINIString(Section, KeyName, Value, FileName)
    Dim INIContents, PosSection, PosEndSection
    
    
    INIContents = ReadFile(FileName)
    
    'Find section
    PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
    If PosSection > 0 Then
        'Section exists. Find end of section
        PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
        '?Is this last section?
        If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
        
        'Separate section contents
        Dim OldsContents, NewsContents, Line
        Dim sKeyName, Found
        OldsContents = Mid(INIContents, PosSection, PosEndSection - PosSection)
        OldsContents = split(OldsContents, vbCrLf)
        
        'Temp variable To find a Key
        sKeyName = LCase(KeyName & "=")
        
        'Enumerate section lines
        For Each Line In OldsContents
            If LCase(Left(Line, Len(sKeyName))) = sKeyName Then
                Line = KeyName & "=" & Value
                Found = True
            End If
            NewsContents = NewsContents & Line & vbCrLf
        Next
        
        If isempty(Found) Then
            'key Not found - add it at the end of section
            NewsContents = NewsContents & KeyName & "=" & Value
        Else
            'remove last vbCrLf - the vbCrLf is at PosEndSection
            NewsContents = Left(NewsContents, Len(NewsContents) - 2)
        End If
        
        'Combine pre-section, new section And post-section data.
        INIContents = Left(INIContents, PosSection-1) & _
        NewsContents & Mid(INIContents, PosEndSection)
    Else        'if PosSection>0 Then
        'Section Not found. Add section data at the end of file contents.
        If Right(INIContents, 2) <> vbCrLf And Len(INIContents)>0 Then 
            INIContents = INIContents & vbCrLf 
        End If
        INIContents = INIContents & "[" & Section & "]" & vbCrLf & _
        KeyName & "=" & Value
    End If      'if PosSection>0 Then
    WriteFile FileName, INIContents
End Sub

'------------------------------------------------
' GetINIString
Function GetINIString(Section, KeyName, Default, FileName)
    Dim INIContents, PosSection, PosEndSection, sContents, Value, Found
    
    
    INIContents = ReadFile(FileName)
    
    'Find section
    PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
    If PosSection > 0 Then
        'Section exists. Find end of section
        PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
        '?Is this last section?
        If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
        
        'Separate section contents
        sContents = Mid(INIContents, PosSection, PosEndSection - PosSection)
        
        If InStr(1, sContents, vbCrLf & KeyName & "=", vbTextCompare)>0 Then
            Found = True
            'Separate value of a key.
            Value = SeparateField(sContents, vbCrLf & KeyName & "=", vbCrLf)
        End If
    End If
    
    If isempty(Found) Then Value = Default
    
    GetINIString = Value
End Function

'------------------------------------------------
' Separates one field between sStart And sEnd
Function SeparateField(ByVal sFrom, ByVal sStart, ByVal sEnd)
    Dim PosB: PosB = InStr(1, sFrom, sStart, 1)
    If PosB > 0 Then
        PosB = PosB + Len(sStart)
        Dim PosE: PosE = InStr(PosB, sFrom, sEnd, 1)
        If PosE = 0 Then PosE = InStr(PosB, sFrom, vbCrLf, 1)
        If PosE = 0 Then PosE = Len(sFrom) + 1
        SeparateField = Mid(sFrom, PosB, PosE - PosB)
    End If
End Function




'*************************************************************************************
'
'
'	INI FILE CLASS AND SUB CLASSES START HERE 
'
'
'*************************************************************************************

'IniFile class used to read and write ini files by loading the file into memory
Class IniFile
    'List of IniSection objects keeps track of all the sections in the INI file
    Private m_pSections
    'Public constructor
    Public Sub Class_Initialize()
        Set m_pSections = CreateObject("Scripting.Dictionary")
		m_pSections.CompareMode = vbTextCompare
    End Sub
    
    'Returns an array of the IniSections in the IniFile
    Public Property Get Sections
         Sections = m_pSections.Items
    End Property

    'Load IniFile object with existing INI Data
    Public Sub Load( ByVal sFileName , ByVal bAppend )
    	If Not bAppend Then RemoveAllSections() ' Clear the object...
    	
        Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject") 
        Dim tempsection : Set tempsection = Nothing
        Dim oReader : Set oReader = objFSO.OpenTextFile( sFileName )
        Dim regexsection : set regexsection = new regexp 
        Dim regexkey : Set regexkey = new regexp
		Dim regexcomment : Set regexcomment = new regexp
          
        regexcomment.Pattern = "^([\s]*#.*)"
		regexcomment.Global = False
		regexcomment.IgnoreCase = True
		regexcomment.MultiLine = False
		
		' Left for history
        'regexsection.Pattern = "\[[\s]*([^\[\s].*[^\s\]])[\s]*\]"
        regexsection.Pattern = "^[\s]*\[[\s]*([^\[\s].*[^\s\]])[\s]*\][\s]*$"
		regexsection.Global = False
		regexsection.IgnoreCase = True
		regexsection.MultiLine = False
		
		regexkey.Pattern = "^\s*([^=\s]*)[^=]*=(.*)" 
		regexkey.Global = False
		regexkey.IgnoreCase = True
		regexkey.MultiLine = False
		
        While Not oReader.AtEndOfStream
        	Dim line : line = oReader.ReadLine()
			If line <> "" Then
            	Dim m
                If regexcomment.Test(line) Then
                    Set m = regexcomment.Execute(line)
		    		'WScript.Echo("Skipping Comment " & m.Item(0).subMatches.Item(0) )
                ElseIf regexsection.Test(line) Then
                    Set m = regexsection.Execute(line)
                    'WScript.Echo("Adding section [" & m.Item(0).subMatches.Item(0) &"]" )
                    Set tempsection = AddSection( m.Item(0).subMatches.Item(0) )
                ElseIf regexkey.Test(line) And Not tempsection Is Nothing Then
                    Set m = regexkey.Execute(line)
                    'WScript.Echo("Adding Key ["& m.Item(0).subMatches.Item(0) &"]=["& m.Item(0).subMatches.Item(1) &"]" )
                    tempsection.AddKey( m.Item(0).subMatches.Item(0) ).Value = m.Item(0).subMatches.Item(1)
				ElseIf Not tempsection Is Nothing Then
					'WScript.Echo("Adding Key ["& line &"]" )
		    		tempsection.AddKey( line )
                'Else
                 '   WScript.Echo("Skipping unknown type of data: " & line)
                End If
            End If
        Wend
        oReader.Close()
    End Sub


    'Allows you to do a save the IniFile resident in memory to file
    Public Sub Save(ByVal sFileName)
    	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
        Dim oWriter : Set oWriter = objFSO.CreateTextFile( sFileName , True )
        Dim s 'IniSection
        Dim k 'IniKey
        For Each s In Sections
            'WScript.Echo("Writing Section: " & s.Name)
            oWriter.WriteLine("[" & s.Name & "]")
            For Each k In s.Keys
				If k.Value <> vbNullString Then
	            	'WScript.Echo("Writing Key: "&k.Name&"="&k.Value)
	                oWriter.WriteLine(k.Name & "="& k.Value )
				Else
					'WScript.Echo("Writing Key: "&k.Name)
					oWriter.WriteLine(k.Name)
				End if
            Next
        Next
        oWriter.Close()
    End Sub

    'Returns the IniSection object associated with a section name
    Public Function GetSection(ByVal sSection )
        Set GetSection = Nothing
        sSection = Trim(sSection) 'Trim spaces
        If Len( sSection ) <> 0 Then
        	If m_pSections.Exists( sSection ) Then
        		Set GetSection = m_pSections.Item(sSection)
        	End If
        End If
    End Function

    ' Adds a section to the IniFile object, returns a IniSection object
    Public Function AddSection(ByVal sSection )
    	Set AddSection = Nothing
		If StrComp(TypeName(sSection),"IniSection",1) = 0 Then 
			If Not sSection Is Nothing Then
				' Only purpose is to be used by child to re-insert
				If Not sSection Is Nothing Then
					If Not m_pSections.Exists( sSection.Name ) Then
	       				Set m_pSections.Item( sSection.Name ) = sSection
	       				Set AddSection = sSection
	    	   		End If
				End If
			End If 
		ElseIf StrComp(TypeName(sSection),"String",1) = 0 Then
        	sSection = Trim(sSection)
        	If Len( sSection ) <> 0 Then 
	        	If m_pSections.Exists( sSection ) Then
        			Set AddSection = m_pSections.Item(sSection)
        		Else
	        		Dim s : Set s = New IniSection
        			Call s.Init( Me , sSection )
        			Set m_pSections.Item(sSection) = s
        			Set AddSection = s
        		End If
	        End If		
		End If
    End Function

    ' Removes all existing sections (clears the object) 
    Public Sub RemoveAllSections()
        Call m_pSections.RemoveAll()
    End Sub

	' Remove a section by name or section object
	Public Function RemoveSection(ByVal Obj)
		RemoveSection = False
		If StrComp(TypeName(Obj),"IniSection",1) = 0 Then 
			If Not Obj Is Nothing Then
				m_pSections.Remove(Obj.Name)
				RemoveSection = True
			End If 
		ElseIf StrComp(TypeName(Obj),"String",1) = 0 Then
			RemoveSection = RemoveSection( GetSection(Obj) )
		End If 
	End Function

	' Remove a key by section namd and key name
	Public Function RemoveKey(ByVal sSection , ByVal sKey)
		RemoveKey = False
        Dim s : Set s = GetSection(sSection)
        If Not s Is Nothing Then
            RemoveKey = s.RemoveKey( sKey )
        End If
	End Function
	
    ' Returns a KeyValue in a certain section
    Public Function GetKeyValue(ByVal sSection , ByVal sKey )
    	GetKeyValue = vbNullString
        Dim s : Set s = GetSection(sSection)
        If Not s Is Nothing Then
            Dim k : Set k = s.GetKey(sKey)
            If Not k Is Nothing Then
                GetKeyValue = k.Value
            End If
        End If
    End Function

    ' Sets a KeyValuePair in a certain section
    Public Function SetKeyValue(ByVal sSection , ByVal sKey , ByVal sValue )
        SetKeyValue = False
        Dim s : Set s = AddSection(sSection)
        If Not s Is Nothing Then
            Dim k : Set k = s.AddKey(sKey)
            If Not s Is Nothing Then
                k.Value = sValue
                SetKeyValue = True
            End If
        End If
    End Function

    ' Renames an existing section returns true on success, false if the section didn't exist or there was another section with the same sNewSection
    Public Function RenameSection(ByVal sSection , ByVal sNewSection)
        ' Note string trims are done in lower calls.
        RenameSection = False
        Dim s : Set s = GetSection(sSection)
        If Not s Is Nothing Then
            RenameSection = s.SetName(sNewSection)
        End If
    End Function

    ' Renames an existing key returns true on success, false if the key didn't exist or there was another section with the same sNewKey
    Public Function RenameKey(ByVal sSection , ByVal sKey, ByVal sNewKey)
        ' Note string trims are done in lower calls.
        RenameKey = False
        Dim s : Set s = GetSection(sSection)
        If Not s Is Nothing Then
            Dim k : Set k = s.GetKey(sKey)
            If Not k Is Nothing Then
                RenameKey = k.SetName(sNewKey)
            End If
        End If
    End Function
	
End Class


'IniSection Class 
Class IniSection
   ' IniFile IniFile object instance
   Private m_pIniFile
   ' Name of the section
   Private m_sSection
   ' List of IniKeys in the section
   Private m_keys

   'Friend constuctor so objects are internally managed
   Public Sub Class_Initialize
       Set m_pIniFile = Nothing
       m_sSection = ""
       Set m_keys = CreateObject("Scripting.Dictionary")
       m_keys.CompareMode = vbTextCompare
   End Sub

   	' Function only works once...
   	Public Sub Init( ByVal oIniFile , ByVal sSection )
   		If m_pIniFile is Nothing Then 
   			Set m_pIniFile = oIniFile
   			m_sSection = sSection
   		End If 	
   	End Sub

    'Returns an array of the IniKeys in the IniFile
    Public Property Get Keys
         Keys = m_keys.Items
    End Property
    
   'Returns the section name
   Public Property Get Name
           name = m_sSection
   End Property
    
   'Set the section name
   'Returns true on success, False if key already exists in the section
   Public Function SetName(ByVal sSection)
       SetName = False  ' Default
       sSection = Trim(sSection)
       If Len( sSection ) <> 0 Then
           Dim s : Set s = m_pIniFile.GetSection(sSection)
           If Not s Is Me And Not s Is Nothing Then Exit Function
           Call m_pIniFile.RemoveSection(Me)
           m_sSection = sSection
           Call m_pIniFile.AddSection(Me)           
           SetName = True
       End If
   End Function

   'Returns the section name
   Public Function GetName()
           GetName = m_sSection
   End Function
   
   'Adds a key to the IniSection object
   'Returns Nothing on failure
	Public Function AddKey(ByVal sKey)
		Set AddKey = Nothing
		' Is this a string or object of IniKey
		If StrComp(TypeName(sKey),"IniKey",1) = 0 Then 
			' Only purpose is to be used by child to re-insert
			If Not sKey Is Nothing Then
				If Not m_keys.Exists( sKey.Name ) Then
	       			Set m_keys.Item(sKey.Name) = sKey
	       			Set AddKey = sKey
	    	   	End If
			End If
		ElseIf StrComp(TypeName(sKey),"String",1) = 0 Then
			' String was passed...
			sKey = Trim(sKey)
			If Len(sKey) <> 0 Then
				If m_keys.Exists( sKey ) Then
		       		Set AddKey = m_keys.Item(sKey)
	    	   	Else
	       			Dim k : Set k = New IniKey
		       		Call k.Init( Me , sKey )
	       			Set m_keys.Item(sKey) = k
	       			Set AddKey = k
	       		End If
       		End If
		End If        
   End Function

   'Returns a IniKey
   'Returns Nothing on failure 
	Public Function GetKey(ByVal sKey)
   		Set GetKey = Nothing
    	sKey = Trim(sKey)
    	If Len(sKey) <> 0 Then
	       	If m_keys.Exists( sKey ) Then
	       		Set GetKey = m_keys.Item(sKey)
	       	End If 
		End If
	End Function

   'Removes all the keys in the section
   Public Sub RemoveAllKeys()
       Call m_keys.RemoveAll()
   End Sub
   
   'Removes a single key by IniKey object by string or object
   Public Function RemoveKey(ByVal Obj)
       	RemoveKey = False
		If StrComp(TypeName(Obj),"IniKey",1) = 0 Then 
			If Not Obj Is Nothing Then
			 	m_keys.Remove(Obj.Name)
				RemoveKey = True
			End If
		ElseIf StrComp(TypeName(Obj),"String",1) = 0 Then
			RemoveKey = RemoveKey( GetKey(Obj) )
		End If        
   End Function

End Class  ' End of IniSection
    
'IniKey Class
Class IniKey
    ' Name of the Key
    Private m_sKey
    ' Value associated
    Private m_sValue
    ' Pointer to the parent CIniSection
    Private m_pSection

    'Friend constuctor so objects are internally managed
    Public Sub Class_Initialize
    	m_sKey = ""
    	m_sValue = ""
        Set m_pSection = Nothing
    End Sub

    'Returns the key's parent IniSection
    Public Sub Init( ByVal oIniSection , ByVal sKey )
           If m_pSection Is Nothing Then
            	Set m_pSection = oIniSection
            	m_sKey = sKey
           End If 
    End Sub
    
    'Returns the name of the Key
    Public Property Get Name
            name = m_sKey
    End Property

'     'Gets\Sets the value associated with the Key
     Public Property Let Value( strKeyValue )
             m_sValue = strKeyValue
     End Property

    'Gets\Sets the value associated with the Key
    Public Property Get Value()
            value = m_sValue
    End Property

    'Sets the key name
    'Returns true on success, fails if the section name sKeyName already exists
    Public Function SetName(ByVal sKey )
    	SetName = False
        sKey = Trim(sKey)
        If Len(sKey) <> 0 Then
            Dim s : Set s = m_pSection.GetKey(sKey)
            If Not s Is Me And Not s Is Nothing Then Exit Function
            Call m_pSection.RemoveKey(Me)
            ' Set our new name
            m_sKey = sKey
            ' Put our own object back
            Call m_pSection.AddKey(Me)
            SetName = True
        End If
    End Function
    
    ' Returns the current key name
    Public Function GetName()
		    GetName = m_sKey
    End Function
    
End Class
