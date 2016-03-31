'------------------------------------------------
' ���ڸ�ʽ��
Function DateToStr(DateTime, ShowType)
    Dim DateMonth, DateDay, DateHour, DateMinute
    DateMonth = Month(DateTime)
    DateDay = Day(DateTime)
    DateHour = Hour(DateTime)
    DateMinute = Minute(DateTime)
    If Len(DateMonth) < 2 Then DateMonth = "0" & DateMonth
    If Len(DateDay) < 2 Then DateDay = "0" & DateDay
    Select Case ShowType
        Case "Y-m"
        DateToStr = Year(DateTime) & "-" & Month(DateTime)
        Case "Y-m-d"
        DateToStr = Year(DateTime) & "-" & DateMonth & "-" & DateDay
        Case "Y-m-d H:I A"
        Dim DateAMPM
        If DateHour > 12 Then
            DateHour = DateHour - 12
            DateAMPM = "PM"
        Else
            DateHour = DateHour
            DateAMPM = "AM"
        End If
        If Len(DateHour) < 2 Then DateHour = "0" & DateHour
        If Len(DateMinute) < 2 Then DateMinute = "0" & DateMinute
        DateToStr = Year(DateTime) & "-" & DateMonth & "-" & DateDay & " " & DateHour & ":" & DateMinute & " " & DateAMPM
        Case "Y-m-d H:I:S"
        Dim DateSecond
        DateSecond = Second(DateTime)
        If Len(DateHour) < 2 Then DateHour = "0" & DateHour
        If Len(DateMinute) < 2 Then DateMinute = "0" & DateMinute
        If Len(DateSecond) < 2 Then DateSecond = "0" & DateSecond
        DateToStr = Year(DateTime) & "-" & DateMonth & "-" & DateDay & " " & DateHour & ":" & DateMinute & ":" & DateSecond
        Case "YmdHIS"
        DateSecond = Second(DateTime)
        If Len(DateHour) < 2 Then DateHour = "0" & DateHour
        If Len(DateMinute) < 2 Then DateMinute = "0" & DateMinute
        If Len(DateSecond) < 2 Then DateSecond = "0" & DateSecond
        DateToStr = Year(DateTime) & DateMonth & DateDay & DateHour & DateMinute & DateSecond
        Case "Ymd"			
        DateToStr = Year(DateTime) & DateMonth & DateDay 
        Case "ym"
        DateToStr = Right(Year(DateTime), 2) & DateMonth
        Case "d"
        DateToStr = DateDay
        Case Else
        If Len(DateHour) < 2 Then DateHour = "0" & DateHour
        If Len(DateMinute) < 2 Then DateMinute = "0" & DateMinute
        DateToStr = Year(DateTime) & "-" & DateMonth & "-" & DateDay & " " & DateHour & ":" & DateMinute
    End Select
End Function



'------------------------------------------------
' ������ݼ��·ݵõ�ÿ�µ�������
Function GetDaysInMonth(iMonth, iYear) 
    Select Case iMonth 
        Case 1, 3, 5, 7, 8, 10, 12 
        GetDaysInMonth = 31 
        Case 4, 6, 9, 11 
        GetDaysInMonth = 30 
        Case 2 
        If IsDate("February 29, " & iYear) Then 
            GetDaysInMonth = 29 
        Else 
            GetDaysInMonth = 28 
        End If 
    End Select 
End Function 

'------------------------------------------------
' �õ�һ���¿�ʼ������
Function GetWeekdayMonthStartsOn(dAnyDayInTheMonth) 
    Dim dTemp 
    dTemp = DateAdd("d", -(Day(dAnyDayInTheMonth) - 1), dAnyDayInTheMonth) 
    GetWeekdayMonthStartsOn = WeekDay(dTemp) 
End Function 

'------------------------------------------------
' �õ���ǰһ���µ���һ����
Function SubtractOneMonth(dDate) 
    SubtractOneMonth = DateAdd("m", -1, dDate) 
End Function 

'------------------------------------------------
' �õ���ǰһ���µ���һ����
Function AddOneMonth(dDate) 
    AddOneMonth = DateAdd("m", 1, dDate) 
End Function 

'------------------------------------------------
' �������ڸ�ʽ��
Function Date2Chinese(iDate)
    Dim num(10)
    Dim iYear
    Dim iMonth
    Dim iDay
    
    num(0) = "��"
    num(1) = "һ"
    num(2) = "��"
    num(3) = "��"
    num(4) = "��"
    num(5) = "��"
    num(6) = "��"
    num(7) = "��"
    num(8) = "��"
    num(9) = "��"
    
    iYear = Year(iDate)
    iMonth = Month(iDate)
    iDay = Day(iDate)
    Date2Chinese = (num(iYear \ 1000) + num((iYear \ 100) Mod 10) + num((iYear\ 10) Mod 10) + num(iYear Mod 10)) & "��"
    
    If iMonth >= 10 Then
        If iMonth = 10 Then
            Date2Chinese = Date2Chinese & "ʮ" & "��"
        Else
            Date2Chinese = Date2Chinese & "ʮ" & num(iMonth Mod 10) & "��"
        End If
    Else
        Date2Chinese = Date2Chinese & num(iMonth Mod 10) & "��"
    End If
    
    If iDay >= 10 Then
        If iDay = 10 Then
            Date2Chinese = Date2Chinese & "ʮ" & "��"
        ElseIf iDay = 20 or iDay = 30 Then
            Date2Chinese = Date2Chinese & num(iDay \ 10) & "ʮ" & "��"
        ElseIf iDay > 20 Then
            Date2Chinese = Date2Chinese & num(iDay \ 10) & "ʮ" & num(iDay Mod 10) & "��"
        Else
            Date2Chinese = Date2Chinese & "ʮ" & num(iDay Mod 10) & "��"
        End If
    Else
        Date2Chinese = Date2Chinese & num(iDay Mod 10) & "��"
    End If
    
End Function

'------------------------------------------------
' Date2ChineseRSS
Function Date2ChineseRSS(iDate)
    Dim num(10)
    Dim iYear
    Dim iMonth
    Dim iDay
    
    num(0) = "��"
    num(1) = "һ"
    num(2) = "��"
    num(3) = "��"
    num(4) = "��"
    num(5) = "��"
    num(6) = "��"
    num(7) = "��"
    num(8) = "��"
    num(9) = "��"
    
    iYear = Year(iDate)
    iMonth = Month(iDate)
    iDay = Day(iDate)
    Date2ChineseRSS = iYear & "��"
    
    If iMonth >= 10 Then
        If iMonth = 10 Then
            Date2ChineseRSS = Date2ChineseRSS & "ʮ" & "��"
        Else
            Date2ChineseRSS = Date2ChineseRSS & "ʮ" & num(iMonth Mod 10) & "��"
        End If
    Else
        Date2ChineseRSS = Date2ChineseRSS & num(iMonth Mod 10) & "��"
    End If
    
End Function


'------------------------------------------------
' Convert a string to a date or datetime
' IN  : sDate (string) : source (format YYYYMMDD HH:MM:SS or YYYYMMDD)
' OUT : (datetime) : destination
Function StringToDate(strDate)
    Dim dDate, sDate
    
    sDate = Trim(strDate)
    Select Case Len(sDate)
        Case 17
        dDate = DateSerial(Left(sDate, 4), Mid(sDate, 5, 2), Mid(sDate, 7, 2)) + TimeSerial(Mid(sDate, 10, 2), Mid(sDate, 13, 2), Mid(sDate, 16, 2))
        Case 8
        dDate = DateSerial(Left(sDate, 4), Mid(sDate, 5, 2), Mid(sDate, 7, 2))
        Case Else
        If isDate(sDate) Then
            dDate = CDate(sDate)
        End If
    End Select
    StringToDate = dDate
End Function
