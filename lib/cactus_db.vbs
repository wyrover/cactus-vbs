'------------------------------------------------
' DBControl Class
Class DBControl

    Private m_connectionString
    Private m_conn
    Private m_dbType
    
    Private Sub Class_Initialize
        m_dbType = "ACCESS"
        m_connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & config("database_connectionstring")
    End Sub

    Private Sub Class_Terminate        
    End Sub

    Public Property Get ConnectionString()
        ConnectionString = m_connectionString 
    End Property

    Public Function Open()
        On Error Resume Next
        Set m_conn = CreateObject(COM_ADO_CONN)
        m_conn.Open m_connectionString	
        If Err Then
            Err.Clear
            Set m_conn = Nothing
            Response.Write "数据库连接出错，请检查连接字串。"
            Response.End
        End If
    End Function


    Public Function Close()
        m_conn.Close
        Set m_conn = Nothing
    End Function


    Public Function CreateRS()
        Set CreateRS = CreateObject(COM_ADO_RECORDSET)
    End Function

    Public Function BeginTrans()
        m_conn.BeginTrans 
        on error resume next
    End Function

    Public Function EndTrans()
        If Err.number = 0 Then
            m_conn.CommitTrans  
        Else
            m_conn.RollbackTrans 
            strerr = Err.Description
            Response.Write "数据库错误！错误日志：<font color=red>"&strerr &"</font>"
            Response.End
        End If
    End Function


    '函数:根据当前数据库类型转换Sql脚本
    '参数:Sql串
    '返回:转换结果Sql串
    Public Function SqlTran(Sql)
        If m_dbType = "ACCESS" Then
            SqlTran = SqlServer_To_Access(Sql)
        Else
            SqlTran = Sql
        End If
    End Function

    '函数:数据库脚本执行(代Sql转换)
    '参数:Sql脚本
    '返回:执行结果
    '说明:本执行可自动根据数据库类型对部分Sql基础语法进行转换执行
    Public Function ExeCute(Sql)        
        If config("isdebug") = 0 Then 
            On Error Resume Next
            Sql = SqlTran(Sql)
            Set ExeCute = m_conn.ExeCute(Sql)
            If Err Then
                    Err.Clear
                    Set m_conn = Nothing
                    Response.Write "查询数据的时候发现错误,请检查您的查询代码是否正确.<br /><li>"
                    Response.Write Sql
                    Response.End
            End If
        Else
            Set ExeCute = m_conn.ExeCute(Sql)
        End If
        SQL_QUERY_NUM = SQL_QUERY_NUM + 1
    End Function

    '函数:数据库脚本执行
    '参数:Sql脚本
    '返回:执行结果
    Public Function ExeCute2(Sql)
        Set ExeCute2 = m_conn.ExeCute(Sql)
    End Function

    Public Function ExeCute3(sql_proc, ByRef parameters)
        Set cmd = CreateObject(COM_ADO_COMMAND)
        With cmd
            .ActiveConnection = m_conn
            .CommandType = &H0004 '存储过程
            .CommandText = sql_proc
        End With
        Set ExeCute3 = cmd.Execute(, parameters)
    End Function

    '函数:SqlServer(97-2000) to Access(97-2000)
    '参数:Sql,数据库类型(ACCESS,SQLSERVER)
    '说明:
    Public Function SqlServer_To_Access(Sql)
        Dim regEx, Matches, Match
        '创建正则对象
        Set regEx = New RegExp
        regEx.IgnoreCase = True
        regEx.Global = True
        regEx.MultiLine = True

        '转:GetDate()
        regEx.Pattern = "(?=[^']?)GETDATE\(\)(?=[^']?)"
        Sql = regEx.Replace(Sql,"NOW()")

        '转:UPPER()
        regEx.Pattern = "(?=[^']?)UPPER\([\s]?(.+?)[\s]?\)(?=[^']?)"
        Sql = regEx.Replace(Sql,"UCASE($1)")

        '转:日期表示方式
        '说明:时间格式必须是2004-23-23 11:11:10 标准格式
        regEx.Pattern = "'([\d]{4,4}\-[\d]{1,2}\-[\d]{1,2}(?:[\s][\d]{1,2}:[\d]{1,2}:[\d]{1,2})?)'"
        Sql = regEx.Replace(Sql,"#$1#")
        
        regEx.Pattern = "DATEDIFF\([\s]?(second|minute|hour|day|month|year)[\s]?\,[\s]?(.+?)[\s]?\,[\s]?(.+?)([\s]?\)[\s]?)"
        Set Matches = regEx.ExeCute(Sql)
        Dim temStr
        For Each Match In Matches
            temStr = "DATEDIFF("
            Select Case lcase(Match.SubMatches(0))
                Case "second" :
                    temStr = temStr & "'s'"
                Case "minute" :
                    temStr = temStr & "'n'"
                Case "hour" :
                    temStr = temStr & "'h'"
                Case "day" :
                    temStr = temStr & "'d'"
                Case "month" :
                    temStr = temStr & "'m'"
                Case "year" :
                    temStr = temStr & "'y'"
            End Select
            temStr = temStr & "," & Match.SubMatches(1) & "," &  Match.SubMatches(2) & Match.SubMatches(3)
            Sql = Replace(Sql,Match.Value,temStr,1,1)
        Next

        '转:Insert函数
        regEx.Pattern = "CHARINDEX\([\s]?'(.+?)'[\s]?,[\s]?'(.+?)'[\s]?\)[\s]?"
        Sql = regEx.Replace(Sql,"INSTR('$2','$1')")

        Set regEx = Nothing
        SqlServer_To_Access = Sql
    End Function    
End Class