'------------------------------------------------
' TagParser Class
Class TagParser

    Private TempContent     ' 临时模版
    Private ResColl         ' 字典对象, 存放标记和标记要替换的内容

    Private Sub Class_Initialize
        Set ResColl = CreateObject(COM_DICT)        
    End Sub

    Private Sub Class_Terminate
        Set ResColl = Nothing
    End Sub

    Public Function Parser(Str)
        TempContent = Str

        ' 开始解析模版
        Tag_Parser()
                
        Parser = TempContent
    End Function


    Private Function Tag_Parser()
        Dim regex, matches, match
        set regex = New RegExp
        regex.IgnoreCase = False
        regex.Global = True
        regex.MultiLine = True

        regex.Pattern = "<cms:file>([^\b]+?)</cms:file>"
        Set matches2 = regex.Execute(TempContent)
        For Each match2 In matches2
            retVal = GetCacheValue(match2.Value)
            If retVal = "" Then 
                If match2.SubMatches(0) <> "" Then
                    retVal = Tag_File_Parser(match2.SubMatches(0))
                End If
            End If
            TempContent = Replace(TempContent, match2.Value,  retVal)
            SetCacheValue match2.Value, retVal, 5
        Next

        regex.Pattern = "<cms:list>([^\b]+?)</cms:list>"
        Set matches = regex.Execute(TempContent)

        Dim strContent, tmpItem
        For Each match In matches
            If match.SubMatches(0) <> "" Then
                'TempContent = match.SubMatches(0)
                'ResColl.Add match.Value Tag_Parser2(match.SubMatches(0))

                TempContent = Replace(TempContent, match.Value,  Tag_Parser2(match.SubMatches(0)))
            End If
        Next

        regex.Pattern = "<cms:function>([^\b]+?)</cms:function>"
        Set matches2 = regex.Execute(TempContent)
        For Each match2 In matches2
            retVal = GetCacheValue(match2.Value)
            If retVal = "" Then 
                If match2.SubMatches(0) <> "" Then
                    Execute("retVal = " & match2.SubMatches(0))
                End If
            End If
            TempContent = Replace(TempContent, match2.Value,  retVal)
            SetCacheValue match2.Value, retVal, 120	
        Next

        regex.Pattern = "<cms:pager>(.*?)</cms:pager>"
        Set matches2 = regex.Execute(TempContent)
        For Each match2 In matches2
            If match2.SubMatches(0) <> "" Then
                TempContent = Replace(TempContent, match2.Value,  "-------------pager-------------------")
            End If
        Next


    End Function

    Private Function Tag_File_Parser(strCommand)
        Dim regex, matches, match, retVal
        set regex = New RegExp
        regex.IgnoreCase = False
        regex.Global = True
        regex.MultiLine = True
        regex.Pattern = "\$([^\b]+?)\$"
        Set matches = regex.Execute(strCommand)
        For Each match In matches
            If match.SubMatches(0) <> "" Then
                filepath = Server.MapPath(".") & "\system\" & Application_PATH & "\views\" & match.SubMatches(0)
                Set filestream = Server.CreateObject(COM_ADOSTREAM)
                    With filestream
                        .Type = 2 '以本模式读取
                        .Mode = 3 
                        .Charset = "utf-8"
                        .Open
                        .Loadfromfile filepath
                        retVal = .readtext
                        .Close
                    End With
                Set filestream = Nothing
            End If
        Next
        Tag_File_Parser = retVal
    End Function

    Private Function Tag_Parser2(strCommand)
        Dim regex, matches, match, retVal, temp
        set regex = New RegExp
        regex.IgnoreCase = False
        regex.Global = True
        regex.MultiLine = True
        regex.Pattern = "<sql>([^\b]+?)</sql>[^\b]*?<template>([^\b]+?)</template>[^\b]*?<cache>([^\b]+?)</cache>"
        Set matches = regex.Execute(strCommand)
        For Each match In matches
            retVal = GetCacheValue(match.Value)
            If retVal = "" Then 
                If match.SubMatches(0) <> "" And match.SubMatches(1) <> "" And match.SubMatches(2) <> "" Then
                    Dim sql, strTemplate, rs, strHTML, strTemplate2
                    sql = match.SubMatches(0)
                    strTemplate = match.SubMatches(1)

                    Dim matches2, match2
                    
                    regex.Pattern = "\$(\w+?)\$"
                    set matches2 = regex.Execute(strTemplate)

                    Dim matches3, match3
                    
                    regex.Pattern = "\$(\w+?)\[(\d+?)\]\$"
                    set matches3 = regex.Execute(strTemplate)

                    Dim matches4, match4

                    regex.Pattern = "\$(\w+?)\((.+?)\)\$"
                    set matches4 = regex.Execute(strTemplate)


                    Set rs = Db.ExeCute(sql)
                    While Not rs.Eof
                        
                        strTemplate2 = strTemplate

                        For Each match4 In matches4
                            'Response.Write match4.SubMatches(1)
                            Dim tempArray, strA
                            tempArray = Split(match4.SubMatches(1), ",")
                            strA = "temp = " & match4.SubMatches(0) & "("
                            For i = 0 To UBound(tempArray)
                                tempArray(i) = rs(Trim(tempArray(i)))
                                If i <> UBound(tempArray) Then
                                    strA = strA & "tempArray(" & i & ")," 
                                Else
                                    strA = strA & "tempArray(" & i & ")"
                                End If
                            Next
                            strA = strA & ")"
                            
                            
                            Execute(strA)
                            strTemplate2 = Replace(strTemplate2, match4.Value, temp)
                        Next

                        For Each match3 In matches3
                            strTemplate2 = Replace(strTemplate2, match3.Value, Left(rs(match3.SubMatches(0)), match3.SubMatches(1)))
                        Next
                                        
                        For Each match2 In matches2
                                strTemplate2 = Replace(strTemplate2, match2.Value, rs(match2.SubMatches(0)))
                        Next

                        strHTML = strHTML & strTemplate2 & vbCrLf
                        rs.MoveNext()
                    Wend
                    rs.Close
                    Set rs = Nothing
                    retVal = strHTML
                    SetCacheValue match.Value, retVal, Int(match.SubMatches(2))
                End If
            End If 
        Next

        Tag_Parser2 = retVal

    End Function
End Class