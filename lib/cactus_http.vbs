'------------------------------------------------
' �����ļ�
Sub DwonloadFile(url,target)    
    Dim http, adodbStream
    Set http = CreateObject(COM_HTTP)
    http.open "GET",url,False
    http.send
    Set adodbStream = createobject(COM_ADOSTREAM)
    adodbStream.Type = adTypeBinary
    adodbStream.Open
    adodbStream.Write http.responseBody
    adodbStream.SaveToFile target
    adodbStream.Close
    Set adodbStream = Nothing
End Sub


'------------------------------------------------
' GetHttp
Function GetHttp(url) 
    Dim xmlhttp
    Set xmlhttp = CreateObject(COM_XMLHTTP)  
    postdata = "" 
    xmlhttp.Open "GET", url, False 
    xmlhttp.setRequestHeader "Authorization", "Basic " & Base64encode("test:pass") 
    'xmlhttp.setRequestHeader("Referer","��·�ľ��Ե�ַ") 
    'xmlhttp.setRequestHeader "Cookie",Cookies   'Cookie 
    xmlhttp.Send postdata 
    Wscript.echo xmlhttp.status & ":" & xmlhttp.statusText 
    respStr = BytesToBstr(xmlhttp.responseBody, "UTF-8") 
    Wscript.echo respStr 
    Set xmlhttp = nothing 
End Function 

'------------------------------------------------
' HttpGet
' @url          URL��ַ
' @charset      ��ҳ����(gb2312, utf-8)
Function HttpGet(url, charset)
    Dim xmlhttp
    Set xmlhttp = CreateObject(COM_XMLHTTP)    
    xmlhttp.Open "GET", url, False     
    xmlhttp.Send() 
    If xmlhttp.readystate <> 4 Then
        Exit Function
    End If
    HttpGet = BytesToBstr(xmlhttp.responseBody, charset)     
    Set xmlhttp = nothing 
End Function


'------------------------------------------------
' PostHttp
Function PostHttp(url) 
    Set xmlhttp = CreateObject(COM_XMLHTTP)  
    postdata = "" 
    xmlhttp.Open "POST", url1, False 
    xmlhttp.setRequestHeader "CONTENT-TYPE","application/x-www-form-urlencoded" 
    xmlhttp.setRequestHeader "Authorization", "Basic " & Base64encode("test:pass") 
    'xmlhttp.setRequestHeader("Referer","��·�ľ��Ե�ַ") 
    'xmlhttp.setRequestHeader "Cookie",Cookies   'Cookie 
    xmlhttp.Send postdata 
    Wscript.echo xmlhttp.status & ":" & xmlhttp.statusText 
    respStr = BytesToBstr(xmlhttp.responseBody, "GB2312") 
    Wscript.echo respStr 
    Set xmlhttp = nothing 
End Function 

'------------------------------------------------
' HttpPostByIE
Function HttpPostByIE()
    Dim objIE, strRet, Stream, postData, strHeaders
    Set objIE = CreateObject("InternetExplorer.Application")

    ' �ݥ��ȥǩ`��
    ' �}���΂������Ť���Ȥ��� "&" �ǤĤʤ�
    Set Stream = CreateObject("ADODB.Stream")
    Stream.Open
    Stream.Charset = "UTF-8"
    Stream.WriteText "dummy=1" ' 1��Ŀ�Υ��`�ϥ��ߥǩ`�������Τǥ��ߩ`����
    Stream.WriteText "&a=�ձ��Z�ƥ���&b=�ƥ���2"
    Stream.Position = 0
    Stream.Type = 1
    postData = Stream.Read
    Stream.Close

    ' POST ���ͤ뤿��α�횥إå��`
    strHeaders = "Content-Type: application/x-www-form-urlencoded" & vbCrLf

    ' ����
    ' 2��Ŀ�� Null �� "_self" �ˉ������ȥ֥饦���Ǳ�ʾ�����
    objIE.Navigate "http://server/test.aspx?ct=0", Null, Null, postData, strHeaders
    Do While objIE.Busy Or objIE.ReadyState <> 4
      WScript.Sleep 100
    Loop
    strRet = objIE.Document.body.innerText
    objIE.Quit

    ' �g�нY��
    WScript.Echo strRet
End Function



'------------------------------------------------
' XML Upload Class
' Example:
'   Dim UploadData
'   Set UploadData = New XMLUpload
'   UploadData.Charset = "utf-8"
'   UploadData.AddForm "content", "Hello world" '�ı�������ƺ�����
'   UploadData.AddFile "file", "test.jpg", "image/jpg", "test.jpg"
'   WScript.Echo UploadData.Upload("http://example.com/takeupload.php")
'   Set UploadData = Nothing
Class XMLUpload
    Private xmlHttp
    Private objTemp
    Private adTypeBinary, adTypeText
    Private strCharset, strBoundary
    
    Private Sub Class_Initialize()
        adTypeBinary = 1
        adTypeText = 2
        Set xmlHttp = CreateObject(COM_HTTP)
        Set objTemp = CreateObject(COM_ADOSTREAM)
        objTemp.Type = adTypeBinary
        objTemp.Open
        strCharset = "utf-8"
        strBoundary = GetBoundary()
    End Sub
    
    Private Sub Class_Terminate()
        objTemp.Close
        Set objTemp = Nothing
        Set xmlHttp = Nothing
    End Sub
    
    'ָ���ַ������ַ���ת�ֽ�����
    Public Function StringToBytes(ByVal strData, ByVal strCharset)
        Dim objFile
        Set objFile = CreateObject(COM_ADOSTREAM)
        objFile.Type = adTypeText
        objFile.Charset = strCharset
        objFile.Open
        objFile.WriteText strData
        objFile.Position = 0
        objFile.Type = adTypeBinary
        If UCase(strCharset) = "UNICODE" Then
            objFile.Position = 2 'delete UNICODE BOM
        ElseIf UCase(strCharset) = "UTF-8" Then
            objFile.Position = 3 'delete UTF-8 BOM
        End If
        StringToBytes = objFile.Read(-1)
        objFile.Close
        Set objFile = Nothing
    End Function
    
    '��ȡ�ļ����ݵ��ֽ�����
    Private Function GetFileBinary(ByVal strPath)
        Dim objFile
        Set objFile = CreateObject(COM_ADOSTREAM)
        objFile.Type = adTypeBinary
        objFile.Open
        objFile.LoadFromFile strPath
        GetFileBinary = objFile.Read(-1)
        objFile.Close
        Set objFile = Nothing
    End Function
    
    '��ȡ�Զ���ı����ݷֽ���
    Private Function GetBoundary()
        Dim ret(12)
        Dim table
        Dim i
        table = "abcdefghijklmnopqrstuvwxzy0123456789"
        Randomize
        For i = 0 To UBound(ret)
            ret(i) = Mid(table, Int(Rnd() * Len(table) + 1), 1)
        Next
        GetBoundary = "---------------------------" & Join(ret, Empty)
    End Function 
    
    '�����ϴ�ʹ�õ��ַ���
    Public Property Let Charset(ByVal strValue)
    strCharset = strValue
    End Property
    
    '����ı�������ƺ�ֵ
    Public Sub AddForm(ByVal strName, ByVal strValue)
        Dim tmp
        tmp = "\r\n--$1\r\nContent-Disposition: form-data; name=""$2""\r\n\r\n$3"
        tmp = Replace(tmp, "\r\n", vbCrLf)
        tmp = Replace(tmp, "$1", strBoundary)
        tmp = Replace(tmp, "$2", strName)
        tmp = Replace(tmp, "$3", strValue)
        objTemp.Write StringToBytes(tmp, strCharset)
    End Sub
    
    '�����ļ��������/�ļ�����/�ļ�MIME����/�ļ�·�����ļ��ֽ�����
    Public Sub AddFile(ByVal strName, ByVal strFileName, ByVal strFileType, ByVal strFilePath)
        Dim tmp
        tmp = "\r\n--$1\r\nContent-Disposition: form-data; name=""$2""; filename=""$3""\r\nContent-Type: $4\r\n\r\n"
        tmp = Replace(tmp, "\r\n", vbCrLf)
        tmp = Replace(tmp, "$1", strBoundary)
        tmp = Replace(tmp, "$2", strName)
        tmp = Replace(tmp, "$3", strFileName)
        tmp = Replace(tmp, "$4", strFileType)
        objTemp.Write StringToBytes(tmp, strCharset)
        objTemp.Write GetFileBinary(strFilePath)
    End Sub
    
    '����multipart/form-data�������
    Private Sub AddEnd()
        Dim tmp
        tmp = "\r\n--$1--\r\n" 
        tmp = Replace(tmp, "\r\n", vbCrLf) 
        tmp = Replace(tmp, "$1", strBoundary)
        objTemp.Write StringToBytes(tmp, strCharset)
        objTemp.Position = 2
    End Sub
    
    '�ϴ���ָ����URL�������ط�����Ӧ��
    Public Function Upload(ByVal strURL)
        Call AddEnd
        xmlHttp.Open "POST", strURL, False
        xmlHttp.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & strBoundary
        'xmlHttp.setRequestHeader "Content-Length", objTemp.size
        xmlHttp.Send objTemp
        Upload = xmlHttp.responseText
    End Function
End Class


'------------------------------------------------
' clsThief Class
Class clsThief
    Private value_      ' ��ȡ��������
    Private src_        ' Ҫ͵��Ŀ��URL��ַ
    Private isGet_      ' �ж��Ƿ��Ѿ�͵��
    Private cookie_ 

    ' ��ֵ��Ҫ͵��Ŀ��URL��ַ/����

    Public Property Let src(Str)
        src_ = Str
    End Property

    '����ֵ��������ȡ��Ӧ���෽���ӹ���������/����

    Public Property Get Value
        Value = value_
    End Property

    Public Property Get Cookie
        Cookie = cookie_
    End Property

    Public Property Get Version
        Version = "�ȷ溣���� Version 2004"
    End Property

    Private Sub class_initialize()
        value_ = ""
        src_ = ""
        isGet_ = False
    End Sub

    Private Sub class_terminate()
    End Sub

    ' ���Ĵ���

    Private Function BytesToBstr(body, Cset)
        Dim objstream
        Set objstream = CreateObject(COM_ADOSTREAM)
        objstream.Type = 1
        objstream.Mode = 3
        objstream.Open
        objstream.Write body
        objstream.Position = 0
        objstream.Type = 2
        objstream.Charset = Cset
        BytesToBstr = objstream.ReadText
        objstream.Close
        Set objstream = Nothing
    End Function

    ' ��ȡĿ��URL��ַ��HTML����/����

    Public Sub steal(encode)
        If src_<>"" Then
            Dim Http
            Set Http = CreateObject(COM_HTTP)
            Http.Open "GET", src_ , false
            Http.send()
            'cookie = Http.getResponseHeader("Set-Cookie")
            If Http.readystate<>4 Then
                Exit Sub
            End If
            value_ = BytesToBSTR(Http.responseBody, encode)
            isGet_ = True
            Set http = Nothing
            If Err.Number<>0 Then Err.Clear
        Else
            response.Write("<script>alert(""��������src���ԣ�"")</script>")
        End If
    End Sub

    ' ɾ��͵��������������Ļ��С��س����Ա��һ���ӹ�/����

    Public Sub noReturn()
        If isGet_ = false Then Call steal()
        value_ = Replace(Replace(value_ , vbCr, ""), vbLf, "")
    End Sub

    ' ��͵���������еĸ����ַ�������ֵ����/����
    ' �����ֱ��Ǿ��ַ���,���ַ���
    Public Sub change(oldStr, Str) 
        If isGet_ = false Then Call steal()
        value_ = Replace(value_ , oldStr, Str)
    End Sub

    ' ��ָ����β�ַ�����͵ȡ�����ݽ��вü�����������β�ַ�����/����
    ' �����ֱ������ַ���,β�ַ���

    Public Sub cut(head, bot)
        If isGet_ = false Then Call steal()
        value_ = Mid(value_ , InStr(value_ , head) + Len(head), InStr(value_ , bot) - InStr(value_ , head) - Len(head))
    End Sub

    ' ��ָ����β�ַ�����͵ȡ�����ݽ��вü���������β�ַ�����/����
    ' �����ֱ������ַ���,β�ַ���

    Public Sub cutX(head, bot)
        If isGet_ = false Then Call steal()
        value_ = Mid(value_ , InStr(value_ , head), InStr(value_ , bot) - InStr(value_ , head) + Len(bot))
    End Sub

    '��ָ����β�ַ���λ��ƫ��ָ���͵ȡ�����ݽ��вü�/����
    '�����ֱ������ַ���,��ƫ��ֵ,β�ַ���,βƫ��ֵ,��ƫ���ø�ֵ,ƫ��ָ�뵥λΪ�ַ���

    Public Sub cutBy(head, headCusor, bot, botCusor)
        If isGet_ = false Then Call steal()
        value_ = Mid(value_ , InStr(value_ , head) + Len(head) + headCusor, InStr(value_ , bot) -1 + botCusor - InStr(value_ , head) - Len(head) - headcusor)
    End Sub

    '��ָ����β�ַ�����͵ȡ����������ֵ�����滻����������β�ַ�����/����
    '�����ֱ������ַ���,β�ַ���,��ֵ,��ֵλ����Ϊ����

    Public Sub filt(head, bot, Str)
        If isGet_ = false Then Call steal()
        value_ = Replace(value_, Mid(value_ , InStr(value_ , head) + Len(head), InStr(value_ , bot) -1), Str)
    End Sub

    '��ָ����β�ַ�����͵ȡ����������ֵ�����滻��������β�ַ�����/����
    '�����ֱ������ַ���,β�ַ���,��ֵ,��ֵΪ����Ϊ����

    Public Sub filtX(head, bot, Str)
        If isGet_ = false Then Call steal()
        value_ = Replace(value_, Mid(value_ , InStr(value_ , head), InStr(value_ , bot) + Len(bot) -1), Str)
    End Sub

    '��ָ����β�ַ���λ��ƫ��ָ���͵ȡ��������ֵ�����滻/����
    '�����ֱ������ַ���,��ƫ��ֵ,β�ַ���,βƫ��ֵ,��ֵ,��ƫ���ø�ֵ,ƫ��ָ�뵥λΪ�ַ���,��ֵΪ����Ϊ����

    Public Sub filtBy(head, headCusor, bot, botCusor, Str)

        If isGet_ = false Then Call steal()
        value_ = Replace(value_ , Mid(value_ , InStr(value_ , head) + Len(head) + headCusor, InStr(value_ , bot) -1 + botCusor - InStr(value_ , head) - Len(head) - headcusor), Str)
    End Sub

    '��͵ȡ�������еľ���URL��ַ��Ϊ������Ե�ַ

    Public Sub local()
        Dim tempReg
        Set tempReg = New RegExp
        tempReg.IgnoreCase = true
        tempReg.Global = true
        tempReg.Pattern = "^(http|https|ftp):(\/\/|////)(\w+.)+(com|net|org|cc|tv|cn|biz|com.cn|net.cn|sh.cn)\/"
        value_ = tempReg.Replace(value_ , "")
        Set tempReg = Nothing
    End Sub

    '��͵���������еķ���������ʽ���ַ�������ֵ�����滻/����
    '���������Զ����������ʽ,��ֵ

    Public Sub replaceByReg(patrn, Str)
        If isGet_ = false Then Call steal()
        Dim tempReg
        Set tempReg = New RegExp
        tempReg.IgnoreCase = true
        tempReg.Global = true
        tempReg = patrn
        value_ = tempReg.Replace(value_ , Str)
        Set tempReg = Nothing
    End Sub

    'Ӧ��������ʽ�Է������������ݽ��зֿ�ɼ������,��������Ϊ��<!--lkstar-->���ϵĴ��ı�/����
    'ͨ������value�õ������ݺ��������split(value,"<!--lkstar-->")�õ�����Ҫ������
    '���������Զ����������ʽ

    Public Sub pickByReg(patrn)
        If isGet_ = false Then Call steal()
        Dim tempReg, match, matches, content
        Set tempReg = New RegExp
        tempReg.IgnoreCase = true
        tempReg.Global = true
        tempReg = patrn
        Set matches = tempReg.Execute(value_)
        For Each match in matches
            content = content&match.Value&"<!--lkstar-->"
        Next
        value_ = content
        Set matches = Nothing
        Set tempReg = Nothing
    End Sub

    '���Ŵ�ģʽ���������ͷ�֮ǰӦ�ô˷���������ʱ�鿴��ػ������HTML�����ҳ����ʾЧ��/����

    Public Sub debug()
        Dim tempstr
        tempstr = "<SCRIPT>function runEx(){var winEx2 = window.open("""", ""winEx2"", ""width=500,height=300,status=yes,menubar=no,scrollbars=yes,resizable=yes""); winEx2.document.open(""text/html"", ""replace""); winEx2.document.write(unescape(event.srcElement.parentElement.children[0].value)); winEx2.document.close(); }function saveFile(){var win=window.open('','','top=10000,left=10000');win.document.write(document.all.asdf.innerText);win.document.execCommand('SaveAs','','javascript.htm');win.close();}</SCRIPT><center><TEXTAREA id=asdf name=textfield rows=32  wrap=VIRTUAL cols=""120"">"&value_&"</TEXTAREA><BR><BR><INPUT name=Button onclick=runEx() type=button value=""�鿴Ч��"">&nbsp;&nbsp;<INPUT name=Button onclick=asdf.select() type=button value=""ȫѡ"">&nbsp;&nbsp;<INPUT name=Button onclick=""asdf.value=''"" type=button value=""���"">&nbsp;&nbsp;<INPUT onclick=saveFile(); type=button value=""�������""></center>"
        'response.Write(tempstr)
        document.Write tempstr
    End Sub

End Class