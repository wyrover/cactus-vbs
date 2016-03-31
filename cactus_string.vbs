'------------------------------------------------
' StrLength
Function StrLength(Str)
    On Error Resume Next
    Dim WINNT_CHINESE
    WINNT_CHINESE = (Len("中国") = 2)
    If WINNT_CHINESE Then
        Dim l, t, c
        Dim i
        l = Len(Str)
        t = l
        For i = 1 To l
            c = Asc(Mid(Str, i, 1))
            If c < 0 Then c = c + 65536
            If c > 255 Then
                t = t + 1
            End If
        Next
        StrLength = t
    Else
    StrLength = Len(Str)
    End If
    If Err.Number <> 0 Then Err.Clear
End Function

'------------------------------------------------
' GetRandomString
Function GetRandomString(digits)

    Dim char_array(80)
    '初始化数字
    For i = 0 To 9
        char_array(i) = CStr(i)
    Next
    '初始化大写字母
    For i = 10 To 35
        char_array(i) = Chr(i + 55)
    Next
    '初始化小写字母
    For i = 36 To 61
        char_array(i) = Chr(i + 61)
    Next
    Randomize   '初始化随机数生成器。
    Do While Len(output) < digits
        num = char_array(Int((62 - 0 + 1) * Rnd + 0))
        output = output + num
    Loop

    gen_key = output
End Function


'------------------------------------------------
' Unicode字符串转utf8字符串
Function UnicodeToUtf8(str)
    Dim i, c, length
    out = ""
    length = Len(str)
    For i = 1 To length
        c = CLng("&H" & Hex(AscW(Mid(str,i,1))))
        If (c >= &H0001) And (c <= &H007F) Then
            out = out & ChrB(c)
        ElseIf c > &H07FF Then
            out = out & ChrB(&HE0 Or (c\(2^12) And &H0F))
            out = out & ChrB(&H80 Or (c\(2^ 6) And &H3F))
            out = out & ChrB(&H80 Or (c\(2^ 0) And &H3F))
        Else
            out = out & ChrB(&HC0 Or (c\(2^ 6) And &H1F))
            out = out & ChrB(&H80 Or (c\(2^ 0) And &H3F))
        End If
    Next
    UnicodeToUtf8 = out
End Function

'------------------------------------------------
' utf8字符串转Unicode字符串
Function Utf8ToUnicode(str)
    Dim i, c, c2, c3, out, length
    out = ""
    i = 1
    length = LenB(str)
    Do While i <= length
        c = AscB(MidB(str,i,1))
        i = i + 1
        Select Case (c \ 2 ^ 4)
            Case 0,1,2,3,4,5,6,7
            out = out & ChrW(c)
            Case 12,13
            c2 = AscB(MidB(str,i,1))
            i = i + 1
            out = out & ChrW(((c And &H1F) * 2 ^ 6) Or (c2 And &H3F))
            Case 14
            c2 = AscB(MidB(str,i,1))
            i = i + 1
            c3 = AscB(MidB(str,i,1))
            i = i + 1
            out = out & ChrW(((c And &H0F) * 2 ^ 12) Or _
            ((c2 And &H3F) * 2 ^ 6) Or _
            ((c3 And &H3F) * 2 ^ 0))
        End Select
    Loop
    Utf8ToUnicode = out
End Function


Function read(path)
    '将Byte()数组转成String字符串
    Dim ado, a(), i, n
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 1 : ado.Open
    ado.LoadFromFile path
    n = ado.Size - 1
    ReDim a(n)
    For i = 0 To n
        a(i) = ChrW(AscB(ado.Read(1)))
    Next
    read = Join(a, "")
End Function

'------------------------------------------------
' 验证字符串是否为utf8编码
' Author: Demon
' Date: 2011/11/10
' Website: http://demon.tw
' Example:
'    s = read("utf-8.txt") '读取文件
'    WScript.Echo is_valid_utf8(s) '判断是否UTF-8
Function is_valid_utf8(ByRef input) 'ByRef以提高效率
    Dim s, re
    Set re = New Regexp
    s = "[\xC0-\xDF]([^\x80-\xBF]|$)"
    s = s & "|[\xE0-\xEF].{0,1}([^\x80-\xBF]|$)"
    s = s & "|[\xF0-\xF7].{0,2}([^\x80-\xBF]|$)"
    s = s & "|[\xF8-\xFB].{0,3}([^\x80-\xBF]|$)"
    s = s & "|[\xFC-\xFD].{0,4}([^\x80-\xBF]|$)"
    s = s & "|[\xFE-\xFE].{0,5}([^\x80-\xBF]|$)"
    s = s & "|[\x00-\x7F][\x80-\xBF]"
    s = s & "|[\xC0-\xDF].[\x80-\xBF]"
    s = s & "|[\xE0-\xEF]..[\x80-\xBF]"
    s = s & "|[\xF0-\xF7]...[\x80-\xBF]"
    s = s & "|[\xF8-\xFB]....[\x80-\xBF]"
    s = s & "|[\xFC-\xFD].....[\x80-\xBF]"
    s = s & "|[\xFE-\xFE]......[\x80-\xBF]"
    s = s & "|^[\x80-\xBF]"
    re.Pattern = s
    is_valid_utf8 = (Not re.Test(input))
End Function

'------------------------------------------------
' 验证字符串是否为utf8编码
Public Function isUTF8(astr)
    Dim c0, c1, c2, c3
    Dim n
    
    isUTF8 = True
    n = 1
    Do While n <= Len(astr)
        c0 = Asc(Mid(astr, n, 1))
        If n <= Len(astr) - 1 Then
            c1 = Asc(Mid(astr, n + 1, 1))
        Else
            c1 = 0
        End If
        If n <= Len(astr) - 2 Then
            c2 = Asc(Mid(astr, n + 2, 1))
        Else
            c2 = 0
        End If
        If n <= Len(astr) - 3 Then
            c3 = Asc(Mid(astr, n + 3, 1))
        Else
            c3 = 0
        End If
        
        If (c0 And 240) = 240 Then
            If (c1 And 128) = 128 And (c2 And 128) = 128 And (c3 And 128) = 128 Then
                n = n + 4
            Else
                isUTF8 = False
                Exit Function
            End If
        ElseIf (c0 And 224) = 224 Then
            If (c1 And 128) = 128 And (c2 And 128) = 128 Then
                n = n + 3
            Else
                isUTF8 = False
                Exit Function
            End If
        ElseIf (c0 And 192) = 192 Then
            If (c1 And 128) = 128 Then
                n = n + 2
            Else
                isUTF8 = False
                Exit Function
            End If
        ElseIf (c0 And 128) = 0 Then
            n = n + 1
        Else
            isUTF8 = False
            Exit Function
        End If
    Loop
End Function

'------------------------------------------------
' utf8 编码
Function Encode_UTF8(astr)
    Dim c
    Dim n
    Dim utftext
    
    utftext = ""
    n = 1
    Do While n <= Len(astr)
        c = AscW(Mid(astr, n, 1))
        If c < 128 Then
            utftext = utftext + Chr(c)
        ElseIf ((c >= 128) And (c < 2048)) Then
            utftext = utftext + Chr(((c \ 64) Or 192))
            utftext = utftext + Chr(((c And 63) Or 128))
        ElseIf ((c >= 2048) And (c < 65536)) Then
            utftext = utftext + Chr(((c \ 4096) Or 224))
            utftext = utftext + Chr((((c \ 64) And 63) Or 128))
            utftext = utftext + Chr(((c And 63) Or 128))
        Else ' c >= 65536
            utftext = utftext + Chr(((c \ 262144) Or 240))
            utftext = utftext + Chr(((((c \ 4096) And 63)) Or 128))
            utftext = utftext + Chr((((c \ 64) And 63) Or 128))
            utftext = utftext + Chr(((c And 63) Or 128))
        End If
        n = n + 1
    Loop
    Encode_UTF8 = utftext
End Function

'------------------------------------------------
' utf8 解码
Function Decode_UTF8(astr)
    Dim c0, c1, c2, c3
    Dim n
    Dim unitext
    
    If isUTF8(astr) = False Then
        Decode_UTF8 = astr
        Exit Function
    End If
    
    unitext = ""
    n = 1
    Do While n <= Len(astr)
        c0 = Asc(Mid(astr, n, 1))
        If n <= Len(astr) - 1 Then
            c1 = Asc(Mid(astr, n + 1, 1))
        Else
            c1 = 0
        End If
        If n <= Len(astr) - 2 Then
            c2 = Asc(Mid(astr, n + 2, 1))
        Else
            c2 = 0
        End If
        If n <= Len(astr) - 3 Then
            c3 = Asc(Mid(astr, n + 3, 1))
        Else
            c3 = 0
        End If
        
        If (c0 And 240) = 240 And (c1 And 128) = 128 And (c2 And 128) = 128 And (c3 And 128) = 128 Then
            unitext = unitext + ChrW((c0 - 240) * 65536 + (c1 - 128) * 4096) + (c2 - 128) * 64 + (c3 - 128)
            n = n + 4
        ElseIf (c0 And 224) = 224 And (c1 And 128) = 128 And (c2 And 128) = 128 Then
            unitext = unitext + ChrW((c0 - 224) * 4096 + (c1 - 128) * 64 + (c2 - 128))
            n = n + 3
        ElseIf (c0 And 192) = 192 And (c1 And 128) = 128 Then
            unitext = unitext + ChrW((c0 - 192) * 64 + (c1 - 128))
            n = n + 2
        ElseIf (c0 And 128) = 128 Then
            unitext = unitext + ChrW(c0 And 127)
            n = n + 1
        Else ' c0 < 128
            unitext = unitext + ChrW(c0)
            n = n + 1
        End If
    Loop
    
    Decode_UTF8 = unitext
End Function

'------------------------------------------------
' 字符串添加引号
Function Quotes(strQuotes)
    Quotes = chr(34) & strQuotes & chr(34)
End Function

'------------------------------------------------
' 获取GUID值
Function NewGUID
    Set TypeLib = CreateObject(COM_TYPELIB) 
    NewGUID = Left(TypeLib.Guid, 38)
    Set TypeLib = Nothing
End Function 

'------------------------------------------------
' 获取GUID值, 不带{}
Function NewGUID2  
    Set TypeLib = CreateObject(COM_TYPELIB)
    NewGUID2 = Mid(TypeLib.Guid, 2, 36)
    Set TypeLib = Nothing
End Function 

'------------------------------------------------
' 区间随机数
' @lowerbound       下限
' @upperbound       上限
Function RandomNum(lowerbound, upperbound)
    Randomize
    RandomNum = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function

'------------------------------------------------
' 创建随机密码
Function CreatePassword(numchar)
    Dim avail, parola, f, i
    
    avail = "abcdefghijklmnopqrstuvwxyz1234567890"
    Randomize
    parola = ""
    for f = 1 to numchar
        i = (CInt(len(avail) * Rnd + 1) mod len(avail)) + 1
        parola = parola & mid(avail, i, 1)
    next
    CreatePassword = parola
End Function


'------------------------------------------------
' 字符串转数字
' @strS         字符串
' @return       Integer (>=0)
Function CID(strS)
    Dim intI
    intI = 0
    If IsNull(strS) Or strS = "" Then
        intI = 0
    Else
        If Not IsNumeric(strS) Then
            intI = 0
        Else
            Dim intk
            On Error Resume Next
            intk = Abs(Clng(strS))
            If Err.Number = 6 Then intk = 0  '数据溢出
            Err.Clear
            intI = intk
        End If
    End If
    CID = intI
End Function

'------------------------------------------------
' 判断用户名是否合法
' @username        用户名
Function IsTrueName(username)
    Dim Hname, I
    IsTrueName = False
    Hname = Array("=", "%", chr(32), "?", "&", ";", ",", "'", ",", chr(34), chr(9), "", "$", "|")
    For I = 0 To Ubound(Hname)
        If InStr(username, Hname(I)) > 0 Then
            Exit Function
        End If
    Next
    IsTrueName = True 
End Function


'------------------------------------------------
' BStr2UStr
Function BStr2UStr(BStr)
    'Byte string to Unicode string conversion
    Dim lngLoop
    BStr2UStr = ""
    For lngLoop = 1 to LenB(BStr)
        BStr2UStr = BStr2UStr & Chr(AscB(MidB(BStr,lngLoop,1))) 
    Next
End Function

'------------------------------------------------
' UStr2Bstr
Function UStr2Bstr(UStr)
    'Unicode string to Byte string conversion
    Dim lngLoop
    Dim strChar
    UStr2Bstr = ""
    For lngLoop = 1 to Len(UStr)
        strChar = Mid(UStr, lngLoop, 1)
        UStr2Bstr = UStr2Bstr & ChrB(AscB(strChar))
    Next
End Function

'------------------------------------------------
' Base64encode
Function Base64Encode(str)  
    Dim CAPIUtil
    Set CAPIUtil = CreateObject(COM_CAPICOM_UTIL)
    Base64encode = CAPIUtil.Base64Encode(str)
    Set CAPIUtil = Nothing
End Function

'------------------------------------------------
' Base64decode
Function Base64Decode(str) 
    Dim CAPIUtil
    Set CAPIUtil = CreateObject(COM_CAPICOM_UTIL)
    Base64Decode = CAPIUtil.Base64Decode(str)
    Set CAPIUtil = Nothing
End Function 

'------------------------------------------------
' ToBase64
Function ToBase64(Src)
    Dim BASE64:BASE64="ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
    "abcdefghijklmnopqrstuvwxyz" & _
    "0123456789+/"
    Dim k
    Dim Bytes
    Dim Code
    Dim Dst
    
    ReDim Bytes(LenB(Src))
    For k=1 To Len(Src)
        Code=AscW(Mid(Src,k,1))
        If Code<0 Then Code=Code+256*256
        Bytes(k*2-1)=Code \ 256
        Bytes(k*2)=Code Mod 256
    Next
    For k=1 To UBound(Bytes) Step 3
        Dst=Dst & Mid(BASE64,(Bytes(k) \ 4)+1,1)
        If k+1>UBound(Bytes) Then
            Dst=Dst & Mid(BASE64,(Bytes(k)*16 Mod 64)+1,1)
        Else
            Dst=Dst & Mid(BASE64,(Bytes(k)*16 Mod 64)+(Bytes(k+1) \ 16)+1,1)
            If k+2>UBound(Bytes) Then
                Dst=Dst & Mid(BASE64,(Bytes(k+1)*4 Mod 64)+1,1)
            Else
                Dst=Dst & Mid(BASE64,(Bytes(k+1)*4 Mod 64)+(Bytes(k+2) \ 64)+1,1)
                Dst=Dst & Mid(BASE64,(Bytes(k+2) Mod 64)+1,1)
            End If
        End If
    Next
    ToBase64=Dst
End Function

'------------------------------------------------
' MD5
Function MD5(str) 
    Dim CAPIHASH
    Set CAPIHASH = CreateObject(COM_CAPICOM_HASH)
    CAPIHASH.Algorithm = CAPICOM_HASH_ALGORITHM_MD5
    CAPIHASH.Hash UStr2Bstr(str)
    MD5 = CAPIHASH.Value
    Set CAPIHASH = Nothing
End Function 

'------------------------------------------------
' MD5_File
Function MD5_File(filename, raw_output)
    Dim HashedData, Utility, Stream
    Set HashedData = CreateObject(COM_CAPICOM_HASH)
    Set Utility = CreateObject(COM_CAPICOM_UTIL)
    Set Stream = CreateObject(COM_ADOSTREAM)
    HashedData.Algorithm = CAPICOM_HASH_ALGORITHM_MD5
    Stream.Type = 1
    Stream.Open
    Stream.LoadFromFile filename
    Do Until Stream.EOS
        HashedData.Hash Stream.Read(1024)
    Loop
    If raw_output Then
        MD5_File = Utility.HexToBinary(HashedData.Value)
    Else
        MD5_File = HashedData.Value
    End If
End Function

'------------------------------------------------
' SHA1
Function SHA1(str) 
    Dim CAPIHASH
    Set CAPIHASH = CreateObject(COM_CAPICOM_HASH)
    CAPIHASH.Algorithm = CAPICOM_HASH_ALGORITHM_SHA1
    CAPIHASH.Hash UStr2Bstr(str)
    SHA1 = CAPIHASH.Value
    Set CAPIHASH = Nothing
End Function 

'------------------------------------------------
' SHA1__File
Function SHA1__File(filename, raw_output)
    Dim HashedData, Utility, Stream
    Set HashedData = CreateObject(COM_CAPICOM_HASH)
    Set Utility = CreateObject(COM_CAPICOM_UTIL)
    Set Stream = CreateObject(COM_ADOSTREAM)
    HashedData.Algorithm = 0
    Stream.Type = 1
    Stream.Open
    Stream.LoadFromFile filename
    Do Until Stream.EOS
        HashedData.Hash Stream.Read(1024)
    Loop
    If raw_output Then
        sha1_file = Utility.HexToBinary(HashedData.Value)
    Else
        sha1_file = HashedData.Value
    End If
End Function


'------------------------------------------------
' SplitURL
Function SplitURL(url, ByRef protocol, ByRef hostname, ByRef port, ByRef pathname, ByRef hash, ByRef search)
    Set Document = CreateObject(COM_HTML)
    Document.write "<html><body><a id=a1 /></body></html>"
    Set Location = Document.body.all.a1

    Location.href = url
    protocol = Location.protocol
    hostname = Location.hostname
    port = Location.port
    pathname = Location.pathname
    hash = Location.hash
    search = Location.search
End Function 



'------------------------------------------------
' URLEncoding
Function URLEncoding(vstrIn) 
    Dim strReturn, ThisChr, innerCode, Hight8, Low8
    strReturn = "" 
    For i = 1 To Len(vstrIn) 
        ThisChr = Mid(vStrIn,i,1) 
        If Abs(Asc(ThisChr)) < &HFF Then 
            strReturn = strReturn & ThisChr 
        Else 
            innerCode = Asc(ThisChr) 
            If innerCode < 0 Then 
                innerCode = innerCode + &H10000 
            End If 
            Hight8 = (innerCode And &HFF00) OR &HFF 
            Low8 = innerCode And &HFF 
            strReturn = strReturn & "%" & Hex(Hight8) &  "%" & Hex(Low8) 
        End If 
    Next 
    URLEncoding = strReturn 
End Function 


'------------------------------------------------
' 过滤html标签
Function FilterHtml(str)
    Dim re    
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.MultiLine = True
    re.Pattern = "<.+?>"
    FilterHtml = re.Replace(str, "")
    Set re = Nothing
End Function

'------------------------------------------------
' 过滤html标签
Function StripHTML(ByRef sHTML)
    Dim re 
    Set re = New RegExp
    re.Pattern = "<[^>]*>" 
    re.IgnoreCase = True  
    re.Global = True    
    StripHTML = re.Replace(sHTML, " ")   
    Set re = Nothing
End Function

'------------------------------------------------
' 过滤指定html标签
Function DecodeFilter(html, filter)
    html = LCase(html)
    filter = split(filter, ",")
    For Each i In filter
        Select Case i
            Case "SCRIPT"   ' 去除所有客户端脚本javascipt,vbscript,jscript,js,vbs,event,...
            html = exeRE("(javascript|jscript|vbscript|vbs):", "#", html)
            html = exeRE("</?script[^>]*>", "", html)
            html = exeRE("on(mouse|exit|error|click|key)", "", html)
            Case "TABLE":   ' 去除表格<table><tr><td><th>
            html = exeRE("</?table[^>]*>", "", html)
            html = exeRE("</?tr[^>]*>", "", html)
            html = exeRE("</?th[^>]*>", "", html)
            html = exeRE("</?td[^>]*>", "", html)
            html = exeRE("</?tbody[^>]*>", "", html)
            Case "CLASS"    ' 去除样式类class=""
            html = exeRE("(<[^>]+) class=[^ |^>]*([^>]*>)", "$1 $2", html) 
            Case "STYLE"    ' 去除样式style=""
            html = exeRE("(<[^>]+) style=""[^""]*""([^>]*>)", "$1 $2", html)
            html = exeRE("(<[^>]+) style='[^']*'([^>]*>)", "$1 $2", html)
            Case "IMG"      ' 去除样式style=""
            html = exeRE("</?img[^>]*>", "", html)
            Case "XML"      ' 去除XML<?xml>
            html = exeRE("<\\?xml[^>]*>", "", html)
            Case "NAMESPACE"    ' 去除命名空间<o:p></o:p>
            html = exeRE("<\/?[a-z]+:[^>]*>", "", html)
            Case "FONT"     ' 去除字体<font></font>
            html = exeRE("</?font[^>]*>", "", html)
            Case "MARQUEE"  ' 去除字幕<marquee></marquee>
            html = exeRE("</?marquee[^>]*>", "", html)
            Case "OBJECT"   ' 去除对象<object><param><embed></object>
            html = exeRE("</?object[^>]*>", "", html)
            html = exeRE("</?param[^>]*>", "", html)
            html = exeRE("</?embed[^>]*>", "", html)
            Case "DIV"      ' 去除对象<object><param><embed></object>
            html = exeRE("</?div([^>])*>", "$1", html)
        End Select
    Next
    'html = Replace(html,"<table","<")
    'html = Replace(html,"<tr","<")
    'html = Replace(html,"<td","<")
    DecodeFilter = html
End Function

'------------------------------------------------
' 字符串转Unicode编码
Function Chinese2Unicode(str) 
    Dim i 
    Dim Str_one 
    Dim Str_unicode 
    For i = 1 To Len(str) 
        Str_one = Mid(str, i, 1) 
        Str_unicode = Str_unicode & chr(38) 
        Str_unicode = Str_unicode & chr(35) 
        Str_unicode = Str_unicode & chr(120) 
        Str_unicode = Str_unicode & Hex(ascw(Str_one)) 
        Str_unicode = Str_unicode & chr(59) 
    Next 
    
    str = Str_unicode
End Function

'------------------------------------------------
' 字符串转Unicode
Function TextToUnicode(strText)
    ' Function to convert a text string into a string of unicode
    ' hexadecimal bytes. The string is first enclosed by quote characters.
    Dim strChar, k

    strText = """" & strText & """"

    TextToUnicode = ""
    For k = 1 To Len(strText)
        strChar = Mid(strText, k, 1)
        TextToUnicode = TextToUnicode & Right("00" & Hex(Asc(strChar)), 2)
        ' Add a "00" byte.
        TextToUnicode = TextToUnicode & "00"
    Next
End Function

'------------------------------------------------
' RegexMatch
Function RegexMatch(Str, Pattern, IgnoreCase)
    On Error Resume Next
    '正規表現
    Dim regex 
    Set regex = New RegExp
    
    '引数チェック
    If IsNull(Pattern) Then
        RegexMatch = False
        Exit Function
    ElseIf IsNull(Value) Then
        RegexMatch = False
        Exit Function
    ElseIf IsNull(IgnoreCase) Then
        RegexMatch = False
        Exit Function
    End If
    
    '正規表現オブジェクトにパターンをセット
    regex.Pattern = Pattern
    regex.IgnoreCase = IgnoreCase '大文字小文字の区別フラグ
    '実行
    'regex.Testメソッドで正規表現にマッチしたらTrueを返す
    If (regex.Test(Str)) Then
        RegexMatch = True
    Else
        RegexMatch = False
    End If
    
    Exit Function    

    If Err.Number = 5017 Then
        ret = MsgBox("不正な正規表現です" & vbCrLf & "メタ文字がエスケープされていません。", vbOKOnly, "Error:" & Err.Number)
        'ex. \
        'ex. aa)
    ElseIf Err.Number = 5018 Then
        ret = MsgBox("不正な正規表現です" & vbCrLf & "量指定子(*?+{})が不正です", vbOKOnly, "Error:" & Err.Number)
        'ex. ?
    ElseIf Err.Number = 5019 Then
        ret = MsgBox("不正な正規表現です" & vbCrLf & "[]が正しく入力されていません", vbOKOnly, "Error:" & Err.Number)
        'ex. [a-z
    ElseIf Err.Number = 5020 Then
        ret = MsgBox("不正な正規表現です" & vbCrLf & "()が正しく入力されていません", vbOKOnly, "Error:" & Err.Number)
        'ex. (a
    ElseIf Err.Number = 5021 Then
        ret = MsgBox("不正な正規表現です" & vbCrLf & "[]内の文字クラスが不正です", vbOKOnly, "Error:" & Err.Number)
        'ex. [a-Z]
    ElseIf Err.Number = 5022 Then
        ret = MsgBox("不正な正規表現です", vbOKOnly, "Error:" & Err.Number)
        'ex.
    Else
        MsgBox (Err.Number & Err.Description)
    End If
    RegexMatch = False
    Exit Function
End Function

'------------------------------------------------
' 正则表达式替换
' @content  文本
' @pattern  正则表达式模式
' @str      替换字符串
Function ReplaceText(content, pattern, str)
    Dim re
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = pattern
    ReplaceText = re.Replace(content, str)
    Set re = Nothing    
End Function


'------------------------------------------------
' HTMLEncode
Function HTMLEncode(text)
    If text = "" or IsNull(text) Then 
        Exit Function
    Else
        If Instr(text, "'") > 0 Then 
            text = replace(text, "'", "&#39;")
        End If
        text = replace(text, ">", "&gt;")
        text = replace(text, "<", "&lt;")
        text = Replace(text, CHR(32), "&nbsp;")
        text = Replace(text, CHR(9), "&nbsp;")
        text = Replace(text, CHR(34), "&quot;")
        text = Replace(text, CHR(13),"")
        text = Replace(text, CHR(10) & CHR(10), "</P><P>")
        text = Replace(text, CHR(10), "<BR>")
        text = Replace(text, CHR(39), "&#39;")
        text = Replace(text, CHR(0), "")
        text = ChkBadWords(text)
        HTMLEncode = text
    End If
End Function


'------------------------------------------------
' HTMLDecode
Public Function HTMLDecode(text)
    If text = "" or IsNull(text) Then 
        Exit Function
    Else
        If Instr(text, "'")>0 Then 
            text = replace(text, "'", "&#39;")
        End If
        text = replace(text, "&gt;", ">")
        text = replace(text, "&lt;", "<")
        text = Replace(text, "&nbsp;", CHR(32))
        text = Replace(text, "&nbsp;", CHR(9))
        text = Replace(text, "&quot;", CHR(34))
        text = Replace(text, "", CHR(13))
        text = Replace(text, "</P><P>", CHR(10) & CHR(10))
        text = Replace(text, "<BR>", CHR(10))
        text = Replace(text, "", CHR(0))
        text = Replace(text, "&#39;", CHR(39))
        text = ChkBadWords(text)
        HTMLDecode = text
    End If
End Function


'------------------------------------------------
' 取字段数据每个汉字的拼音首字母
Function getpychar(char)
    tmp = 65536 + Asc(char)
    If (tmp>= 45217 And tmp<= 45252) Then
        getpychar = "A"
    ElseIf (tmp>= 45253 And tmp<= 45760) Then
        getpychar = "B"
    ElseIf (tmp>= 47761 And tmp<= 46317) Then
        getpychar = "C"
    ElseIf (tmp>= 46318 And tmp<= 46825) Then
        getpychar = "D"
    ElseIf (tmp>= 46826 And tmp<= 47009) Then
        getpychar = "E"
    ElseIf (tmp>= 47010 And tmp<= 47296) Then
        getpychar = "F"
    ElseIf (tmp>= 47297 And tmp<= 47613) Then
        getpychar = "G"
    ElseIf (tmp>= 47614 And tmp<= 48118) Then
        getpychar = "H"
    ElseIf (tmp>= 48119 And tmp<= 49061) Then
        getpychar = "J"
    ElseIf (tmp>= 49062 And tmp<= 49323) Then
        getpychar = "K"
    ElseIf (tmp>= 49324 And tmp<= 49895) Then
        getpychar = "L"
    ElseIf (tmp>= 49896 And tmp<= 50370) Then
        getpychar = "M"
    ElseIf (tmp>= 50371 And tmp<= 50613) Then
        getpychar = "N"
    ElseIf (tmp>= 50614 And tmp<= 50621) Then
        getpychar = "O"
    ElseIf (tmp>= 50622 And tmp<= 50905) Then
        getpychar = "P"
    ElseIf (tmp>= 50906 And tmp<= 51386) Then
        getpychar = "Q"
    ElseIf (tmp>= 51387 And tmp<= 51445) Then
        getpychar = "R"
    ElseIf (tmp>= 51446 And tmp<= 52217) Then
        getpychar = "S"
    ElseIf (tmp>= 52218 And tmp<= 52697) Then
        getpychar = "T"
    ElseIf (tmp>= 52698 And tmp<= 52979) Then
        getpychar = "W"
    ElseIf (tmp>= 52980 And tmp<= 53640) Then
        getpychar = "X"
    ElseIf (tmp>= 53689 And tmp<= 54480) Then
        getpychar = "Y"
    ElseIf (tmp>= 54481 And tmp<= 62289) Then
        getpychar = "Z"
    Else '如果不是中文，则不处理
        getpychar = char
    End If
End Function

'------------------------------------------------
' 获取拼音 
Function GetPinYin(Str)
    Dim I
    For I = 1 To Len(Str)
        GetPinYin = GetPinYin & getpychar(Mid(Str, i, 1))
    Next
End Function

'------------------------------------------------
' 验证Email 
Function CheckEmail(Str)
    CheckEmail = False
    Dim re, match
    Set re = New RegExp
    re.Pattern = "^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$"
    re.IgnoreCase = True
    Set match = re.Execute(Str)
    If match.Count Then CheckEmail = True
    Set re = Nothing
End Function

'------------------------------------------------
' 验证用户名
Function CheckUserName(str)
    Dim re
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.MultiLine = True
    re.Pattern = "^[a-z0-9_]{2,20}$"
    CheckUserName = re.Test(str)
    Set re = Nothing
End Function

'------------------------------------------------
' 获取计算机名
Function GetComputerName()
    Dim shell, regpath
    Set shell = CreateObject(COM_SHELL)
    regpath = "HKLM\System\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName"
    GetComputerName = shell.RegRead(regpath)
End Function

'------------------------------------------------
' EncodeTextAndBase64
Function EncodeTextAndBase64(text, charset)
    Set CAPIUtil = CreateObject(COM_CAPICOM_UTIL)
    Set Stream = CreateObject(COM_ADOSTREAM)
    Set StreamBin = CreateObject(COM_ADOSTREAM)

    '******************************
    ' Base64 エンコード
    '******************************
    Stream.Open
    'Stream.Charset = "shift_jis"
    ' shift_jis で入力文字を書き込む
    'Stream.WriteText "日本語表示OK"
    Stream.Charset = charset    
    Stream.WriteText text
    Stream.Position = 0

    ' バイナリで開く
    StreamBin.Open
    StreamBin.Type = 1

    ' テキストをバイナリに変換
    Stream.CopyTo StreamBin
    Stream.Close

    ' 読み込みの為にデータポインタを先頭にセット
    StreamBin.Position = 0

    ' 変換
    strBinaryString = CAPIUtil.ByteArrayToBinaryString( StreamBin.Read )
    strBase64 = CAPIUtil.Base64Encode( strBinaryString )
    ' 長い文字列は仕様として、(\r\n を含めて 76) 改行されます
    EncodeTextAndBase64 = Replace(strBase64,vbCrLf,"")
End Function

'------------------------------------------------
' EncodeTextAndHash
Function EncodeTextAndHash(text, charset, hash)
    Set CAPIUtil = CreateObject(COM_CAPICOM_UTIL)
    Set HashedData = CreateObject(COM_CAPICOM_HASH)
    Set Stream = CreateObject(COM_ADOSTREAM)
    Set StreamBin = CreateObject(COM_ADOSTREAM)
 
    Stream.Open
    Stream.Charset = charset
    ' shift_jis で入力文字を書き込む
    Stream.WriteText text
    Stream.Position = 0

    ' バイナリで開く
    StreamBin.Open
    StreamBin.Type = 1

    ' テキストをバイナリに変換
    Stream.CopyTo StreamBin
    Stream.Close

    ' 読み込みの為にデータポインタを先頭にセット
    StreamBin.Position = 0

    strBinaryString = CAPIUtil.ByteArrayToBinaryString( StreamBin.Read )

    '***********************************************************
    ' SHA1 と MD5 と SHA256
    '***********************************************************
    ' SHA1

    Select Case hash
        Case "md5"
            HashedData.Algorithm = 3
        Case "sha1"
            HashedData.Algorithm = 0
        Case "sha256"
            HashedData.Algorithm = 4
    End Select    

    HashedData.Hash(strBinaryString)
    EncodeTextAndHash = LCase(HashedData.Value)
End Function

'------------------------------------------------
' TextToBin
'    msgbox TextToBin("テスト","UTF-8") ←文字化ける
'    msgbox TextToBin("テスト","UTF-16") ←文字化けない
Function TextToBin(TextData, CharSet) 
    Const adTypeBinary = 1 
    Const adTypeText = 2 
    Dim objStream 
    Set objStream = CreateObject(COM_ADOSTREAM) 
    objStream.Type = adTypeText 
    objStream.Charset = CharSet 
    objStream.Open 
    objStream.WriteText TextData 
    objStream.Position = 0 
    objStream.Type = adTypeBinary 
    Select Case UCase(CharSet) 
        Case "UNICODE","UTF-16" 
        objStream.Position = 2 
        Case "UTF-8" 
        objStream.Position = 3 
    End Select 
    TextToBin = objStream.Read 
    objStream.Close 
    Set objStream = Nothing 
End Function

'------------------------------------------------
' SJIStoUTF8
Function SJIStoUTF8(strSJIS)

    strUNICODE = ASCW(strSJIS)  'ASCWでユニコードにする

    'コードを２進にしてワークに代入する
    strWORK2 = HEX16toSTR2(HEX(strUNICODE))

    '切り取って、UTF8の２進数を作成する
    'xxxx xxxx xxxx xxxx を下記に割り当てる
    '1110xxxx 10xxxxxx 10xxxxxx
    strUTF8CODE = "1110" & Mid(strWORK2, 1, 4)
    strUTF8CODE = strUTF8CODE & "10" & Mid(strWORK2, 5, 6)
    strUTF8CODE = strUTF8CODE & "10" & Mid(strWORK2, 11, 6)

    '作成した２進数を１６進数に戻す
    strWORK16 = STR2toHEX16(strUTF8CODE)

    '%を付けて格納
    strRET = ""  'リターン値を初期化
    strRET = strRET & "%" & Mid(strWORK16, 1, 2) '%を付けた文字列を作成
    strRET = strRET & "%" & Mid(strWORK16, 3, 2) 
    strRET = strRET & "%" & Mid(strWORK16, 5, 2) 

    'リターン値を代入
    SJIStoUTF8 = strRET

End Function

'------------------------------------------------
' HEX16toSTR2
' HEX16進文字列を受け取り2進数文字列を返す
Function HEX16toSTR2(strHEX)

    Dim n       'ループカウンタ
    Dim i       'ループのカウンタ
    Dim n8421   '8 4 2 1の数値計算用
    Dim str2STR 

    Dim nCHK

    str2STR = ""  '結果のエリアを初期化する

    '文字数分ループする
    For n = 1 To Len(strHEX)
        On Error Resume Next   'エラー発生時次の行へ
        nCHK = 0 '0で初期化
        nCHK = CInt("&h" & Mid(strHEX, n, 1)) 'n文字目を数値変換
        On Error Goto 0        'エラー処理を通常に戻す

        n8421 = 8  '初期値に８を代入する(上からチェックしたいので)
        For i = 1 To 4  '４回まわるよ
            If (nCHK And n8421) = 0 Then 'Andでビットをチェックする
                str2STR = str2STR & "0"  'ビットは立ってないよ
            Else
                str2STR = str2STR & "1"  'ビットは立ってるよ
            End If
            '次のビットをチェックしたいので２で割る
            n8421 = n8421 / 2
        Next 
    Next

    'リターン値をセットして終了
    HEX16toSTR2 = str2STR

End Function

'------------------------------------------------
' STR2toHEX16
' 2進文字列を受け取り16進文字列を返す
Function STR2toHEX16(str2)

    Dim strHEX
    Dim n       'ループカウンタ
    Dim i       'ループのカウンタ
    Dim n8421   '8 4 2 1の数値計算用
    Dim nBYTE

    '頭４文字単位かチェックする
    n = Len(str2) Mod 4      '足りない文字数を計算する
    If n <> 0 Then 
       str2 = String(4 - n, "0") & str2 '頭に文字０を追加する
    End If

    strHEX = ""   '結果のエリアを初期化する

    '文字数分ループする
    For n = 1 To Len(str2) Step 4  '４文字(1バイト)単位にループを作る
        n8421 = 8  '初期値に８を代入する(上から計算したいので)
        nBYTE = 0  '1バイト計算用変数を初期化
        For i = 0 To 3  '４回まわるよ(4ビット分)
            'ビットが立っているかチェックする
            If Mid(str2, n + i, 1) = "1" Then
                nBYTE = nBYTE + n8421   'ビットに対応した数値を＋する
            End If
            '次のビットを計算したいので２で割る
            n8421 = n8421 / 2
        Next 
        '計算して、１倍との数値が完成したので１６進文字にしてセットする
        strHEX = strHEX & Hex(nBYTE)
    Next 

    'リターン値をセットして関数を抜ける
    STR2toHEX16 = strHEX

End Function

Function TrimL(s)
    Dim r : Set r = New RegExp
    r.Global = False
    r.IgnoreCase = True
    r.Pattern = "^\s*"
    TrimL = r.Replace(s, "")
End Function

Function TrimR(s)
    Dim r : Set r = New RegExp
    r.Global = False
    r.IgnoreCase = True
    r.Pattern = "\s*$"
    TrimR = r.Replace(s, "")
End Function

Function TrimB(s)
    TrimB = TrimR(TrimL(s))
End Function


'------------------------------------------------
' StringBuilder Class
Class StringBuilder
    Private strArray()
    Private intGrowRate
    Private intItemCount
    
    Private Sub Class_Initialize()
        intGrowRate = 50
        intItemCount = 0
    End Sub
    
    Public Property Get GrowRate
        GrowRate = intGrowRate
    End Property
    
    Public Property Let GrowRate(value)
        intGrowRate = value
    End Property
    
    Private Sub InitArray()
        Redim Preserve strArray(intGrowRate)
    End Sub
    
    Public Sub Clear()
        intItemCount = 0
        Erase strArray
    End Sub
    
    Public Sub Append(str)
        
        If intItemCount = 0 Then
            Call InitArray
        ElseIf intItemCount > UBound(strArray) Then
            Redim Preserve strArray(Ubound(strArray) + intGrowRate)
        End If
        
        strArray(intItemCount) = str
        
        intItemCount = intItemCount + 1
        
    End Sub
    
    Public Function FindString(str)
        Dim x,mx
        mx = intItemCount - 1
        For x = 0 To mx
            If strArray(x) = str Then
                FindString = x
                Exit Function
            End If
        Next
        FindString = -1
    End Function
    
    Public Function ToString2(sep)
        If intItemCount = 0 Then
            ToString2 = ""
        Else
            Redim Preserve strArray(intItemCount)
            ToString2 = Join(strArray,sep)
        End If
    End Function
    
    Public Default Function ToString()
        If intItemCount = 0 Then
            ToString = ""
        Else
            ToString = Join(strArray,"")
        End If
    End Function

End Class


'------------------------------------------------
' PageCode Class
' Example:
'    Dim obj_page_code, arr
'    Set obj_page_code = New PageCode
'    arr = obj_page_code.enum_pagecode()
'    For Each s in arr
'        Echo s
'    Next
Class PageCode
    Public Function enum_pagecode()
        Dim reg, arr
        key = "MIME\Database\Charset"
	    Set reg = GetObject("Winmgmts:\root\default:StdRegProv")
	    reg.EnumKey HKCR, key, arr
        enum_pagecode = arr
    End Function

    Function ChangeCode(strFile, code1, code2)
        Set ADOStrm = CreateObject(COM_ADOSTREAM)
        ADOStrm.Type = 2
        ADOStrm.Mode = 3
        ADOStrm.CharSet = code1
        ADOStrm.Open
        ADOStrm.LoadFromFile strFile
        data= ADOStrm.ReadText
        ADOStrm.Position = 0
        ADOStrm.CharSet = code2
        ADOStrm.WriteText data
        ADOStrm.SetEOS
        ADOStrm.SaveToFile strFile & "_" & code2 & ".txt", 2
        ADOStrm.Close
    End Function

End Class 