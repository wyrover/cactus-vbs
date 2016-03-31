'------------------------------------------------
' Pager Class
Class Pager
    Private m_id 
    Private m_currentpage 
    Private m_recordcount  
    Private m_pagecount  
    Private m_pagesize 
    Private m_endfix

    Public  Function Init(id, currentpage, recordcount, pagesize, endfix)
        m_id = id
        m_currentpage = currentpage
        m_recordcount = recordcount
        m_pagesize = pagesize
        m_endfix = endfix

        If recordcount mod pagesize <> 0 Then
            m_pagecount = Int((recordcount / pagesize) + 1)
        Else 
            m_pagecount = Int(recordcount / pagesize)
        End If
    End Function

    Public Function PageSize()
        PageSize = Int(m_pagesize)
    End Function

    Public Function getHTML() 
        If m_currentpage < 1 Then
            m_currentpage = 1
        End If
        If m_pagecount < 1 Then
            m_pagecount = 1
        End If
        If m_currentpage > m_pagecount Then
            m_currentpage = m_pagecount
        End If


        Dim prevpage 
        prevpage =  m_currentpage - 1 

        Dim nextpage  
        nextpage =  m_currentpage + 1 



        Dim retval 
        Dim sbPager 
        Set sbPager =  New StringBuilder
        sbPager.Append("<span class=""count"">Pages: ")
        sbPager.Append(m_currentpage)
        sbPager.Append("/")
        sbPager.Append(m_pagecount)
        sbPager.Append("</span>")

        sbPager.Append("<b>")

        If prevpage < 1 Then
            sbPager.Append(" &laquo; First")
            sbPager.Append(" &laquo;")
        Else 
            sbPager.Append(" <a href=""" & m_id & "1" & m_endfix & """>&laquo; First</a>")
            sbPager.Append(" <a href=""" & m_id & prevpage & m_endfix & """>&laquo;</a>")
        End If


            Dim startpage 
            If (m_currentpage mod 10) = 0 Then
                startpage = m_currentpage - 9
            Else 
                startpage = m_currentpage - CInt((m_currentpage mod 10)) + 1
            End If

            If startpage > 10 Then
                sbPager.Append(" <a href=""" & m_id & (startpage - 1) & m_endfix & """>...</a>")
            End If

            Dim i 
            For  i = startpage To  startpage + 10- 1  Step  i + 1
                If i > m_pagecount Then
                    Exit For
                End If
                If i = m_currentpage Then
                    sbPager.Append(" <span title=""Page " & i & """>" & i & "</span>")
                Else 
                    sbPager.Append(" <a href=""" & m_id & i & m_endfix & """>" & i & "</a>")
                End If
            Next

            If m_pagecount >= m_currentpage + 10 Then
                sbPager.Append(" <a href=""" & m_id & (startpage + 10) & m_endfix & """>...</a>")
            End If


        If nextpage > m_pagecount Then
            sbPager.Append(" &raquo;")
            sbPager.Append(" Last &raquo;")
        Else 
            sbPager.Append(" <a href=""" & m_id & nextpage & m_endfix & """>&raquo;</a>")
            sbPager.Append(" <a href=""" & m_id & m_pagecount & m_endfix & """>Last &raquo;</a>")
        End If

        sbPager.Append("</b>")

        retval = sbPager.ToString()
        getHTML = retval
    End Function
End Class