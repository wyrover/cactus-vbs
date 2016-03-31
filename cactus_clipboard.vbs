Function FetchClipboardData()
    Dim objMSIE
    Set objMSIE = CreateObject(COM_IE)
    objMSIE.Navigate("about:blank")
    FetchClipboardData = objMSIE.document.parentwindow.clipboardData.GetData("text")
    objMSIE.Quit
End Function