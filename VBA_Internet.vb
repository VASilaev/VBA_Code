Public Function GetHTMLPageByURL(URL$) As String
' Функция возвращает тело переданной в URL страницы
'#param URL$ - URL страницы
'----------------------------------------------------------------------------------------
Dim oXMLHTTP As Object
On Error GoTo GetHTMLPageByURL_Err
 
    Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    With oXMLHTTP
        .Open "GET", URL, False
        .send
        If .Status = 200 Then
            GetHTMLPageByURL = .ResponseText ' Тело HTML стр.
        End If
    End With
 
GetHTMLPageByURL_End:
    On Error Resume Next
    Set oXMLHTTP = Nothing
    Err.Clear
    Exit Function
 
GetHTMLPageByURL_Err:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function" & _
           "GetHTMLPageByURL - 00_Tests.", vbCritical, "Произошла ошибка!"
    'Debug.Print "GetHTMLPageByURL_Line: " & Erl & "."
    Err.Clear
    Resume GetHTMLPageByURL_End
 
End Function

Public Function getbybarcode(sBarcode)
'Возвращает наименование товара по его штрихкоду используя информацию с ru.disai.org
'#param sBarcode - Штрихкод

getbybarcode = regexp_replace( _
                 Replace( _
                   regexp_substr( _
                     GetHTMLPageByURL("https://ru.disai.org/?search_query=" & sBarcode), _
                     "<tr bgcolor=#e0e8f2 valign=top>((?:.|\n|\r)+?)</tr>" _
                   ), _
                   "</td>", _
                   vbTab _
                 ), _
                 "\g<[^>]+>", _
                 "" _
               )

getbybarcode = Replace(getbybarcode, "&nbsp;", " ", 1, -1, 1)
getbybarcode = Replace(getbybarcode, "&quot;", """", 1, -1, 1)
getbybarcode = Replace(getbybarcode, "&lt;", "<", 1, -1, 1)
getbybarcode = Replace(getbybarcode, "&gt;", ">", 1, -1, 1)
getbybarcode = Replace(getbybarcode, "&reg;", "®", 1, -1, 1)
getbybarcode = Replace(getbybarcode, "&copy;", "©", 1, -1, 1)
getbybarcode = Replace(getbybarcode, "&bull;", "·", 1, -1, 1)
getbybarcode = Replace(getbybarcode, "&amp;", "&", 1, -1, 1)
getbybarcode = Replace(getbybarcode, vbCr, "")
getbybarcode = Replace(getbybarcode, vbLf, "")


getbybarcode = Split(getbybarcode, vbTab)

End Function
